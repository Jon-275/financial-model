import assert from 'node:assert/strict';
import { test } from 'node:test';
import { reviewPullRequest } from '../tools/gemini_review.mjs';
import { collectBatches, patchLimit } from '../tools/gemini_review_batches.mjs';

const sha = '1234567'.padEnd(40, '0');
const source = {
  filename: 'tools/model.py',
  status: 'modified',
  additions: 1,
  deletions: 1,
  patch: '@@ -1 +1 @@\n-old\n+new',
};
const finding = {
  path: source.filename,
  line: 1,
  side: 'RIGHT',
  severity: 'high',
  issue: 'The new value drops a required tax adjustment.',
  evidence: 'The new line removes the tax adjustment.',
  fix: 'Keep the adjustment in the calculation.',
};

async function run({
  files = [source],
  changedFiles = files.length,
  findings = [],
  apiKey = 'fake-key',
  httpStatus = 200,
  finishReason = 'STOP',
  responseText = JSON.stringify({ findings }),
  latestSha = sha,
  manual = false,
  earlierThreads = [],
  threadPages = [earlierThreads],
  threadError = false,
  reviewError = false,
  reply,
  beforeResponse,
  afterLog,
} = {}) {
  const pr = {
    number: 17,
    state: 'open',
    draft: false,
    changed_files: changedFiles,
    base: { ref: 'main' },
    head: { sha, repo: { fork: true } },
  };
  const context = {
    payload: manual ? { inputs: { 'pull-request': '17' } } : { pull_request: pr },
    repo: { owner: 'owner', repo: 'model' },
    serverUrl: 'https://github.com',
    runId: 39,
  };
  const results = { statuses: [], sent: [], reviews: [], errors: [], logs: [] };
  let reads = 0;
  const github = {
    paginate: (method, args) => method(args),
    graphql: async (query, { owner, repo, number, after }) => {
      assert.match(query, /reviewThreads/);
      assert.deepEqual({ owner, repo, number }, { owner: 'owner', repo: 'model', number: pr.number });
      if (threadError) throw new Error('GitHub refused to read the review threads.');
      const index = after === null ? 0 : Number(after);
      return { repository: { pullRequest: { reviewThreads: {
        nodes: threadPages[index],
        pageInfo: {
          hasNextPage: index + 1 < threadPages.length,
          endCursor: String(index + 1),
        },
      } } } };
    },
    rest: {
      pulls: {
        get: async ({ pull_number }) => {
          assert.equal(pull_number, pr.number);
          return { data: { ...pr, head: { ...pr.head, sha: reads++ ? latestSha : sha } } };
        },
        listFiles: async () => files,
        createReview: async args => {
          if (reviewError) throw new Error('GitHub refused the review.');
          results.reviews.push(args);
        },
      },
      repos: {
        createCommitStatus: async args => results.statuses.push(args),
      },

    },
  };
  const fetchApi = async (_url, options) => {
    assert.equal(options.headers['x-goog-api-key'], apiKey);
    results.sent.push(JSON.parse(options.body));
    await beforeResponse?.(options.signal, results.sent.length);
    const next = reply?.(results.sent.at(-1), results.sent.length) ?? {};
    return {
      ok: (next.status ?? httpStatus) === 200,
      status: next.status ?? httpStatus,
      json: async () => ({
        candidates: [{ finishReason: next.finishReason ?? finishReason, content: { parts: [{ text: next.text ?? responseText }] } }],
      }),
    };
  };
  await reviewPullRequest({
    github, context, apiKey, fetchApi,
    core: { notice() {}, info: message => { results.logs.push(message); afterLog?.(); }, setFailed: message => results.errors.push(message) },
  });
  assert.equal(results.statuses.length, 2);
  assert.equal(results.statuses[0].state, 'pending');
  for (const status of results.statuses) {
    assert.equal(status.context, 'Gemini Review');
    assert.equal(status.sha, sha);
    assert.equal(status.target_url, 'https://github.com/owner/model/actions/runs/39');
  }
  return results;
}

test('a clean code review posts one review on the exact commit', async () => {
  const runResult = await run();
  assert.equal(runResult.statuses[1].state, 'success');
  assert.equal(runResult.reviews[0].pull_number, 17);
  assert.equal(runResult.reviews[0].commit_id, sha);
  assert.equal(runResult.reviews[0].event, 'COMMENT');
  assert.deepEqual(runResult.reviews[0].comments, []);
  assert.match(runResult.reviews[0].body, /no actionable issues/);
  assert.match(runResult.sent[0].contents[0].parts[0].text, /\+new/);
});

test('serious findings fail the required check and open a thread with evidence', async () => {
  for (const severity of ['critical', 'high']) {
    const runResult = await run({ findings: [{ ...finding, severity }] });
    assert.equal(runResult.statuses[1].state, 'failure');
    const [comment] = runResult.reviews[0].comments;
    assert.equal(comment.path, source.filename);
    assert.equal(comment.line, 1);
    assert.equal(comment.side, 'RIGHT');
    assert.match(comment.body, new RegExp(severity.toUpperCase()));
    assert.match(comment.body, /removes the tax adjustment/);
  }
});

test('smaller findings open inline threads even when the review status passes', async () => {
  const result = await run({ findings: [{ ...finding, severity: 'medium' }] });
  assert.equal(result.statuses[1].state, 'success');
  assert.match(result.statuses[1].description, /Resolve the finding threads/);
  assert.equal(result.reviews[0].comments.length, 1);
});

test('multiple findings attach to original deletion and context lines', async () => {
  const file = {
    ...source,
    filename: 'tools/renamed.py',
    previous_filename: source.filename,
    status: 'renamed',
    patch: '@@ -10,2 +20,2 @@\n-old\n+new\n context\n@@ -30 +40 @@\n left unchanged',
  };
  const result = await run({
    files: [file],
    findings: [
      { ...finding, line: 10, side: 'LEFT' },
      { ...finding, issue: 'A separate issue.', line: 38, side: 'RIGHT' },
    ],
  });
  const [oldLine, newLine] = result.reviews[0].comments;
  assert.equal(oldLine.path, file.filename);
  assert.equal(oldLine.line, 10);
  assert.equal(oldLine.side, 'LEFT');
  assert.equal(newLine.line, 40);
  assert.equal(newLine.side, 'RIGHT');
  assert.match(newLine.body, /Reported at `tools\/model\.py:38`/);
});

test('only unresolved, current workflow threads suppress duplicate findings', async () => {
  const first = await run({ findings: [finding] });
  const body = first.reviews[0].comments[0].body;
  const thread = {
    isResolved: false, isOutdated: false,
    comments: { nodes: [{ author: { __typename: 'Bot', login: 'github-actions' }, body }] },
  };
  const repeated = await run({ findings: [{ ...finding, line: 4 }], threadPages: [[], [thread]] });
  assert.deepEqual(repeated.reviews[0].comments, []);
  assert.equal(repeated.statuses[1].state, 'failure');
  for (const earlier of [
    { ...thread, isResolved: true },
    { ...thread, isOutdated: true },
    { ...thread, comments: { nodes: [{ author: { __typename: 'User', login: 'github-actions' }, body }] } },
  ]) {
    const result = await run({ findings: [finding], earlierThreads: [earlier] });
    assert.equal(result.reviews[0].comments.length, 1);
  }
});

test('findings with the same title and different evidence each open a thread', async () => {
  const result = await run({ findings: [finding, { ...finding, evidence: 'A second failure.' }] });
  assert.equal(result.reviews[0].comments.length, 2);
});

test('the required check fails if GitHub cannot publish a finding thread', async () => {
  const result = await run({ findings: [{ ...finding, severity: 'low' }], reviewError: true });
  assert.equal(result.statuses[1].state, 'failure');
  assert.match(result.errors[0], /GitHub refused the review/);
});

test('the required check fails if the open review threads cannot be read', async () => {
  const result = await run({ findings: [finding], threadError: true });
  assert.equal(result.statuses[1].state, 'failure');
  assert.equal(result.reviews.length, 0);
  assert.match(result.errors[0], /refused to read the review threads/);
});

test('workbook and documentation changes can pass without sending data to Gemini', async () => {
  const runResult = await run({
    files: [{ filename: 'workbook/example.xlsx' }, { filename: 'README.md' }],
    apiKey: '',
    manual: true,
  });
  assert.equal(runResult.statuses[1].state, 'success');
  assert.equal(runResult.sent.length, 0);
  assert.equal(runResult.reviews.length, 0);
});

test('an incomplete or unavailable review fails instead of passing', async () => {
  const cases = [
    { options: { apiKey: '' }, reason: /secret is missing/ },
    { options: { httpStatus: 429 }, reason: /HTTP 429/ },
    { options: { finishReason: 'MAX_TOKENS' }, reason: /did not complete/ },
    { options: { responseText: 'not JSON' }, reason: /invalid review response/ },
    { options: { findings: [{ ...finding, path: 'unrelated.py' }] }, reason: /could not be validated/ },
    { options: { findings: [{ ...finding, line: 0 }] }, reason: /could not be validated/ },
    { options: { findings: [{ ...finding, side: 'MISSING' }] }, reason: /could not be validated/ },
    { options: { files: [{ ...source, additions: 0, deletions: 0, patch: '' }], findings: [finding] },
      reason: /without an available diff line/ },
    { options: { changedFiles: 2 }, reason: /incomplete changed-file list/ },
    { options: { files: [{ ...source, additions: 2 }] }, reason: /incomplete patch/ },
  ];
  for (const { options, reason } of cases) {
    const runResult = await run(options);
    assert.equal(runResult.statuses[1].state, 'failure');
    assert.match(runResult.errors[0], reason);
    assert.equal(runResult.reviews.length, 0);
  }
});

test('an older review cannot post approval after the PR head changes', async () => {
  const runResult = await run({ latestSha: 'new commit' });
  assert.equal(runResult.statuses[1].state, 'error');
  assert.equal(runResult.reviews.length, 0);
});

function newFile(filename, lines) {
  return {
    filename, status: 'added', deletions: 0, additions: lines.length,
    patch: `@@ -0,0 +1,${lines.length} @@\n` + lines.map(line => '+' + line).join('\n'),
  };
}

function textOf(request) {
  return request.contents[0].parts[0].text;
}

function cleanBatchReply(request, index) {
  const requiresNote = request.generationConfig.responseJsonSchema.required.includes('overview');
  return { text: JSON.stringify(requiresNote ? { findings: [], overview: 'Checked module contracts ' + index }
    : { findings: [] }) };
}

const largeFiles = ['tools/a.py', 'tools/b.bas'].map(name => newFile(name, ['x'.repeat(patchLimit * 0.6)]));

test('one large hunk and a single long Unicode line are completely preserved in bounded requests', () => {
  for (const lines of [Array.from({ length: 150 }, (_,i) => 'number_' + i + '_'.repeat(18)),
                       ['🧾'.repeat(240), 'tail']]) {
    const file = newFile('tools/large.bas', lines);
    const { batches } = collectBatches([file], 1, 200);
    assert.ok(batches.length > 1);
    const fragments = batches.map(({ text }) => {
      assert.ok(text.length <= 200);
      assert.ok(text.isWellFormed());
      assert.ok(text.startsWith('File: "tools/large.bas" (added)\n'));
      return text.replace(/^File:[^\n]+\n(?:Continuing @@[^\n]+\n)?/, '');
    });
    assert.equal(fragments.join(''), file.patch);
    if (lines.length === 2) assert.ok(batches.some(({ text }) => text.includes('inside diff line "+"')));
  }
});

test('splits at complete hunks when possible and preserves absolute lines, context and no-newline markers', () => {
  const file = { ...source, additions: 2, deletions: 2, patch: [
    '@@ -10,2 +20,2 @@ first', '-first change', '+new first', ' context',
    '@@ -51 +68 @@ second', '-second change', '+new second', '\\ No newline at end of file',
  ].join('\n') };
  const { batches, paths } = collectBatches([file], 1, 115);
  assert.equal(batches.length, 2);
  assert.match(batches[1].text, /\n@@ -51 \+68 @@ second/);
  assert.deepEqual(paths.get(file.filename).anchors, { LEFT: [10,51], RIGHT: [20,21,68] });
});

test('reviews production VBA, test fragments, helpers and dependency pins but never exports workbook or household data', async () => {
  const allowed = ['tools/vba/Run.bas','tools/vba/Account.CLS','tools/vba/tests/Runner.fragment',
    '.githooks/pre-commit','.github/workflows/check.yaml','tools/requirements.txt','tools/build.sh','tools/test.mjs'];
  const excluded = ['workbook/template.xlsm','docs/example.pdf','README.md','tests/scenarios.json','private/example.xlsx'];
  const result = await run({ files: [...allowed.map(path => newFile(path, ['ok'])),
    ...excluded.map(path => newFile(path, ['sensitive fixture literal']))] });
  assert.equal(result.statuses[1].state, 'success');
  const sent = result.sent.map(textOf).join('\n');
  for (const path of allowed) assert.ok(sent.includes(JSON.stringify(path)));
  for (const path of excluded) assert.ok(!sent.includes(path));
  assert.ok(!sent.includes('sensitive fixture literal'));
});

test('batches are deterministic and each reviewed source segment is submitted before cross-file review', async () => {
  const forward = collectBatches(largeFiles, 2);
  const reverse = collectBatches(largeFiles.toReversed(), 2);
  assert.deepEqual(forward, reverse);
  const result = await run({ files: largeFiles.toReversed(), reply: cleanBatchReply });
  assert.equal(result.statuses[1].state, 'success');
  assert.equal(result.reviews.length, 1);
  assert.equal(result.sent.length, 3);
  for (const [i,file] of largeFiles.entries()) {
    assert.ok(textOf(result.sent[i]).includes(file.patch));
    assert.ok(result.sent[i].generationConfig.responseJsonSchema.required.includes('overview'));
  }
  assert.match(textOf(result.sent[2]), /contracts 1[\s\S]+contracts 2/);
  assert.match(result.sent[2].systemInstruction.parts[0].text, /spanning files or patches/);
  assert.ok(result.sent.every(r => textOf(r).length <= patchLimit));
});

test('serious findings from an early batch and the integration pass both block merging', async () => {
  const issue = { ...finding, path: largeFiles[0].filename };
  const cross = { ...finding, path: largeFiles[1].filename, issue: 'An incompatible argument crosses modules.' };
  const reply = (request, i) => ({
    text: JSON.stringify(i === 3 ? { findings: [cross] } :
      { findings: i === 1 ? [issue] : [], overview: 'Reviewed the caller and its argument.' }),
  });
  const result = await run({ files: largeFiles, reply });
  assert.equal(result.statuses[1].state, 'failure');
  assert.deepEqual(result.reviews[0].comments.map(c => c.path), [issue.path, cross.path]);
});

test('API failures or invalid findings in any batch cannot publish a partial review', async () => {
  const badReplies = [
    [2, { status: 503 }], [2, { finishReason: 'MAX_TOKENS' }],
    [1, { text: JSON.stringify({ findings: [{ ...finding, path: largeFiles[1].filename }], overview: 'Invalid location.' }) }],
    [3, { text: '{}' }], [3, { text: 'incomplete JSON' }],
  ];
  for (const [at, response] of badReplies) {
    const result = await run({ files: largeFiles,
      reply: (request,i) => i === at ? response : cleanBatchReply(request, i) });
    assert.equal(result.statuses[1].state, 'failure');
    assert.equal(result.reviews.length, 0);
    assert.equal(result.sent.length, at);
  }
});

test('a complete note beyond the requested length reaches cross-file review without truncation', async () => {
  const note = 'Longer complete note '.repeat(150) + 'unique last words';
  const result = await run({ files: largeFiles, reply: (request, i) =>
    i === 1 ? { text: JSON.stringify({ findings: [], overview: note }) } : cleanBatchReply(request, i) });
  assert.equal(result.statuses[1].state, 'success');
  assert.equal(result.sent.length, 3);
  assert.ok(textOf(result.sent[2]).includes(note));
});

test('one note repair keeps findings from the first response and the retry', async () => {
  const first = { ...finding, path: largeFiles[0].filename }, second = { ...first, issue: 'Second separate defect' };
  const result = await run({ files: largeFiles, reply: (request, i) => {
    if (i === 1) return { text: JSON.stringify({ findings: [first] }) };
    if (i === 2) return { text: JSON.stringify({ findings: [second], overview: 'Both call the funding module.' }) };
    return cleanBatchReply(request, i);
  } });
  assert.equal(result.sent.length, 4);
  assert.equal(textOf(result.sent[0]), textOf(result.sent[1]));
  assert.match(result.sent[1].systemInstruction.parts[0].text, /100 words/);
  assert.equal(result.statuses[1].state, 'failure');
  assert.equal(result.reviews[0].comments.length, 2);
});

test('invalid notes on the retry fail without posting partial results', async () => {
  for (const note of [undefined, '  ', 'x'.repeat(patchLimit)]) {
    const result = await run({ files: largeFiles, reply: () => ({ text: JSON.stringify({ findings: [], overview: note }) }) });
    assert.equal(result.sent.length, 2);
    assert.equal(result.statuses[1].state, 'failure');
    assert.equal(result.reviews.length, 0);
    assert.match(result.errors[0], /unusable cross-file note twice/);
  }
});

test('a truncated hunk, duplicate filename or missing VBA patch fails before any source leaves the workflow', async () => {
  for (const files of [
    [{ ...source, patch: '@@ -7,3 +7,3 @@\n-old\n+new' }],
    [source, source],
    [{ ...source, filename: 'tools/vba/Runner.bas', patch: undefined }],
  ]) {
    const result = await run({ files });
    assert.equal(result.statuses[1].state, 'failure');
    assert.equal(result.sent.length, 0);
  }
});

test('a partial file review cannot cite material outside its own hunk fragment', async () => {
  const file = newFile('tools/runner.bas', Array.from({ length: 300 }, (_,i) => 'a'.repeat(300) + i));
  const result = await run({ files: [file], reply: () => ({
    text: JSON.stringify({ findings: [{...finding, path: file.filename, line: 300}], overview: 'Reviewed this fragment.' }),
  }) });
  assert.equal(result.sent.length, 1);
  assert.equal(result.statuses[1].state, 'failure');
  assert.match(result.errors[0], /could not be validated/);
});

test('integration findings can cite either end of a file split over multiple requests', async () => {
  const file = newFile('tools/runner.bas', Array.from({ length: 300 }, (_,i) => 'a'.repeat(300) + i));
  const result = await run({ files: [file], reply: (request,i) => {
    if (request.generationConfig.responseJsonSchema.required.includes('overview')) return cleanBatchReply(request,i);
    return { text: JSON.stringify({ findings: [{ ...finding, path: file.filename, line: 300 }] }) };
  } });
  assert.equal(result.sent.length, 4);
  assert.equal(result.reviews[0].comments[0].line, 300);
  assert.equal(result.statuses[1].state, 'failure');
});

test('metadata-only source changes and pure renames are reviewed without inventing a diff line', async () => {
  const renamed = { filename: 'tools/moved.bas', previous_filename: 'tools/old.bas', status: 'renamed', additions: 0,
    deletions: 0 };
  for (const file of [renamed, { filename: 'tools/empty.py', status: 'added', additions: 0, deletions: 0, patch: '' }]) {
    const result = await run({ files: [file] });
    assert.equal(result.statuses[1].state, 'success');
    assert.ok(textOf(result.sent[0]).includes(JSON.stringify(file.filename)));
    if (file.previous_filename) assert.ok(textOf(result.sent[0]).includes(JSON.stringify(file.previous_filename)));
  }
});

test('renaming an old path and introducing a new file at that name preserves each canonical anchor', () => {
  const added = newFile('tools/a.py', ['new']);
  const moved = { ...source, filename: 'tools/b.py', previous_filename: 'tools/a.py', status: 'renamed',
    patch: '@@ -15 +15 @@\n-before\n+after' };
  const { paths } = collectBatches([moved, added], 2);
  assert.equal(paths.get(added.filename).path, added.filename);
  assert.equal(paths.get(moved.filename).path, moved.filename);
  assert.deepEqual(paths.get(moved.filename).anchors.RIGHT, [15]);
});

test('cross-file notes are reduced in bounded passes without dropping findings or any final notes', async () => {
  const files = Array.from({ length: 19 }, (_,i) => newFile('tools/file-' + String(i).padStart(2,'0') + '.py',
    ['x'.repeat(patchLimit * 0.6)]));
  const result = await run({files, reply: (request,i) => {
    const integration = request.systemInstruction.parts[0].text.includes('spanning files or patches');
    const final = !request.generationConfig.responseJsonSchema.required.includes('overview');
    const note = (integration ? 'group-' : 'leaf-') + i;
    return { text: JSON.stringify(final ? { findings: [] } :
      { findings: i === 20 ? [{ ...finding, path: files[0].filename, severity: 'medium' }] : [],
        overview: note.padEnd(integration ? 15 : 2_395, '.') }) };
  } });
  assert.equal(result.sent.length, 22); // 19 patch batches, two groups of notes, one final review.
  assert.ok(result.sent.every(req => textOf(req).length <= patchLimit));
  const groups = textOf(result.sent[19]) + textOf(result.sent[20]);
  for (let i = 1; i <= 19; i++) assert.ok(groups.includes('leaf-' + i + '.'), String(i));
  assert.match(textOf(result.sent[21]), /group-20[\s\S]*group-21/);
  assert.equal(result.reviews[0].comments.length, 1);
  assert.equal(result.statuses[1].state, 'success');
  assert.match(result.reviews[0].body, /3 cross-file/);
});

test('an aborted fetch after a successful batch fails the check and never posts a partial review', async t => {
  let controller;
  t.mock.method(AbortSignal, 'timeout', ms => {
    assert.ok(ms > 0 && ms <= 240_000);
    controller = new AbortController();
    return controller.signal;
  });
  const result = await run({
    files: largeFiles, reply: cleanBatchReply,
    beforeResponse: (signal, i) => {
      if (i === 1) return;
      return new Promise((_resolve, reject) => {
        signal.addEventListener('abort', () => reject(signal.reason), { once: true });
        controller.abort(new DOMException('The review request timed out.', 'TimeoutError'));
      });
    },
  });
  assert.equal(result.sent.length, 2);
  assert.equal(result.statuses[1].state, 'failure');
  assert.match(result.errors[0], /request timed out/);
  assert.equal(result.reviews.length, 0);
});

test('expiration of the shared deadline between batches fails before another request is sent', async t => {
  let time = 1_000;
  t.mock.method(Date, 'now', () => time);
  const result = await run({
    files: largeFiles, reply: cleanBatchReply,
    afterLog: () => { time += 25 * 60_000; },
  });
  assert.equal(result.sent.length, 1);
  assert.equal(result.statuses[1].state, 'failure');
  assert.match(result.errors[0], /before the deadline/);
  assert.equal(result.reviews.length, 0);
});

test('a patch fits exactly at the 40,000-character limit and splits above it', () => {
  assert.equal(patchLimit, 40_000);
  const file = newFile('tools/boundary.py', ['']);
  const overhead = collectBatches([file], 1).batches[0].text.length;
  const contentLength = patchLimit - overhead;
  for (const extra of [0, 1]) {
    const input = newFile(file.filename, ['a'.repeat(contentLength + extra)]);
    const { batches } = collectBatches([input], 1);
    assert.equal(batches.length === 1, extra === 0);
    assert.ok(batches.every(batch => batch.text.length <= patchLimit));
    const patch = batches.map(batch => batch.text.replace(/^File:[^\n]+\n(?:Continuing @@[^\n]+\n)?/, '')).join('');
    assert.equal(patch, input.patch);
  }
});

test('a request receives up to four minutes and caps its timeout at the remaining total budget', async t => {
  let time = 0;
  const timeouts = [];
  t.mock.method(Date, 'now', () => time);
  t.mock.method(AbortSignal, 'timeout', ms => {
    timeouts.push(ms);
    return new AbortController().signal;
  });
  const files = Array.from({ length: 7 }, (_, i) => newFile('tools/part-' + i + '.py', ['x'.repeat(24_000)]));
  const result = await run({ files, reply: cleanBatchReply, beforeResponse: (_signal, i) => {
    time += i < 7 ? 230_000 : 10_000;
  } });
  assert.deepEqual(timeouts, [240_000,240_000,240_000,240_000,240_000,240_000,120_000,110_000]);
  assert.equal(result.statuses[1].state, 'success');
  assert.deepEqual(result.logs, [...Array.from({ length: 6 }, (_, i) => 'Patch batch ' + (i+1) + '/7: 230000 ms.'),
    'Patch batch 7/7: 10000 ms.', 'Cross-file batch 1/1: 10000 ms.']);
});


test('a final response arriving at the shared deadline cannot publish a successful review', async t => {
  let time = 0;
  t.mock.method(Date, 'now', () => time);
  const result = await run({ beforeResponse: () => { time = 25 * 60_000; } });
  assert.equal(result.sent.length, 1);
  assert.equal(result.statuses[1].state, 'failure');
  assert.equal(result.reviews.length, 0);
  assert.match(result.errors[0], /before the deadline/);
});

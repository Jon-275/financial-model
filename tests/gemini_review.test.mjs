import assert from 'node:assert/strict';
import { test } from 'node:test';
import { reviewPullRequest } from '../tools/gemini_review.mjs';

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
  const results = { statuses: [], sent: [], reviews: [], errors: [] };
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
    return {
      ok: httpStatus === 200,
      status: httpStatus,
      json: async () => ({
        candidates: [{ finishReason, content: { parts: [{ text: responseText }] } }],
      }),
    };
  };
  await reviewPullRequest({
    github, context, apiKey, fetchApi,
    core: { notice() {}, setFailed: message => results.errors.push(message) },
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

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
  comments = [],
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
  const results = { statuses: [], sent: [], comments: [], errors: [] };
  let reads = 0;
  const github = {
    paginate: (method, args) => method(args),
    rest: {
      pulls: {
        get: async ({ pull_number }) => {
          assert.equal(pull_number, pr.number);
          return { data: { ...pr, head: { ...pr.head, sha: reads++ ? latestSha : sha } } };
        },
        listFiles: async () => files,
      },
      repos: {
        createCommitStatus: async args => results.statuses.push(args),
      },
      issues: {
        listComments: async () => comments,
        createComment: async args => results.comments.push(args),
        updateComment: async args => results.comments.push(args),
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

test('a clean code review passes on the PR head and posts a comment', async () => {
  const runResult = await run();
  assert.equal(runResult.statuses[1].state, 'success');
  assert.equal(runResult.comments[0].issue_number, 17);
  assert.match(runResult.comments[0].body, /no actionable issues/);
  assert.match(runResult.sent[0].contents[0].parts[0].text, /\+new/);
});

test('serious findings fail the required check with evidence in the comment', async () => {
  for (const severity of ['critical', 'high']) {
    const runResult = await run({ findings: [{ ...finding, severity }] });
    assert.equal(runResult.statuses[1].state, 'failure');
    assert.match(runResult.comments[0].body, /removes the tax adjustment/);
  }
});

test('smaller findings pass and update the previous review comment', async () => {
  const comments = [{
    id: 87,
    user: { login: 'github-actions[bot]' },
    body: '<!-- financial-model-gemini-review -->\n\nEarlier review',
  }];
  const runResult = await run({ findings: [{ ...finding, severity: 'medium' }], comments });
  assert.equal(runResult.statuses[1].state, 'success');
  assert.equal(runResult.comments[0].comment_id, 87);
  assert.match(runResult.comments[0].body, /Passed with suggestions/);
});

test('workbook and documentation changes can pass without sending data to Gemini', async () => {
  const runResult = await run({
    files: [{ filename: 'workbook/example.xlsx' }, { filename: 'README.md' }],
    apiKey: '',
    manual: true,
  });
  assert.equal(runResult.statuses[1].state, 'success');
  assert.equal(runResult.sent.length, 0);
  assert.equal(runResult.comments.length, 0);
});

test('an incomplete or unavailable review fails instead of passing', async () => {
  const cases = [
    { options: { apiKey: '' }, reason: /secret is missing/ },
    { options: { httpStatus: 429 }, reason: /HTTP 429/ },
    { options: { finishReason: 'MAX_TOKENS' }, reason: /did not complete/ },
    { options: { responseText: 'not JSON' }, reason: /invalid review response/ },
    { options: { findings: [{ ...finding, path: 'unrelated.py' }] }, reason: /could not be validated/ },
    { options: { changedFiles: 2 }, reason: /incomplete changed-file list/ },
    { options: { files: [{ ...source, additions: 2 }] }, reason: /incomplete patch/ },
  ];
  for (const { options, reason } of cases) {
    const runResult = await run(options);
    assert.equal(runResult.statuses[1].state, 'failure');
    assert.match(runResult.errors[0], reason);
    assert.equal(runResult.comments.length, 0);
  }
});

test('an older review cannot post approval after the PR head changes', async () => {
  const runResult = await run({ latestSha: 'new commit' });
  assert.equal(runResult.statuses[1].state, 'error');
  assert.equal(runResult.comments.length, 0);
});

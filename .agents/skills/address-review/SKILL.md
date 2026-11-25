---
name: address-review
description: Work through one review round on a pull request (wait for every reviewer on the latest commit, fix or answer every thread with an accepted or rejected label, resolve it, push once). Use when a PR has review findings or a failing Review labels status; not for reviewing your own branch before opening a PR (review-locally).
---

# Address a review round

Every push re-runs CI and every reviewer, so a round ends in exactly one push. The labels
on your replies feed the lessons and the reviewers' examples, so every thread gets one.

## Steps

1. **Wait.** Every reviewer must have finished on the PR's latest commit. Check with
   `gh pr checks <number>`; a review still pending means the round has not started.
2. **Read every open thread**, from every reviewer, before changing anything. Group
   duplicates: one fix can answer several threads.
3. **For each finding, decide:**
   - **Right:** fix it with a test that fails when the fix is reverted.
   - **Wrong:** do not change the code; know why it is wrong.
4. **Run the gates** from `AGENTS.md` once all fixes are in.
5. **Push once.**
6. **Reply on every thread in one line** and resolve it. Start with the project's label
   (`accepted:` or `rejected:` by default; `devloop.toml` `labels.prefixes` lists them):
   - `accepted: fixed in <sha>, test added`
   - `rejected: <why the finding is wrong>`
   A resolved thread without a labelled reply leaves the `Review labels` status failing.
7. **Merge the default branch in** only when GitHub reports a conflict or the base changed
   code this PR depends on, and only once, before the final round.

## When not to use

- Before a PR exists: use `review-locally`.
- Never push once per finding, and never resolve a thread without a labelled reply.
- Never reject a finding to make a check pass; reject only what is wrong.

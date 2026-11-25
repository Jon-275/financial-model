---
name: pr-workflow
description: Take one change from a fresh branch to a merged pull request (worktree, gates, PR, one push per review round, merge only when every check is green and every thread labelled). Use for any code or docs change; not for a change over the size limit (split-stack) or for answering review threads (address-review).
---

# Pull request workflow

One concern per pull request. CI enforces every gate below; this skill is the order that
gets a change through them first time.

## Steps

1. **Start clean.** Fetch, then create a worktree on a new branch from the default branch:
   `git worktree add -b <type>/<slug> .worktrees/<slug> origin/main`. Never commit to the
   default branch, and never switch the branch of a checkout another agent is using.
2. **Read first.** Run the `lessons-intake` skill for the area you are changing.
3. **Tests first.** Write the test that fails without the change, watch it fail, then make
   it pass. A stateful change follows the `stateful-change` skill.
4. **Run the gates** listed under "Commands" in `AGENTS.md` until all pass, and
   `devloop doctor --check` when you touched skills, `AGENTS.md` or `devloop.toml`.
5. **Add a changelog fragment** under the changelog folder (`changes/` by default), named
   `<slug>.<section>.md`, and check it with `devloop changelog check`.
6. **Check the size** with `devloop size --base origin/main`, which counts the committed
   branch with the same exclusions CI uses. Over `pr.max_lines`, stop and run
   `split-stack`.
7. **Review locally** with the `review-locally` skill and fix what it finds.
8. **Commit** with a conventional message (`feat:`, `fix:`, `docs:`, `chore:`), then run
   `devloop pr-open` before pushing: never push to a branch whose PR already merged or
   closed. `devloop pr-open --install` makes git run it on every push.
9. **Open the PR** with `gh pr create`. Say what changed and why; for a stateful change,
   put the states, identities and invariants in the description.
10. **Review rounds** follow the `address-review` skill: wait for every reviewer, answer
    every thread, push once.
11. **Merge** only when the required checks pass and every thread is resolved with a
    labelled reply. Merge with a merge commit. In a stack, move the next PR's base to the
    default branch before deleting the merged branch.

## When not to use

- A change over the size limit: plan the stack first with `split-stack`.
- Releasing a version: use `release`.
- Never bypass a gate (`--no-verify`, skipping a check): fix what it reports.

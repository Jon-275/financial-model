---
name: split-stack
description: Plan a change that is over the pull request size limit as a stack of small PRs that each merge on their own, then build and merge them in order. Use when a planned or finished change exceeds pr.max_lines; not for a change that fits in one PR (pr-workflow).
---

# Split a change into a stack

Large PRs take days and dozens of review threads; small ones merge in hours. Plan the
stack before writing code, not after.

## Steps

1. **Measure.** The limit is `pr.max_lines` in `devloop.toml`, counting added plus deleted
   lines, with generated and lock files left out. On a finished branch,
   `devloop size --base origin/main` gives the size as CI counts it.
2. **Cut by concern, in dependency order.** Each PR must build, pass every gate and be
   useful or harmless on its own once merged. Good cuts:
   - a pure function or data model with its tests, before the code that calls it;
   - a new command or module, before the workflow or docs that use it;
   - a refactor with no behaviour change, before the change that needs it.
   Never cut a fix away from its test, or a migration away from the code that reads it.
3. **Write the plan** in the first PR's description: each PR, its concern and its base.
4. **Build the stack.** PR 1 targets the default branch; each later PR targets the branch
   before it. Keep each under the limit; if one grows past it, split it again.
5. **Review rounds** run per PR with `address-review`. A fix lower in the stack is merged
   up into the later branches once, before their next round.
6. **Merge in order** with a merge commit. Before deleting a merged branch, move the next
   PR's base to the default branch (`gh pr edit <next> --base main`): deleting a PR's base
   branch closes that PR.

## When not to use

- A change that fits under the limit: use `pr-workflow`.
- A change that truly cannot be split: label it `pr.large_label` (default
  `large-pr-approved`) and say why in the description, instead of a stack.

---
name: planner-pr-workflow
description: Prepare and review Retirement Planner pull requests with dev-loop and choose the workbook checks needed for formula or runner changes.
---

# Retirement Planner pull requests

Use the maintainer guide at `tests/README.md` from the repository root. It lists the checks and release evidence for the planner. Use fictional workbooks in test and review evidence.

1. Create a topic branch from `origin/main` in its own worktree. Describe the change and how to verify it. For saved financial state, specify the states, identities, legal transitions and invariants before changing behavior. Exercise invalidation, interrupted runs and retry when they apply. An interrupted runner retains the prior batch for inspection. Only a complete matching run appears current.
2. Run the applicable tests and add a changelog fragment when the behavior or workflow changes. Check the fragment with `devloop changelog check`, and run `devloop doctor --check` when editing its configuration or skills. For workbook or VBA changes, record independent comparison and applicable desktop evidence from the maintainer guide.
3. Commit the candidate on the topic branch. Run `devloop size --base origin/main` **after committing**, because it reads the commit graph. For a stack use the actual predecessor ref instead of `origin/main`. Split an oversized change into independently verifiable PRs or document why an exception was approved. The `large-pr-approved` label applies to the hosted size check. Review the committed diff locally and run `devloop pr-open` before pushing the branch for review.
4. Wait for the configured GitHub reviews on the latest commit. Address all findings, run the checks again, commit the fixes, and recheck the committed size before pushing the review round together. Reply with `accepted:` and evidence or `rejected:` and the reason, then resolve the threads. Merge through the PR once the required checks and thread resolution pass.

For a stack, put shared behavior and its tests first. Use merge commits, retarget the dependent PR before deleting its predecessor branch, then check size against `origin/main`. Verify fresh required checks against `main`. If retargeting does not trigger them, merge the updated `main` into the dependent branch and push once to start the checks. For a release, use `docs/publishing.md` and the workbook checks in the maintainer guide.

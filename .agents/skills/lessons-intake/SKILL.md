---
name: lessons-intake
description: Before changing code, read the project's top review lessons and the decisions that bind the area, so the change avoids the bug classes review has already found here. Use at the start of every task that changes code; not for replacing the tests or the review.
---

# Lessons intake

Review finds the same classes of bug again and again. `devloop lessons` ranks them from
labelled review findings and council notes; reading them first is the cheapest review.

## Steps

1. **Read the top lessons** in the `AGENTS.md` block between the `lessons:start` and
   `lessons:end` markers, and the full page (`lessons.doc` in `devloop.toml`, default
   `docs/lessons-learned.md`) for the classes that touch your area.
2. **Read the decisions** in `docs/decisions/` for the area you are changing.
3. **Note which lessons apply** to this change in one line each, and keep them in mind:
   for example, a fix needs a test that fails when the fix is reverted; a check and its
   action run under the same lock; an error never becomes "empty".
4. If a lesson or decision conflicts with the task, stop and ask the owner rather than
   choosing.

## When not to use

- It replaces nothing: the tests, gates and reviews still run.
- To change the lessons themselves: they are generated; run `devloop lessons`.

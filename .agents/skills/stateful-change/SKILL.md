---
name: stateful-change
description: Make a change to stateful code (external writes, queues or outboxes, rollback, persisted evidence, reported figures) by writing its states, identities and invariants first, then an end-state property test, then the code. Use before writing such code; not for stateless helpers, docs or a test-only change.
---

# Stateful change

Bugs in stateful code hide in sequences: a retry after a timeout, a second worker, a
crash between two writes. Example tests miss them; a property over every sequence finds
them. The design comes first, so the test can.

## Steps

1. **Write the design** in the PR description, before any code:
   - **States:** every state a record can be in, and the moves between them.
   - **Identities:** every id a record is known by (local id, remote id, version), and the
     one function that resolves them. Compare whole keys, never part of one.
   - **Invariants:** what must hold after any sequence of operations and faults, for
     example "every field ends at its newest value" or "no write lands after its lease
     expired".
2. **Write the end-state test first**: a property test that runs random sequences of
   operations with injected faults (timeouts, crashes, retries, a second worker, an
   upgrade from data written by the previous release) and asserts every invariant on the
   final state. Watch it fail against a naive implementation.
3. **Implement** until the property holds. Each bug the property finds becomes a fixed
   example test too.
4. **Locks and races:** check and act under the same lock; take locks in one global order,
   proved by a deterministic two-thread test (events, not sleeps).
5. **Upgrades:** data written by the default branch must still load; migrations are
   additive, idempotent and safe if interrupted.
6. **Errors:** never turn a failed read into "empty" or "healthy", and never record a
   lasting verdict from a read that failed.

## When not to use

- Stateless code, docs or tests alone: use `pr-workflow`.
- Do not skip the design because the change looks small; small stateful changes caused
  most of the long review rounds this skill exists to prevent.

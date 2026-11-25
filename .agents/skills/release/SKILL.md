---
name: release
description: Release a new version (bump the version, build the changelog from its fragments, open and merge the release PR, tag the merge commit and publish the release). Use when the owner asks for a release; not for adding a changelog entry to a change (pr-workflow).
---

# Release

A release is its own small PR, merged like any other, then a tag on its merge commit.

## Steps

1. **Pick the version** with the owner: a new feature is a minor bump, a fix a patch, a
   breaking change a major (with an upgrade note).
2. **Branch** from the latest default branch: `release/v<version>` in its own worktree.
3. **Bump the version** everywhere the project records it (the package metadata, the
   lockfile entry, any install line in the README, and `framework.version` in
   `devloop.toml` when this project is devloop itself).
4. **Build the changelog:** `devloop changelog check`, then `devloop changelog build`.
   Move the new entries under a `## <version> - <date>` heading and leave the empty
   `[Unreleased]` headings in place.
5. **Run every gate**, including `devloop doctor --check`, then open the PR with
   `pr-workflow` and wait for its checks. When this project is devloop itself, also run
   `devloop bench` locally and put its scorecard in the PR description.
6. **Merge**, then tag the merge commit and publish the release with that version's
   changelog section as the notes:
   `gh release create v<version> --target <merge sha> --notes-file <notes>`.
7. **Check** the default branch's CI on the merge commit, and install from the new tag
   once to confirm it reports the new version.

## When not to use

- Adding an entry for a change: add a fragment in that change's PR.
- Never tag a commit that is not on the default branch, and never move a published tag.

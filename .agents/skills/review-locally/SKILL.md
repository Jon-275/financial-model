---
name: review-locally
description: Run the project's model reviewers on the current branch with devloop review --local before opening or updating a pull request, then fix or knowingly answer every finding. Use before each push that starts a review round; not for answering threads already on a PR (address-review).
---

# Review locally

The reviewers that gate the pull request can run on your branch first, so a round on the
PR finds less. A local review posts nothing; it prints the summary and each finding.

## Steps

1. Commit your work; the review reads the branch through git, not the working tree.
2. For each reviewer named in `devloop.toml` under `[review.reviewers]`, run
   `devloop review --local --provider <name>` (add `--base <ref>` when the branch does not
   start from `origin/main`). Each needs only its own API key in the environment.
3. For each finding:
   - **Real:** write a test that fails because of it, fix it, and re-run the gates.
   - **Not real:** note why in one line; you will need the reason if the PR's reviewer
     raises it again.
4. Re-run the review once after the fixes. Stop when it raises nothing new; do not loop on
   wording-only findings.
5. If the change touches workflows, secrets or anything a pull request can influence, also
   do a security review of the diff and fix what it finds.

## When not to use

- Threads already posted on the PR: use `address-review`.
- A docs-only change in a tier the project does not review: there is nothing to run.
- Never paste an API key into a command line that is logged; set it in the environment.

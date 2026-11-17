# History cleanup and clone replacement

November 17, 2026

The retained repository history was rewritten to remove earlier workbook copies. The current template and example were then restored from reviewed synthetic artifacts. The current source tree was verified against the reviewed release before the rewrite. The owner's commit email now uses the public GitHub alias.

All 12 completed topic branches were deleted after matching their final commits to merged pull requests. GitHub now deletes branches automatically after a merge; this behavior was verified with the publication cleanup pull request. The Workbook, Greptile and Gemini checks remain required on main.

## Replace an earlier clone

1. Stop using the earlier clone for pushes. Do not merge its branches into the cleaned repository.
2. Save any uncommitted work privately. Review individual source changes before bringing them into the new clone. Keep personalized workbooks outside Git.
3. Clone the existing repository into a new folder:

   ```sh
   git clone https://github.com/Jon-275/financial-model.git
   cd financial-model
   git config --local core.hooksPath .githooks
   python3 -S -B tools/check_workbook_commit.py --tracked
   ```

4. Confirm that the audit verified both tracked workbook files. Set this clone's commit email to the noreply address shown in your GitHub email settings. Use `git config --local user.email` to check the address before committing.
5. Remove the earlier clone after preserving any work you need. A maintainer who must retain a recovery copy should keep it in restricted storage with pushes disabled.

## Verification and remaining exposure

A fresh clone of main passed the workbook content audit. Its reachable history contains only the approved template and example. Twelve earlier task clones were disabled for normal pushes. The owner's checkout was replaced with a fresh clone, and its complete earlier contents and planning edits were preserved in restricted recovery storage.

GitHub's read-only pull-request references and cached views require separate cleanup. They are outside the retained branch history. The owner authorized the prepared Support request, and GitHub confirmed its submission on November 17, 2026. Support's disposition remains pending. The colleague's earlier review clone also needs replacement and confirmation. Previously downloaded copies cannot be verified from this repository.

The maintainer keeps the affected references, recovery location and Support handoff outside Git. Retire recovery copies after Support cleanup and known clone replacements have been verified. Keep any necessary user files before deleting a recovery copy.

For subsequent releases, audit workbook content before updating approval checksums. The commit hook and CI both reject changes to the approved synthetic artifacts. [Release procedure](../tests/README.md) · [GitHub history-removal guidance](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)

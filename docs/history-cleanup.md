# History cleanup and clone replacement

Updated November 18, 2026

The retained repository history was rewritten to remove earlier workbook copies. The current template and example were then restored from reviewed synthetic artifacts. The current source tree was verified against the reviewed release before the rewrite. The owner's commit email now uses the public GitHub alias.

All 12 completed topic branches were deleted after matching their final commits to merged pull requests. The publication cleanup pull request confirmed that GitHub now deletes branches automatically after a merge. The Workbook, Greptile and Gemini checks remain required on main.

## Replace an earlier clone

If you kept a clone from before the history rewrite, replace it before contributing again. Anyone sharing the reviewed workbook from the cleaned repository can use that copy directly.

1. Stop using the earlier clone for pushes. Do not merge its branches into the cleaned repository.
2. Save any uncommitted work privately. Review individual source changes before bringing them into the new clone. Keep personalized workbooks outside Git.
3. Clone the existing repository into a new folder.

   ```sh
   git clone https://github.com/Jon-275/financial-model.git
   cd financial-model
   git config --local core.hooksPath .githooks
   python3 -S -B tools/check_workbook_commit.py --tracked
   ```

4. Confirm that the audit verified both tracked workbook files. Set this clone's commit email to the noreply address shown in your GitHub email settings. Use `git config --local user.email` to check the address before committing.
5. Remove the earlier clone after preserving any work you need. A maintainer who must retain a recovery copy should keep it in restricted storage with pushes disabled.

## Cleanup record

A fresh clone of main passed the workbook content audit. Its reachable history contains only the approved template and example. Twelve earlier task clones were disabled for normal pushes. The owner's checkout was replaced with a fresh clone, and its complete earlier contents and planning edits were preserved in restricted recovery storage.

GitHub received the owner's Support request about earlier pull-request references and cached views on November 17, 2026. On November 18, the owner confirmed that GitHub had deleted the affected history. The fresh-clone audit covered the repository’s reachable history. Individual cached views and previously downloaded workbooks were outside its scope.

One metadata item remains open because a retained squash merge on main lists an author email outside GitHub’s noreply service.

The maintainer keeps the affected references, recovery location and Support correspondence outside Git. Before retiring a restricted recovery copy, preserve any user files that are still needed.

For subsequent releases, audit workbook content before updating approval checksums. The commit hook and CI both reject changes to the approved synthetic artifacts. [Release procedure](../tests/README.md) · [GitHub history-removal guidance](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository)

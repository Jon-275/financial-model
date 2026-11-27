# Website and releases

For maintainers. [Documentation index](README.md) · [Workbook checks](../tests/README.md)

## Website

The [public website](https://jon-275.github.io/financial-model/) uses plain HTML and CSS in `site/`. It has no JavaScript, external fonts or analytics. It links to the workbook guides and packages the approved workbook, example PDF and license as downloads.

Build and preview it locally:

```sh
python3 tools/build_site.py
python3 -m http.server 8000 --directory scratch/site
```

Open `http://localhost:8000`. The builder checks local file links and same-page anchors. It verifies downloads and report previews against the [public release manifest](../site/release-manifest.json) before packaging and again on the packaged files. The build publishes the manifest and a matching `SHA256SUMS.txt` in `downloads/`. The checksums are sorted by filename and use the same format as the GitHub release.

In `site/index.html`, use a quoted GitHub link under `/blob/main/docs/` for each public product guide. The builder pins these links in the generated site to the manifest's `guides_revision` and preserves their page anchors. Keep that revision on the reviewed version of the guides. It can differ from the source commit for the released workbook.

Verify the packaged downloads locally:

```sh
(cd scratch/site/downloads && shasum -a 256 -c SHA256SUMS.txt)
```

The build is a check of the published bytes. Use the [workbook checks](../tests/README.md) to validate calculations and Excel behavior. Keep personal files outside the generated site directory, which the builder replaces.

The **Website** workflow builds pull requests and deploys from `main`. In **Settings → Pages**, keep **GitHub Actions** as the source. Keep the public URL in the repository's About section.

## Report previews

The images in `docs/images/` show pages 1, 3 and 5 of the fictional example PDF. After updating the PDF, regenerate them:

```sh
pdftoppm -f 1 -l 1 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/summary
pdftoppm -f 3 -l 3 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/market-risk
pdftoppm -f 5 -l 5 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/input-decisions
```

Inspect each image. Use fictional figures and label report pages clearly; they are not screenshots of editable inputs.

`docs/images/social-preview.png` is the website's link-preview image. Upload the same image in **Settings → General → Social preview** for the repository.

## Workbook releases

1. Complete the [workbook release checks](../tests/README.md#share-an-updated-workbook), including native Excel checks as needed. A documentation change does not establish new workbook validation.
2. Choose the reviewed artifact source commit and a version tag. Use a patch version for corrections, a minor version for compatible features and a major version for incompatible changes. The manifest identifies the commit that supplies the release bytes. If those bytes change in the activation PR, put them in an earlier commit so the manifest can point to that commit.
3. Update `site/release-manifest.json` with the approved version, artifact source commit, guide commit, file digests and platform evidence. Use the exact tested Mac build and the recorded reconciliation counts. The `v1.0.0` manifest records the observed Apple Silicon run matrix and Windows as unverified. Keep the public guide and site claims within the recorded scope.
4. Review the draft GitHub release for that exact commit. Attach `retirement-model.xlsm`, `retirement-plan-example.pdf`, `LICENSE.txt` and `SHA256SUMS.txt` from the reviewed site build. Record changes, limits, checks performed, Excel version and checks still open.
5. Publish the reviewed release and verify that its tag resolves to the manifest's source commit. Download the files from the versioned release and verify their checksums and the checksum file against the reviewed site build. Activate the version on Pages through the reviewed PR after this check passes.
6. After Pages deploys, download the public files from both channels into separate directories and verify each with its checksum file. Compare both checksum files to the reviewed build. Compare the Pages previews with the manifest and smoke-test the download and guide links from the live site.

Run the publication smoke check from the same worktree as the reviewed site build. It uses `curl` to compare the versioned release downloads with the staged files. Add `--pages` after deployment to compare the published HTML, styles, release metadata, downloads and previews. Open the live guide links separately to confirm the pinned pages and headings.

```sh
python3 tools/check_public_release.py
python3 tools/check_public_release.py --pages
```

If activation needs to be reversed, restore the last verified manifest, matching artifacts, preview images and public guide links through a PR. Keep the original release tag and checksum record intact. Publish corrected files under a new version.

## Repository details

Keep the description, website link and topics current. Current topics include `retirement-planning`, `financial-modeling`, `excel`, `vba`, `python`, `scenario-analysis` and `personal-finance`.

Use fictional figures in issues, screenshots and shared workbooks. Check the README, first-use guide and site for matching requirements and limits before each release.

# Maintain the website and downloads

## Website

The marketing site is plain HTML and CSS in `site/`. It uses no JavaScript, external fonts, analytics or build dependencies. The site links to the full GitHub guides and packages the tracked workbook, example report and license for download.

```sh
python3 tools/build_site.py
python3 -m http.server 8000 --directory scratch/site
```

Open `http://localhost:8000`. The builder checks local asset links and verifies that downloads match their tracked sources. It does not execute the workbook or validate financial assumptions.

In repository **Settings → Pages**, select **GitHub Actions** as the source. The Website workflow builds on pull requests and deploys only from `main`. The expected project URL is `https://jon-275.github.io/financial-model/`. Once deployed, add that URL to the repository's About website field.

## Keep previews accurate

`docs/images/summary.png`, `market-risk.png` and `input-decisions.png` are renders of pages 1, 3 and 5 of the fictional example PDF. They are report previews, not editable-input screenshots. Regenerate after updating the PDF:

```sh
pdftoppm -f 1 -l 1 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/summary
pdftoppm -f 3 -l 3 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/market-risk
pdftoppm -f 5 -l 5 -scale-to 1500 -singlefile -png docs/examples/retirement-plan-example.pdf docs/images/input-decisions
```

Visually inspect the renders. Do not substitute personal reports or invented results. `docs/images/social-preview.png` is a repository social-preview asset; upload it through **Settings → General → Social preview**. The website also references it in Open Graph metadata.

## Versioned releases

Use semantic versions: patch for corrections, minor for compatible features and major for incompatible model changes. Record changes and limitations in each release. Publish a fresh fictional workbook and matching PDF after completing the [workbook release checks](../tests/README.md#share-an-updated-workbook) and native Excel verification as appropriate. A website-only change does not establish new workbook validation.

1. Choose the reviewed commit on `main` and a version tag, such as `v0.1.0` for the initial public workbook.
2. Run the release checks and record the source commit, validation performed, Excel version and any open limitations in the release notes.
3. Attach `retirement-model.xlsm`, `retirement-plan-example.pdf`, `LICENSE.txt` and `SHA256SUMS.txt`. Keep the workbook filename stable for direct release links.
4. Create a draft GitHub Release with the tag targeted at that exact commit. Review assets and notes before publishing. The release remains a draft until explicitly published; a draft does not appear to public visitors.
5. After publication, verify the download, then use `https://github.com/Jon-275/financial-model/releases/latest/download/retirement-model.xlsm` for a latest-release link if desired.

The website's direct download is packaged from the deployed commit. The README downloads from `main` so it works before the first release is published. The Versions link lets users find immutable published releases when available.

## Repository presentation

Suggested About description: **Excel retirement planner for exploring spending, taxes, account access and market risk. Built for Excel for Mac with inspectable calculations.**

Topics: `retirement-planning`, `financial-modeling`, `excel`, `vba`, `python`, `scenario-analysis`, `personal-finance`.

Keep personal data out of issues and screenshots. The issue chooser provides help and bug-report forms with this reminder. Before a release, check that the README, first-use guide, screenshots and site agree about compatibility and limitations.

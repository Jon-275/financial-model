# Website and releases

For maintainers. [Documentation index](README.md) · [Workbook checks](../tests/README.md)

## Website

The [public website](https://jon-275.github.io/financial-model/) uses plain HTML and CSS in `site/`. It has no JavaScript, external fonts or analytics. It links to the workbook guides and packages the tracked workbook, example PDF and license as downloads.

Build and preview it locally:

```sh
python3 tools/build_site.py
python3 -m http.server 8000 --directory scratch/site
```

Open `http://localhost:8000`. The builder checks local links and verifies that the downloads match the tracked files. It does not run the workbook or check financial results. Rebuilding clears the previous generated site; keep personal files elsewhere.

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
2. Choose the reviewed commit on `main` and a version tag. Use a patch version for corrections, a minor version for compatible features and a major version for incompatible changes.
3. Create a draft release for that exact commit. Record changes, limits, checks performed, Excel version and any checks still open.
4. Attach `retirement-model.xlsm`, the matching `retirement-plan-example.pdf`, `LICENSE.txt` and `SHA256SUMS.txt`.
5. Review the notes and files, publish the release, then verify the downloads.

The README links to the latest published workbook release. The website packages the workbook from its deployed commit. Keep the workbook, report, guides and previews consistent.

## Repository details

Keep the description, website link and topics current. Current topics include `retirement-planning`, `financial-modeling`, `excel`, `vba`, `python`, `scenario-analysis` and `personal-finance`.

Use fictional figures in issues, screenshots and shared workbooks. Check the README, first-use guide and site for matching requirements and limits before each release.

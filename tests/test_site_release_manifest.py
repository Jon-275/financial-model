"""Checks for the release metadata and byte-for-byte public site package."""

import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import patch

import build_site


class SiteReleaseManifestTests(unittest.TestCase):
    def setUp(self) -> None:
        """Create a small site and release with independently recorded digests."""
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.output = self.root / "published"
        self.manifest_path = self.root / "site/release-manifest.json"
        self.downloads = {
            "LICENSE.txt": "LICENSE",
            "retirement-model.xlsm": "workbook/retirement-model.xlsm",
            "retirement-plan-example.pdf": "docs/examples/retirement-plan-example.pdf",
        }
        self.previews = ("input-decisions.png", "market-risk.png", "social-preview.png", "summary.png")
        self.manifest = {
            "schema_version": 1,
            "version": "v1.0.0",
            "source_commit": "1" * 40,
            "guides_revision": "2" * 40,
            "compatibility": {
                "mac": {
                    "status": "supported",
                    "label": "Desktop Excel for Mac",
                    "evidence": {
                        "date": "2026-11-25",
                        "architecture": "Apple Silicon",
                        "chip": "Apple M5 Pro",
                        "macos": "26.7.3 (25G342)",
                        "excel_build": "16.115.26112215",
                        "excel_update_channel": "unobserved",
                        "reconciled_runs": 24,
                        "total_runs": 24,
                        "record": "https://example.com/desktop-verification",
                    },
                },
                "windows": {"status": "unverified", "label": "Windows desktop Excel"},
            },
            "downloads": {name: self.create_file(path) for name, path in self.downloads.items()},
            "previews": {name: self.create_file(f"docs/images/{name}") for name in self.previews},
        }
        self.create_file("site/styles.css")
        self.create_file("site/favicon.svg")
        self.create_file(
            "site/index.html",
            '<a href="downloads/retirement-model.xlsm">Download</a>'
            '<a href="downloads/SHA256SUMS.txt">Checksums</a>'
            '<a href="downloads/release-manifest.json">Release manifest</a>',
        )
        self.write_manifest()
        root = patch.object(build_site, "ROOT", self.root)
        root.start()
        self.addCleanup(root.stop)

    def create_file(self, path: str, content: str | None = None) -> str:
        """Write a fixture file and return its actual digest."""
        target = self.root / path
        target.parent.mkdir(parents=True, exist_ok=True)
        data = (content if content is not None else path).encode()
        target.write_bytes(data)
        return hashlib.sha256(data).hexdigest()

    def write_manifest(self) -> None:
        """Save fixture metadata as the reviewed input file."""
        self.manifest_path.write_text(json.dumps(self.manifest), encoding="utf-8")

    def test_build_emits_stable_metadata_and_checksums(self) -> None:
        """Build twice and check the public checksum format and all artifact bytes."""
        build_site.build(self.output)
        build_site.build(self.output)
        downloads = self.output / "downloads"
        self.assertEqual((downloads / "release-manifest.json").read_bytes(), self.manifest_path.read_bytes())
        checksums = "".join(
            f"{self.manifest['downloads'][name]}  {name}\n" for name in sorted(self.downloads)
        ).encode("ascii")
        self.assertEqual((downloads / "SHA256SUMS.txt").read_bytes(), checksums)
        for name, source in self.downloads.items():
            self.assertEqual((downloads / name).read_bytes(), (self.root / source).read_bytes())
        for name in self.previews:
            self.assertEqual(
                (self.output / "assets" / name).read_bytes(),
                (self.root / "docs/images" / name).read_bytes(),
            )

    def test_rejects_drift_before_removing_last_build(self) -> None:
        """Reject a different workbook, PDF, license or public preview before rebuilding."""
        build_site.build(self.output)
        checksums = (self.output / "downloads/SHA256SUMS.txt").read_bytes()
        for path in (*self.downloads.values(), "docs/images/summary.png"):
            with self.subTest(path=path):
                self.create_file(path, "different bytes")
                with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                    build_site.build(self.output)
                self.create_file(path)
        self.assertEqual((self.output / "downloads/SHA256SUMS.txt").read_bytes(), checksums)

    def test_rejects_missing_artifact_digest(self) -> None:
        """Require an approved digest for every public release file."""
        del self.manifest["downloads"]["retirement-plan-example.pdf"]
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, "downloads"):
            build_site.build(self.output)
        self.assertFalse(self.output.exists())

    def test_rejects_invalid_release_metadata(self) -> None:
        """Require a real release version and commit, one supported build and unverified Windows."""
        changed = (
            ("version", "latest", "version"),
            ("source_commit", "632e942", "source_commit"),
            ("guides_revision", "main", "guides_revision"),
            (
                "compatibility",
                {**self.manifest["compatibility"], "windows": {"status": "verified"}},
                "compatibility for windows",
            ),
        )
        for key, value, field in changed:
            with self.subTest(field=field):
                original = self.manifest[key]
                self.manifest[key] = value
                self.write_manifest()
                with self.assertRaisesRegex(ValueError, field):
                    build_site.build(self.output)
                self.manifest[key] = original
        self.assertFalse(self.output.exists())

    def test_pins_product_guides_and_preserves_link_fragments(self) -> None:
        """Rewrite only project guide placeholders in the generated HTML."""
        author_html = (
            '<a href="https://github.com/Jon-275/financial-model/blob/main/docs/getting-started.md">Start</a>'
            "<a href='https://github.com/Jon-275/financial-model/blob/main/docs/risk-analysis.md"
            "?plain=1#measured-runtime'>Runtime</a>"
            '<a href="https://github.com/Jon-275/financial-model/blob/main/README.md">Project</a>'
            '<a href="https://github.com/example/repo/blob/main/docs/index.md">Other project</a>'
        )
        self.create_file("site/index.html", author_html)
        build_site.build(self.output)
        published = (self.output / "index.html").read_text(encoding="utf-8")
        self.assertIn(f'/blob/{"2" * 40}/docs/getting-started.md', published)
        self.assertIn(f'/blob/{"2" * 40}/docs/risk-analysis.md?plain=1#measured-runtime', published)
        self.assertIn("/financial-model/blob/main/README.md", published)
        self.assertIn("/example/repo/blob/main/docs/index.md", published)
        self.assertEqual((self.root / "site/index.html").read_text(encoding="utf-8"), author_html)

    def test_rejects_unpinned_guide_placeholder(self) -> None:
        """A guide link outside the quoted placeholder format must fail the build."""
        self.create_file(
            "site/index.html",
            "<a href=https://github.com/Jon-275/financial-model/blob/main/docs/README.md>Guides</a>",
        )
        with self.assertRaisesRegex(ValueError, "quoted guide link"):
            build_site.build(self.output)

    def test_rejects_incomplete_mac_evidence(self) -> None:
        """Mac support requires an observed build and fully reconciled recorded runs."""
        evidence = self.manifest["compatibility"]["mac"]["evidence"]
        for key, value in (("excel_build", ""), ("reconciled_runs", 23)):
            with self.subTest(field=key):
                original = evidence[key]
                evidence[key] = value
                self.write_manifest()
                with self.assertRaisesRegex(ValueError, "compatibility evidence"):
                    build_site.build(self.output)
                evidence[key] = original

    def test_rejects_corruption_during_copy(self) -> None:
        """Check packaged bytes before emitting public checksum files."""
        original_copy = shutil.copy2

        def corrupt_copy(source: Path, target: Path) -> Path:
            """Simulate corruption after a source file passed preflight."""
            original_copy(source, target)
            if target.name == "retirement-plan-example.pdf":
                target.write_bytes(b"changed on copy")
            return target

        with patch.object(build_site.shutil, "copy2", side_effect=corrupt_copy):
            with self.assertRaisesRegex(ValueError, "SHA-256 mismatch"):
                build_site.build(self.output)
        self.assertFalse((self.output / "downloads/SHA256SUMS.txt").exists())


if __name__ == "__main__":
    unittest.main()

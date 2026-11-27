"""Compare a staged Pages build with the public release and deployed site."""

import argparse
from filecmp import cmp
from pathlib import Path
import subprocess
from tempfile import TemporaryDirectory

from build_site import ROOT, checksum_bytes, read_manifest, verify_digest

RELEASES = "https://github.com/Jon-275/financial-model/releases/download"
PAGES = "https://jon-275.github.io/financial-model"


def verify_public_file(url: str, staged: Path) -> None:
    """Fetch a public file over bounded HTTPS and compare all bytes."""
    with TemporaryDirectory() as temp:
        received = Path(temp) / "received"
        command = [
            "curl", "--disable", "--fail", "--silent", "--show-error", "--location",
            "--proto", "=https", "--proto-redir", "=https",
            "--connect-timeout", "5", "--max-time", "30",
            "--max-filesize", str(staged.stat().st_size),
            "--output", str(received), url,
        ]
        try:
            subprocess.run(command, check=True, timeout=35)
        except subprocess.CalledProcessError as error:
            raise ValueError(f"Could not fetch {url} (curl exit {error.returncode})") from error
        except subprocess.TimeoutExpired as error:
            raise ValueError(f"Timed out fetching {url}") from error
        if not cmp(staged, received, shallow=False):
            raise ValueError(f"Publication mismatch at {url} against {staged}")


def check(site: Path, *, pages: bool) -> None:
    """Verify staged and published files before reporting the active release."""
    downloads = site / "downloads"
    version, _, digests, previews = read_manifest((downloads / "release-manifest.json").read_bytes())
    if (downloads / "SHA256SUMS.txt").read_bytes() != checksum_bytes(digests):
        raise ValueError("The staged checksum file does not match the release manifest")
    for name, digest in digests.items():
        verify_digest(downloads / name, digest)

    release_files = ("SHA256SUMS.txt", *sorted(digests))
    for name in release_files:
        verify_public_file(f"{RELEASES}/{version}/{name}", downloads / name)
    print(f"GitHub release {version} matches {downloads}")

    if pages:
        for name, digest in previews.items():
            verify_digest(site / "assets" / name, digest)
        public_paths = (
            "index.html",
            "styles.css",
            "assets/favicon.svg",
            *(f"assets/{name}" for name in sorted(previews)),
            *(f"downloads/{name}" for name in (*release_files, "release-manifest.json")),
        )
        for path in public_paths:
            url_path = "" if path == "index.html" else path
            verify_public_file(f"{PAGES}/{url_path}", site / path)
        print(f"GitHub Pages matches {site}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--site", default=ROOT / "scratch/site", type=Path, help="local build output")
    parser.add_argument("--pages", action="store_true", help="also compare the deployed Pages files")
    args = parser.parse_args()
    try:
        check(args.site, pages=args.pages)
    except (OSError, ValueError, subprocess.SubprocessError) as error:
        parser.exit(1, f"{error}\n")

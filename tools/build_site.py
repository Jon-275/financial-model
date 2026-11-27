"""Build the dependency-free GitHub Pages site from tracked release artifacts."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
from html.parser import HTMLParser
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]
DOWNLOAD_SOURCES = {
    "LICENSE.txt": "LICENSE",
    "retirement-model.xlsm": "workbook/retirement-model.xlsm",
    "retirement-plan-example.pdf": "docs/examples/retirement-plan-example.pdf",
}
PREVIEWS = {"input-decisions.png", "market-risk.png", "social-preview.png", "summary.png"}
GUIDE_REPO = "https://github.com/Jon-275/financial-model/blob/"
GUIDE_PLACEHOLDER = f"{GUIDE_REPO}main/docs/"


def read_manifest(manifest_bytes: bytes) -> tuple[str, str, dict[str, str], dict[str, str]]:
    """Validate release metadata and return the version, guides revision and digests."""
    manifest = json.loads(manifest_bytes)
    if not isinstance(manifest, dict):
        raise ValueError("Invalid release manifest object")
    schema_version = manifest.get("schema_version")
    if type(schema_version) is not int or schema_version != 1:
        raise ValueError("Invalid release manifest schema_version")
    fields = (
        ("version", r"v[0-9]+\.[0-9]+\.[0-9]+"),
        ("source_commit", r"[0-9a-f]{40}"),
        ("guides_revision", r"[0-9a-f]{40}"),
    )
    for field, pattern in fields:
        value = manifest.get(field)
        if not isinstance(value, str) or not re.fullmatch(pattern, value):
            raise ValueError(f"Invalid release manifest {field}")

    compatibility = manifest.get("compatibility")
    if not isinstance(compatibility, dict):
        raise ValueError("Invalid release manifest compatibility")
    for platform, status in (("mac", "supported"), ("windows", "unverified")):
        entry = compatibility.get(platform)
        if not isinstance(entry, dict) or entry.get("status") != status:
            raise ValueError(f"Invalid release manifest compatibility for {platform}")
        label = entry.get("label")
        if not isinstance(label, str) or not label.strip():
            raise ValueError(f"Invalid release manifest compatibility for {platform}")
    evidence = compatibility["mac"].get("evidence")
    evidence_fields = ("date", "architecture", "chip", "macos", "excel_build", "excel_update_channel", "record")
    if not isinstance(evidence, dict) or any(
        not isinstance(evidence.get(field), str) or not evidence[field].strip() for field in evidence_fields
    ):
        raise ValueError("Invalid release manifest compatibility evidence for Mac")
    reconciled = evidence.get("reconciled_runs")
    total = evidence.get("total_runs")
    if type(reconciled) is not int or type(total) is not int or reconciled <= 0 or reconciled != total:
        raise ValueError("Invalid release manifest compatibility evidence run counts")

    for field, names in (("downloads", set(DOWNLOAD_SOURCES)), ("previews", PREVIEWS)):
        digests = manifest.get(field)
        if not isinstance(digests, dict) or set(digests) != names:
            raise ValueError(f"Invalid release manifest {field}; expected {', '.join(sorted(names))}")
        for name, digest in digests.items():
            if not isinstance(digest, str) or not re.fullmatch(r"[0-9a-f]{64}", digest):
                raise ValueError(f"Invalid SHA-256 in release manifest {field}: {name}")
    return manifest["version"], manifest["guides_revision"], manifest["downloads"], manifest["previews"]


def checksum_bytes(download_digests: dict[str, str]) -> bytes:
    """Format the public SHA-256 list in stable filename order."""
    return "".join(f"{digest}  {name}\n" for name, digest in sorted(download_digests.items())).encode("ascii")


def verify_digest(path: Path, expected: str) -> None:
    """Reject bytes that differ from the approved release artifact."""
    with path.open("rb") as stream:
        actual = hashlib.file_digest(stream, "sha256").hexdigest()
    if actual != expected:
        raise ValueError(f"SHA-256 mismatch for {path}: expected {expected}, got {actual}")


def build(output: Path) -> None:
    """Build the site after verifying the approved release inputs."""
    output = output.resolve()
    marker = output / ".retirement-planner-site"
    if output.exists() and any(output.iterdir()):
        if not marker.is_file():
            raise ValueError("Choose an empty directory or an existing generated site directory")

    manifest_bytes = (ROOT / "site/release-manifest.json").read_bytes()
    _, guides_revision, download_digests, preview_digests = read_manifest(manifest_bytes)
    artifacts = {
        **{f"downloads/{name}": (ROOT / source, download_digests[name]) for name, source in DOWNLOAD_SOURCES.items()},
        **{f"assets/{name}": (ROOT / "docs/images" / name, digest) for name, digest in preview_digests.items()},
    }
    for source, digest in artifacts.values():
        verify_digest(source, digest)

    if marker.is_file():
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)
    marker.touch()
    index = (ROOT / "site/index.html").read_text(encoding="utf-8")
    index = re.sub(
        r"""(\shref\s*=\s*['"])""" + re.escape(GUIDE_PLACEHOLDER),
        rf"\g<1>{GUIDE_REPO}{guides_revision}/docs/",
        index,
        flags=re.IGNORECASE,
    )
    (output / "index.html").write_bytes(index.encode("utf-8"))
    shutil.copy2(ROOT / "site/styles.css", output / "styles.css")
    assets = output / "assets"
    assets.mkdir(exist_ok=True)
    shutil.copy2(ROOT / "site/favicon.svg", assets / "favicon.svg")
    downloads = output / "downloads"
    downloads.mkdir(exist_ok=True)
    for target, (source, digest) in artifacts.items():
        shutil.copy2(source, output / target)
        verify_digest(output / target, digest)

    (downloads / "release-manifest.json").write_bytes(manifest_bytes)
    (downloads / "SHA256SUMS.txt").write_bytes(checksum_bytes(download_digests))
    (output / ".nojekyll").touch()

    class Links(HTMLParser):
        def __init__(self) -> None:
            """Collect same-page link fragments and IDs for a complete-document check."""
            super().__init__()
            self.ids: set[str] = set()
            self.fragments: set[str] = set()

        def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
            """Check local paths and collect same-page link fragments."""
            for key, value in attrs:
                if key == "id" and value:
                    self.ids.add(value)
                if key not in ("href", "src") or not value:
                    continue
                if key == "href" and value.lower().startswith(GUIDE_PLACEHOLDER.lower()):
                    raise ValueError(f"Use a quoted guide link so it can be pinned: {value}")
                url = urlsplit(value)
                if url.scheme or url.netloc:
                    continue
                target = (output / unquote(url.path)).resolve()
                if not target.is_relative_to(output) or not target.exists():
                    raise ValueError(f"Missing or unsafe site target: {value}")
                if key == "href" and url.fragment and target in (output, output / "index.html"):
                    self.fragments.add(unquote(url.fragment))

    links = Links()
    links.feed(index)
    links.close()
    missing = sorted(links.fragments - links.ids)
    if missing:
        raise ValueError("Missing site fragment: " + ", ".join(f"#{fragment}" for fragment in missing))
    print(f"Built site and verified local links and release artifact hashes: {output}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "scratch/site")
    build(parser.parse_args().output)

"""Build the dependency-free GitHub Pages site from tracked release artifacts."""
from pathlib import Path
import argparse
import hashlib
import shutil
from html.parser import HTMLParser
from urllib.parse import unquote, urlsplit

ROOT = Path(__file__).resolve().parents[1]


def build(output):
    output = output.resolve()
    marker = output / ".retirement-planner-site"
    if output.exists() and any(output.iterdir()):
        if not marker.is_file():
            raise ValueError("Choose an empty directory or an existing generated site directory")
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)
    marker.touch()
    for name in ("index.html", "styles.css"):
        shutil.copy2(ROOT / "site" / name, output / name)
    assets = output / "assets"
    assets.mkdir(exist_ok=True)
    for path in (ROOT / "docs/images").glob("*.png"):
        shutil.copy2(path, assets / path.name)
    shutil.copy2(ROOT / "site/favicon.svg", assets / "favicon.svg")
    downloads = output / "downloads"
    downloads.mkdir(exist_ok=True)
    sources = {
        "retirement-model.xlsm": ROOT / "workbook/retirement-model.xlsm",
        "retirement-plan-example.pdf": ROOT / "docs/examples/retirement-plan-example.pdf",
        "LICENSE.txt": ROOT / "LICENSE",
    }
    for name, source in sources.items():
        shutil.copy2(source, downloads / name)
        assert hashlib.sha256(source.read_bytes()).digest() == hashlib.sha256((downloads / name).read_bytes()).digest()
    (output / ".nojekyll").touch()

    class Links(HTMLParser):
        def handle_starttag(self, tag, attrs):
            for key, value in attrs:
                if key not in ("href", "src") or not value:
                    continue
                url = urlsplit(value)
                if url.scheme or url.netloc or not url.path:
                    continue
                target = (output / unquote(url.path)).resolve()
                if not target.is_relative_to(output) or not target.exists():
                    raise ValueError(f"Missing or unsafe site target: {value}")
    Links().feed((output / "index.html").read_text())
    print(f"Built site and verified local links and download hashes: {output}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "scratch/site")
    build(parser.parse_args().output)

"""Inline presentation SVG illustrations into ai-for-project-managers.html."""

from __future__ import annotations

import re
from pathlib import Path

PRESENTATION_DIR = Path(__file__).resolve().parent.parent
HTML_PATH = PRESENTATION_DIR / "ai-for-project-managers.html"
IMAGES_DIR = PRESENTATION_DIR / "images"

IMG_PATTERN = re.compile(
    r'<img src="images/([^"]+\.svg)" alt="" width="[^"]*" height="[^"]*" />'
)


def uniquify_svg(svg: str, suffix: int) -> str:
    """Avoid duplicate id/marker conflicts when the same SVG appears twice."""
    ids = re.findall(r'\bid="([^"]+)"', svg)
    for id_val in ids:
        new_id = f"{id_val}-{suffix}"
        svg = svg.replace(f'id="{id_val}"', f'id="{new_id}"')
        svg = svg.replace(f"url(#{id_val})", f"url(#{new_id})")
    return svg


def embed() -> None:
    html = HTML_PATH.read_text(encoding="utf-8")
    counter = 0

    def replace_img(match: re.Match[str]) -> str:
        nonlocal counter
        counter += 1
        filename = match.group(1)
        svg_path = IMAGES_DIR / filename
        if not svg_path.exists():
            raise FileNotFoundError(f"Missing illustration: {svg_path}")
        svg = svg_path.read_text(encoding="utf-8").strip()
        if svg.startswith("<?"):
            svg = svg.split(">", 1)[1].strip()
        svg = uniquify_svg(svg, counter)
        return svg

    updated = IMG_PATTERN.sub(replace_img, html)
    if counter == 0:
        raise RuntimeError("No image tags found to embed.")
    HTML_PATH.write_text(updated, encoding="utf-8")
    print(f"Embedded {counter} inline SVG(s) into {HTML_PATH.name}")


if __name__ == "__main__":
    embed()

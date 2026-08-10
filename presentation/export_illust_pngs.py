"""Render deck illustration SVGs to PNGs for PowerPoint embed."""
from __future__ import annotations

from pathlib import Path

from playwright.sync_api import sync_playwright

IMG = Path(__file__).resolve().parent / "images"
JOBS = [
    ("illust-title.svg", "illust-title.png", 560, 400, 3),
    ("illust-you-lead.svg", "illust-you-lead.png", 960, 560, 2),
    ("illust-narrative.svg", "illust-narrative.png", 800, 340, 3),
]


def main() -> None:
    with sync_playwright() as p:
        browser = p.chromium.launch()
        for svg_name, png_name, w, h, scale in JOBS:
            context = browser.new_context(
                viewport={"width": w, "height": h},
                device_scale_factor=scale,
            )
            page = context.new_page()
            svg = (IMG / svg_name).read_text(encoding="utf-8")
            html = f"""<!DOCTYPE html>
<html><head>
<meta charset="utf-8"/>
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600&display=swap" rel="stylesheet"/>
<style>
  html,body{{margin:0;padding:0;background:#fff;}}
  .wrap{{width:{w}px;height:{h}px;display:flex;align-items:center;justify-content:center;}}
  svg{{width:100%;height:100%;}}
</style>
</head><body><div class="wrap">{svg}</div></body></html>"""
            page.set_content(html, wait_until="networkidle")
            page.wait_for_timeout(600)
            out = IMG / png_name
            page.locator(".wrap").screenshot(path=str(out))
            print(f"Wrote {out.name}: {out.stat().st_size} bytes @ {scale}x")
            context.close()
        browser.close()


if __name__ == "__main__":
    main()

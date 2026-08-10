"""Generate circular Pantone-style fan wheel SVG."""
from __future__ import annotations

import math
from pathlib import Path

SEGMENTS = 48
RINGS = 6
CX, CY = 400, 400
R_INNER = 26
R_OUTER = 390

out_path = Path(__file__).resolve().parent.parent / "assets" / "pantone-wheel.svg"

parts: list[str] = [
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 800" aria-hidden="true">',
    '<rect width="800" height="800" fill="none"/>',
    f'<circle cx="{CX}" cy="{CY}" r="18" fill="#ffffff" stroke="#e0e0e0" stroke-width="1"/>',
]

for i in range(SEGMENTS):
    a0 = (i / SEGMENTS) * 2 * math.pi - math.pi / 2
    a1 = ((i + 1) / SEGMENTS) * 2 * math.pi - math.pi / 2
    hue = (i / SEGMENTS) * 360
    for ring in range(RINGS):
        r0 = R_INNER + (ring / RINGS) * (R_OUTER - R_INNER)
        r1 = R_INNER + ((ring + 1) / RINGS) * (R_OUTER - R_INNER)
        sat = 52 + ring * 7
        light = 74 - ring * 8
        color = f"hsl({hue:.0f}, {sat}%, {light}%)"
        x0o, y0o = CX + r0 * math.cos(a0), CY + r0 * math.sin(a0)
        x1o, y1o = CX + r0 * math.cos(a1), CY + r0 * math.sin(a1)
        x0x, y0x = CX + r1 * math.cos(a0), CY + r1 * math.sin(a0)
        x1x, y1x = CX + r1 * math.cos(a1), CY + r1 * math.sin(a1)
        large = 0 if (a1 - a0) <= math.pi else 1
        d = (
            f"M {x0o:.1f} {y0o:.1f} A {r0:.1f} {r0:.1f} 0 {large} 1 {x1o:.1f} {y1o:.1f} "
            f"L {x1x:.1f} {y1x:.1f} A {r1:.1f} {r1:.1f} 0 {large} 0 {x0x:.1f} {y0x:.1f} Z"
        )
        parts.append(f'<path d="{d}" fill="{color}" stroke="#ffffff" stroke-width="0.55"/>')

parts.append('<g transform="translate(400,400) rotate(20)">')
parts.append('<rect x="52" y="-16" width="118" height="32" rx="2" fill="#8a7768" stroke="#fff" stroke-width="1.5"/>')
parts.append('<circle cx="52" cy="0" r="3" fill="#c0c0c0"/>')
parts.append("</g>")
parts.append("</svg>")

out_path.write_text("\n".join(parts), encoding="utf-8")
print(out_path)

"""Generate compact Pantone-style wheel for site header logo."""
from __future__ import annotations

import math
from pathlib import Path

SEGMENTS = 36
RINGS = 5
SIZE = 64
CX = CY = SIZE / 2
R_INNER = 5
R_OUTER = SIZE / 2 - 1

out_path = Path(__file__).resolve().parent.parent / "assets" / "logo-wheel.svg"

parts: list[str] = [
    f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {SIZE} {SIZE}" aria-hidden="true">',
    f'<circle cx="{CX}" cy="{CY}" r="4.5" fill="#ffffff" stroke="#e8e8e8" stroke-width="0.75"/>',
]

for i in range(SEGMENTS):
    a0 = (i / SEGMENTS) * 2 * math.pi - math.pi / 2
    a1 = ((i + 1) / SEGMENTS) * 2 * math.pi - math.pi / 2
    hue = (i / SEGMENTS) * 360
    for ring in range(RINGS):
        r0 = R_INNER + (ring / RINGS) * (R_OUTER - R_INNER)
        r1 = R_INNER + ((ring + 1) / RINGS) * (R_OUTER - R_INNER)
        sat = 54 + ring * 8
        light = 72 - ring * 9
        color = f"hsl({hue:.0f}, {sat}%, {light}%)"
        large = 0 if (a1 - a0) <= math.pi else 1
        x0o, y0o = CX + r0 * math.cos(a0), CY + r0 * math.sin(a0)
        x1o, y1o = CX + r0 * math.cos(a1), CY + r0 * math.sin(a1)
        x0x, y0x = CX + r1 * math.cos(a0), CY + r1 * math.sin(a0)
        x1x, y1x = CX + r1 * math.cos(a1), CY + r1 * math.sin(a1)
        d = (
            f"M {x0o:.2f} {y0o:.2f} A {r0:.2f} {r0:.2f} 0 {large} 1 {x1o:.2f} {y1o:.2f} "
            f"L {x1x:.2f} {y1x:.2f} A {r1:.2f} {r1:.2f} 0 {large} 0 {x0x:.2f} {y0x:.2f} Z"
        )
        parts.append(f'<path d="{d}" fill="{color}" stroke="#ffffff" stroke-width="1.1"/>')

parts.append("</svg>")
out_path.write_text("\n".join(parts), encoding="utf-8")
print(out_path)

# -*- coding: utf-8 -*-
"""Generate app.ico from the groma_cross SVG design.

SVG geometry (100x100 viewBox):
- Horizontal bar: rect(18, 46, 64, 8), rx=1
- Vertical bar:   rect(46, 18, 8, 64), rx=1
- Center void:    rect(44, 44, 12, 12) [white fill]
- Center dot:     circle(50, 50, r=2.5)

Foreground color: currentColor → we'll pick a solid dark color.
"""
from PIL import Image, ImageDraw

SIZES = [16, 32, 48, 64, 128, 256]
FG = (30, 35, 45, 255)      # near-black, slight blue tint
BG = (255, 255, 255, 0)      # transparent
VOID = (255, 255, 255, 0)    # transparent void (matches bg)


def render(size):
    # Render at 4x super-sampling then downscale for crisper edges
    scale = 4
    S = size * scale
    img = Image.new("RGBA", (S, S), BG)
    draw = ImageDraw.Draw(img)

    def u(v):
        """SVG unit (0..100) → pixels at current size."""
        return v / 100.0 * S

    def rect(x, y, w, h, fill, radius=0):
        if radius > 0:
            draw.rounded_rectangle(
                [u(x), u(y), u(x + w), u(y + h)],
                radius=u(radius),
                fill=fill,
            )
        else:
            draw.rectangle([u(x), u(y), u(x + w), u(y + h)], fill=fill)

    # Horizontal bar
    rect(18, 46, 64, 8, FG, radius=1)
    # Vertical bar
    rect(46, 18, 8, 64, FG, radius=1)
    # Center void (punch out to transparent)
    rect(44, 44, 12, 12, VOID)
    # Center precision dot
    r = 2.5
    draw.ellipse(
        [u(50 - r), u(50 - r), u(50 + r), u(50 + r)],
        fill=FG,
    )

    return img.resize((size, size), Image.LANCZOS)


def build_ico(path):
    # Render a large master image; Pillow will create each ICO size from it
    master = render(256)
    master.save(path, format="ICO", sizes=[(s, s) for s in SIZES])


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else r"C:\temp\app.ico"
    build_ico(out)
    # Also save a big PNG for preview
    preview = render(512)
    preview.save(out.replace(".ico", "_preview.png"))
    import os
    print(f"Wrote {out} ({os.path.getsize(out)} bytes)")

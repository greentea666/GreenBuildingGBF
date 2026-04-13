# -*- coding: utf-8 -*-
"""Generate a retro-style app icon for GreenBuildingGBF"""
from PIL import Image, ImageDraw, ImageFont
import os

SIZES = [16, 32, 48, 64, 128, 256]
OUT = os.path.join(os.path.dirname(__file__), "app.ico")


def draw_icon(size):
    """Draw a single icon at given size."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # ── Background: warm off-white rounded square ──
    margin = max(1, size // 16)
    bg_rect = [margin, margin, size - margin - 1, size - margin - 1]
    # Rounded rect background
    r = max(2, size // 8)
    d.rounded_rectangle(bg_rect, radius=r, fill="#F5F3EE", outline="#1A1A1A",
                        width=max(1, size // 64))

    # ── Amber accent bar at top ──
    bar_h = max(2, size // 8)
    bar_rect = [margin + 1, margin + 1, size - margin - 2, margin + bar_h]
    d.rectangle(bar_rect, fill="#F0A500")

    # ── Text: "GBF" ──
    try:
        font_size = max(8, size * 38 // 100)
        font = ImageFont.truetype("consola.ttf", font_size)
    except Exception:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/consola.ttf", font_size)
        except Exception:
            font = ImageFont.load_default()

    text = "GBF"
    bbox = d.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (size - tw) // 2
    ty = margin + bar_h + (size - margin * 2 - bar_h - th) // 2 - max(1, size // 20)
    d.text((tx, ty), text, fill="#1A1A1A", font=font)

    # ── Small leaf/green accent at bottom-right ──
    leaf_size = max(3, size // 6)
    lx = size - margin - leaf_size - max(2, size // 10)
    ly = size - margin - leaf_size - max(2, size // 12)
    d.ellipse([lx, ly, lx + leaf_size, ly + leaf_size], fill="#2D8A4E")

    return img


# Generate all sizes
icons = [draw_icon(s) for s in SIZES]

# Save as .ico with multiple sizes
icons[0].save(OUT, format="ICO", sizes=[(s, s) for s in SIZES],
              append_images=icons[1:])
print(f"Icon saved: {OUT}")

"""
Generates the PO Data Extractor logo and Windows icon.

    python make_assets.py

Writes into the "assets" folder:
    app.ico       - multi-size icon (16-256 px) for the .exe, window and taskbar
    logo_<n>.png  - logo at several sizes; the app picks the one that best
                    fits the screen's DPI (header and About dialog)

The design is drawn large and scaled down, and is kept deliberately bold
(solid tile, high-contrast white document, bright green badge) so it stays
recognisable even at 16 px in the taskbar.
"""

import os
from PIL import Image, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

SIZE = 1024
LOGO_SIZES = (48, 56, 64, 72, 80, 96, 112, 128)
TILE_TOP = (24, 140, 80)      # green
TILE_BOTTOM = (10, 84, 45)    # deep green
DOC = (255, 255, 255)
DOC_FOLD = (200, 230, 212)
LINE = (170, 178, 189)
GRID = (16, 124, 65)
BADGE = (31, 35, 40)          # charcoal
BADGE_RING = (255, 255, 255)


def vertical_gradient(size, top, bottom):
    grad = Image.new("RGB", (1, size))
    for y in range(size):
        t = y / (size - 1)
        grad.putpixel((0, y), tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return grad.resize((size, size))


def draw_logo(size=SIZE):
    s = size / 1024
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    # Rounded tile with gradient.
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, size - 1, size - 1],
                                           radius=int(220 * s), fill=255)
    img.paste(vertical_gradient(size, TILE_TOP, TILE_BOTTOM), (0, 0), mask)

    # Soft shadow under the document.
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(
        [int(230 * s), int(170 * s), int(770 * s), int(880 * s)],
        radius=int(50 * s), fill=(0, 0, 0, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(int(28 * s)))
    img = Image.alpha_composite(img, shadow)

    d = ImageDraw.Draw(img)

    # Document with folded corner.
    left, top, right, bottom = 210 * s, 140 * s, 750 * s, 850 * s
    fold = 170 * s
    d.rounded_rectangle([left, top, right, bottom], radius=int(48 * s), fill=DOC)
    # Cut the top-right corner and draw the fold.
    d.polygon([(right - fold, top - 2), (right + 2, top - 2), (right + 2, top + fold)],
              fill=TILE_TOP)
    d.polygon([(right - fold, top), (right, top + fold), (right - fold, top + fold)],
              fill=DOC_FOLD)

    # Header text lines ("PO" heading feel).
    d.rounded_rectangle([left + 60 * s, top + 70 * s, left + 290 * s, top + 120 * s],
                        radius=int(25 * s), fill=TILE_BOTTOM)
    d.rounded_rectangle([left + 60 * s, top + 160 * s, left + 360 * s, top + 195 * s],
                        radius=int(17 * s), fill=LINE)

    # Spreadsheet grid (the "to Excel" part).
    gx0, gy0, gx1, gy1 = left + 60 * s, top + 260 * s, right - 60 * s, bottom - 70 * s
    w = int(26 * s)
    d.rounded_rectangle([gx0, gy0, gx1, gy1], radius=int(24 * s), outline=GRID, width=w)
    rows = 4
    for i in range(1, rows):
        y = gy0 + (gy1 - gy0) * i / rows
        d.line([(gx0, y), (gx1, y)], fill=GRID, width=int(18 * s))
    cx = gx0 + (gx1 - gx0) * 0.38
    d.line([(cx, gy0), (cx, gy1)], fill=GRID, width=int(18 * s))

    # Green check badge, bottom-right.
    r = 185 * s
    bx, by = 760 * s, 770 * s
    d.ellipse([bx - r - 26 * s, by - r - 26 * s, bx + r + 26 * s, by + r + 26 * s],
              fill=BADGE_RING)
    d.ellipse([bx - r, by - r, bx + r, by + r], fill=BADGE)
    d.line([(bx - 85 * s, by + 5 * s), (bx - 22 * s, by + 70 * s), (bx + 95 * s, by - 65 * s)],
           fill=(255, 255, 255), width=int(56 * s), joint="curve")

    return img


def main():
    os.makedirs(ASSETS, exist_ok=True)
    big = draw_logo()

    for size in LOGO_SIZES:
        big.resize((size, size), Image.LANCZOS).save(os.path.join(ASSETS, f"logo_{size}.png"))
    big.resize((256, 256), Image.LANCZOS).save(
        os.path.join(ASSETS, "app.ico"),
        sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48),
               (64, 64), (128, 128), (256, 256)])
    print(f"Assets written to {ASSETS}")


if __name__ == "__main__":
    main()

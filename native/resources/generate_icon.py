#!/usr/bin/env python3
"""Generate Nexel app icon — bold spreadsheet with chart cards on the left."""

from PIL import Image, ImageDraw, ImageFont
import os
import subprocess
import math

SIZE = 1024
CORNER_RADIUS = 220


def rounded_rectangle_mask(size, radius):
    mask = Image.new('L', (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return mask


def draw_icon(size=SIZE):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))

    # Background: clean deep green
    for y in range(size):
        t = y / size
        r = int(20 + 15 * t)
        g = int(100 - 20 * t)
        b = int(60 - 10 * t)
        for x in range(size):
            img.putpixel((x, y), (r, g, b, 255))

    draw = ImageDraw.Draw(img)

    # === FONTS ===
    try:
        font_hdr = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.028))
        font_num = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.022))
        font_sel = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.030))
    except Exception:
        font_hdr = ImageFont.load_default()
        font_num = font_hdr
        font_sel = font_hdr

    # ==========================================
    #  TABLE CARD — top-right, ~75% of the icon
    # ==========================================
    t_left = int(size * 0.26)
    t_top = int(size * 0.08)
    t_right = size - int(size * 0.06)
    t_bottom = size - int(size * 0.08)
    t_rad = int(size * 0.035)

    # Shadow
    shadow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    s_draw = ImageDraw.Draw(shadow)
    for i in range(20):
        a = int(25 * (1 - i / 20))
        s_draw.rounded_rectangle(
            [t_left + i, t_top + i + 6, t_right + i, t_bottom + i + 6],
            radius=t_rad + i, fill=(0, 0, 0, a)
        )
    img = Image.alpha_composite(img, shadow)
    draw = ImageDraw.Draw(img)

    # Card
    draw.rounded_rectangle([t_left, t_top, t_right, t_bottom],
                           radius=t_rad, fill=(255, 255, 255, 252))

    tw = t_right - t_left
    th = t_bottom - t_top

    # Grid dimensions
    ncols = 4
    nrows = 9
    hdr_h = int(th * 0.08)
    cell_w = tw / ncols
    d_top = t_top + hdr_h
    cell_h = (t_bottom - d_top) / nrows

    # Header bar
    hdr_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    hd = ImageDraw.Draw(hdr_ov)
    hd.rounded_rectangle([t_left, t_top, t_right, t_top + hdr_h],
                         radius=t_rad, fill=(33, 115, 70, 40))
    hd.rectangle([t_left, t_top + t_rad, t_right, t_top + hdr_h],
                 fill=(33, 115, 70, 40))
    img = Image.alpha_composite(img, hdr_ov)
    draw = ImageDraw.Draw(img)

    # Header labels
    for ci, lbl in enumerate("ABCD"):
        cx = t_left + int((ci + 0.5) * cell_w)
        cy = t_top + hdr_h // 2
        bb = draw.textbbox((0, 0), lbl, font=font_hdr)
        tw2, th2 = bb[2] - bb[0], bb[3] - bb[1]
        draw.text((cx - tw2 // 2, cy - th2 // 2 - 1), lbl,
                  fill=(33, 115, 70, 180), font=font_hdr)

    # Grid lines
    gc = (180, 190, 185, 40)
    draw.line([(t_left + 4, d_top), (t_right - 4, d_top)],
              fill=(33, 115, 70, 55), width=2)
    for r in range(1, nrows + 1):
        y = d_top + int(r * cell_h)
        if y < t_bottom:
            draw.line([(t_left + 4, y), (t_right - 4, y)], fill=gc, width=1)
    for c in range(1, ncols):
        x = t_left + int(c * cell_w)
        draw.line([(x, t_top + 4), (x, t_bottom - 4)], fill=gc, width=1)

    # Alternating rows
    row_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    rd = ImageDraw.Draw(row_ov)
    for r in range(nrows):
        if r % 2 == 0:
            y1 = d_top + int(r * cell_h)
            y2 = d_top + int((r + 1) * cell_h)
            if y2 <= t_bottom:
                rd.rectangle([t_left + 1, y1 + 1, t_right - 1, y2 - 1],
                             fill=(33, 115, 70, 6))
    img = Image.alpha_composite(img, row_ov)
    draw = ImageDraw.Draw(img)

    # Cell numbers (subtle grey text in cells)
    import random
    random.seed(99)
    num_color = (100, 110, 105, 110)
    for r in range(nrows):
        for c in range(ncols):
            cx = t_left + int((c + 0.5) * cell_w)
            cy = d_top + int((r + 0.5) * cell_h)
            val = str(random.randint(10, 999))
            bb = draw.textbbox((0, 0), val, font=font_num)
            tw2, th2 = bb[2] - bb[0], bb[3] - bb[1]
            draw.text((cx - tw2 // 2, cy - th2 // 2), val,
                      fill=num_color, font=font_num)

    # Selected cell (B3) — bright green border
    sr, sc = 2, 1
    sx1 = t_left + int(sc * cell_w) + 1
    sy1 = d_top + int(sr * cell_h) + 1
    sx2 = t_left + int((sc + 1) * cell_w) - 1
    sy2 = d_top + int((sr + 1) * cell_h) - 1
    draw.rounded_rectangle([sx1, sy1, sx2, sy2], radius=4, fill=(255, 255, 255, 230))
    draw.rounded_rectangle([sx1 - 2, sy1 - 2, sx2 + 2, sy2 + 2],
                           radius=6, outline=(34, 197, 94, 255), width=4)
    sel_val = "42"
    bb = draw.textbbox((0, 0), sel_val, font=font_sel)
    tw2, th2 = bb[2] - bb[0], bb[3] - bb[1]
    draw.text(((sx1 + sx2) // 2 - tw2 // 2, (sy1 + sy2) // 2 - th2 // 2),
              sel_val, fill=(33, 115, 70, 255), font=font_sel)

    # ==========================================
    #  CHART CARDS — left side, overlapping table
    # ==========================================
    chart_ov = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    cd = ImageDraw.Draw(chart_ov)

    # --- BAR CHART (upper card) ---
    bc_l, bc_t = int(size * 0.05), int(size * 0.09)
    bc_r, bc_b = int(size * 0.38), int(size * 0.47)
    cr = 20

    # Shadow
    for i in range(14):
        a = int(30 * (1 - i / 14))
        cd.rounded_rectangle([bc_l + i - 1, bc_t + i + 4, bc_r + i - 1, bc_b + i + 4],
                             radius=cr + i, fill=(0, 0, 0, a))
    # Card bg
    cd.rounded_rectangle([bc_l, bc_t, bc_r, bc_b], radius=cr, fill=(255, 255, 255, 245))

    # Bars
    ba_l = bc_l + 18
    ba_r = bc_r - 14
    ba_t = bc_t + 18
    ba_b = bc_b - 14
    ba_w = ba_r - ba_l
    ba_h = ba_b - ba_t

    # Axis
    cd.line([(ba_l, ba_b), (ba_r, ba_b)], fill=(200, 200, 200, 80), width=2)

    vals = [0.42, 0.75, 0.55, 0.92, 0.68, 0.85]
    colors = [
        (59, 130, 246),   # blue
        (34, 197, 94),    # green
        (251, 191, 36),   # amber
        (239, 68, 68),    # red
        (139, 92, 246),   # purple
        (14, 165, 233),   # sky
    ]
    nb = len(vals)
    bsp = ba_w / nb
    bw = bsp * 0.65

    for i, (v, clr) in enumerate(zip(vals, colors)):
        bx = ba_l + i * bsp + (bsp - bw) / 2
        bh = v * (ba_h - 4)
        by = ba_b - bh
        cd.rounded_rectangle([bx, by, bx + bw, ba_b], radius=5,
                             fill=(*clr, 210))

    # Guide lines
    for f in [0.33, 0.66]:
        gy = ba_b - f * (ba_h - 4)
        cd.line([(ba_l, gy), (ba_r, gy)], fill=(200, 200, 200, 40), width=1)

    # --- LINE/AREA CHART (lower card) ---
    lc_l, lc_t = int(size * 0.04), int(size * 0.52)
    lc_r, lc_b = int(size * 0.40), int(size * 0.90)

    # Shadow
    for i in range(14):
        a = int(30 * (1 - i / 14))
        cd.rounded_rectangle([lc_l + i - 1, lc_t + i + 4, lc_r + i - 1, lc_b + i + 4],
                             radius=cr + i, fill=(0, 0, 0, a))
    # Card bg
    cd.rounded_rectangle([lc_l, lc_t, lc_r, lc_b], radius=cr, fill=(255, 255, 255, 245))

    la_l = lc_l + 18
    la_r = lc_r - 14
    la_t = lc_t + 18
    la_b = lc_b - 14
    la_w = la_r - la_l
    la_h = la_b - la_t

    # Axis
    cd.line([(la_l, la_b), (la_r, la_b)], fill=(200, 200, 200, 80), width=2)

    # Guide lines
    for f in [0.33, 0.66]:
        gy = la_b - f * (la_h - 4)
        cd.line([(la_l, gy), (la_r, gy)], fill=(200, 200, 200, 40), width=1)

    # Two line series
    series = [
        ([0.25, 0.40, 0.35, 0.60, 0.52, 0.75, 0.88], (59, 130, 246), 45),
        ([0.45, 0.38, 0.52, 0.44, 0.62, 0.58, 0.70], (34, 197, 94), 35),
    ]
    for lv, lc_color, aa in series:
        npts = len(lv)
        pts = []
        for i, v in enumerate(lv):
            px = la_l + 6 + i * (la_w - 12) / (npts - 1)
            py = la_b - 4 - v * (la_h - 10)
            pts.append((int(px), int(py)))

        # Area fill
        poly = [(pts[0][0], la_b)] + pts + [(pts[-1][0], la_b)]
        cd.polygon(poly, fill=(*lc_color, aa))

        # Line
        for i in range(len(pts) - 1):
            cd.line([pts[i], pts[i + 1]], fill=(*lc_color, 230), width=3)

        # Dots
        for px, py in pts:
            cd.ellipse([px - 5, py - 5, px + 5, py + 5], fill=(255, 255, 255, 245))
            cd.ellipse([px - 3, py - 3, px + 3, py + 3], fill=(*lc_color, 240))

    img = Image.alpha_composite(img, chart_ov)

    # Apply rounded mask
    mask = rounded_rectangle_mask(size, CORNER_RADIUS)
    img.putalpha(mask)
    return img


def create_icns(png_path, icns_path):
    iconset_dir = png_path.replace('.png', '.iconset')
    os.makedirs(iconset_dir, exist_ok=True)
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    src = Image.open(png_path)
    for s in sizes:
        resized = src.resize((s, s), Image.LANCZOS)
        resized.save(os.path.join(iconset_dir, f'icon_{s}x{s}.png'))
        if s <= 512:
            double = src.resize((s * 2, s * 2), Image.LANCZOS)
            double.save(os.path.join(iconset_dir, f'icon_{s}x{s}@2x.png'))
    subprocess.run(['iconutil', '-c', 'icns', iconset_dir, '-o', icns_path], check=True)
    import shutil
    shutil.rmtree(iconset_dir)


if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon = draw_icon()
    png_path = os.path.join(script_dir, 'icon.png')
    icns_path = os.path.join(script_dir, 'AppIcon.icns')
    icon.save(png_path, 'PNG')
    print(f"Saved {png_path}")
    create_icns(png_path, icns_path)
    print(f"Saved {icns_path}")

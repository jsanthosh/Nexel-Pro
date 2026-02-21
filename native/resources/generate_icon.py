#!/usr/bin/env python3
"""Generate Nexel app icon — professional spreadsheet icon design."""

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
    draw = ImageDraw.Draw(img)

    # Background: deep green gradient (top-left bright → bottom-right dark)
    for y in range(size):
        t = y / size
        r = int(18 * (1 - t * 0.4))
        g = int(120 - 50 * t)
        b = int(65 - 25 * t)
        draw.line([(0, y), (size - 1, y)], fill=(r, g, b, 255))

    # Subtle lighter overlay in top-center for depth
    overlay = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    ov_draw = ImageDraw.Draw(overlay)
    cx, cy = size // 2, int(size * 0.3)
    max_r = int(size * 0.55)
    for radius in range(max_r, 0, -3):
        alpha = int(18 * (1 - radius / max_r))
        ov_draw.ellipse(
            [cx - radius, cy - radius, cx + radius, cy + radius],
            fill=(100, 200, 130, alpha)
        )
    img = Image.alpha_composite(img, overlay)
    draw = ImageDraw.Draw(img)

    # === Grid parameters ===
    margin = int(size * 0.13)
    grid_left = margin
    grid_top = margin
    grid_right = size - margin
    grid_bottom = size - int(size * 0.25)

    num_cols = 4
    num_rows = 4
    cell_w = (grid_right - grid_left) / num_cols
    cell_h = (grid_bottom - grid_top) / num_rows
    gap = 4
    radius = 10

    # === Draw cells ===
    for row in range(num_rows):
        for col in range(num_cols):
            x1 = grid_left + col * cell_w + gap
            y1 = grid_top + row * cell_h + gap
            x2 = grid_left + (col + 1) * cell_w - gap
            y2 = grid_top + (row + 1) * cell_h - gap

            if row == 0:
                # Header: solid vibrant green
                fill = (30, 155, 85, 240)
            else:
                # Data cells: frosted glass with alternating tint
                if row % 2 == 1:
                    fill = (255, 255, 255, 95)
                else:
                    fill = (255, 255, 255, 70)

            draw.rounded_rectangle([x1, y1, x2, y2], radius=radius, fill=fill)

    # === Selected cell highlight (B2) ===
    sel_r, sel_c = 1, 1
    sx1 = grid_left + sel_c * cell_w + gap
    sy1 = grid_top + sel_r * cell_h + gap
    sx2 = grid_left + (sel_c + 1) * cell_w - gap
    sy2 = grid_top + (sel_r + 1) * cell_h - gap
    # Bright cell fill
    draw.rounded_rectangle([sx1, sy1, sx2, sy2], radius=radius, fill=(255, 255, 255, 180))
    # Green selection border
    draw.rounded_rectangle(
        [sx1 - 2, sy1 - 2, sx2 + 2, sy2 + 2],
        radius=radius + 2, outline=(16, 200, 90, 255), width=5
    )

    # === Header text (A B C D) ===
    try:
        hdr_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.052))
    except Exception:
        hdr_font = ImageFont.load_default()

    for col, lbl in enumerate("ABCD"):
        cx = grid_left + col * cell_w + cell_w / 2
        cy = grid_top + cell_h / 2
        bbox = draw.textbbox((0, 0), lbl, font=hdr_font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - tw / 2, cy - th / 2 - 2), lbl,
                  fill=(255, 255, 255, 255), font=hdr_font)

    # === Cell data: small colored horizontal bars (abstract data) ===
    # These look good at any icon size — better than tiny text
    bar_data = [
        # (row, col, width_fraction, color)
        (1, 0, 0.65, (255, 255, 255, 160)),
        (1, 2, 0.50, (255, 255, 255, 140)),
        (1, 3, 0.75, (255, 255, 255, 160)),
        (2, 0, 0.80, (255, 255, 255, 150)),
        (2, 1, 0.45, (255, 255, 255, 130)),
        (2, 2, 0.60, (255, 255, 255, 150)),
        (2, 3, 0.40, (255, 255, 255, 120)),
        (3, 0, 0.55, (255, 255, 255, 140)),
        (3, 1, 0.70, (255, 255, 255, 160)),
        (3, 2, 0.85, (255, 255, 255, 170)),
        (3, 3, 0.50, (255, 255, 255, 130)),
    ]
    bar_h = int(size * 0.016)
    for row, col, frac, color in bar_data:
        cx = grid_left + col * cell_w + cell_w / 2
        cy = grid_top + row * cell_h + cell_h / 2
        bw = (cell_w - gap * 4) * frac
        draw.rounded_rectangle(
            [cx - bw / 2, cy - bar_h / 2, cx + bw / 2, cy + bar_h / 2],
            radius=bar_h // 2, fill=color
        )

    # Selected cell: show "42" as a value (visible even at small sizes)
    try:
        val_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.055))
    except Exception:
        val_font = hdr_font
    val_text = "42"
    bbox = draw.textbbox((0, 0), val_text, font=val_font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    vcx = grid_left + sel_c * cell_w + cell_w / 2
    vcy = grid_top + sel_r * cell_h + cell_h / 2
    draw.text((vcx - tw / 2, vcy - th / 2 - 1), val_text,
              fill=(16, 110, 60, 255), font=val_font)

    # === Mini bar chart in bottom area ===
    chart_left = grid_left + int(size * 0.04)
    chart_right = grid_right - int(size * 0.04)
    chart_bottom = grid_bottom + int(size * 0.13)
    chart_top = grid_bottom + int(size * 0.03)
    chart_h = chart_bottom - chart_top

    bar_values = [0.4, 0.7, 0.5, 0.9, 0.6, 0.8, 0.45, 0.95]
    n = len(bar_values)
    total_w = chart_right - chart_left
    bw = total_w / n * 0.65
    spacing = total_w / n

    for i, v in enumerate(bar_values):
        bx = chart_left + i * spacing + (spacing - bw) / 2
        bh = v * chart_h
        by = chart_bottom - bh
        # Gradient-like effect: lighter bars for higher values
        alpha = int(120 + 100 * v)
        draw.rounded_rectangle(
            [bx, by, bx + bw, chart_bottom],
            radius=4,
            fill=(120, 230, 160, min(alpha, 255))
        )

    # === "Nexel" text ===
    try:
        title_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", int(size * 0.095))
    except Exception:
        title_font = ImageFont.load_default()

    text = "Nexel"
    bbox = draw.textbbox((0, 0), text, font=title_font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (size - tw) / 2
    ty = chart_bottom + int(size * 0.025)

    # Shadow
    draw.text((tx + 2, ty + 2), text, fill=(0, 30, 15, 70), font=title_font)
    # Text
    draw.text((tx, ty), text, fill=(255, 255, 255, 255), font=title_font)

    # Apply mask
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

#!/usr/bin/env python3
# CRUDマトリクス・アイコン生成
from PIL import Image, ImageDraw, ImageFont
import os

os.makedirs('icons', exist_ok=True)
INK = (32, 48, 79)        # indigo
CELLS = [
    ('C', (46, 139, 87)),   # green  (top-left)
    ('R', (45, 109, 181)),  # blue   (top-right)
    ('U', (181, 132, 12)),  # amber  (bottom-left)
    ('D', (192, 57, 43)),   # red    (bottom-right)
]

def font(sz):
    for p in ['/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
              '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf']:
        if os.path.exists(p):
            return ImageFont.truetype(p, sz)
    return ImageFont.load_default()

def rounded(draw, box, r, fill):
    draw.rounded_rectangle(box, radius=r, fill=fill)

def make(size, maskable=False, out='icon.png'):
    img = Image.new('RGB', (size, size), INK)
    d = ImageDraw.Draw(img)
    # 方眼グリッド（うっすら）
    step = max(8, size // 16)
    grid = (44, 62, 96)
    for x in range(0, size, step):
        d.line([(x, 0), (x, size)], fill=grid, width=1)
    for y in range(0, size, step):
        d.line([(0, y), (size, y)], fill=grid, width=1)

    # コンテンツ領域（maskableは内側80%に収める）
    pad = int(size * 0.16) if maskable else int(size * 0.13)
    area = size - pad * 2
    gap = max(3, int(area * 0.04))
    cell = (area - gap) // 2
    r = max(4, int(cell * 0.16))
    positions = [(0, 0), (1, 0), (0, 1), (1, 1)]
    fnt = font(int(cell * 0.62))
    for (label, color), (cx, cy) in zip(CELLS, positions):
        x0 = pad + cx * (cell + gap)
        y0 = pad + cy * (cell + gap)
        rounded(d, [x0, y0, x0 + cell, y0 + cell], r, color)
        # 文字（白、中央）
        bbox = d.textbbox((0, 0), label, font=fnt)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        d.text((x0 + cell / 2 - tw / 2 - bbox[0], y0 + cell / 2 - th / 2 - bbox[1]),
               label, font=fnt, fill=(255, 255, 255))
    img.save(out)
    print('wrote', out, size)

make(192, False, 'icons/icon-192.png')
make(512, False, 'icons/icon-512.png')
make(512, True,  'icons/icon-maskable-512.png')

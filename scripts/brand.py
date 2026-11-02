"""Toolkit brand pictures, made once from the Help ones (same GroupChatOn bubble, blue + white toolbox instead of teal + ?).
   python -X utf8 scripts/brand.py      (needs Pillow; writes static/brand/*)
   The teal bubble becomes blue, the ? hole is filled, and a toolbox is drawn in its place."""
import os, re
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(HERE, '..', 'groupchaton-tools', 'assets', 'brand')
OUT = os.path.join(HERE, 'static', 'brand')
TEAL, BLUE = (13, 148, 136), (37, 99, 235)

def near(c, ref, d=70): return sum((a - b) ** 2 for a, b in zip(c, ref)) < d * d

def toolbox(draw, cx, cy, s, width):
    """The toolbox sign in 100-unit space centred on (50, 52), scaled by s (pixels per unit)."""
    P = lambda x, y: (cx + (x - 50) * s, cy + (y - 52) * s)
    w = max(1, round(width * s))
    draw.rounded_rectangle([P(16, 38), P(84, 82)], radius=7 * s, outline='white', width=w)
    draw.line([P(38, 38), P(38, 26), P(62, 26), P(62, 38)], fill='white', width=w, joint='curve')
    draw.line([P(16, 57), P(84, 57)], fill='white', width=w)
    draw.rounded_rectangle([P(43, 51), P(57, 64)], radius=3 * s, fill='white')

def convert(name, out):
    im = Image.open(os.path.join(SRC, name)).convert('RGBA')
    W, H = im.size
    px = im.load()
    teal = [(x, y) for y in range(H) for x in range(W) if px[x, y][3] > 200 and near(px[x, y][:3], TEAL, 40)]
    x0, x1 = min(p[0] for p in teal), max(p[0] for p in teal)
    y0, y1 = min(p[1] for p in teal), max(p[1] for p in teal)
    is_teal = lambda c: c[3] > 40 and near(c[:3], TEAL, 90)
    # holes inside the bubble = non-teal pixels the bubble's edge can't reach (the ? and its dot)
    seen = set()
    stack = [(x, y) for x in range(x0, x1 + 1) for y in (y0, y1)] + [(x, y) for y in range(y0, y1 + 1) for x in (x0, x1)]
    while stack:
        x, y = stack.pop()
        if (x, y) in seen or not (x0 <= x <= x1 and y0 <= y <= y1) or is_teal(px[x, y]): continue
        seen.add((x, y)); stack += [(x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)]
    holes = [(x, y) for y in range(y0, y1 + 1) for x in range(x0, x1 + 1) if (x, y) not in seen and not is_teal(px[x, y])]
    for x, y in holes: px[x, y] = BLUE + (255,)
    # teal (and its blends with black, white edges or transparency) → blue
    for y in range(y0 - 2, min(H, y1 + 3)):
        for x in range(max(0, x0 - 2), min(W, x1 + 3)):
            r, g, b, a = px[x, y]
            if a == 0: continue
            if near((r, g, b), TEAL, 90):
                px[x, y] = BLUE + (a,)
            elif g > r + 25 and b > r + 15 and abs(g - b) < 40:   # teal mixed with black/white at an edge
                k = (g + b) / (TEAL[1] + TEAL[2])
                px[x, y] = tuple(min(255, round(c * k)) for c in BLUE) + (a,)
    hx0, hx1 = min(p[0] for p in holes), max(p[0] for p in holes)
    hy0, hy1 = min(p[1] for p in holes), max(p[1] for p in holes)
    s = (hy1 - hy0 + 1) / 64          # the ? is about 64 units tall
    # draw 4× larger, then shrink, for smooth edges
    big = Image.new('RGBA', (W * 4, H * 4), (0, 0, 0, 0))
    toolbox(ImageDraw.Draw(big), (hx0 + hx1) / 2 * 4, (hy0 + hy1) / 2 * 4 + 2 * s * 4, s * 4 * 1.05, 9)
    im = Image.alpha_composite(im, big.resize((W, H), Image.LANCZOS))
    im.save(os.path.join(OUT, out))
    return im

os.makedirs(OUT, exist_ok=True)
convert('help-logo-light.png', 'logo-light.png')
convert('help-logo-dark.png', 'logo-dark.png')
ic = convert('help-favicon-512.png', 'favicon-512.png')
for n in (192, 48, 32): ic.resize((n, n), Image.LANCZOS).save(os.path.join(OUT, 'favicon-%d.png' % n))
ic.resize((180, 180), Image.LANCZOS).save(os.path.join(OUT, 'apple-touch-icon.png'))
ic.save(os.path.join(OUT, 'favicon.ico'), sizes=[(16, 16), (32, 32), (48, 48)])
# maskable app icon: the mark on a white tile with a safe margin
m = Image.new('RGBA', (512, 512), (255, 255, 255, 255))
m.alpha_composite(ic.resize((360, 360), Image.LANCZOS), (76, 76))
m.save(os.path.join(OUT, 'maskable-512.png'))
# vector mark for headers/footers
svg = open(os.path.join(SRC, 'help-mark.svg'), encoding='utf-8').read().replace('#0D9488', '#2563EB').replace('GroupChatOn Help', 'Toolkit by GroupChatOn')
svg = re.sub(r'<g transform="translate\(390 270\) scale\(5.0\)">.*</g></svg>',
    '<g transform="translate(390 270) scale(5.0)"><g fill="none" stroke="#fff" stroke-width="9" stroke-linejoin="round">'
    '<rect x="16" y="38" width="68" height="44" rx="7"/><path d="M38 38V26h24v12M16 57h68"/></g>'
    '<rect x="43" y="51" width="14" height="13" rx="3" fill="#fff"/></g></svg>', svg, flags=re.S)
open(os.path.join(OUT, 'mark.svg'), 'w', encoding='utf-8').write(svg)
print('brand pictures written to', OUT)

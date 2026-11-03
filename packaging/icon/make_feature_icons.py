"""Draw feature icons for Resources/Icons (16-256 px), in the flat style of the others there.

    python packaging/icon/make_feature_icons.py [name...]      (needs Pillow; no name = all of them)

Icons:
    Sync    two white arrows going round on a blue disc (Folder Sync in the search results)

Like make_icon.py, everything is drawn from simple shapes (no fonts, no third-party artwork), and
small sizes get thicker strokes instead of being scaled down from the large one. A feature icon of
our own looks the same on every Windows version: the icon numbers inside imageres.dll are not stable
(the "sync" icon there turned into a printer on another Windows build).
"""
import math
import os
import sys

from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_icon import SIZES, SS, write_ico   # noqa: E402  (same folder)


def disc(n, color):
    img = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse([0, 0, n - 1, n - 1], fill=color)
    return img


def arrow_arc(draw, c, r, start, end, width, head, fill):
    """Arc from start to end degrees (clockwise, 0 = 3 o'clock) with an arrowhead at the end."""
    draw.arc([c - r, c - r, c + r, c + r], start, end, fill=fill, width=int(width))
    t = math.radians(end)
    px, py = c + r * math.cos(t), c + r * math.sin(t)          # the arc's end
    tx, ty = -math.sin(t), math.cos(t)                         # going on clockwise
    nx, ny = math.cos(t), math.sin(t)                          # outwards
    tip = (px + tx * head, py + ty * head)
    draw.polygon([tip, (px + nx * head * 0.8, py + ny * head * 0.8), (px - nx * head * 0.8, py - ny * head * 0.8)], fill=fill)


def design_sync(n, weight):
    img = disc(n, (2, 132, 199, 255))
    d = ImageDraw.Draw(img)
    c, r = n / 2, n * 0.27
    width, head = n * 0.085 * weight, n * 0.15 * min(weight, 1.25)
    arrow_arc(d, c, r, 195, 320, width, head, 'white')         # over the top, to the right
    arrow_arc(d, c, r, 15, 140, width, head, 'white')          # under the bottom, to the left
    return img


DESIGNS = {'Sync': design_sync}


def render(design, size):
    weight = 1.45 if size <= 16 else 1.3 if size <= 24 else 1.15 if size <= 40 else 1.0
    big = DESIGNS[design](size * SS, weight)
    return big.resize((size, size), Image.LANCZOS)


def main():
    names = sys.argv[1:] or list(DESIGNS)
    root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
    for name in names:
        path = os.path.join(root, 'Resources', 'Icons', name + '.ico')
        write_ico([render(name, s) for s in SIZES], path)
        print('%s (%s px)' % (path, ', '.join(map(str, SIZES))))


if __name__ == '__main__':
    main()

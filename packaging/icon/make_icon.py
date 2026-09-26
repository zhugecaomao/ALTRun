"""Draw the ALTRun icon: Resources/ALTRun.ico (16-256 px) and docs/images/logo.png (256 px).

    python packaging/icon/make_icon.py [design]      (needs Pillow; design: prompt / lens / a)

The program uses "prompt"; the other two were the alternatives when choosing.

Everything is drawn from simple shapes (no fonts, no third-party artwork), so the icon
is ALTRun's own and can be used without licence questions. Small sizes are drawn with
thicker strokes instead of being scaled down from the large one, so they stay sharp.
"""
import io
import os
import struct
import sys

from PIL import Image, ImageDraw

SIZES = [16, 20, 24, 32, 40, 48, 64, 256]
SS = 8                                     # supersampling


def lerp(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(len(a)))


def background(n, top, bottom, radius):
    """Rounded square with a diagonal gradient, n x n pixels."""
    grad = Image.new('RGBA', (n, n))
    px = grad.load()
    for y in range(n):
        for x in range(n):
            px[x, y] = lerp(top, bottom, (x + y) / (2 * (n - 1))) + (255,)
    mask = Image.new('L', (n, n), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, n - 1, n - 1], radius=radius, fill=255)
    out = Image.new('RGBA', (n, n), (0, 0, 0, 0))
    out.paste(grad, (0, 0), mask)
    return out


def thick_line(draw, points, width, fill):
    """Polyline with round joins and caps."""
    draw.line(points, fill=fill, width=int(width), joint='curve')
    r = width / 2
    for x, y in (points[0], points[-1]):
        draw.ellipse([x - r, y - r, x + r, y + r], fill=fill)


def design_prompt(n, weight):
    """Chevron + caret, like a command prompt."""
    img = background(n, (47, 128, 255), (88, 64, 230), n * 0.22)
    d = ImageDraw.Draw(img)
    w = n * 0.105 * weight
    thick_line(d, [(n * 0.26, n * 0.30), (n * 0.47, n * 0.50), (n * 0.26, n * 0.70)], w, 'white')
    thick_line(d, [(n * 0.55, n * 0.70), (n * 0.76, n * 0.70)], w, 'white')
    return img


def design_lens(n, weight):
    """Magnifier with a chevron in the lens."""
    img = background(n, (47, 128, 255), (88, 64, 230), n * 0.22)
    d = ImageDraw.Draw(img)
    w = n * 0.09 * weight
    cx, cy, r = n * 0.44, n * 0.43, n * 0.22
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline='white', width=int(w))
    thick_line(d, [(cx + r * 0.72, cy + r * 0.72), (n * 0.78, n * 0.78)], w * 1.25, 'white')
    s = r * 0.42
    thick_line(d, [(cx - s * 0.5, cy - s), (cx + s * 0.55, cy), (cx - s * 0.5, cy + s)], w * 0.8, 'white')
    return img


def design_a(n, weight):
    """An "A" made of a chevron, its crossbar running on as a caret."""
    img = background(n, (20, 184, 166), (37, 99, 235), n * 0.22)
    d = ImageDraw.Draw(img)
    w = n * 0.105 * weight
    thick_line(d, [(n * 0.24, n * 0.76), (n * 0.46, n * 0.24), (n * 0.68, n * 0.76)], w, 'white')
    thick_line(d, [(n * 0.35, n * 0.58), (n * 0.80, n * 0.58)], w * 0.9, 'white')
    return img


DESIGNS = {'prompt': design_prompt, 'lens': design_lens, 'a': design_a}


def render(design, size):
    weight = 1.45 if size <= 16 else 1.3 if size <= 24 else 1.15 if size <= 40 else 1.0
    big = DESIGNS[design](size * SS, weight)
    return big.resize((size, size), Image.LANCZOS)


def dib(img):
    """32-bit BGRA DIB (BITMAPINFOHEADER + pixels + AND mask) for an .ico entry."""
    w, h = img.size
    header = struct.pack('<IiiHHIIiiII', 40, w, h * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    rows = []
    for y in range(h - 1, -1, -1):
        rows.append(b''.join(struct.pack('BBBB', b, g, r, a)
                             for r, g, b, a in (img.getpixel((x, y)) for x in range(w))))
    mask_row = ((w + 31) // 32) * 4
    return header + b''.join(rows) + b'\0' * (mask_row * h)


def write_ico(images, path):
    entries, blobs = [], []
    for img in images:
        size = img.size[0]
        if size >= 256:
            buf = io.BytesIO()
            img.save(buf, 'PNG')
            data = buf.getvalue()
        else:
            data = dib(img)
        entries.append((size, data))
    offset = 6 + 16 * len(entries)
    out = struct.pack('<HHH', 0, 1, len(entries))
    for size, data in entries:
        dim = 0 if size >= 256 else size
        out += struct.pack('<BBBBHHII', dim, dim, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
        blobs.append(data)
    with open(path, 'wb') as f:
        f.write(out + b''.join(blobs))


def main():
    design = sys.argv[1] if len(sys.argv) > 1 else 'prompt'
    root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
    ico = os.path.join(root, 'Resources', 'ALTRun.ico')
    logo = os.path.join(root, 'docs', 'images', 'logo.png')
    write_ico([render(design, s) for s in SIZES], ico)
    render(design, 256).save(logo, optimize=True)
    print('%s (%s px)\n%s' % (ico, ', '.join(map(str, SIZES)), logo))


if __name__ == '__main__':
    main()

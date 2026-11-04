"""Draw the Microsoft Store package's images: packaging/msix/Assets/*.png, from the same shapes as the program icon.

    python packaging/icon/make_store_assets.py      (needs Pillow)

Square44x44Logo   the app list, taskbar and Start; the targetsize-N versions are used as they are at each size
                  (altform-unplated: on the taskbar and in Start without a plate behind them)
Square150x150Logo the Start tile; the icon takes about two thirds of it, as Windows' own tiles do
StoreLogo         the package's logo in the Store and in installers
"""
import os
import sys

from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_icon import render   # noqa: E402  (same folder)

TARGET_SIZES = [16, 20, 24, 30, 32, 36, 40, 48, 60, 64, 72, 80, 96, 256]


def padded(size, fraction):
    """The icon at fraction of size, centered on a transparent square."""
    inner = max(16, round(size * fraction))
    out = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    out.paste(render('prompt', inner), ((size - inner) // 2, (size - inner) // 2))
    return out


def main():
    root = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..'))
    out = os.path.join(root, 'packaging', 'msix', 'Assets')
    os.makedirs(out, exist_ok=True)
    images = {'Square44x44Logo.png': render('prompt', 44), 'Square150x150Logo.png': padded(150, 0.66), 'StoreLogo.png': render('prompt', 50)}
    for size in TARGET_SIZES:
        images['Square44x44Logo.targetsize-%d.png' % size] = render('prompt', size)
        images['Square44x44Logo.targetsize-%d_altform-unplated.png' % size] = render('prompt', size)
    for name, img in sorted(images.items()):
        img.save(os.path.join(out, name), optimize=True)
    print('%s: %d images' % (out, len(images)))


if __name__ == '__main__':
    main()

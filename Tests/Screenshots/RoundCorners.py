"""RoundCorners.py - Windows 11-style rounded corners for the screenshots

The Windows VM of GitHub Actions has no GPU acceleration, so the system draws no rounded window corners and the shots come out square.
This makes the four corners of each screenshot transparent (radius 8 pixels, as Windows 11 at 100% scaling, with smooth edges),
so they look right on light and dark pages. For shots at 150% scaling (environment variable SHOT_SCALE=1.5) the radius scales too.
Images that already have rounded corners (transparent top left) and the animations demo.png / demo-*.png are left alone.

    python Tests/Screenshots/RoundCorners.py [screenshot folder]      default docs/images/screenshots
"""
import os
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw

RADIUS = 8
SCALE = 4                                    # draw the rounded rectangle at 4x and scale it down: smooth edges


def rounded_mask(size, radius=RADIUS):
    """A rounded rectangle mask of this size (mode L: 255 inside, 0 in the corners)"""
    width, height = size
    big = Image.new("L", (width * SCALE, height * SCALE), 0)
    ImageDraw.Draw(big).rounded_rectangle((0, 0, width * SCALE - 1, height * SCALE - 1), radius * SCALE, fill=255)
    return big.resize(size, Image.LANCZOS)


def round_image(image, radius=RADIUS):
    """The image with rounded corners, as RGBA"""
    image = image.convert("RGBA")
    alpha = ImageChops.multiply(image.getchannel("A"), rounded_mask(image.size, radius))
    image.putalpha(alpha)
    return image


def main(folder):
    count = 0
    for path in sorted(Path(folder).glob("*.png")):
        with Image.open(path) as image:
            image.load()
        if path.stem == "demo" or path.stem.startswith("demo-") or (image.mode == "RGBA" and image.getpixel((0, 0))[3] == 0):
            continue
        round_image(image, round(RADIUS * float(os.environ.get("SHOT_SCALE") or 1))).save(path, optimize=True)
        count += 1
    print(f"rounded {count} screenshot(s) in {folder}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "docs/images/screenshots")

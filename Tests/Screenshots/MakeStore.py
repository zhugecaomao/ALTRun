"""MakeStore.py - the Microsoft Store screenshots: packaging/msix/Screenshots/*.png

Partner Center wants desktop screenshots of at least 1366 x 768, but the README's screenshots are window-sized. This puts each
chosen screenshot (rounded corners already, see RoundCorners.py) on a 1920 x 1080 background, the same blurred picture as the
animations (MakeDemo.background), centered with a shadow. Shots taken at 100% (Preferences, Quick Switch) are scaled up to
about the size of the 150% ones. English interface and made-up data, as all of ALTRun's screenshots.
The order is the order in the Store; the captions are typed in Partner Center (packaging/msix/listing.md).

    python Tests/Screenshots/MakeStore.py [screenshot folder] [output folder]
"""
import sys
from pathlib import Path

from PIL import Image, ImageFilter

from MakeDemo import background

SIZE = (1920, 1080)
SHOTS = ["search", "actions", "calculator", "clipboard", "files", "quickswitch", "theme-dark", "prefs-appearance"]
MAX_W, MAX_H = 1600, 920                 # the window's largest size on the background
MAX_SCALE = 1.5                          # a 100% shot at most 1.5 times (like the 150% ones), never larger than it fits
SHADOW_BLUR, SHADOW_OFFSET, SHADOW_OPACITY = 28, 14, 120
ROOT = Path(__file__).resolve().parents[2]


def compose(shot, canvas):
    """The shot (RGBA) centered on the canvas (RGB), with a soft shadow below it"""
    scale = min(MAX_SCALE if shot.width < 900 else 1, MAX_W / shot.width, MAX_H / shot.height)
    if scale != 1:
        shot = shot.resize((round(shot.width * scale), round(shot.height * scale)), Image.LANCZOS)
    x, y = (SIZE[0] - shot.width) // 2, (SIZE[1] - shot.height) // 2
    alpha = shot.getchannel("A")
    shadow = Image.new("RGBA", SIZE, (0, 0, 0, 0))
    shadow.paste((0, 0, 0, 255), (x, y + SHADOW_OFFSET), alpha.point(lambda v: v * SHADOW_OPACITY // 255))
    out = Image.alpha_composite(canvas.convert("RGBA"), shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR)))
    out.alpha_composite(shot, (x, y))
    return out.convert("RGB")


def main():
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "docs" / "images" / "screenshots"
    target = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "packaging" / "msix" / "Screenshots"
    target.mkdir(parents=True, exist_ok=True)
    canvas = background(SIZE)
    for index, name in enumerate(SHOTS, 1):
        shot = Image.open(source / (name + ".png")).convert("RGBA")
        compose(shot, canvas).save(target / ("%02d-%s.png" % (index, name)), optimize=True)
    print("%s: %d screenshots, %d x %d" % (target, len(SHOTS), *SIZE))


if __name__ == "__main__":
    main()

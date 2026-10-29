"""MakeDemo.py - animations (APNG): demo.png on the README front page and the feature animations demo-*.png

Turns the frames a "demo" scene of TakeScreenshots.ahk saved (frames.txt: each line "file name Tab ms") into a looping animated PNG:
in each frame it finds the search window (the area that differs from the plain backdrop of the shots), adds rounded corners and a shadow,
and puts it on a blurred background. The window's top left stays put, so the animation keeps its size when the window gets taller or shorter.
With a "windows" file in the frame folder (several windows: Notepad and the search window, a dialog and its folder panel) it keeps
the area any window covers in any frame instead, and the backdrop between the windows becomes the blurred background, with their shadow.

Background: Tests/Screenshots/backdrop.jpg (or .png) when there is one (scaled, cropped and blurred), otherwise an abstract blue-purple picture
(a few colored light spots + Gaussian blur, like the Windows 11 default wallpaper).
APNG rather than GIF: GIF has only 256 colors, so the background gradient would be dithered (grainy) and the rounded corners couldn't be semi-transparent (jaggies).
Browsers and GitHub play APNG directly; Pillow stores only the part of each frame that changed, so the files stay small.
When the shots are taken at 150% scaling (environment variable SHOT_SCALE=1.5), the margins, shadow and corners scale up too.

    python Tests/Screenshots/MakeDemo.py <frame folder> <output png>
"""
import os
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

from RoundCorners import RADIUS, round_image, rounded_mask

SCALE = float(os.environ.get("SHOT_SCALE") or 1)
BACKDROP = (0x8A, 0x9B, 0xB0)        # background color behind the window in TakeScreenshots.ahk shots
MARGIN = round(56 * SCALE)           # background visible around the window
SHADOW_BLUR, SHADOW_OFFSET, SHADOW_OPACITY = round(18 * SCALE), round(10 * SCALE), 110
CORNER = round(RADIUS * SCALE)
HERE = Path(__file__).resolve().parent


def window_box(image):
    """The area of the frame that differs from the backdrop color (the search window)"""
    diff = ImageChops.difference(image, Image.new("RGB", image.size, BACKDROP)).convert("L").point(lambda v: 255 if v > 8 else 0)
    return diff.getbbox()


def background(size):
    """The blurred background picture"""
    for name in ("backdrop.jpg", "backdrop.png"):
        path = HERE / name
        if path.exists():
            image = Image.open(path).convert("RGB")
            scale = max(size[0] / image.width, size[1] / image.height)
            image = image.resize((round(image.width * scale) + 1, round(image.height * scale) + 1), Image.LANCZOS)
            left, top = (image.width - size[0]) // 2, (image.height - size[1]) // 2
            return image.crop((left, top, left + size[0], top + size[1])).filter(ImageFilter.GaussianBlur(24))
    width, height = size
    image = Image.new("RGB", size)
    draw = ImageDraw.Draw(image)
    top, bottom = (24, 38, 92), (58, 32, 104)                          # dark blue -> dark purple
    for y in range(height):
        t = y / max(1, height - 1)
        draw.line([(0, y), (width, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(top, bottom)))
    glow = Image.new("RGB", size, (0, 0, 0))
    spots = ImageDraw.Draw(glow)
    for cx, cy, r, color in [(0.18, 0.25, 0.42, (40, 120, 255)), (0.85, 0.15, 0.38, (150, 90, 255)),
                             (0.70, 0.85, 0.45, (255, 90, 170)), (0.10, 0.95, 0.35, (30, 200, 230)),
                             (0.50, 0.50, 0.30, (90, 110, 255))]:
        x, y, rr = cx * width, cy * height, r * max(width, height)
        spots.ellipse((x - rr, y - rr, x + rr, y + rr), fill=color)
    glow = glow.filter(ImageFilter.GaussianBlur(max(width, height) // 6)).point(lambda v: v * 3 // 5)   # softer light spots
    return ImageChops.screen(image, glow)                               # light spots on top of the gradient


def with_shadow(canvas_bg, window_size):
    """The background with the window's shadow (the window itself isn't on it yet)"""
    shadow = Image.new("RGBA", canvas_bg.size, (0, 0, 0, 0))
    mask = rounded_mask(window_size, CORNER).point(lambda v: v * SHADOW_OPACITY // 255)
    shadow.paste((0, 0, 0, 255), (MARGIN, MARGIN + SHADOW_OFFSET), mask)
    shadow = shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))
    return Image.alpha_composite(canvas_bg.convert("RGBA"), shadow).convert("RGB")


def several_windows(images):
    """The frames cropped to the area any window covers; the backdrop color around and between them shows the background"""
    boxes = [window_box(image) or (0, 0, image.width, image.height) for image in images]
    box = (min(b[0] for b in boxes), min(b[1] for b in boxes), max(b[2] for b in boxes), max(b[3] for b in boxes))
    size = (box[2] - box[0] + 2 * MARGIN, box[3] - box[1] + 2 * MARGIN)
    canvas_bg = background(size)
    frames = []
    for image in images:
        crop = image.crop(box)
        mask = ImageChops.difference(crop, Image.new("RGB", crop.size, BACKDROP)).convert("L").point(lambda v: 255 if v > 8 else 0)
        shadow = Image.new("RGBA", size, (0, 0, 0, 0))
        shadow.paste((0, 0, 0, 255), (MARGIN, MARGIN + SHADOW_OFFSET), mask.point(lambda v: v * SHADOW_OPACITY // 255))
        frame = Image.alpha_composite(canvas_bg.convert("RGBA"), shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))).convert("RGB")
        frame.paste(crop, (MARGIN, MARGIN), mask)
        frames.append(frame)
    return frames


def main(frame_dir, output):
    frame_dir = Path(frame_dir)
    entries = []                                                        # (image, ms, keys: no longer used)
    for line in (frame_dir / "frames.txt").read_text(encoding="utf-8").splitlines():
        if line.strip():
            fields = line.split("\t") if "\t" in line else line.rsplit(" ", 1)
            fields += [""] * (3 - len(fields))
            entries.append((Image.open(frame_dir / fields[0]).convert("RGB"), int(fields[1]), fields[2]))
    if not entries:
        raise SystemExit("no frames")

    if (frame_dir / "windows").exists():
        frames = several_windows([entry[0] for entry in entries])
        save(frames, entries, output)
        return
    windows = []
    for image, *_ in entries:
        box = window_box(image) or (0, 0, image.width, image.height)
        windows.append(image.crop(box))
    width = max(w.width for w in windows) + 2 * MARGIN
    height = max(w.height for w in windows) + 2 * MARGIN
    canvas_bg = background((width, height))
    frames = []
    for window in windows:
        frame = with_shadow(canvas_bg, window.size).convert("RGBA")
        frame.alpha_composite(round_image(window, CORNER), (MARGIN, MARGIN))     # semi-transparent rounded edges, no jaggies
        frames.append(frame.convert("RGB"))
    save(frames, entries, output)


def save(frames, entries, output):
    frames[0].save(output, format="PNG", save_all=True, append_images=frames[1:], duration=[entry[1] for entry in entries],
                   loop=0, disposal=0, blend=0, optimize=True)
    print(f"{output}: {len(frames)} frames, {frames[0].width}x{frames[0].height}, {Path(output).stat().st_size // 1024} KB")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

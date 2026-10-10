"""MakeDemoGif.py - README 首页的动图 demo.gif

把 TakeScreenshots.ahk 的 "demo" 场景存下的帧 (frames.txt: 每行 "文件名 毫秒") 合成循环播放的 GIF:
每一帧里找出搜索窗口 (和截图时的纯色背景不同的范围), 加上圆角和阴影, 放在一张高斯模糊的背景图上。
窗口的左上角固定, 窗口变高变矮时动图大小不变。

背景: Tests/Screenshots/backdrop.jpg (或 .png) 存在时用它 (缩放裁剪后模糊), 否则生成一张蓝紫色调的抽象图
(几团彩色光斑 + 高斯模糊, 和 Windows 11 的默认壁纸类似)。不加颗粒: GIF 只有 256 色, 颗粒会挤掉界面文字的颜色, 文件也大很多。

    python Tests/Screenshots/MakeDemoGif.py <帧所在的文件夹> <输出的 gif>
"""
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

from RoundCorners import rounded_mask

BACKDROP = (0x8A, 0x9B, 0xB0)        # TakeScreenshots.ahk 里截图时背景的颜色
MARGIN = 56                          # 窗口四周露出的背景
SHADOW_BLUR, SHADOW_OFFSET, SHADOW_OPACITY = 18, 10, 110
UI_COLORS = 176                      # 256 色里给窗口的颜色数
HERE = Path(__file__).resolve().parent


def window_box(image):
    """帧里和背景色不同的范围 (搜索窗口)"""
    diff = ImageChops.difference(image, Image.new("RGB", image.size, BACKDROP)).convert("L").point(lambda v: 255 if v > 8 else 0)
    return diff.getbbox()


def background(size):
    """高斯模糊的背景图"""
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
    top, bottom = (24, 38, 92), (58, 32, 104)                          # 深蓝 -> 深紫
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
    glow = glow.filter(ImageFilter.GaussianBlur(max(width, height) // 6)).point(lambda v: v * 3 // 5)   # 光斑柔和一些
    return ImageChops.screen(image, glow)                               # 光斑叠在渐变上


def with_shadow(canvas_bg, window_size):
    """背景 + 窗口的阴影 (窗口本身还没放上去)"""
    shadow = Image.new("RGBA", canvas_bg.size, (0, 0, 0, 0))
    mask = rounded_mask(window_size).point(lambda v: v * SHADOW_OPACITY // 255)
    shadow.paste((0, 0, 0, 255), (MARGIN, MARGIN + SHADOW_OFFSET), mask)
    shadow = shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))
    return Image.alpha_composite(canvas_bg.convert("RGBA"), shadow).convert("RGB")


def palette_of(images, colors):
    """几张图共用的调色板: [r, g, b, ...] 正好 colors 种颜色 (不够时重复最后一种, 不补黑色)"""
    sample = Image.new("RGB", (max(i.width for i in images), sum(i.height for i in images)))
    y = 0
    for image in images:
        sample.paste(image, (0, y))
        y += image.height
    values = sample.quantize(colors=colors, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE).getpalette()[:colors * 3]
    while len(values) < colors * 3:
        values += values[-3:]
    return values


def indexed(image, values, dither):
    """按给定的调色板转成颜色索引 (L 模式, 0 起)"""
    palette = Image.new("P", (1, 1))
    palette.putpalette(values + values[-3:] * (256 - len(values) // 3))
    return Image.frombytes("L", image.size, image.quantize(palette=palette, dither=dither).tobytes())


def main(frame_dir, output):
    frame_dir = Path(frame_dir)
    entries = []
    for line in (frame_dir / "frames.txt").read_text(encoding="utf-8").splitlines():
        if line.strip():
            name, ms = line.rsplit(" ", 1)
            entries.append((Image.open(frame_dir / name).convert("RGB"), int(ms)))
    if not entries:
        raise SystemExit("no frames")

    windows = []
    for image, _ in entries:
        box = window_box(image) or (0, 0, image.width, image.height)
        windows.append(image.crop(box))
    width = max(w.width for w in windows) + 2 * MARGIN
    height = max(w.height for w in windows) + 2 * MARGIN
    canvas_bg = background((width, height))
    shaded = [with_shadow(canvas_bg, window.size) for window in windows]

    # GIF 只有 256 色: 窗口 (文字、图标) 用 UI_COLORS 种, 不抖动, 文字清楚; 背景和阴影用其余的颜色, 抖动, 渐变没有色带。
    # 所有帧共用这套颜色: 背景没变的地方每一帧都一样, GIF 只存变化的部分
    ui = palette_of(windows, UI_COLORS)
    bg = palette_of(shaded, 256 - UI_COLORS)
    frames = []
    for window, back in zip(windows, shaded):
        index = indexed(back, bg, Image.Dither.FLOYDSTEINBERG).point(lambda i: i + UI_COLORS)
        index.paste(indexed(window, ui, Image.Dither.NONE), (MARGIN, MARGIN), rounded_mask(window.size).point(lambda v: 255 if v >= 128 else 0))
        frame = index.convert("P")
        frame.putpalette(ui + bg)
        frames.append(frame)
    frames[0].save(output, save_all=True, append_images=frames[1:], duration=[ms for _, ms in entries], loop=0, optimize=True, disposal=1)
    print(f"{output}: {len(frames)} frames, {width}x{height}, {Path(output).stat().st_size // 1024} KB")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

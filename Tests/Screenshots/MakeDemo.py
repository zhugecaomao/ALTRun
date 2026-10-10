"""MakeDemo.py - README 首页的动图 demo.png (APNG)

把 TakeScreenshots.ahk 的 "demo" 场景存下的帧 (frames.txt: 每行 "文件名 Tab 毫秒 Tab 按键") 合成循环播放的动画 PNG:
每一帧里找出搜索窗口 (和截图时的纯色背景不同的范围), 加上圆角和阴影, 放在一张高斯模糊的背景图上。
不再在右下角画这一步按的键: 搜索窗口底部的状态栏已经显示 Enter / Ctrl+K 等按键 (frames.txt 的第三列不用了)。
窗口的左上角固定, 窗口变高变矮时动图大小不变。

背景: Tests/Screenshots/backdrop.jpg (或 .png) 存在时用它 (缩放裁剪后模糊), 否则生成一张蓝紫色调的抽象图
(几团彩色光斑 + 高斯模糊, 和 Windows 11 的默认壁纸类似)。
用 APNG 而不是 GIF: GIF 只有 256 色, 背景的渐变要抖动 (有颗粒), 窗口的圆角也不能半透明 (有锯齿)。
浏览器和 GitHub 都直接播放 APNG; Pillow 只存每一帧变化的部分, 文件不大。
截图按 150% 缩放截的时候 (环境变量 SHOT_SCALE=1.5), 边距、阴影、圆角也跟着放大。

    python Tests/Screenshots/MakeDemo.py <帧所在的文件夹> <输出的 png>
"""
import os
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

from RoundCorners import RADIUS, round_image, rounded_mask

SCALE = float(os.environ.get("SHOT_SCALE") or 1)
BACKDROP = (0x8A, 0x9B, 0xB0)        # TakeScreenshots.ahk 里截图时背景的颜色
MARGIN = round(56 * SCALE)           # 窗口四周露出的背景
SHADOW_BLUR, SHADOW_OFFSET, SHADOW_OPACITY = round(18 * SCALE), round(10 * SCALE), 110
CORNER = round(RADIUS * SCALE)
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
    mask = rounded_mask(window_size, CORNER).point(lambda v: v * SHADOW_OPACITY // 255)
    shadow.paste((0, 0, 0, 255), (MARGIN, MARGIN + SHADOW_OFFSET), mask)
    shadow = shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))
    return Image.alpha_composite(canvas_bg.convert("RGBA"), shadow).convert("RGB")


def main(frame_dir, output):
    frame_dir = Path(frame_dir)
    entries = []                                                        # (图片, 毫秒, 按键)
    for line in (frame_dir / "frames.txt").read_text(encoding="utf-8").splitlines():
        if line.strip():
            fields = line.split("\t") if "\t" in line else line.rsplit(" ", 1)
            fields += [""] * (3 - len(fields))
            entries.append((Image.open(frame_dir / fields[0]).convert("RGB"), int(fields[1]), fields[2]))
    if not entries:
        raise SystemExit("no frames")

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
        frame.alpha_composite(round_image(window, CORNER), (MARGIN, MARGIN))     # 圆角边缘半透明, 不会有锯齿
        frames.append(frame.convert("RGB"))
    frames[0].save(output, format="PNG", save_all=True, append_images=frames[1:], duration=[entry[1] for entry in entries],
                   loop=0, disposal=0, blend=0, optimize=True)
    print(f"{output}: {len(frames)} frames, {width}x{height}, {Path(output).stat().st_size // 1024} KB")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

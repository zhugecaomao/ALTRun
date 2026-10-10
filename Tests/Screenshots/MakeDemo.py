"""MakeDemo.py - README 首页的动图 demo.png (APNG)

把 TakeScreenshots.ahk 的 "demo" 场景存下的帧 (frames.txt: 每行 "文件名 Tab 毫秒 Tab 小标题 Tab 按键") 合成循环播放的动画 PNG:
每一帧里找出搜索窗口 (和截图时的纯色背景不同的范围), 加上圆角和阴影, 放在一张高斯模糊的背景图上;
窗口上方是这个场景的小标题, 右下角是这一步按的键 (例如 "Ctrl+Alt+C" 画成三个键帽)。
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

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont

from RoundCorners import RADIUS, round_image, rounded_mask

SCALE = float(os.environ.get("SHOT_SCALE") or 1)
BACKDROP = (0x8A, 0x9B, 0xB0)        # TakeScreenshots.ahk 里截图时背景的颜色
MARGIN = round(56 * SCALE)           # 窗口四周露出的背景
SHADOW_BLUR, SHADOW_OFFSET, SHADOW_OPACITY = round(18 * SCALE), round(10 * SCALE), 110
CORNER = round(RADIUS * SCALE)
CAPTION_SIZE, KEY_SIZE = round(22 * SCALE), round(17 * SCALE)
TOP = MARGIN + round(34 * SCALE)     # 窗口上方多留出小标题的位置
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


def font(names, size):
    """Windows 的 Segoe UI (截图工作流在 Windows 上运行); 别的系统上用 DejaVu"""
    for name in names + ["DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def draw_caption(frame, text):
    """窗口左上方的小标题: 白字, 下面一点淡淡的阴影"""
    if not text:
        return
    caption_font = font(["seguisb.ttf", "segoeui.ttf", "DejaVuSans-Bold.ttf"], CAPTION_SIZE)
    layer = Image.new("RGBA", frame.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    y = (TOP - round(10 * SCALE)) - CAPTION_SIZE - round(6 * SCALE)
    draw.text((MARGIN, y + max(1, round(SCALE))), text, font=caption_font, fill=(0, 0, 0, 90))
    layer = layer.filter(ImageFilter.GaussianBlur(max(1, round(2 * SCALE))))
    ImageDraw.Draw(layer).text((MARGIN, y), text, font=caption_font, fill=(255, 255, 255, 245))
    frame.alpha_composite(layer)


def draw_keys(frame, keys):
    """右下角的键帽: "Ctrl+Alt+C" -> [Ctrl] + [Alt] + [C]"""
    if not keys:
        return
    key_font = font(["segoeui.ttf", "DejaVuSans.ttf"], KEY_SIZE)
    draw = ImageDraw.Draw(frame)
    pad_x, pad_y, gap = round(10 * SCALE), round(6 * SCALE), round(8 * SCALE)
    radius, depth = round(6 * SCALE), max(2, round(3 * SCALE))
    parts = [part for part in keys.split("+") if part]
    widths = [max(draw.textlength(part, font=key_font) + 2 * pad_x, KEY_SIZE + 2 * pad_y) for part in parts]
    plus = draw.textlength("+", font=key_font)
    height = KEY_SIZE + 2 * pad_y
    x = frame.width - round(24 * SCALE) - sum(widths) - (len(parts) - 1) * (plus + 2 * gap)
    y = frame.height - round(24 * SCALE) - height - depth
    for i, (part, width) in enumerate(zip(parts, widths)):
        box = (round(x), y, round(x + width), y + height)
        draw.rounded_rectangle((box[0], box[1] + depth, box[2], box[3] + depth), radius, fill=(180, 186, 198, 255))   # 键帽的侧面
        draw.rounded_rectangle(box, radius, fill=(250, 251, 253, 255), outline=(205, 210, 220, 255))
        draw.text(((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), part, font=key_font, fill=(40, 44, 52), anchor="mm")
        x += width
        if i < len(parts) - 1:
            draw.text((x + gap, y + height / 2), "+", font=key_font, fill=(255, 255, 255, 230), anchor="lm")
            x += plus + 2 * gap


def with_shadow(canvas_bg, window_size):
    """背景 + 窗口的阴影 (窗口本身还没放上去)"""
    shadow = Image.new("RGBA", canvas_bg.size, (0, 0, 0, 0))
    mask = rounded_mask(window_size, CORNER).point(lambda v: v * SHADOW_OPACITY // 255)
    shadow.paste((0, 0, 0, 255), (MARGIN, TOP + SHADOW_OFFSET), mask)
    shadow = shadow.filter(ImageFilter.GaussianBlur(SHADOW_BLUR))
    return Image.alpha_composite(canvas_bg.convert("RGBA"), shadow).convert("RGB")


def main(frame_dir, output):
    frame_dir = Path(frame_dir)
    entries = []                                                        # (图片, 毫秒, 小标题, 按键)
    for line in (frame_dir / "frames.txt").read_text(encoding="utf-8").splitlines():
        if line.strip():
            fields = line.split("\t") if "\t" in line else line.rsplit(" ", 1)
            fields += [""] * (4 - len(fields))
            entries.append((Image.open(frame_dir / fields[0]).convert("RGB"), int(fields[1]), fields[2], fields[3]))
    if not entries:
        raise SystemExit("no frames")

    windows = []
    for image, *_ in entries:
        box = window_box(image) or (0, 0, image.width, image.height)
        windows.append(image.crop(box))
    width = max(w.width for w in windows) + 2 * MARGIN
    height = max(w.height for w in windows) + TOP + MARGIN
    canvas_bg = background((width, height))
    frames = []
    for window, (_, _, caption, keys) in zip(windows, entries):
        frame = with_shadow(canvas_bg, window.size).convert("RGBA")
        frame.alpha_composite(round_image(window, CORNER), (MARGIN, TOP))        # 圆角边缘半透明, 不会有锯齿
        draw_caption(frame, caption)
        draw_keys(frame, keys)
        frames.append(frame.convert("RGB"))
    frames[0].save(output, format="PNG", save_all=True, append_images=frames[1:], duration=[entry[1] for entry in entries],
                   loop=0, disposal=0, blend=0, optimize=True)
    print(f"{output}: {len(frames)} frames, {width}x{height}, {Path(output).stat().st_size // 1024} KB")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

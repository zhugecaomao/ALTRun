"""RoundCorners.py - 截图加上 Windows 11 那样的圆角

GitHub Actions 的 Windows 虚拟机没有显卡加速, 系统不画窗口圆角, 截出来是直角。
这里把每张截图的四个角做成透明 (半径 8 像素, 和 Windows 11 在 100% 缩放下一样, 边缘抗锯齿),
在浅色和深色页面上都好看。按 150% 缩放截图时 (环境变量 SHOT_SCALE=1.5) 半径也乘上这个比例。
已经是圆角的图 (左上角透明) 和动图 demo.png 不处理。

    python Tests/Screenshots/RoundCorners.py [截图文件夹]      默认 docs/images/screenshots
"""
import os
import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw

RADIUS = 8
SCALE = 4                                    # 先按 4 倍画圆角矩形再缩小: 边缘平滑


def rounded_mask(size, radius=RADIUS):
    """size 大小的圆角矩形遮罩 (L 模式, 里面 255, 角上 0)"""
    width, height = size
    big = Image.new("L", (width * SCALE, height * SCALE), 0)
    ImageDraw.Draw(big).rounded_rectangle((0, 0, width * SCALE - 1, height * SCALE - 1), radius * SCALE, fill=255)
    return big.resize(size, Image.LANCZOS)


def round_image(image, radius=RADIUS):
    """返回 RGBA 的圆角图片"""
    image = image.convert("RGBA")
    alpha = ImageChops.multiply(image.getchannel("A"), rounded_mask(image.size, radius))
    image.putalpha(alpha)
    return image


def main(folder):
    count = 0
    for path in sorted(Path(folder).glob("*.png")):
        with Image.open(path) as image:
            image.load()
        if path.stem == "demo" or (image.mode == "RGBA" and image.getpixel((0, 0))[3] == 0):
            continue
        round_image(image, round(RADIUS * float(os.environ.get("SHOT_SCALE") or 1))).save(path, optimize=True)
        count += 1
    print(f"rounded {count} screenshot(s) in {folder}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "docs/images/screenshots")

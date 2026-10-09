"""MakeDemoGif.py - README 首页的动图 demo.gif

把 TakeScreenshots.ahk 的 "demo" 场景存下的帧 (frames.txt: 每行 "文件名 毫秒") 合成循环播放的 GIF。
每一帧都是同一块屏幕区域; 先找出所有帧里和背景色不同的范围 (搜索窗口最高时的大小), 四周留一点背景再裁剪,
这样窗口变高变矮时动图大小不变。

    python Tests/Screenshots/MakeDemoGif.py <帧所在的文件夹> <输出的 gif>
"""
import sys
from pathlib import Path

from PIL import Image, ImageChops

BACKDROP = (0x8A, 0x9B, 0xB0)        # TakeScreenshots.ahk 里背景的颜色
PADDING = 16


def main(frame_dir, output):
    frame_dir = Path(frame_dir)
    entries = []
    for line in (frame_dir / "frames.txt").read_text(encoding="utf-8").splitlines():
        if line.strip():
            name, ms = line.rsplit(" ", 1)
            entries.append((Image.open(frame_dir / name).convert("RGB"), int(ms)))
    if not entries:
        raise SystemExit("no frames")

    box = None
    for image, _ in entries:
        diff = ImageChops.difference(image, Image.new("RGB", image.size, BACKDROP)).convert("L").point(lambda v: 255 if v > 8 else 0)
        found = diff.getbbox()
        if found:
            box = found if box is None else (min(box[0], found[0]), min(box[1], found[1]), max(box[2], found[2]), max(box[3], found[3]))
    width, height = entries[0][0].size
    if box:
        box = (max(0, box[0] - PADDING), max(0, box[1] - PADDING), min(width, box[2] + PADDING), min(height, box[3] + PADDING))
    else:
        box = (0, 0, width, height)

    frames = [image.crop(box).quantize(colors=128, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE) for image, _ in entries]
    frames[0].save(output, save_all=True, append_images=frames[1:], duration=[ms for _, ms in entries], loop=0, optimize=True, disposal=1)
    print(f"{output}: {len(frames)} frames, {box[2] - box[0]}x{box[3] - box[1]}, {Path(output).stat().st_size // 1024} KB")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])

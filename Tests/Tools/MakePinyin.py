"""Makes Resources/Pinyin.txt: the pinyin initial of every CJK character U+3400 - U+9FFF.

One line of 27648 characters, one per code point from U+3400: "A" ... "Z", or "-" when the
character has no reading. Pinyin.Initials (Lib/Util.ahk) reads it to match "jsb" with 记事本.

Which letter a character gets:
  1. GB2312 level-1 characters (the 3755 common ones, sorted by pinyin in GB2312) keep the
     letter of their GB2312 range, as ALTRun has always matched them (polyphones such as
     长 stay "C").
  2. Traditional characters that Resources/Kanji.txt maps to such a character take its letter
     (長 = 长 = "C").
  3. Everything else takes the first letter of its most common reading in pypinyin
     (MIT license, https://github.com/mozillazg/python-pinyin).

    pip install pypinyin
    python Tests/Tools/MakePinyin.py
"""
import os
from pypinyin import Style, lazy_pinyin

FIRST, LAST = 0x3400, 0x9FFF
REPO = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", ".."))

# Same ranges as Pinyin._ranges in Lib/Util.ahk: GB2312 code - 65536 -> initial
RANGES = [(-20319, -20284, "A"), (-20283, -19776, "B"), (-19775, -19219, "C"), (-19218, -18711, "D"),
          (-18710, -18527, "E"), (-18526, -18240, "F"), (-18239, -17923, "G"), (-17922, -17418, "H"),
          (-17417, -16475, "J"), (-16474, -16213, "K"), (-16212, -15641, "L"), (-15640, -15166, "M"),
          (-15165, -14923, "N"), (-14922, -14915, "O"), (-14914, -14631, "P"), (-14630, -14150, "Q"),
          (-14149, -14091, "R"), (-14090, -13319, "S"), (-13318, -12839, "T"), (-12838, -12557, "W"),
          (-12556, -11848, "X"), (-11847, -11056, "Y"), (-11055, -10247, "Z")]


def gb2312_initial(ch):
    try:
        code = ch.encode("gb2312")
    except UnicodeEncodeError:
        return ""
    if len(code) != 2:
        return ""
    value = (code[0] << 8) + code[1] - 65536
    for low, high, letter in RANGES:
        if low <= value <= high:
            return letter
    return ""


def pypinyin_initial(ch):
    letter = lazy_pinyin(ch, style=Style.FIRST_LETTER)[0]
    return letter.upper() if len(letter) == 1 and "a" <= letter.lower() <= "z" else ""


def main():
    simplified = {}
    with open(os.path.join(REPO, "Resources", "Kanji.txt"), encoding="utf-8") as f:
        for pair in f.read().split():
            if len(pair) == 2:
                simplified[pair[1]] = pair[0]
    letters = []
    for code in range(FIRST, LAST + 1):
        ch = chr(code)
        letter = gb2312_initial(ch) or gb2312_initial(simplified.get(ch, "")) or pypinyin_initial(ch)
        letters.append(letter or "-")
    with open(os.path.join(REPO, "Resources", "Pinyin.txt"), "w", encoding="ascii", newline="") as f:
        f.write("".join(letters))
    print(f"{len(letters)} characters, {sum(c != '-' for c in letters)} with an initial")


if __name__ == "__main__":
    main()

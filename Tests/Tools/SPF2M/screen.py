"""Read the text on a DOSBox screenshot (80x25 text mode, 640x400, 8x16 VGA font).

Every character cell is turned into a 128-bit pattern (pixel differs from the cell's
background or not) and looked up in a glyph table. The table is learned once from a
screenshot of CHARS.TXT (all printable ASCII characters, 60 per line) shown with
"type CHARS.TXT" at the top of a cleared screen:

    python screen.py learn chars.png glyphs.json
    python screen.py decode screen.png glyphs.json
"""
import json
import sys

from PIL import Image

CELL_W, CELL_H = 8, 16
PRINTABLE = ''.join(chr(c) for c in range(33, 127))
CHARS_PER_LINE = 60
FIRST_ROW = 2                           # "C:\>type CHARS.TXT" is on row 1 after cls
UNKNOWN = '\u2588'                      # shown for a pattern not in the table (e.g. the cursor)


def cells(image):
    img = image.convert('RGB')
    px = img.load()
    width, height = img.size
    grid = []
    for r in range(height // CELL_H):
        line = []
        for c in range(width // CELL_W):
            bg = px[c * CELL_W, r * CELL_H]
            line.append(''.join('1' if px[c * CELL_W + x, r * CELL_H + y] != bg else '0'
                                for y in range(CELL_H) for x in range(CELL_W)))
        grid.append(line)
    return grid


def learn(png, table_path):
    grid = cells(Image.open(png))
    lines = [PRINTABLE[i:i + CHARS_PER_LINE] for i in range(0, len(PRINTABLE), CHARS_PER_LINE)]
    table = {'0' * CELL_W * CELL_H: ' '}
    for i, text in enumerate(lines):
        for j, ch in enumerate(text):
            table[grid[FIRST_ROW + i][j]] = ch
    if len(table) != len(PRINTABLE) + 1:
        raise SystemExit('expected %d different glyphs, got %d - is CHARS.TXT at the top of the screen?'
                         % (len(PRINTABLE) + 1, len(table)))
    with open(table_path, 'w') as f:
        json.dump(table, f)
    return table


def decode(png, table):
    return [''.join(table.get(bits, UNKNOWN) for bits in line).rstrip() for line in cells(Image.open(png))]


if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[1] not in ('learn', 'decode'):
        raise SystemExit(__doc__)
    if sys.argv[1] == 'learn':
        learn(sys.argv[2], sys.argv[3])
    else:
        print('\n'.join(decode(sys.argv[2], json.load(open(sys.argv[3])))))

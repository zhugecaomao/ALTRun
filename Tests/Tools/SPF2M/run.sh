#!/usr/bin/env bash
# Run every case in cases.json through SPF2M.EXE and rebuild Tests/Fixtures/SPF2M-Reference.json.
#
#   Tests/Tools/SPF2M/run.sh /path/to/SPF2M.EXE [cases.json] [output.json]
#
# Without arguments after SPF2M.EXE: all cases, written to Tests/Fixtures/SPF2M-Reference.json.
# To try a few cases, give your own cases file and an output file somewhere else.
#
# Linux only: needs dosbox, Xvfb, xdotool, ImageMagick (import) and Python 3 with Pillow
#   sudo apt install dosbox xvfb xdotool imagemagick && pip install pillow
# SPF2M runs on a virtual display (:5), nothing appears on screen. ~140 cases take ~15 min.
set -euo pipefail

here="$(cd "$(dirname "$0")" && pwd)"
exe="${1:?usage: run.sh /path/to/SPF2M.EXE [cases.json] [output.json]}"
cases="${2:-$here/cases.json}"
output="${3:-$here/../../Fixtures/SPF2M-Reference.json}"
if [ -n "${2:-}" ] && [ -z "${3:-}" ]; then
    echo "give an output file too, so a few cases do not replace the whole fixture" >&2; exit 1
fi
work="$(mktemp -d)"
export DISPLAY=:5
trap 'pkill -P $$ || true; rm -rf "$work"' EXIT

cp "$exe" "$work/SPF2M.EXE"
cp "$here/CHARS.TXT" "$work/"

pgrep -f "Xvfb :5" >/dev/null || { Xvfb :5 -screen 0 1024x768x24 >/dev/null 2>&1 & sleep 1; }

# DOSBox shows the 80x25 text screen 1:1 (640x400), so each character is one 8x16 cell
start_dosbox() {           # $1 = command to run after mounting the work folder
    cat > "$work/dosbox.conf" <<EOF
[sdl]
fullscreen=false
output=surface
[render]
scaler=none
aspect=false
frameskip=0
[cpu]
cycles=max
[autoexec]
mount c "$work"
c:
cls
$1
EOF
    dosbox -conf "$work/dosbox.conf" >/dev/null 2>&1 &
    dosbox_pid=$!
    for _ in $(seq 1 50); do
        window="$(xdotool search --pid "$dosbox_pid" 2>/dev/null | tail -1 || true)"
        [ -n "$window" ] && break
        sleep 0.2
    done
    [ -n "$window" ] || { echo "DOSBox window not found" >&2; exit 1; }
    sleep 2
}

# 1. learn the font: show every printable character, read the glyph patterns
start_dosbox "type CHARS.TXT"
import -window "$window" "$work/chars.png"
python3 "$here/screen.py" learn "$work/chars.png" "$work/glyphs.json"
kill "$dosbox_pid"; wait "$dosbox_pid" 2>/dev/null || true

# 2. type the cases into SPF2M
start_dosbox "SPF2M.EXE"
python3 "$here/drive.py" "$window" "$work/glyphs.json" "$cases" "$work/screens.json"
kill "$dosbox_pid"; wait "$dosbox_pid" 2>/dev/null || true

# 3. screens -> fixture (the support-interval case was typed in by hand, see manual.json)
python3 "$here/build_fixture.py" "$work/screens.json" --manual "$here/manual.json" \
    -o "$output"

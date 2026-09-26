"""Type test cases into SPF2M.EXE running in DOSBox and record every screen.

    python drive.py <window-id> <glyphs.json> <cases.json> <output.json> [first-case-index]

<window-id> is the DOSBox X11 window (xdotool search --class dosbox). SPF2M must be at its
first question ("Enter Type of Tendon Profile") or at a DOS prompt in its folder.
A case is {"profile", "tendon", "start", "end", "dist"} plus optional "radius" and "contra";
a missing optional answer is left empty (Enter = SPF2M's default).

The output is re-written after every case, so a crash loses nothing: re-run from the failed
case with first-case-index. SPF2M sometimes falls back to DOS on impossible input; that is
recorded as "crashed" and SPF2M is started again for the next case.
"""
import json
import os
import subprocess
import sys
import tempfile
import time

from screen import decode


class Spf2m:
    def __init__(self, window, table):
        self.window = window
        self.table = table
        self.png = os.path.join(tempfile.gettempdir(), 'spf2m-screen.png')

    def _xdo(self, *args):
        subprocess.run(['xdotool', *args], check=True)

    def screen(self):
        subprocess.run(['import', '-window', self.window, self.png], check=True)
        return decode(self.png, self.table)

    def answer(self, text, wait=0.5):
        if text != '':
            self._xdo('type', '--window', self.window, '--delay', '40', str(text))
        self._xdo('key', '--window', self.window, 'Return')
        time.sleep(wait)

    def wait_for(self, word, timeout=6):
        end = time.time() + timeout
        while True:
            lines = self.screen()
            if any(word in line for line in lines):
                return lines
            if time.time() > end:
                raise RuntimeError('timeout waiting for %r:\n%s' % (word, '\n'.join(lines)))
            time.sleep(0.2)

    def at_dos_prompt(self, lines):
        return any(line.rstrip(' \u2588').endswith('C:\\>') for line in lines[-4:])

    def ensure_running(self):
        lines = self.screen()
        if self.at_dos_prompt(lines) and not any('Enter Type' in line for line in lines):
            self.answer('cls', 0.2)
            self.answer('SPF2M.EXE', 1.5)

    def run(self, case):
        self.ensure_running()
        self.wait_for('Enter Type of Tendon Profile')
        self.answer(case['profile'])
        self.wait_for('Enter Type of Tendon')
        self.answer(case['tendon'])
        self.wait_for('MIN. RADIUS OF CURVATURE = [')
        self.answer(case.get('radius', ''))
        self.wait_for('Level of Starting Point')
        self.answer(case['start'])
        self.wait_for('Level of Ending')
        self.answer(case['end'])
        self.wait_for('Total Horizontal Distance')
        self.answer(case['dist'], 0.8)
        lines = self.screen()
        record = {'case': case, 'screens': [lines]}

        if any('Contraflexure = [' in line for line in lines):
            self.answer(case.get('contra', ''), 0.8)
            lines = self.screen()
            record['screens'].append(lines)
            # an entered contraflexure shows the resulting radius and asks again; Enter accepts it
            for _ in range(3):
                if any('Radius of Curvature =' in line for line in lines) and \
                        not any('Change Support Interval' in line or 'CONTINUE' in line for line in lines):
                    self.answer('', 0.8)
                    lines = self.screen()
                    record['screens'].append(lines)

        for _ in range(40):
            lines = self.screen()
            if self.at_dos_prompt(lines):
                record['crashed'] = True
                record['final'] = lines
                return record
            if any('CONTINUE ON FOR ANOTHER PROFILE' in line for line in lines):
                break
            if any('Change Support Interval' in line for line in lines):
                record['screens'].append(lines)
                self.answer('N', 0.8)
                continue
            time.sleep(0.3)
        record['final'] = self.screen()
        self.answer('Y', 0.5)
        return record


def main():
    if len(sys.argv) not in (5, 6):
        raise SystemExit(__doc__)
    window, glyphs, cases_path, out_path = sys.argv[1:5]
    first = int(sys.argv[5]) if len(sys.argv) == 6 else 0
    spf = Spf2m(window, json.load(open(glyphs)))
    cases = json.load(open(cases_path))[first:]
    records = []
    for case in cases:
        try:
            records.append(spf.run(case))
            print(json.dumps(case), file=sys.stderr)
        except Exception as e:                          # noqa: BLE001 - record and stop, re-run from here
            records.append({'case': case, 'fail': str(e)})
            print('FAIL', json.dumps(case), e, file=sys.stderr)
            break
        finally:
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(records, f, indent=1, ensure_ascii=False)


if __name__ == '__main__':
    main()

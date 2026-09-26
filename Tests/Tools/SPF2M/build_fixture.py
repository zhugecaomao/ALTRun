"""Build Tests/Fixtures/SPF2M-Reference.json from the screens recorded by drive.py.

    python build_fixture.py out1.json out2.json ... [--manual manual.json] [-o SPF2M-Reference.json]

Later files win when the same case appears more than once (re-runs after a crash or a fix).
Each case in the fixture:
    Profile, Tendon, Start, End, Distance      the input (Radius / Contraflexure when entered)
    Rows                [distance, interval, actual, beam, slab] as SPF2M lists them
    ShownContraflexure  "Dist. to Point of Contraflexure" printed under the table
    ShownRadius         the radius SPF2M used, when a radius or contraflexure was entered
    Expect: "error"     SPF2M said NOT ACHIEVABLE, printed no table, or fell back to DOS
"""
import argparse
import json
import re

ROW = re.compile(r'^\s+(\d+)\s+(\d+)\s+(-?\d+)\s+(-?\d+)\s+(-?\d+)\s*$')


def parse_screens(record):
    """Rows and the numbers printed around them, from the final screen (and earlier ones)."""
    final = record.get('final', [])
    rows = [[int(v) for v in m.groups()] for m in map(ROW.match, final) if m]
    info = {}
    for line in final + sum(record.get('screens', []), []):
        m = re.search(r'Contraflexure\s+=\s+(-?\d+)\s*$', line)
        if m:
            info['contra'] = int(m.group(1))
        m = re.search(r'Radius of Curvature\s+=\s+(\d+)\s*$', line)
        if m:
            info['radius'] = int(m.group(1))
    info['error'] = any('NOT ACHIEVABLE' in line or '!!' in line for line in final)
    return rows, info


def fixture_case(record):
    case = record['case']
    out = {'Profile': case['profile'], 'Tendon': case['tendon'],
           'Start': case['start'], 'End': case['end'], 'Distance': case['dist']}
    if 'radius' in case:
        out['Radius'] = case['radius']
    if 'contra' in case:
        out['Contraflexure'] = case['contra']
    if 'intervals' in case:
        out['Intervals'] = case['intervals']
    rows, info = parse_screens(record)
    if record.get('crashed') or info['error'] or not rows:
        out['Expect'] = 'error'
        return out
    out['Rows'] = rows
    if 'contra' in info:
        out['ShownContraflexure'] = info['contra']
    if ('radius' in case or 'contra' in case) and 'radius' in info:
        out['ShownRadius'] = info['radius']
    return out


def case_key(case):
    return json.dumps(case, sort_keys=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('outputs', nargs='+', help='files written by drive.py')
    ap.add_argument('--manual', action='append', default=[],
                    help='records typed in by hand, same format as drive.py output')
    ap.add_argument('-o', '--output', default='SPF2M-Reference.json')
    args = ap.parse_args()

    records = {}
    for path in args.outputs + args.manual:
        for record in json.load(open(path, encoding='utf-8')):
            if 'fail' not in record:                    # driver gave up on this one; re-run it
                records[case_key(record['case'])] = record

    fixture = {
        'Source': 'SPF2M.EXE (Turbo Basic, Oct 93) run in DOSBox 0.74; screens decoded from the 80x25 text display',
        'Cases': [fixture_case(r) for r in records.values()],
    }
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(fixture, f, indent=1, ensure_ascii=False)
        f.write('\n')
    print('%d cases -> %s' % (len(fixture['Cases']), args.output))


if __name__ == '__main__':
    main()

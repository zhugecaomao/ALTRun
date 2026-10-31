"""bench_stats.py - sums up BenchDriver's results (one line label|metric|value; label = <version>-r<round>[s<runner>]) as Markdown tables;
the results of several runners are pooled (their files put together), and the candidate / baseline ratio of each runner is shown"""
import sys, statistics, collections, re

# Idle CPU budget, Mcycles per minute (median): after #291 / #295 / #299 ALTRun measured 18 to 25, a script with no timer 0.8,
# and a single SetTimer left set brings back AutoHotkey's 10 ms timer at about 330. Over the budget the candidate is flagged
IDLE_METRIC, IDLE_BUDGET = "idle cpu Mcycles per minute", 60

def pct(values, p):
    s = sorted(values)
    k = (len(s) - 1) * p / 100
    f = int(k)
    return s[f] if f + 1 >= len(s) else s[f] + (s[f + 1] - s[f]) * (k - f)

samples = collections.defaultdict(list)
by_runner = collections.defaultdict(list)                      # (runner, version, metric) -> values
fails = collections.Counter()
fail_examples = collections.defaultdict(list)
res = collections.defaultdict(dict)
notes = collections.defaultdict(list)                          # version -> notes from the log (e.g. the time of each step of saving the icon cache)
background = collections.defaultdict(list)                     # (version, task) -> timings over 50 ms (ALTRun log's "Perf: slow")
versions = []
for line in open(sys.argv[1], encoding="utf-8"):
    parts = line.lstrip("\ufeff").rstrip("\r\n").split("|", 2)  # each runner's file starts with a BOM, also inside the joined file
    if len(parts) != 3:
        continue
    label, metric, value = parts
    version = label.rsplit("-r", 1)[0]
    m = re.search(r"-r\d+s(\d+)$", label)
    runner = int(m.group(1)) if m else 1
    if version not in versions:
        versions.append(version)
    if metric.startswith("res "):
        res[label][metric[4:]] = value
    elif metric == "note":
        notes[version].append(value)
    elif metric.startswith("bg "):
        background[(version, metric[3:])].append(float(value))
    elif metric.endswith(" fail") or metric in ("fatal", "log error") or metric.startswith("slow "):
        fails[(version, metric)] += 1
        if len(fail_examples[(version, metric)]) < 12:
            fail_examples[(version, metric)].append(value)
    else:
        try:
            samples[(version, metric)].append(float(value))
            by_runner[(runner, version, metric)].append(float(value))
        except ValueError:
            pass

out = []
metrics = []
for (v, m) in samples:
    if m not in metrics:
        metrics.append(m)
out.append("| metric | " + " | ".join(f"{v} n / median / p90 / p99 / max" for v in versions) + " |")
out.append("|---|" + "---|" * len(versions))
for m in sorted(metrics):
    cells = []
    for v in versions:
        vals = samples.get((v, m), [])
        if vals:
            cells.append(f"{len(vals)} / {statistics.median(vals):.1f} / {pct(vals, 90):.1f} / {pct(vals, 99):.1f} / {max(vals):.1f}")
        else:
            cells.append("-")
    out.append(f"| {m} | " + " | ".join(cells) + " |")
out.append("")
runners = sorted({r for (r, v, m) in by_runner})
if len(runners) > 1 and len(versions) == 2:
    # Each runner compares both versions on its own machine: the ratios should agree; one runner far off the others means
    # that machine was disturbed, and a difference seen on only one runner is not to be trusted
    a, b = versions
    out.append(f"{b} / {a} median on each runner (below 1 = {b} faster):")
    out.append("")
    out.append("| metric | " + " | ".join(f"runner {r}" for r in runners) + " |")
    out.append("|---|" + "---|" * len(runners))
    for m in sorted(metrics):
        cells = []
        for r in runners:
            va, vb = by_runner.get((r, a, m), []), by_runner.get((r, b, m), [])
            ma, mb = (statistics.median(va), statistics.median(vb)) if va and vb else (0, 0)
            cells.append(f"{mb / ma:.2f}" if ma >= 1 else ("-" if not (va and vb) else "~"))   # medians under 1 ms: noise, no ratio
        out.append(f"| {m} | " + " | ".join(cells) + " |")
    out.append("")
out.append("Failures: " + (", ".join(f"{v} {m}: {n} (e.g. {'; '.join(fail_examples[(v, m)])})" for (v, m), n in sorted(fails.items())) or "none"))
out.append("")
if background:
    out.append("Background tasks over 50 ms (from the ALTRun log):")
    out.append("")
    out.append("| version | task | n | median | max |")
    out.append("|---|---|---|---|---|")
    for (v, task), vals in sorted(background.items(), key=lambda kv: -max(kv[1])):
        out.append(f"| {v} | {task} | {len(vals)} | {statistics.median(vals):.0f} | {max(vals):.0f} |")
    out.append("")
for v, items in notes.items():
    out.append(f"Notes ({v}): " + "; ".join(items[:12]))
    out.append("")
keys = []
for label in res:
    for k in res[label]:
        if k not in keys:
            keys.append(k)
out.append("| round | " + " | ".join(keys) + " |")
out.append("|---|" + "---|" * len(keys))
for label in res:
    out.append(f"| {label} | " + " | ".join(res[label].get(k, "") for k in keys) + " |")
candidate = "candidate" if "candidate" in versions else (versions[-1] if versions else "")   # bench.yml names the versions baseline and candidate
idle = samples.get((candidate, IDLE_METRIC), [])
if idle and statistics.median(idle) > IDLE_BUDGET:
    message = f"Idle CPU over budget: {candidate} median {statistics.median(idle):.1f} Mcycles per minute (budget {IDLE_BUDGET}); is a SetTimer left set while idle?"
    out.insert(0, f"**{message}**\n")
    out.append(f"::warning::{message}")                          # an annotation on the run; bench.yml keeps this line out of the summary
print("\n".join(out))

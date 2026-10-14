"""bench_stats.py - 汇总 BenchDriver 的结果 (每行 label|metric|value; label = <版本>-r<轮>) 成 Markdown 表"""
import sys, statistics, collections

def pct(values, p):
    s = sorted(values)
    k = (len(s) - 1) * p / 100
    f = int(k)
    return s[f] if f + 1 >= len(s) else s[f] + (s[f + 1] - s[f]) * (k - f)

samples = collections.defaultdict(list)
fails = collections.Counter()
fail_examples = collections.defaultdict(list)
res = collections.defaultdict(dict)
versions = []
for line in open(sys.argv[1], encoding="utf-8-sig"):
    parts = line.rstrip("\n").split("|", 2)
    if len(parts) != 3:
        continue
    label, metric, value = parts
    version = label.rsplit("-r", 1)[0]
    if version not in versions:
        versions.append(version)
    if metric.startswith("res "):
        res[label][metric[4:]] = value
    elif metric.endswith(" fail") or metric == "fatal" or metric.startswith("slow "):
        fails[(version, metric)] += 1
        if len(fail_examples[(version, metric)]) < 12:
            fail_examples[(version, metric)].append(value)
    else:
        try:
            samples[(version, metric)].append(float(value))
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
out.append("Failures: " + (", ".join(f"{v} {m}: {n} (e.g. {'; '.join(fail_examples[(v, m)])})" for (v, m), n in sorted(fails.items())) or "none"))
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
print("\n".join(out))

import json
import re
from collections import Counter
from pathlib import Path

p = Path(__file__).with_name("basic_maths_unit_test_l01_l18.json")
data = json.loads(p.read_text(encoding="utf-8"))
qs = data["questions"]

assert data["title"] == "Basic Maths Unit Test"
assert data["lectureNumber"] == 18
assert data["durationMinutes"] == 90
assert data["marksPerQuestion"] == 2
assert len(qs) == 45

bad = []
bands = {f"L{i:02d}": 0 for i in range(1, 19)}
ascii_bad = []

for i, q in enumerate(qs, 1):
    for key in ("topic", "prompt", "options", "correctIndex", "explanation"):
        if key not in q:
            bad.append((i, "missing", key))
    if len(q.get("options", [])) != 4:
        bad.append((i, "options", len(q.get("options", []))))
    ci = q.get("correctIndex", -1)
    if not isinstance(ci, int) or not (0 <= ci <= 3):
        bad.append((i, "correctIndex", ci))
    m = re.search(r"L(\d{2})", q["topic"])
    if m:
        bands[f"L{m.group(1)}"] += 1
    # also count L12 / L13 style
    for n in re.findall(r"L(\d{2})", q["topic"]):
        bands[f"L{n}"] += 0  # already counted
    blob = (q["prompt"] + " " + q["explanation"]).lower()
    for term in ("sqrt(", "theta", "i-hat"):
        if term in blob:
            ascii_bad.append((i, term))

# For combined tags like "L12 / L13 rates", count both
for q in qs:
    for n in re.findall(r"L(\d{2})", q["topic"]):
        bands[f"L{n}"] = bands.get(f"L{n}", 0)  # ensure
# recount properly including multi tags
bands = {f"L{i:02d}": 0 for i in range(1, 19)}
for q in qs:
    for n in set(re.findall(r"L(\d{2})", q["topic"])):
        bands[f"L{n}"] += 1

missing = [k for k, v in bands.items() if v == 0]


def ln_first(t: str) -> int:
    m = re.search(r"L(\d{2})", t)
    return int(m.group(1)) if m else 0


bucket = Counter()
for q in qs:
    n = ln_first(q["topic"])
    if 1 <= n <= 5:
        bucket["trig_L01-05"] += 1
    elif 6 <= n <= 9:
        bucket["ap_graphs_L06-09"] += 1
    elif 10 <= n <= 14:
        bucket["diff_L10-14"] += 1
    elif 15 <= n <= 18:
        bucket["int_L15-18"] += 1
    else:
        bucket[f"other_{n}"] += 1

# L09 items are in ap_graphs by first tag; plan wanted some L09 in both ap and diff.
# Adjust: move L09 differentiation/exponential into diff mentally — report raw first-tag buckets.

print("count", len(qs))
print("total_marks", data["marksPerQuestion"] * len(qs))
print("bands", bands)
print("missing_bands", missing)
print("buckets_by_first_tag", dict(bucket))
print("bad", bad)
print("ascii_bad", ascii_bad)

# Plan: every lecture band at least once. Fill gaps if any by failing clearly.
if missing:
    raise SystemExit(f"Missing lecture coverage: {missing}")
if bad:
    raise SystemExit(f"Schema errors: {bad}")
if ascii_bad:
    raise SystemExit(f"ASCII leftovers: {ascii_bad}")
print("VALIDATION_OK")

import json
from pathlib import Path

data = json.loads(Path(__file__).with_name("basic_maths_unit_test_l01_l18.json").read_text(encoding="utf-8"))
lines = [f"{data['title']} — {len(data['questions'])} questions", ""]
for i, q in enumerate(data["questions"], 1):
    lines.append(f"{i:02d}. [{q['topic']}]")
    lines.append(f"    {q['prompt']}")
    for j, opt in enumerate(q["options"]):
        mark = "*" if j == q["correctIndex"] else " "
        lines.append(f"   {mark} {chr(65 + j)}. {opt}")
    lines.append("")
out = Path(__file__).resolve().parents[2] / "_qs_list.txt"
out.write_text("\n".join(lines), encoding="utf-8")
print(out)

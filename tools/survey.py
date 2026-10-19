# -*- coding: utf-8 -*-
"""先摸清原始文本结构，为正式解析器定位边界情况。"""
import re
from pathlib import Path

RAW = Path(__file__).with_name("_raw_all.txt")
text = RAW.read_text(encoding="utf-8")

lines = []
page = 0
for ln in text.split("\n"):
    m = re.match(r"^=====\s*PAGE\s+(\d+)\s*=====$", ln.strip())
    if m:
        page = int(m.group(1))
        continue
    if ln.strip():
        lines.append((ln.rstrip(), page))

print("non-empty lines:", len(lines))

# 1) 行首题号（带点）
qdot = [(i, ln, p) for i, (ln, p) in enumerate(lines) if re.match(r"^\s*\d{1,3}\s*[.．、]", ln)]
print("lines starting with NUMBER.:", len(qdot))

# 2) 行首题号（不带点，可能是粘连）
qnodot = [(i, ln, p) for i, (ln, p) in enumerate(lines)
          if re.match(r"^\s*\d{1,3}(?![\d])\s*[^A-Za-z0-9.．、\s]", ln) and not re.match(r"^\s*\d{1,3}\s*[.．、]", ln)]
print("lines starting with NUMBER w/o dot:", len(qnodot))
for i, ln, p in qnodot[:20]:
    print(f"  p{p} | {ln[:90]}")

# 3) 单行出现两个以上选项标记
multi_opt_line = [(i, ln) for i, (ln, p) in enumerate(lines) if len(re.findall(r"[A-D]\s*[.．、]", ln)) >= 2]
print("\nlines with >=2 option markers:", len(multi_opt_line))
for i, ln in multi_opt_line[:15]:
    print("  |", ln[:110])

# 4) 选项标记形态统计
opt_forms = {}
for ln, p in lines:
    m = re.match(r"^\s*([A-D])\s*([.．、])\s*", ln)
    if m:
        opt_forms[m.group(2)] = opt_forms.get(m.group(2), 0) + 1
print("\noption marker chars:", opt_forms)

# 5) 答案括号里的值
ans_vals = {}
for ln, p in lines:
    for m in re.finditer(r"[（(]\s*([^（）()]{0,8}?)\s*[）)]", ln):
        v = m.group(1).strip()
        if re.fullmatch(r"[A-D]{1,4}|对|错|正确|错误|√|×|T|F", v):
            ans_vals[v] = ans_vals.get(v, 0) + 1
print("\nanswer-bracket values (top 30):", sorted(ans_vals.items(), key=lambda x: -x[1])[:30])

# 6) 题型前缀
for t in ["单选题", "多选题", "判断题"]:
    print(t, "count:", text.count(t))

# 7) 每题块里是否一定出现 A./B./C./D.

#!/usr/bin/env python3
"""檢查 md 能否回答題目，並檢查圖片表格是否出現在對應標題之後（順序）。"""
import json
import sys
from pathlib import Path

md = Path(sys.argv[1]).read_text(encoding="utf-8")
questions = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))

score = 0
for q in questions:
    hit = q["answer"] in md
    score += hit
    tag = "圖片內" if q["in_image"] else "儲存格"
    print(f"[{'PASS' if hit else 'FAIL'}] ({tag}) {q['q']} → {q['answer']}")

# 順序檢查：表1 標題 < 438 < 結論 < 表2 標題 < 88.1
marks = ["表1：", "438", "結論：", "表2：", "88.1"]
pos = [md.find(m) for m in marks]
ordered = all(p >= 0 for p in pos) and pos == sorted(pos)
print(f"[{'PASS' if ordered else 'FAIL'}] 順序：文字與圖片表格依原始位置排列 {pos}")
print(f"\n分數：{score + ordered}/{len(questions) + 1}")

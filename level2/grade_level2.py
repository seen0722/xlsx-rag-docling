#!/usr/bin/env python3
"""Level 2 評分：逐題檢查轉出的 md，輸出每種圖片類型的通過率。

結果分三種：
  PASS   值在正確位置
  STRUCT 值有出現在該圖片段落，但不在正確的列／欄（結構錯）
  MISS   值根本沒出現（OCR 讀不到或讀錯）

用法: python -I grade_level2.py converted.md level2_exam.json [report.md]
"""
import json
import re
import sys
import unicodedata
from collections import OrderedDict
from pathlib import Path


def norm(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)   # 全形→半形、℃→°C、µ→μ
    return re.sub(r"\s+", "", s).replace("°C", "℃")


def split_segments(md: str) -> dict:
    """依 <!-- image N --> 切段，回傳 {'01': text, ...}。"""
    parts = re.split(r"<!-- image (\d+) -->", md)
    segs = {}
    for i in range(1, len(parts) - 1, 2):
        n, body = int(parts[i]), parts[i + 1]
        # 切到下一張圖的說明文字為止，避免下一題的 caption 混進本段造成誤判
        nxt = re.search(rf"圖{n + 1}[：:]", body)
        segs[f"{n:02d}"] = body[:nxt.start()] if nxt else body
    return segs


def table_rows(seg: str):
    return [[norm(c) for c in line.strip().strip("|").split("|")]
            for line in seg.splitlines() if line.strip().startswith("|")]


def check_cell(seg, chk):
    accept = {norm(a) for a in chk["accept"]}
    label = norm(chk["row"])
    for cells in table_rows(seg):
        if label in cells:
            i = cells.index(label) + chk["offset"]
            if i < len(cells) and cells[i] in accept:
                return "PASS"
    flat = norm(seg)
    return "STRUCT" if any(a in flat for a in accept) else "MISS"


def check_tokens(seg, chk):
    toks = [norm(t) for t in chk["tokens"]]
    flat = norm(seg)
    kind = chk["type"]
    if kind == "contains":
        return "PASS" if all(t in flat for t in toks) else "MISS"
    if not all(t in flat for t in toks):
        return "MISS"
    if kind == "order":
        pos = [flat.find(t) for t in toks]
        return "PASS" if pos == sorted(pos) else "STRUCT"
    if kind == "same_line":
        lines = [norm(l) for l in seg.splitlines()]
        return "PASS" if any(all(t in l for t in toks) for l in lines) else "STRUCT"
    raise ValueError(f"unknown check type: {kind}")


def grade(md_path, exam_path):
    segs = split_segments(Path(md_path).read_text(encoding="utf-8"))
    exam = json.loads(Path(exam_path).read_text(encoding="utf-8"))
    results = []
    for e in exam:
        seg = segs.get(e["image"], "")
        chk = e["check"]
        status = check_cell(seg, chk) if chk["type"] == "cell" else check_tokens(seg, chk)
        results.append({**e, "status": status})
    return results


def report(results) -> str:
    by_cat = OrderedDict()
    for r in results:
        by_cat.setdefault((r["image"], r["category"]), []).append(r["status"])
    lines = ["| 圖 | 類型 | 通過 | PASS | STRUCT | MISS |", "|---|---|---|---|---|---|"]
    for (img, cat), sts in by_cat.items():
        p = sts.count("PASS")
        lines.append(f"| {img} | {cat} | {p}/{len(sts)} | {p} | "
                     f"{sts.count('STRUCT')} | {sts.count('MISS')} |")
    total = sum(r["status"] == "PASS" for r in results)
    lines.append(f"\n**總分：{total}/{len(results)}（{total / len(results):.0%}）**\n")
    lines.append("| # | 類型 | 題目 | 標準答案 | 結果 |\n|---|---|---|---|---|")
    for r in results:
        lines.append(f"| {r['id']} | {r['category']} | {r['question']} | {r['answer']} | {r['status']} |")
    return "\n".join(lines)


if __name__ == "__main__":
    rep = report(grade(sys.argv[1], sys.argv[2]))
    print(rep)
    if len(sys.argv) > 3:
        Path(sys.argv[3]).write_text(rep + "\n", encoding="utf-8")

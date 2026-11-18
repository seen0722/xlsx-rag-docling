#!/usr/bin/env python3
"""Level 2 測試檔：12 種圖片類型嵌入同一份 xlsx，每題附標準答案與自動檢查規則。

用法: python -I make_level2.py <輸出目錄>
產出: level2.xlsx、images/*.png|jpg、level2_exam.json、level2_exam.csv

檢查規則（給 grade_level2.py）:
  cell     : md 表格中找到 row 標籤那格，往右 offset 格的值 ∈ accept
  contains : 圖片段落內包含所有 tokens
  order    : tokens 依序出現在圖片段落內
  same_line: tokens 出現在同一行（同一列表格）
"""
import csv
import json
import math
import random
import sys
from pathlib import Path

import numpy as np
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageFilter, ImageFont

FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"
WHITE, BLACK, HEAD = "white", "black", "#dde6f0"
RED, GREEN = "#f4b6b6", "#b9e4c2"


def font(size):
    return ImageFont.truetype(FONT_PATH, size)


# ---------- 共用：一般表格 ----------
def draw_table(rows, col_w, row_h=44, size=22, grid=True, fills=None, pad=12):
    """rows: 2D list；fills: {(r, c): color}；回傳 PIL Image。"""
    ft = font(size)
    widths = col_w if isinstance(col_w, list) else [col_w] * len(rows[0])
    w, h = sum(widths) + 2, row_h * len(rows) + 2
    img = Image.new("RGB", (w, h), WHITE)
    d = ImageDraw.Draw(img)
    fills = fills or {}
    for r, row in enumerate(rows):
        x = 0
        for c, val in enumerate(row):
            box = [x, r * row_h, x + widths[c], (r + 1) * row_h]
            fill = fills.get((r, c), HEAD if (r == 0 and grid) else None)
            if fill:
                d.rectangle(box, fill=fill)
            if grid:
                d.rectangle(box, outline=BLACK, width=2)
            d.text((x + pad, r * row_h + (row_h - size) // 2 - 2), str(val), fill=BLACK, font=ft)
            x += widths[c]
    if not grid:  # 無格線表格只留表頭下的一條線（常見報表樣式）
        d.line([0, row_h, w, row_h], fill=BLACK, width=1)
    return img


# ---------- 1. 合併儲存格 + 多層表頭 ----------
def img_merged():
    cw, rh, ft = [150, 150, 150, 150, 150], 44, font(22)
    data = [["T10", "12", "3", "7", "1"], ["T11", "25", "6", "14", "2"],
            ["T12-Pro", "9", "0", "4", "5"]]
    img = Image.new("RGB", (sum(cw) + 2, rh * 5 + 2), WHITE)
    d = ImageDraw.Draw(img)
    xs = np.cumsum([0] + cw)

    def cell(x0, y0, x1, y1, text, fill=None):
        if fill:
            d.rectangle([x0, y0, x1, y1], fill=fill)
        d.rectangle([x0, y0, x1, y1], outline=BLACK, width=2)
        tw = d.textlength(text, font=ft)
        d.text(((x0 + x1 - tw) / 2, (y0 + y1) / 2 - 14), text, fill=BLACK, font=ft)

    cell(xs[0], 0, xs[1], rh * 2, "機型", HEAD)                 # 跨兩列
    cell(xs[1], 0, xs[3], rh, "Android 14", HEAD)               # 跨兩欄
    cell(xs[3], 0, xs[5], rh, "Android 15", HEAD)
    for i, t in enumerate(["CTS fail", "VTS fail", "CTS fail", "VTS fail"]):
        cell(xs[i + 1], rh, xs[i + 2], rh * 2, t, HEAD)
    for r, row in enumerate(data):
        for c, v in enumerate(row):
            cell(xs[c], rh * (r + 2), xs[c + 1], rh * (r + 3), v)
    return img


# ---------- 2. 無格線表格 ----------
def img_nogrid():
    rows = [["測試情境", "平均電流(mA)", "持續時間(s)"],
            ["Boot", "812.4", "32.1"], ["Idle", "145.7", "600"],
            ["Video", "523.9", "1800"], ["Camera", "967.2", "300"]]
    return draw_table(rows, [170, 210, 200], grid=False, row_h=40)


# ---------- 3. 只用顏色表示 PASS/FAIL ----------
def img_color():
    rows = [["測項", "T10", "T11", "T12-Pro"],
            ["Suspend/Resume", "", "", ""], ["Camera", "", "", ""],
            ["Wi-Fi", "", "", ""], ["Bluetooth", "", "", ""]]
    fails = {(2, 1), (1, 2), (4, 3)}  # T10 Camera, T11 Suspend, T12-Pro BT
    fills = {(r, c): (RED if (r, c) in fails else GREEN)
             for r in range(1, 5) for c in range(1, 4)}
    img = draw_table(rows, [230, 130, 130, 130], fills=fills)
    d = ImageDraw.Draw(img)
    d.text((10, img.height - 2), "", font=font(16))
    legend = Image.new("RGB", (img.width, 40), WHITE)
    ld = ImageDraw.Draw(legend)
    ld.rectangle([10, 10, 30, 30], fill=GREEN, outline=BLACK)
    ld.text((38, 8), "PASS", fill=BLACK, font=font(18))
    ld.rectangle([110, 10, 130, 30], fill=RED, outline=BLACK)
    ld.text((138, 8), "FAIL", fill=BLACK, font=font(18))
    out = Image.new("RGB", (img.width, img.height + 40), WHITE)
    out.paste(img, (0, 0))
    out.paste(legend, (0, img.height))
    return out


# ---------- 4. 特殊符號與單位 ----------
def img_symbols():
    rows = [["項目", "規格", "實測", "結果"],
            ["Suspend 電流", "≤ 5.0 µA", "3.8 µA", "✓"],
            ["溫升", "≤ 15 ℃", "12.5 ℃", "✓"],
            ["充電效率", "≥ 85%", "83.2%", "✗"],
            ["電池容量", "5,000 mAh", "4,987 mAh", "✓"],
            ["電壓誤差", "±0.05 V", "±0.03 V", "✓"]]
    return draw_table(rows, [190, 170, 170, 100])


# ---------- 5. 易混字元 ----------
def img_confusable():
    rows = [["序號", "料號", "批號", "最低溫度(℃)"],
            ["1", "B01-O8", "L1-I0-7", "-5.0"],
            ["2", "S5-8B0", "O0-Q9", "-12.5"],
            ["3", "Z2-7T1", "I1l-60", "0.5"]]
    return draw_table(rows, [100, 170, 170, 190])


# ---------- 6. 低解析度 + JPEG 壓縮 ----------
def img_lowres(out_dir):
    rows = [["站別", "投入數", "良率"], ["SMT", "1250", "98.72%"],
            ["FATP", "1232", "97.15%"], ["OQC", "1197", "99.40%"]]
    img = draw_table(rows, [150, 150, 150])
    small = img.resize((int(img.width * 0.45), int(img.height * 0.45)), Image.BILINEAR)
    path = out_dir / "06_lowres.jpg"
    small.save(path, "JPEG", quality=25)
    return path


# ---------- 7. 手機翻拍（透視、旋轉、陰影、雜訊）----------
def _perspective_coeffs(src, dst):
    m = []
    for (x, y), (u, v) in zip(dst, src):
        m.append([x, y, 1, 0, 0, 0, -u * x, -u * y])
        m.append([0, 0, 0, x, y, 1, -v * x, -v * y])
    return np.linalg.solve(np.array(m, float), np.array(src, float).reshape(8))


def img_photo():
    rows = [["版本", "發布日期", "Build ID"],
            ["v1.0.3", "2026-09-14", "UKQ1.260914.001"],
            ["v1.1.0", "2026-10-02", "UKQ1.261002.003"],
            ["v1.2.0", "2026-11-05", "UKQ1.261105.002"]]
    tbl = draw_table(rows, [140, 190, 260])
    pad = 60
    canvas = Image.new("RGB", (tbl.width + pad * 2, tbl.height + pad * 2), "#e9e4da")
    canvas.paste(tbl, (pad, pad))
    w, h = canvas.size
    src = [(0, 0), (w, 0), (w, h), (0, h)]
    dst = [(40, 25), (w - 10, 0), (w - 30, h - 5), (0, h - 30)]   # 梯形透視
    img = canvas.transform((w, h), Image.PERSPECTIVE, _perspective_coeffs(src, dst),
                           Image.BICUBIC, fillcolor="#7a7468")
    img = img.rotate(3.5, expand=False, fillcolor="#7a7468", resample=Image.BICUBIC)
    arr = np.asarray(img).astype(float)
    shade = np.linspace(1.0, 0.62, w)[None, :, None]               # 左亮右暗的陰影
    rng = np.random.default_rng(7)
    arr = arr * shade + rng.normal(0, 9, arr.shape)
    img = Image.fromarray(np.clip(arr, 0, 255).astype("uint8"))
    return img.filter(ImageFilter.GaussianBlur(0.8))


# ---------- 8. Excel 截圖（列號、欄標、UI 雜訊）----------
def img_excel_ui():
    data = [["機型", "RAM", "ROM", "售價(NT$)"], ["T10", "4GB", "64GB", "8,990"],
            ["T11", "6GB", "128GB", "11,490"], ["T12-Pro", "8GB", "256GB", "15,990"]]
    rn_w, cw, rh, top = 40, 140, 32, 70
    w, h = rn_w + cw * 4 + 2, top + rh * (len(data) + 2)
    img = Image.new("RGB", (w, h), WHITE)
    d, ft, sm = ImageDraw.Draw(img), font(18), font(15)
    d.rectangle([0, 0, w, 32], fill="#217346")                     # 綠色功能區
    d.text((10, 6), "常用   插入   頁面配置   公式   資料", fill=WHITE, font=sm)
    d.rectangle([0, 32, w, top - rh + 30], fill="#f3f3f3")
    d.text((8, 40), "B3", fill=BLACK, font=sm)
    d.text((70, 40), "fx   T11", fill=BLACK, font=sm)
    y0 = top
    d.rectangle([0, y0, w, y0 + rh], fill="#e6e6e6")
    for i, col in enumerate("ABCD"):
        d.text((rn_w + cw * i + cw // 2 - 5, y0 + 6), col, fill="#444", font=sm)
    for r in range(len(data) + 1):
        y = y0 + rh * (r + 1)
        d.rectangle([0, y, rn_w, y + rh], fill="#e6e6e6", outline="#c8c8c8")
        d.text((14, y + 7), str(r + 1), fill="#444", font=sm)
        for c in range(4):
            x = rn_w + cw * c
            d.rectangle([x, y, x + cw, y + rh], outline="#d4d4d4")
            if r < len(data):
                d.text((x + 8, y + 5), data[r][c], fill=BLACK, font=ft)
    d.rectangle([rn_w + cw * 0, y0 + rh * 3, rn_w + cw, y0 + rh * 4], outline="#217346", width=3)
    return img


# ---------- 9. 密集大表（20 列 × 8 欄小字）----------
DENSE_MODULES = ["Display", "Touch", "Audio", "Camera", "Sensor"]


def dense_value(r, c):
    return f"{(r * 37 + c * 53) % 900 / 10 + 10:.1f}"


def img_dense():
    head = ["TC ID", "模組", "T10", "T11", "T12", "T13", "T14", "T15"]
    rows = [head] + [[f"TC-{r:03d}", DENSE_MODULES[r % 5]] +
                     [dense_value(r, c) for c in range(2, 8)] for r in range(1, 21)]
    return draw_table(rows, [100, 100] + [78] * 6, row_h=24, size=14, pad=6)


# ---------- 10. 長條圖（無數值標籤）----------
def img_chart():
    vals = {"T10": 42, "T11": 35, "T12-Pro": 58}
    w, h, left, bottom, top = 560, 380, 70, 320, 40
    img = Image.new("RGB", (w, h), WHITE)
    d, ft = ImageDraw.Draw(img), font(18)
    d.text((left, 6), "待機功耗 (mW)", fill=BLACK, font=font(20))
    scale = (bottom - top) / 60
    for t in range(0, 61, 20):
        y = bottom - t * scale
        d.line([left, y, w - 20, y], fill="#dddddd")
        d.text((left - 36, y - 11), str(t), fill=BLACK, font=ft)
    d.line([left, top, left, bottom], fill=BLACK, width=2)
    d.line([left, bottom, w - 20, bottom], fill=BLACK, width=2)
    for i, (k, v) in enumerate(vals.items()):
        x = left + 40 + i * 150
        d.rectangle([x, bottom - v * scale, x + 80, bottom], fill="#4a7bb7")
        d.text((x + 40 - d.textlength(k, font=ft) / 2, bottom + 8), k, fill=BLACK, font=ft)
    return img


# ---------- 11. 流程圖（蛇形排列：閱讀順序 ≠ 流程順序）----------
def img_flow():
    w, h = 760, 360
    img = Image.new("RGB", (w, h), WHITE)
    d, ft = ImageDraw.Draw(img), font(20)
    bw, bh = 170, 56
    pos = {"Power Key": (40, 40), "PBL": (295, 40), "XBL": (550, 40),
           "ABL/UEFI": (550, 180), "Kernel": (295, 180), "init": (40, 180),
           "EDL mode": (550, 300)}
    for name, (x, y) in pos.items():
        fill = "#fde2c8" if name == "EDL mode" else "#e3ecf7"
        d.rounded_rectangle([x, y, x + bw, y + bh], 10, fill=fill, outline=BLACK, width=2)
        d.text((x + bw / 2 - d.textlength(name, font=ft) / 2, y + 14), name, fill=BLACK, font=ft)

    def arrow(p, q, label=None):
        d.line([p, q], fill=BLACK, width=3)
        ang = math.atan2(q[1] - p[1], q[0] - p[0])
        for s in (2.6, -2.6):
            d.line([q, (q[0] - 14 * math.cos(ang + s / 6), q[1] - 14 * math.sin(ang + s / 6))],
                   fill=BLACK, width=3)
        if label:
            d.text(((p[0] + q[0]) / 2 - 80, (p[1] + q[1]) / 2 + 40), label, fill="#b00000", font=font(16))

    arrow((210, 68), (295, 68))
    arrow((465, 68), (550, 68))
    arrow((635, 96), (635, 180))
    arrow((550, 208), (465, 208))
    arrow((295, 208), (210, 208))
    arrow((720, 96), (720, 300), "boot fail")
    return img


# ---------- 12. 紅框 + 箭頭標註 ----------
def img_annot():
    rows = [["問題項目", "嚴重度", "負責角色"], ["Wi-Fi 斷線", "Major", "CP"],
            ["Camera 黑屏", "Critical", "BSP"], ["Setup Wizard 卡住", "Major", "AP"],
            ["OTA 失敗", "Minor", "SCM"]]
    tbl = draw_table(rows, [250, 150, 150])
    img = Image.new("RGB", (tbl.width + 230, tbl.height + 20), WHITE)
    img.paste(tbl, (10, 10))
    d = ImageDraw.Draw(img)
    y0, y1 = 10 + 44 * 2, 10 + 44 * 3
    d.rectangle([4, y0 - 4, tbl.width + 16, y1 + 4], outline="#e00000", width=4)
    ax = tbl.width + 20
    d.line([(ax + 70, y0 + 22), (ax, y0 + 22)], fill="#e00000", width=4)
    d.polygon([(ax, y0 + 22), (ax + 16, y0 + 12), (ax + 16, y0 + 32)], fill="#e00000")
    d.text((ax + 76, y0 + 8), "需優先處理", fill="#e00000", font=font(22))
    return img


# ---------- 題目 ----------
def C(row, offset, *accept):
    return {"type": "cell", "row": row, "offset": offset, "accept": list(accept)}


def T(kind, *tokens):
    return {"type": kind, "tokens": list(tokens)}


IMAGES = [  # (id, 類別, 說明文字, 產生函式)
    ("01", "合併儲存格+多層表頭", "圖1：各 Android 版本 xTS fail 數", img_merged),
    ("02", "無格線表格", "圖2：功耗量測（無格線報表）", img_nogrid),
    ("03", "顏色表示狀態", "圖3：smoke test 結果（綠＝PASS，紅＝FAIL）", img_color),
    ("04", "特殊符號與單位", "圖4：電源規格驗證", img_symbols),
    ("05", "易混字元", "圖5：料號與批號清單", img_confusable),
    ("06", "低解析度+JPEG", "圖6：產線良率（低解析截圖）", None),
    ("07", "手機翻拍", "圖7：release 版本紀錄（手機拍螢幕）", img_photo),
    ("08", "Excel 截圖含 UI", "圖8：機型規格（Excel 畫面截圖）", img_excel_ui),
    ("09", "密集大表", "圖9：效能測試 20 項（ms）", img_dense),
    ("10", "長條圖", "圖10：待機功耗比較", img_chart),
    ("11", "流程圖", "圖11：Qualcomm 開機流程", img_flow),
    ("12", "紅框箭頭標註", "圖12：issue 清單（紅框為主管標註）", img_annot),
]

QUESTIONS = {
    "01": [("Android 15 的 T11 CTS fail 數是多少？", "14", C("T11", 3, "14")),
           ("Android 15 的 T12-Pro VTS fail 數是多少？", "5", C("T12-Pro", 4, "5")),
           ("Android 14 的 T10 VTS fail 數是多少？", "3", C("T10", 2, "3"))],
    "02": [("Idle 情境的平均電流是多少 mA？", "145.7 mA", C("Idle", 1, "145.7")),
           ("Video 情境持續多久（秒）？", "1800 秒", C("Video", 2, "1800")),
           ("Camera 情境的平均電流是多少 mA？", "967.2 mA", C("Camera", 1, "967.2"))],
    "03": [("smoke test 中 T10 哪一項 FAIL？", "Camera",
            C("Camera", 1, "FAIL", "✗", "X", "NG", "紅")),
           ("smoke test 中 T11 哪一項 FAIL？", "Suspend/Resume",
            C("Suspend/Resume", 2, "FAIL", "✗", "X", "NG", "紅"))],
    "04": [("Suspend 電流實測值是多少？", "3.8 µA", C("Suspend 電流", 2, "3.8µA", "3.8μA", "3.8uA")),
           ("充電效率驗證結果是通過還是不通過？", "不通過（✗，83.2% < 85%）",
            C("充電效率", 3, "✗", "X", "x", "×", "FAIL")),
           ("電池實測容量是多少？", "4,987 mAh", C("電池容量", 2, "4,987mAh")),
           ("溫升實測值是多少？", "12.5 ℃", C("溫升", 2, "12.5℃", "12.5°C"))],
    "05": [("序號 1 的料號是什麼？", "B01-O8（O 是英文字母）", C("1", 1, "B01-O8")),
           ("序號 2 的批號是什麼？", "O0-Q9", C("2", 2, "O0-Q9")),
           ("序號 2 的最低溫度是多少？", "-12.5 ℃", C("2", 3, "-12.5")),
           ("序號 3 的批號是什麼？", "I1l-60（大寫 I、數字 1、小寫 l）", C("3", 2, "I1l-60"))],
    "06": [("SMT 站的良率是多少？", "98.72%", C("SMT", 2, "98.72%")),
           ("FATP 站投入多少台？", "1232", C("FATP", 1, "1232")),
           ("OQC 站的良率是多少？", "99.40%", C("OQC", 2, "99.40%"))],
    "07": [("v1.1.0 的 Build ID 是什麼？", "UKQ1.261002.003", C("v1.1.0", 2, "UKQ1.261002.003")),
           ("v1.2.0 是哪一天發布的？", "2026-11-05", C("v1.2.0", 1, "2026-11-05"))],
    "08": [("T11 的 ROM 容量是多少？", "128GB", C("T11", 2, "128GB")),
           ("T12-Pro 的售價是多少？", "NT$15,990", C("T12-Pro", 3, "15,990")),
           ("T10 的 RAM 是多少？", "4GB", C("T10", 1, "4GB"))],
    "09": [("TC-013 在 T11 的數值是多少 ms？", dense_value(13, 3), C("TC-013", 3, dense_value(13, 3))),
           ("TC-017 在 T13 的數值是多少 ms？", dense_value(17, 5), C("TC-017", 5, dense_value(17, 5))),
           ("TC-020 在 T15 的數值是多少 ms？", dense_value(20, 7), C("TC-020", 7, dense_value(20, 7)))],
    "10": [("T11 的待機功耗大約是多少 mW？", "約 35 mW", T("contains", "35")),
           ("T12-Pro 的待機功耗大約是多少 mW？", "約 58 mW", T("contains", "58")),
           ("T10 的待機功耗大約是多少 mW？", "約 42 mW", T("contains", "42"))],
    "11": [("開機流程中，Kernel 的前一個階段是什麼？", "ABL/UEFI",
            T("order", "Power Key", "PBL", "XBL", "ABL", "Kernel", "init")),
           ("XBL 開機失敗時會進入什麼模式？", "EDL mode", T("contains", "EDL"))],
    "12": [("issue 清單中被紅框標註的是哪一項？", "Camera 黑屏",
            T("same_line", "Camera", "紅框")),
           ("標註「需優先處理」的是哪一項？", "Camera 黑屏（Critical，BSP 負責）",
            T("same_line", "Camera", "需優先處理"))],
}


def build(out_dir: Path):
    img_dir = out_dir / "images"
    img_dir.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Level2"
    ws["A1"] = "Level 2 圖片辨識測試：12 種圖片類型"
    row, exam = 3, []
    for img_id, cat, caption, fn in IMAGES:
        if fn is None:
            path = img_lowres(img_dir)
        else:
            path = img_dir / f"{img_id}.png"
            fn().save(path)
        ws.cell(row=row, column=1, value=caption)
        xl = XLImage(path)
        ws.add_image(xl, f"A{row + 1}")
        row += math.ceil(xl.height / 20) + 3          # 預設列高約 20px
        for q, a, check in QUESTIONS[img_id]:
            exam.append({"id": len(exam) + 1, "image": img_id, "category": cat,
                         "question": q, "answer": a, "source": caption.split("：")[0],
                         "check": check})
    wb.save(out_dir / "level2.xlsx")
    (out_dir / "level2_exam.json").write_text(
        json.dumps(exam, ensure_ascii=False, indent=2), encoding="utf-8")
    with open(out_dir / "level2_exam.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["id", "category", "source", "question", "answer"])
        for e in exam:
            w.writerow([e["id"], e["category"], e["source"], e["question"], e["answer"]])
    print(f"ok: {out_dir / 'level2.xlsx'}（{len(IMAGES)} 張圖、{len(exam)} 題）")


if __name__ == "__main__":
    random.seed(0)
    build(Path(sys.argv[1]))

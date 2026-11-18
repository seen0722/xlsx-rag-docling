#!/usr/bin/env python3
"""產生已知答案的測試 xlsx：儲存格文字 + 兩張嵌入的表格圖片。"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageFont

FONT = "/System/Library/Fonts/STHeiti Medium.ttc"
OUT = Path(sys.argv[1])
OUT.mkdir(parents=True, exist_ok=True)

TABLE1 = [["機型", "影片播放(hr)", "網頁瀏覽(hr)", "待機(hr)"],
          ["T10", "12.5", "9.8", "410"],
          ["T11", "13.2", "10.4", "438"],
          ["T12-Pro", "11.7", "8.9", "385"]]
TABLE2 = [["機型", "0-50% (min)", "0-100% (min)"],
          ["T10", "27.5", "96.3"],
          ["T11", "31.2", "104.6"],
          ["T12-Pro", "24.8", "88.1"]]


def render_table(rows, path, col_w=190, row_h=48):
    font = ImageFont.truetype(FONT, 24)
    w, h = col_w * len(rows[0]) + 2, row_h * len(rows) + 2
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    for r, row in enumerate(rows):
        if r == 0:
            d.rectangle([0, 0, w, row_h], fill="#dde6f0")
        for c, val in enumerate(row):
            x, y = c * col_w, r * row_h
            d.rectangle([x, y, x + col_w, y + row_h], outline="black", width=2)
            d.text((x + 12, y + 10), val, fill="black", font=font)
    img.save(path)


render_table(TABLE1, OUT / "table1.png")
render_table(TABLE2, OUT / "table2.png")

wb = Workbook()
ws = wb.active
ws.title = "電池測試"
ws["A1"] = "平板電池續航測試報告"
ws["A3"] = "表1：各機型續航時間，測試條件：螢幕亮度 200 nits、Wi-Fi 開啟"
ws.add_image(XLImage(OUT / "table1.png"), "A4")
ws["A14"] = "結論：T11 影片播放續航最長；T12-Pro 待機最短，需追查 suspend 耗電。"
ws["A16"] = "表2：充電時間，使用原廠 45W 充電器"
ws.add_image(XLImage(OUT / "table2.png"), "A17")
ws["A27"] = "備註：T12-Pro 充電最快，但續航最短。"
wb.save(OUT / "battery_test.xlsx")

# 題目與標準答案（答案只存在於圖片中的，標記 in_image）
questions = [
    {"q": "T11 的待機時間是多少？", "answer": "438", "in_image": True},
    {"q": "T12-Pro 充到 100% 要幾分鐘？", "answer": "88.1", "in_image": True},
    {"q": "T10 網頁瀏覽續航？", "answer": "9.8", "in_image": True},
    {"q": "表1 的測試亮度？", "answer": "200 nits", "in_image": False},
    {"q": "表2 用什麼充電器？", "answer": "45W", "in_image": False},
]
(OUT / "questions.json").write_text(
    json.dumps(questions, ensure_ascii=False, indent=2), encoding="utf-8")
print("ok:", OUT / "battery_test.xlsx")

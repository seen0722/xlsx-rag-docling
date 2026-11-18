#!/usr/bin/env python3
"""xlsx → md（Docling 混合做法）
1. Docling 轉 xlsx：取得儲存格文字與圖片位置（<!-- image --> placeholder）
2. 每張圖片再丟給 Docling 的 image pipeline（OCR + TableFormer）轉成 md 表格
3. 依順序把圖片的 md 填回 placeholder

用法: python -I xlsx2md_docling.py input.xlsx output.md
"""
import sys
import tempfile
from pathlib import Path

from docling.datamodel.base_models import InputFormat
from docling.datamodel.pipeline_options import OcrMacOptions, PdfPipelineOptions
from docling.document_converter import DocumentConverter, ImageFormatOption
from docling_core.types.doc import ImageRefMode

PLACEHOLDER = "<!-- image -->"


def build_image_converter(full_page_ocr: bool = False) -> DocumentConverter:
    opts = PdfPipelineOptions(do_ocr=True, do_table_structure=True)
    # macOS Vision OCR；full_page_ocr 用於圖片被判定為「圖形」時，強制擷取文字
    opts.ocr_options = OcrMacOptions(lang=["zh-Hant", "en-US"],
                                     force_full_page_ocr=full_page_ocr)
    return DocumentConverter(
        format_options={InputFormat.IMAGE: ImageFormatOption(pipeline_options=opts)})


def raw_ocr_lines(path: Path, line_tol: float = 0.02) -> str:
    """直接用 macOS Vision OCR，依位置由上而下、由左而右輸出文字（同一行以 ' | ' 分隔）。"""
    from ocrmac import ocrmac
    items = ocrmac.OCR(str(path), language_preference=["zh-Hant", "en-US"]).recognize()
    # bbox = (x, y, w, h)，y 原點在下方 → 取 1 - y 讓上方排前面
    items = sorted(items, key=lambda t: (round((1 - t[2][1]) / line_tol), t[2][0]))
    lines, cur, cur_key = [], [], None
    for text, _conf, (x, y, w, h) in items:
        key = round((1 - y) / line_tol)
        if cur and key != cur_key:
            lines.append(" | ".join(cur))
            cur = []
        cur.append(text)
        cur_key = key
    if cur:
        lines.append(" | ".join(cur))
    return "\n".join(lines)


def _convert(conv: DocumentConverter, path: Path) -> str:
    md = conv.convert(path).document.export_to_markdown(image_mode=ImageRefMode.PLACEHOLDER)
    return md.replace(PLACEHOLDER, "").strip()   # 移除巢狀 placeholder，避免回填時錯位


def picture_to_md(convs: tuple, pil_img, tmpdir: Path, idx: int) -> str:
    path = tmpdir / f"pic_{idx}.png"
    pil_img.save(path)
    try:
        body = _convert(convs[0], path)
        if not body:  # 被判定為圖形（流程圖、圖表）：Docling 不匯出圖形內文字，改直接 OCR
            body = raw_ocr_lines(path)
            if body:
                body = f"> ℹ️ 圖片 {idx} 非表格，以下為 OCR 文字（版面關係可能遺失）\n\n{body}"
        return body or f"> ⚠️ 圖片 {idx} 未辨識出內容"
    except Exception as e:  # 明確標記失敗，不靜默跳過
        return f"> ⚠️ 圖片 {idx} 轉換失敗：{e}"


def main(src: str, dst: str) -> None:
    doc = DocumentConverter().convert(src).document
    md = doc.export_to_markdown(image_mode=ImageRefMode.PLACEHOLDER)

    pictures = list(doc.pictures)
    chunks = md.split(PLACEHOLDER)   # 先切成固定片段，不在修改中的字串上重複搜尋
    if len(chunks) - 1 != len(pictures):
        sys.exit(f"placeholder 數 {len(chunks) - 1} ≠ 圖片數 {len(pictures)}，停止以免錯位")

    convs = (build_image_converter(), build_image_converter(full_page_ocr=True))
    out = [chunks[0]]
    with tempfile.TemporaryDirectory() as td:
        for i, pic in enumerate(pictures, 1):
            pil = pic.get_image(doc)
            body = (picture_to_md(convs, pil, Path(td), i) if pil
                    else f"> ⚠️ 圖片 {i} 無法取出")
            out += [f"<!-- image {i} -->\n{body}", chunks[i]]

    Path(dst).write_text("".join(out), encoding="utf-8")
    print(f"done: {dst}（{len(pictures)} 張圖片）")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

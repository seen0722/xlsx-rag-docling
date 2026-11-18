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


def build_image_converter() -> DocumentConverter:
    opts = PdfPipelineOptions(do_ocr=True, do_table_structure=True)
    opts.ocr_options = OcrMacOptions(lang=["zh-Hant", "en-US"])  # macOS Vision OCR
    return DocumentConverter(
        format_options={InputFormat.IMAGE: ImageFormatOption(pipeline_options=opts)})


def picture_to_md(img_conv: DocumentConverter, pil_img, tmpdir: Path, idx: int) -> str:
    path = tmpdir / f"pic_{idx}.png"
    pil_img.save(path)
    try:
        md = img_conv.convert(path).document.export_to_markdown()
        return md.strip() or f"> ⚠️ 圖片 {idx} 未辨識出內容"
    except Exception as e:  # 明確標記失敗，不靜默跳過
        return f"> ⚠️ 圖片 {idx} 轉換失敗：{e}"


def main(src: str, dst: str) -> None:
    doc = DocumentConverter().convert(src).document
    md = doc.export_to_markdown(image_mode=ImageRefMode.PLACEHOLDER)

    pictures = list(doc.pictures)
    if md.count(PLACEHOLDER) != len(pictures):
        sys.exit(f"placeholder 數 {md.count(PLACEHOLDER)} ≠ 圖片數 {len(pictures)}，停止以免錯位")

    img_conv = build_image_converter()
    with tempfile.TemporaryDirectory() as td:
        for i, pic in enumerate(pictures, 1):
            pil = pic.get_image(doc)
            body = (picture_to_md(img_conv, pil, Path(td), i) if pil
                    else f"> ⚠️ 圖片 {i} 無法取出")
            md = md.replace(PLACEHOLDER, f"<!-- image {i} -->\n{body}", 1)

    Path(dst).write_text(md, encoding="utf-8")
    print(f"done: {dst}（{len(pictures)} 張圖片）")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])

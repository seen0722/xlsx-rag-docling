# xlsx-rag-docling

Convert Excel files that contain **embedded table images** into Markdown for RAG, using [Docling](https://github.com/docling-project/docling), and test the result with a question set whose answers are known in advance.

## Why

Many enterprise spreadsheets paste tables as screenshots next to the text cells that explain them. A plain xlsx → md conversion loses those tables, so a RAG system can't answer questions about their numbers.

Test results on the included sample file (`testcase/battery_test.xlsx`):

| Approach | Score | Notes |
|---|---|---|
| Docling default | 2/6 | Each image becomes `<!-- image -->`, in the right place |
| Docling `--ocr` | 2/6 | Ignored for xlsx, with no error |
| Docling `--enrich-picture-description` | 2/6 | Also ignored for xlsx |
| **Hybrid (`xlsx2md_docling.py`)** | **6/6** | All 15 numbers in both image tables are correct and in the right place |

**Root cause:** xlsx goes through Docling's `SimplePipeline`. The OCR and table-structure models only run in the PDF and image pipelines, so the CLI flags above have no effect on xlsx.

**Hybrid approach:**
1. Convert the xlsx with Docling. This gives the cell text, with a placeholder where each image sits.
2. Run each embedded image through Docling's image pipeline (OCR + TableFormer) to get a Markdown table.
3. Put each table back into its placeholder, in order.

## Files

| File | Purpose |
|---|---|
| `xlsx2md_docling.py` | Hybrid converter: xlsx → md, with image tables filled in |
| `make_testcase.py` | Builds the sample xlsx (text cells + 2 table images) |
| `grade.py` | Scores an md file against `questions.json` (5 questions + 1 ordering check) |
| `make_exam.py` | Builds the 20-question RAG exam (`rag_exam.json` / `rag_exam.csv`) |
| `testcase/` | Sample xlsx, the source images, and the question sets |
| `output/battery_test.md` | Hybrid converter output (6/6) |

## Usage

```bash
uv venv --python 3.12 .venv
uv pip install --python .venv/bin/python -r requirements.txt

# Convert
.venv/bin/python xlsx2md_docling.py testcase/battery_test.xlsx out.md

# Score
.venv/bin/python grade.py out.md testcase/questions.json

# Rebuild the sample file and the exam
.venv/bin/python make_testcase.py testcase
.venv/bin/python make_exam.py testcase
```

## RAG exam (20 questions)

`testcase/rag_exam.csv` (opens in Excel) and `testcase/rag_exam.json` (for automated scoring) contain 20 questions in five categories:

| Category | Count | What it tests |
|---|---|---|
| Image lookup | 6 | Did the image tables get converted to text and indexed? |
| Text lookup | 3 | Basic retrieval of cell text |
| Text + image | 4 | Does the explanatory text stay in the same chunk as its table? |
| Calculation | 4 | Can the system collect several numbers and compute or rank them? |
| Refusal | 3 | Does the system say "not in the document" instead of making up an answer? |

The `keywords` field supports automated scoring: an answer counts as correct only if it contains every keyword. Refusal answers vary in wording, so score those by hand or with an LLM judge.

## Caveats

- **macOS only as written:** the OCR engine is `ocrmac` (Apple Vision), and `make_testcase.py` uses a macOS CJK font. On Linux, use `EasyOcrOptions` or `RapidOcrOptions` and check how well they handle Traditional Chinese.
- **Ideal test images:** the sample images are clean renders. Real screenshots (low resolution, compression artifacts, merged cells, missing gridlines) will score lower.
- **OCR changes characters:** half-width `(hr)` came back as full-width `（hr）`. Normalize before indexing if you rely on keyword search.
- **Noise:** Docling wraps each text cell in a single-cell table. Consider turning these back into plain paragraphs to save tokens.

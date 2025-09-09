# pdf-epub-to-xml (Windows)

Batch convert PDF & EPUB to normalized XML.

## Run
1. Python 3.10+
2. Install: `pip install -r requirements.txt`
3. Put files in `input/`, then: `python convert_both.py`
4. See results in `output/` + `batch_report.csv`

## Notes
- Requires Poppler (`pdftohtml`) on PATH (Windows).
- Scanned PDFs are flagged as `NEEDS_OCR`.

# Third-Party Notices — PDF Tables to Excel

This product builds on open-source software. Versions are the exact pins from
`install.bat` (the only runtime dependency list of this product). Every
license type below was checked against the installed package's own metadata
and license files in this session's test environment (2026-08-26) — not
copied from memory.

## Runtime dependencies

| Package | Version | License | Use |
|---|---|---|---|
| camelot-py | 0.11.0 | MIT | The table-extraction engine (lattice + stream modes) |
| opencv-python-headless | 4.10.0.84 | Apache-2.0 | Image handling for table detection |
| openpyxl | 3.1.5 | MIT | Writing the .xlsx spreadsheets |
| pypdfium2 | 4.30.0 | (Apache-2.0 OR BSD-3-Clause) AND PdfiumThirdParty | PDF page counting + rasterizing pages for OCR |
| pytesseract | 0.3.13 | Apache-2.0 | Python wrapper for the Tesseract OCR engine |
| Pillow | 10.4.0 | HPND (historical permissive notice) | Image format support for OCR rendering |

Short attributions:

- **camelot-py** — MIT License. Copyright (c) 2019-2021 Camelot Developers.
  Full text: `licenses/camelot-py-0.11.0-LICENSE.txt`.
- **opencv-python-headless** — the OpenCV core is Apache-2.0 (per the
  package's own metadata). Full texts as shipped in the wheel:
  `licenses/opencv-python-headless-4.10.0.84-LICENSE.txt` (packaging notice)
  and the wheel's `LICENSE-3RD-PARTY.txt` for OpenCV's bundled third-party
  components.
- **pytesseract** — Apache License 2.0.
  Full text: `licenses/pytesseract-0.3.13-LICENSE.txt`.
- **Pillow** — HPND-style historical permission notice (the license Pillow
  has always shipped under).
  Full text: `licenses/Pillow-10.4.0-LICENSE.txt`.
- **openpyxl** — MIT per the package's own metadata. This wheel ships no
  license text file, so none could be vendored here; the authoritative text
  lives upstream with the openpyxl project (foss.heptapod.net/openpyxl).
- **pypdfium2** — dual Apache-2.0 / BSD-3-Clause for the Python bindings,
  which bundle PDFium; PDFium carries its own third-party notices.
  Shipped verbatim: `licenses/pypdfium2-4.30.0-PdfiumThirdParty.txt`.

## External tools (optional, free, NOT bundled)

- **Tesseract OCR** — Apache-2.0. Needed only for the optional scanned-page
  fallback (`--ocr`). Installed separately by the user; this product never
  bundles or redistributes it.
- **Ghostscript** — AGPL-3.0. Needed only for the lined-grid ("lattice")
  reading mode. Installed separately by the user; this product never bundles
  or redistributes it.

## Development-only (never part of the shipped product)

- **fpdf2** 2.8.8 — LGPL-3.0 (checked in the installed wheel). Used ONLY by
  `scripts/selftest_pdf_tables.py` to generate sample PDFs when running the
  self-test. It is not installed by `install.bat`, does not ship with the
  product, and nothing in the runtime depends on it.

If you are a rights holder and believe something above misstates your terms,
the mistake will be fixed on notice.

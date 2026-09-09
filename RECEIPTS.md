# PDF Tables to Excel 1.0.0 — Verification Receipts

All commands run on **2026-08-26**, Windows, from `products/pdf-tables-excel/`.
Interpreter: CPython **3.12.13** (via uv; the installer pins 3.11.9 — both are
inside the supported 3.10+ range stated on `tables.py`). Dependencies were
installed at the exact pins from `install.bat`:

```
uv venv .venv --python 3.12
uv pip install --python .venv\Scripts\python.exe camelot-py==0.11.0 opencv-python-headless==4.10.0.84 openpyxl==3.1.5 pypdfium2==4.30.0 pytesseract==0.3.13 pillow==10.4.0 fpdf2==2.8.8
```

(`fpdf2` is dev-only — see "Why generated sample PDFs" below.) Every
transcript below is pasted from a real run this session. QA can rerun:

```
.venv\Scripts\python.exe scripts\selftest_pdf_tables.py
```

---

## 1) Why generated sample PDFs (tooling choice, honestly)

The self-test needs PDFs whose table contents are KNOWN in advance, so cell
equality can be asserted. Two options were considered: writing raw PDF bytes
by hand, or generating them with a small library. Choice: **fpdf2 pinned in
`requirements-dev.txt`** (a dev/selftest-only block). Raw PDF construction is
error-prone in ways that would test our writer instead of the converter;
fpdf2 produces standard, well-formed pages and keeps the test honest.

License correction found while pinning: the wheel ships **LGPL-3.0**, not MIT
as first assumed. It stays out of the shipped product entirely — it is not in
`install.bat`, only in `requirements-dev.txt`, used solely by the self-test
at development time. Recorded in `THIRD-PARTY-NOTICES.md`.

## 2) Self-test — full transcript (10/10 PASS)

Command:

```
.venv\Scripts\python.exe scripts\selftest_pdf_tables.py
```

Observed (verbatim):

```
[info] Tesseract found: C:\Program Files\Tesseract-OCR
[info] workspace: C:\Users\me\AppData\Local\Temp\pdftables_selftest_kmj8lmkw
$ python.exe tables.py lined.pdf -o lined.xlsx
[run] lined.pdf
  [warn] Ghostscript is missing (needed for the lined-grid mode). Install it free: winget install ArtifexSoftware.GhostScript - or rerun with --mode stream.
  [done] 1 table(s) -> lined.xlsx (HIGH 1 / MEDIUM 0 / LOW 0) in 0.7s

Summary: 1 file(s) converted, 1 table(s) total, 0 failed.
[PASS] lined-table conversion exit code 0
[PASS] lined table matches known grid cell-by-cell
[PASS] lined table carries an honest grade label - Spacing guess, HIGH
[info] lined.pdf engine=Spacing guess grade=HIGH accuracy=100
$ python.exe tables.py spaced.pdf -o spaced.xlsx --mode stream
[run] spaced.pdf
  [done] 1 table(s) -> spaced.xlsx (HIGH 1 / MEDIUM 0 / LOW 0) in 0.7s

Summary: 1 file(s) converted, 1 table(s) total, 0 failed.
[PASS] stream-mode conversion exit code 0
[PASS] lineless table matches known grid cell-by-cell
[info] spaced.pdf engine=Spacing guess grade=HIGH accuracy=100  (stream CAN score HIGH >= 90 on clean layouts)
$ python.exe tables.py scan.pdf -o scan.xlsx --ocr
[run] scan.pdf
  page 1: OCR fallback used (graded LOW, always proofread)
  [done] 1 table(s) -> scan.xlsx (HIGH 0 / MEDIUM 0 / LOW 1) in 1.0s

Summary: 1 file(s) converted, 1 table(s) total, 0 failed.
[PASS] OCR conversion exit code 0
[PASS] OCR path produced results on the image-only page - 1 OCR table(s)
[PASS] EVERY OCR result is labeled LOW (never silently trusted) - LOW
[PASS] no OCR result claims HIGH/MEDIUM confidence
[PASS] OCR recovered most known tokens (content sanity) - 100% of expected tokens; missing=[]
============================================================
SELFTEST PASSED: 10/10 checks.
(Sample PDFs live only in the temp workspace; rerun regenerates them.)
===EXIT:0===
```

What each case proves:

- **lined.pdf** — a table drawn with ruling lines, converted through the
  real product CLI (`python tables.py lined.pdf`). The resulting .xlsx was
  opened with openpyxl and every cell equals the known source grid. This is
  also the "real-world-ish conversion" receipt: an invoice-style table going
  in, a graded spreadsheet coming out.
- **spaced.pdf** — no lines at all, columns implied purely by spacing,
  forced through `--mode stream`. Cell-by-cell exact match, observed
  accuracy **100 → HIGH**. This is the receipt behind the README claim that
  stream mode CAN reach HIGH (>=90) on clean layouts.
- **scan.pdf** — an image-ONLY page (the table exists as pixels; zero text
  objects), run with `--ocr`. Tesseract OCR really executed, produced rows
  covering 100% of the expected tokens, and every resulting table carries
  the **LOW** label with no accuracy number claimed. The product never
  presents OCR output as trusted.

## 3) Environment notes (so nothing above oversells)

- **Tesseract** is installed on this machine but OFF the default PATH
  (`C:\Program Files\Tesseract-OCR`). The self-test prepends that folder to
  the child process PATH when present — the same situation as a buyer who
  installed it normally. No OCR results were fabricated; the transcript
  shows the converter's own log line for the OCR fallback.
- **Ghostscript is NOT installed here**, so the lattice ("lined grid") mode
  could NOT be exercised end to end in this session. The transcript shows
  the converter handling that absence gracefully: it prints its friendly
  warning and the spacing engine still recovered the table exactly. Lattice
  end-to-end verification remains open until a machine with Ghostscript runs
  the same self-test (the self-test accepts whichever engine fires and
  reports it truthfully).

## 4) Copy files aligned this session

- README price $39 → **$29** everywhere; competitor section replaced with a
  hedged category line pointing to `data/research/verification-queue-2026-08-26.md`.
- Refund stance replaced: **14 days, no forms, via the store's message button**.
- Grade table rewritten so stream-mode HIGH (>=90) is stated truthfully.
- `STORE_LISTING.md` written (disclosure-first; FAQ covers Ghostscript,
  privacy/offline, OCR honesty, best-fit inputs, refunds; $29 launch /
  planned $49 anchor; no competitor brand names).
- `THIRD-PARTY-NOTICES.md` written; license texts vendored into `licenses/`
  from the exact wheels installed above.

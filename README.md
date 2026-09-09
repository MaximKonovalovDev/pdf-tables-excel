# PDF Tables to Excel

Turn the tables inside any PDF into clean Excel spreadsheets — on your own
PC, in seconds, even for a whole folder of files at once. Offline, no
account, no telemetry.

Bank statements, invoices, supplier price lists, and reports arrive as PDFs
with the numbers locked inside. This tool gets them out: two extraction
engines (ruling-line tables first, text-spacing columns second), an optional
OCR fallback for scanned pages, batch mode, a preview page before you trust
anything, and an honest confidence grade on every table.

## What it does

- **Reads tables from any PDF** using two engines: one that follows the
  ruling lines of drawn tables, one that figures out columns from text
  spacing when there are no lines. On clean layouts the spacing engine
  scores 90+ and lands a HIGH grade — see the self-test receipts in
  `RECEIPTS.md`.
- **Optional OCR fallback** for scanned pages (off by default, needs the
  free Tesseract program).
- **Batch mode**: point it at a folder, get one spreadsheet per PDF.
- **Preview before you trust it**: every table found lands in a
  `preview.html` you can eyeball in a browser.
- **An honest grade on every table** — HIGH (usually perfect), MEDIUM
  (spot-check merged cells), LOW (proofread every number). Grades appear
  while it runs, in the preview page, and on a Summary sheet inside each
  spreadsheet.

## Your files never leave this computer

No upload. No cloud. No account. No telemetry. It works with the internet
cable unplugged — try it.

## Quick start (Windows)

1. Double-click **install.bat** once. It sets up everything the product
   needs, in its own folder, touching nothing else on your system.
2. Start it:
   - Desktop window: `.venv\Scripts\python tables.py`
   - Or straight from the command line:

```text
:: one file -> invoice.xlsx + invoice.preview.html
python tables.py invoice.pdf

:: a whole folder of statements -> spreadsheets in .\done
python tables.py "C:\statements" --out-dir done

:: scanned pages too (needs free Tesseract)
python tables.py scan.pdf --ocr

:: only pages 2-4 of a fat report
python tables.py report.pdf --pages 2-4
```

## Verification

`RECEIPTS.md` holds the full self-test transcript: **10/10 checks passed**
on generated sample PDFs with known contents — lined tables, lineless
spacing tables, and OCR pages — including the open note that the
Ghostscript-dependent lattice path was not exercised on that machine.

## Honest limits

- **Garbage in, garbage out.** A blurry phone photo of a crumpled receipt
  will not turn into a clean spreadsheet. Neither will handwriting.
- Scanned PDFs need the OCR fallback (free Tesseract, one command during
  install). OCR output is always graded LOW — proofread it.
- Password-protected PDFs must be unlocked first.
- Very unusual layouts (tables split across two pages, nested headers) may
  need a quick manual touch-up in Excel. The grades tell you exactly
  which ones.
- Windows 10/11 recommended (macOS/Linux work via the same Python
  commands).

## What's in the box

- `tables.py` — the whole product: desktop window + command line
- `install.bat` — one-time installer with exact pinned versions
- `examples/README.txt` — five-minute starter guide
- `RECEIPTS.md` — the verification transcript
- `THIRD-PARTY-NOTICES.md` — dependency licenses

## License

MIT — see `LICENSE`.

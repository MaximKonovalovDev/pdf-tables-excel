# PDF Tables -> Excel extractor

> Portfolio showcase extracted from private studio work. Measured receipts inside.

# PDF Tables to Excel

**$29 once. Yours forever. No subscription.**

Turn the tables inside any PDF into clean Excel spreadsheets — on your own PC,
in seconds, even for a whole folder of files at once.

---

## The problem you know too well

Bank statements, invoices, supplier price lists, and reports arrive as PDFs.
The numbers are in there — locked. So you either retype them by hand, or you
pay a monthly fee to an online converter and hand your documents to a stranger's
server.

There is a third option.

## What it does

- **Reads tables from any PDF** using two engines: one that follows the
  ruling lines of drawn tables (the first one to try when a table has
  lines), and one that figures out columns from text spacing when there
  are no lines. On clean layouts the spacing engine scores 90+ and lands
  a HIGH grade - see the selftest receipts in this folder.
- **Optional OCR fallback** for scanned pages (paper that was photographed or
  scanned as an image). One checkbox, off by default.
- **Batch mode**: point it at a folder, get one spreadsheet per PDF.
- **A preview page before you trust it**: every table it found lands in a
  `preview.html` you can open in your browser and eyeball.
- **An honest grade on every table** — see below. No tool on earth gets every
  table right. The difference is whether the tool tells you which ones to check.
  This one does.
- **Desktop window or command line**, whichever you prefer. Same engine inside.

## Every table comes with a grade

| Grade | How it was read | What to do |
|-------|-----------------|------------|
| **HIGH** | Straight from the table's own grid lines, or from clean text spacing scoring 90 or more | Usually perfect |
| **MEDIUM** | Column layout inferred from text spacing, score 70-89 | Spot-check merged cells |
| **LOW** | A rough guess, or scanned-page OCR | Proofread every number |

The grades appear while it runs, in the preview page, and on a Summary sheet
inside each spreadsheet. If something looks wrong, it will not pretend otherwise.

## Your files never leave this computer

No upload. No cloud. No account. No telemetry. The app works with the internet
cable unplugged — try it. Your bank statements and invoices stay where they
belong: with you.

## $29 once vs. monthly subscriptions

Popular web converters run monthly subscriptions — roughly $5-30/mo for one
person, which adds up year after year for reading PDFs (vendor pricing pages
checked 2026-08-26; sources listed in the repo research file
`data/research/verification-queue-2026-08-26.md`). This is one payment of
**$29**, permanent updates to the 1.x series included, on every PC you
personally use.

## Quick start (Windows)

1. Double-click **install.bat** once. It sets up everything the product needs,
   in its own folder, touching nothing else on your system.
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

## Honest limits

No sales pitch survives contact with every PDF, so here is the truth up front:

- **Garbage in, garbage out.** A blurry phone photo of a crumpled receipt will
  not turn into a clean spreadsheet. Neither will handwriting.
- Scanned PDFs need the OCR fallback, which needs the free Tesseract program
  (one command, listed during install). OCR output is always graded LOW —
  proofread it.
- Password-protected PDFs must be unlocked first.
- Very unusual layouts (tables split across two pages, nested headers) may need
  a quick manual touch-up in Excel. The grades tell you exactly which ones.
- Windows 10/11 recommended (macOS/Linux work via the same Python commands).

## What's in the box

- `tables.py` — the whole product: desktop window + command line
- `install.bat` — one-time installer with exact pinned versions
- `examples/README.txt` — five-minute starter guide

## 14-day guarantee

Use it for two weeks. If it doesn't save you hours of retyping, send a message
through the store's message button and get every cent back. No forms,
no interrogation.


#!/usr/bin/env python3
"""
Selftest for PDF Tables to Excel (tables.py).

What it does - no shortcuts:

1. Generates small sample PDFs with KNOWN tables using fpdf2
   (dev-only dependency, see requirements-dev.txt).
2. Runs the real product command line (`python tables.py ...`) as a
   subprocess, exactly like a buyer would.
3. Opens the produced .xlsx files with openpyxl and asserts the sheet
   cells EQUAL the known grids, cell by cell.
4. Exercises the OCR path on an image-only page (a picture of a table,
   no text layer) and asserts every OCR result is honestly labeled LOW -
   the product must never pass OCR guesses off as reliable.

Run it from the product folder:

    .venv\\Scripts\\python.exe scripts\\selftest_pdf_tables.py

Exit code 0 = all PASS lines printed and nothing failed.
"""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

PRODUCT_ROOT = Path(__file__).resolve().parent.parent

# Known grid 1: drawn WITH ruling lines (lattice-style table).
GRID_LINED = [
    ["SKU", "QTY", "PRICE"],
    ["A1001", "3", "19.99"],
    ["B2002", "12", "4.50"],
    ["C3003", "7", "110.00"],
]

# Known grid 2: NO lines at all; columns only implied by spacing.
GRID_SPACED = [
    ["ID", "CITY", "TOTAL"],
    ["N21", "PARIS", "88.10"],
    ["N22", "MADRID", "129.99"],
    ["N23", "OSLO", "64.75"],
]

# Tesseract install dirs a buyer (or this test PC) may have. The product
# README tells buyers to install Tesseract; here we accept the standard
# Windows locations so the OCR path can really run end to end.
TESSERACT_DIRS = [
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR"),
]


def find_python():
    """Prefer the product's own .venv interpreter."""
    venv_py = PRODUCT_ROOT / ".venv" / "Scripts" / "python.exe"
    return str(venv_py) if venv_py.exists() else sys.executable


def make_lined_pdf(path):
    """Table with visible ruling lines around every cell."""
    from fpdf import FPDF

    pdf = FPDF(format="A4")
    pdf.add_page()
    pdf.set_font("Courier", size=12)
    x0, y0 = 40.0, 40.0
    col_w = [60.0, 45.0, 45.0]     # wide gutters: content <= 36 pt per token
    row_h = 14.0
    for r, row in enumerate(GRID_LINED):
        x = x0
        y = y0 + r * row_h
        for c, cell in enumerate(row):
            pdf.rect(x, y, col_w[c], row_h)
            pdf.set_xy(x + 2, y + 3.5)
            pdf.cell(0, 0, cell)
            x += col_w[c]
    pdf.output(str(path))


def make_spaced_pdf(path):
    """Same idea, zero lines: wide even gutters carry the column layout."""
    from fpdf import FPDF

    pdf = FPDF(format="A4")
    pdf.add_page()
    pdf.set_font("Courier", size=12)
    col_x = [40.0, 105.0, 160.0]   # generous fixed gutters between columns
    y = 40.0
    for row in GRID_SPACED:
        for x, cell in zip(col_x, row):
            pdf.set_xy(x, y)
            pdf.cell(0, 0, cell)
        y += 14.0
    pdf.output(str(path))


def make_scanned_pdf(path):
    """An IMAGE-ONLY page: the table exists only as pixels, like a scan.

    Rendered with PIL, embedded by fpdf2 - the PDF carries no text objects,
    so text-based extraction has nothing to read and OCR is the only path.
    Plain text on white, no ruling boxes (a photographed price list), which
    is also what OCR engines handle most realistically.
    """
    from PIL import Image, ImageDraw, ImageFont
    from fpdf import FPDF

    img = Image.new("RGB", (1400, 380), "white")
    dr = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("consola.ttf", 34)
        font_b = ImageFont.truetype("consolab.ttf", 36)
    except OSError:
        font = font_b = ImageFont.load_default()
    xs = [80, 560, 900]
    y = 40
    for row_i, row in enumerate(GRID_LINED):
        f = font_b if row_i == 0 else font
        for x, cell in zip(xs, row):
            dr.text((x, y), cell, fill="black", font=f)
        y += 58
    png = str(path) + ".src.png"
    img.save(png)

    pdf = FPDF(unit="pt", format=(img.width, img.height))
    pdf.add_page()
    pdf.image(png, x=0, y=0, w=img.width, h=img.height)
    pdf.output(str(path))
    os.remove(png)


def run_converter(pdf_path, out_xlsx, extra_args, env):
    cmd = [find_python(), str(PRODUCT_ROOT / "tables.py"),
           str(pdf_path), "-o", str(out_xlsx)] + extra_args
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env,
                          cwd=str(PRODUCT_ROOT), timeout=600)
    print(f"$ {' '.join(os.path.basename(c) for c in cmd)}")
    if proc.stdout.strip():
        print(proc.stdout.rstrip())
    if proc.stderr.strip():
        print("[stderr]", proc.stderr.rstrip())
    return proc.returncode


def read_sheet_rows(xlsx_path):
    """Return {sheet_name: [[cell,...],...]} for every non-Summary sheet."""
    from openpyxl import load_workbook

    wb = load_workbook(str(xlsx_path), data_only=True)
    out = {}
    for name in wb.sheetnames:
        if name == "Summary":
            continue
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else str(v).strip() for v in row])
        out[name] = rows
    return out


def read_summary(xlsx_path):
    """First table row of the Summary sheet, or None when nothing was found."""
    from openpyxl import load_workbook

    ws = load_workbook(str(xlsx_path), data_only=True)["Summary"]
    rows = [["" if v is None else str(v).strip() for v in r]
            for r in ws.iter_rows(values_only=True)]
    header_idx = next((i for i, r in enumerate(rows) if r and r[0] == "Sheet"),
                      None)
    if header_idx is None:
        return None
    header = rows[header_idx]
    for r in rows[header_idx + 1:]:
        if not any(r):
            continue
        return dict(zip(header, r))
    return None


def assert_grid_exact(name, got_rows, expected):
    problems = []
    if len(got_rows) != len(expected):
        problems.append(f"{name}: row count {len(got_rows)} != {len(expected)}")
        return problems
    for r, (got, want) in enumerate(zip(got_rows, expected)):
        width = max(len(got), len(want))
        for c in range(width):
            g = got[c] if c < len(got) else ""
            w = want[c] if c < len(want) else ""
            if g != w:
                problems.append(
                    f"{name}: cell[{r}][{c}] {g!r} != expected {w!r}")
    return problems


def token_coverage(got_rows, expected):
    """Share of expected non-empty tokens that appear anywhere in output."""
    haystack = {cell for row in got_rows for cell in row if cell}
    wanted = {c for row in expected for c in row if c}
    hit = sum(1 for t in wanted if t in haystack)
    return hit / max(1, len(wanted)), sorted(wanted - haystack)


def main():
    failures = []
    passes = 0

    def check(name, cond, detail=""):
        nonlocal passes
        status = "PASS" if cond else "FAIL"
        print(f"[{status}] {name}" + (f" - {detail}" if detail else ""))
        if cond:
            passes += 1
        else:
            failures.append(name)

    env = dict(os.environ)
    tess_dir = next((d for d in TESSERACT_DIRS if os.path.isfile(
        os.path.join(d, "tesseract.exe"))), None)
    if tess_dir:
        env["PATH"] = tess_dir + os.pathsep + env.get("PATH", "")
        print(f"[info] Tesseract found: {tess_dir}")
    else:
        print("[info] Tesseract not found on this machine - "
              "the OCR case will be reported honestly below.")

    tmp = Path(tempfile.mkdtemp(prefix="pdftables_selftest_"))
    print(f"[info] workspace: {tmp}")

    # ---- Case 1: lined table -------------------------------------------
    pdf1 = tmp / "lined.pdf"
    make_lined_pdf(pdf1)
    rc1 = run_converter(pdf1, tmp / "lined.xlsx", [], env)
    check("lined-table conversion exit code 0", rc1 == 0)
    if rc1 == 0:
        sheets = read_sheet_rows(tmp / "lined.xlsx")
        summary = read_summary(tmp / "lined.xlsx")
        if not summary:
            check("lined table found", False, "Summary sheet lists no table")
        else:
            method = summary["Method"]
            grade = summary["Confidence"]
            acc = summary["Accuracy %"]
            got = sheets.get(summary["Sheet"])
            probs = assert_grid_exact("lined", got or [], GRID_LINED) if got \
                else ["lined: no table sheet found"]
            check("lined table matches known grid cell-by-cell", not probs,
                  "; ".join(probs[:4]))
            check("lined table carries an honest grade label",
                  grade in ("HIGH", "MEDIUM", "LOW"), f"{method}, {grade}")
            print(f"[info] lined.pdf engine={method} grade={grade} "
                  f"accuracy={acc}")

    # ---- Case 2: lineless (stream) table -------------------------------
    pdf2 = tmp / "spaced.pdf"
    make_spaced_pdf(pdf2)
    rc2 = run_converter(pdf2, tmp / "spaced.xlsx", ["--mode", "stream"], env)
    check("stream-mode conversion exit code 0", rc2 == 0)
    if rc2 == 0:
        sheets = read_sheet_rows(tmp / "spaced.xlsx")
        summary = read_summary(tmp / "spaced.xlsx")
        if not summary:
            check("lineless table found", False, "Summary sheet lists no table")
        else:
            method = summary["Method"]
            grade = summary["Confidence"]
            acc = summary["Accuracy %"]
            got = sheets.get(summary["Sheet"])
            probs = assert_grid_exact("spaced", got or [], GRID_SPACED) if got \
                else ["spaced: no table sheet found"]
            check("lineless table matches known grid cell-by-cell", not probs,
                  "; ".join(probs[:4]))
            print(f"[info] spaced.pdf engine={method} grade={grade} "
                  f"accuracy={acc}"
                  "  (stream CAN score HIGH >= 90 on clean layouts)")

    # ---- Case 3: scanned page through the OCR fallback ------------------
    pdf3 = tmp / "scan.pdf"
    make_scanned_pdf(pdf3)
    if tess_dir:
        rc3 = run_converter(pdf3, tmp / "scan.xlsx", ["--ocr"], env)
        check("OCR conversion exit code 0", rc3 == 0)
        if rc3 == 0:
            sheets = read_sheet_rows(tmp / "scan.xlsx")
            s = read_summary(tmp / "scan.xlsx")
            ocr_rows = [s] if (s and s["Method"] == "Scanned-page OCR") else []
            check("OCR path produced results on the image-only page",
                  bool(ocr_rows),
                  f"{len(ocr_rows)} OCR table(s)")
            check("EVERY OCR result is labeled LOW (never silently trusted)",
                  bool(ocr_rows) and all(s["Confidence"] == "LOW"
                                         for s in ocr_rows),
                  ", ".join(sorted({s['Confidence'] for s in ocr_rows}) or {"-"}))
            check("no OCR result claims HIGH/MEDIUM confidence",
                  all(s["Confidence"] not in ("HIGH", "MEDIUM")
                      for s in ocr_rows))
            got = sheets.get(ocr_rows[0]["Sheet"]) if ocr_rows else None
            cov, missing = token_coverage(got or [], GRID_LINED)
            check("OCR recovered most known tokens (content sanity)",
                  cov >= 0.8, f"{cov:.0%} of expected tokens; missing={missing}")
        else:
            check("OCR conversion exit code 0", False, f"rc={rc3}")
    else:
        check("OCR honesty case SKIPPED (no Tesseract binary)", False,
              "install free Tesseract and rerun - see README")

    # ---- wrap up ---------------------------------------------------------
    print("=" * 60)
    total = passes + len(failures)
    if failures:
        print(f"SELFTEST FAILED: {len(failures)} of {total} checks: "
              + ", ".join(failures))
        return 1
    print(f"SELFTEST PASSED: {passes}/{total} checks.")
    print("(Sample PDFs live only in the temp workspace; rerun regenerates them.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

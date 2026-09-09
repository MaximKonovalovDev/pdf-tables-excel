PDF Tables to Excel - five-minute starter
==========================================

No sample files are included (keeps the download small). Any PDF with a table
works: a bank statement, an invoice, a supplier price list, a report.

Try it in 3 steps
-----------------
1) Copy any PDF with a table into this folder and name it test.pdf.
2) Open a terminal here (type "cmd" in the folder's address bar) and run:

   ..\.venv\Scripts\python ..\tables.py test.pdf

   You get two files:
   - test.xlsx            the spreadsheet
   - test.preview.html    open it in your browser to see every table, graded

3) Read the grades (also printed while it runs):

   HIGH    read from the PDF's own ruling lines - usually perfect
   MEDIUM  figured out from text spacing - spot-check merged cells
   LOW     a rough guess or scanned-page OCR - proofread every number

More things to try
------------------
Whole folder at once (one spreadsheet per PDF):

   ..\.venv\Scripts\python ..\tables.py "C:\Users\you\Downloads\statements" --out-dir done

Scanned pages (pictures of paper) - add the OCR flag (needs free Tesseract,
see the note at the end of install.bat):

   ..\.venv\Scripts\python ..\tables.py scan.pdf --ocr

Just some pages of a big report:

   ..\.venv\Scripts\python ..\tables.py report.pdf --pages 2-4

The desktop window instead of typing:

   ..\.venv\Scripts\python ..\tables.py

Tip: check preview.html BEFORE you send the numbers anywhere. That is what
it is for.

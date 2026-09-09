@echo off
setlocal EnableExtensions
title PDF Tables to Excel - one-time installer
cd /d "%~dp0"

REM Pinned tool versions. Re-run this file any time to repair the install.
set "UV_VERSION=0.5.29"
set "PY_VERSION=3.11.9"

echo =====================================================
echo  PDF Tables to Excel - installer
echo  One-time setup. Internet is needed for this step only.
echo  The app itself runs fully offline afterwards.
echo =====================================================
echo.

set "UV=uv"
where uv >nul 2>nul
if errorlevel 1 (
  echo [1/4] Downloading uv %UV_VERSION% ^(a small helper that installs Python packages^)...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/%UV_VERSION%/install.ps1 | iex"
  if errorlevel 1 (
    echo ERROR: could not download uv. Check your internet, or install Python from
    echo        python.org and run this manually:
    echo            pip install camelot-py==0.11.0 opencv-python-headless==4.10.0.84 openpyxl==3.1.5 pypdfium2==4.30.0 pytesseract==0.3.13 pillow==10.4.0
    pause
    exit /b 1
  )
  if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV=%USERPROFILE%\.local\bin\uv.exe"
) else (
  echo [1/4] uv already installed - skipping download.
)

echo.
echo [2/4] Getting Python %PY_VERSION% (kept private inside this product folder)...
"%UV%" python install %PY_VERSION%
if errorlevel 1 goto :fail

echo.
echo [3/4] Installing the PDF engine (exact pinned versions)...
if exist ".venv" rmdir /s /q ".venv"
"%UV%" venv ".venv" --python %PY_VERSION%
if errorlevel 1 goto :fail
"%UV%" pip install --python ".venv\Scripts\python.exe" camelot-py==0.11.0 opencv-python-headless==4.10.0.84 openpyxl==3.1.5 pypdfium2==4.30.0 pytesseract==0.3.13 pillow==10.4.0
if errorlevel 1 goto :fail

echo.
echo [4/4] Checking two optional free helpers...
where gswin64c >nul 2>nul
if errorlevel 1 where gs >nul 2>nul
if errorlevel 1 (
  echo   NOTE: Ghostscript not found. Without it, the most accurate "lined grid"
  echo         mode stays off; spacing mode and OCR still work fine.
  echo         Free install:  winget install ArtifexSoftware.GhostScript
)
where tesseract >nul 2>nul
if errorlevel 1 (
  echo   NOTE: Tesseract not found. Without it, the OCR fallback for scanned
  echo         PDFs stays off. Only needed if you scan paper documents.
  echo         Free install:  winget install UB-Mannheim.TesseractOCR
)

echo.
echo DONE. To start the app:
echo     .venv\Scripts\python tables.py
echo.
pause
exit /b 0

:fail
echo.
echo ERROR: setup failed - read the messages above.
pause
exit /b 1

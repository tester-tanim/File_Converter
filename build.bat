@echo off
title File Converter - EXE Builder
color 0A

echo ============================================
echo   File Converter - EXE Builder
echo ============================================
echo.

:: ── Step 1: Check Python ──────────────────────
echo [1/6] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install from https://python.org
    pause
    exit /b 1
)
echo       OK
echo.

:: ── Step 2: Install Python dependencies ───────
echo [2/6] Installing Python dependencies...
pip install ^
    pyinstaller ^
    pandas ^
    openpyxl ^
    matplotlib ^
    pillow ^
    pypdf ^
    pdf2image ^
    img2pdf ^
    python-docx ^
    reportlab ^
    --quiet
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo       OK
echo.

:: ── Step 3: Check Poppler ─────────────────────
echo [3/6] Checking for Poppler...
if not exist "poppler" (
    echo.
    echo -----------------------------------------------
    echo  Poppler not found!
    echo.
    echo  Please do the following:
    echo  1. Download Poppler from:
    echo     https://github.com/oschwartz10612/poppler-windows/releases
    echo  2. Extract the ZIP file
    echo  3. Find the folder containing pdftoppm.exe
    echo  4. Copy that folder into this directory and rename it to: poppler
    echo.
    echo  Your folder should look like:
    echo    File_Converter\
    echo      poppler\
    echo        pdftoppm.exe
    echo        pdfinfo.exe
    echo        ...
    echo      file_converter.py
    echo      build.bat
    echo -----------------------------------------------
    echo.
    set /p CONT="Skip Poppler and build without PDF support? (y/n): "
    if /i "%CONT%" neq "y" (
        echo Build cancelled. Add poppler folder and try again.
        pause
        exit /b 1
    )
    echo       Skipping Poppler - PDF to Image will need system Poppler.
) else (
    echo       Found poppler folder OK
)
echo.

:: ── Step 4: Clean old build ───────────────────
echo [4/6] Cleaning old build files...
if exist "dist"              rmdir /s /q dist
if exist "build"             rmdir /s /q build
if exist "FileConverter.spec" del /q FileConverter.spec
echo       OK
echo.

:: ── Step 5: Build EXE ─────────────────────────
echo [5/6] Building EXE (this may take 1-3 minutes)...
echo.

if exist "poppler" (
    pyinstaller --onefile --windowed ^
        --name "FileConverter" ^
        --add-data "poppler;poppler" ^
        --hidden-import pandas ^
        --hidden-import openpyxl ^
        --hidden-import matplotlib ^
        --hidden-import matplotlib.backends.backend_agg ^
        --hidden-import PIL ^
        --hidden-import PIL.Image ^
        --hidden-import PIL.ImageOps ^
        --hidden-import pdf2image ^
        --hidden-import img2pdf ^
        --hidden-import pypdf ^
        --hidden-import docx ^
        --hidden-import reportlab ^
        --hidden-import reportlab.platypus ^
        --hidden-import reportlab.lib.styles ^
        --hidden-import reportlab.lib.pagesizes ^
        --hidden-import reportlab.lib.units ^
        --hidden-import reportlab.lib.colors ^
        --hidden-import reportlab.lib.enums ^
        --collect-submodules reportlab ^
        --collect-submodules docx ^
        file_converter.py
) else (
    pyinstaller --onefile --windowed ^
        --name "FileConverter" ^
        --hidden-import pandas ^
        --hidden-import openpyxl ^
        --hidden-import matplotlib ^
        --hidden-import matplotlib.backends.backend_agg ^
        --hidden-import PIL ^
        --hidden-import PIL.Image ^
        --hidden-import PIL.ImageOps ^
        --hidden-import pdf2image ^
        --hidden-import img2pdf ^
        --hidden-import pypdf ^
        --hidden-import docx ^
        --hidden-import reportlab ^
        --hidden-import reportlab.platypus ^
        --hidden-import reportlab.lib.styles ^
        --hidden-import reportlab.lib.pagesizes ^
        --hidden-import reportlab.lib.units ^
        --hidden-import reportlab.lib.colors ^
        --hidden-import reportlab.lib.enums ^
        --collect-submodules reportlab ^
        --collect-submodules docx ^
        file_converter.py
)

if errorlevel 1 (
    echo.
    echo ERROR: Build failed. Check the output above.
    pause
    exit /b 1
)
echo.

:: ── Step 6: Done ──────────────────────────────
echo [6/6] Done!
echo.
echo ============================================
echo   SUCCESS! Your EXE is ready:
echo   dist\FileConverter.exe
echo ============================================
echo.
echo You can now share dist\FileConverter.exe
echo with anyone - no Python needed!
echo.

:: Open the dist folder
explorer dist

pause

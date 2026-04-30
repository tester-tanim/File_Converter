@echo off
setlocal EnableDelayedExpansion
title File Converter - EXE Builder
color 0A

echo.
echo  =====================================================
echo    FILE CONVERTER ^| EXE Builder
echo  =====================================================
echo.

:: ════════════════════════════════════════════════════════
::  STEP 1 — Check Python
:: ════════════════════════════════════════════════════════
echo  [1/7]  Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERROR] Python not found!
    echo          Install Python 3.8+ from https://python.org
    echo          Make sure to check "Add Python to PATH" during install.
    echo.
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo         Found: %%v
echo.

:: ════════════════════════════════════════════════════════
::  STEP 2 — Install Python dependencies
:: ════════════════════════════════════════════════════════
echo  [2/7]  Installing Python dependencies...
echo         (this may take a minute on first run)
echo.
pip install --quiet --upgrade ^
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
    docx2pdf

if errorlevel 1 (
    echo.
    echo  [ERROR] Failed to install one or more packages.
    echo          Check your internet connection and try again.
    echo.
    pause & exit /b 1
)
echo         All packages installed OK
echo.

:: ════════════════════════════════════════════════════════
::  STEP 3 — Check DOCX to PDF converter
:: ════════════════════════════════════════════════════════
echo  [3/7]  Checking DOCX to PDF converter...

set DOCX_METHOD=fallback

:: check LibreOffice first (best quality, free)
if exist "C:\Program Files\LibreOffice\program\soffice.exe"       set DOCX_METHOD=libreoffice
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" set DOCX_METHOD=libreoffice
where soffice >nul 2>&1
if not errorlevel 1 set DOCX_METHOD=libreoffice

:: check Microsoft Word via docx2pdf (pywin32 must be importable)
if "!DOCX_METHOD!"=="fallback" (
    python -c "import win32com.client" >nul 2>&1
    if not errorlevel 1 set DOCX_METHOD=word
)

if "!DOCX_METHOD!"=="libreoffice" (
    echo         LibreOffice found - perfect DOCX to PDF conversion ^(free^)
)
if "!DOCX_METHOD!"=="word" (
    echo         Microsoft Word found - perfect DOCX to PDF via docx2pdf
)
if "!DOCX_METHOD!"=="fallback" (
    echo.
    echo         [INFO] Neither LibreOffice nor Microsoft Word detected.
    echo         docx2pdf is installed but needs Microsoft Word to run.
    echo.
    echo         DOCX to PDF will use the built-in fallback converter
    echo         ^(works for most documents, basic formatting^).
    echo.
    echo         For perfect 1:1 conversion, install one of:
    echo           - Microsoft Word ^(already have it? make sure its installed^)
    echo           - LibreOffice ^(free^): https://www.libreoffice.org
    echo.
    echo         You can continue now - the EXE will automatically use
    echo         Word or LibreOffice if installed later.
    echo.
)
echo.

:: ════════════════════════════════════════════════════════
::  STEP 4 — Check Poppler
:: ════════════════════════════════════════════════════════
echo  [4/7]  Checking Poppler (for PDF to Images)...

if exist "poppler" (
    if exist "poppler\pdftoppm.exe" (
        echo         Poppler found - PDF to Images will work perfectly
        set POPPLER_OK=1
    ) else (
        echo         [WARNING] poppler\ folder exists but pdftoppm.exe not found inside.
        echo         Make sure you copied the right folder ^(the one with pdftoppm.exe^).
        set POPPLER_OK=0
    )
) else (
    echo.
    echo         [WARNING] Poppler not found.
    echo         PDF to Images conversion will NOT work in the EXE.
    echo.
    echo         To add Poppler:
    echo           1. Download from:
    echo              https://github.com/oschwartz10612/poppler-windows/releases
    echo           2. Extract the ZIP
    echo           3. Find the folder containing pdftoppm.exe
    echo           4. Copy it here and rename it to: poppler
    echo.
    echo         Your folder should look like:
    echo           File_Converter\
    echo             poppler\
    echo               pdftoppm.exe
    echo               pdfinfo.exe  ...
    echo             file_converter.py
    echo             build.bat
    echo.
    set /p POP_SKIP="  Continue without Poppler? (y/n): "
    if /i "!POP_SKIP!" neq "y" (
        echo  Build cancelled.
        pause & exit /b 1
    )
    set POPPLER_OK=0
)
echo.

:: ════════════════════════════════════════════════════════
::  STEP 5 — Clean old build artifacts
:: ════════════════════════════════════════════════════════
echo  [5/7]  Cleaning old build files...
if exist "dist"                rmdir /s /q dist
if exist "build"               rmdir /s /q build
if exist "FileConverter.spec"  del   /q   FileConverter.spec
echo         OK
echo.

:: ════════════════════════════════════════════════════════
::  STEP 6 — Build the EXE
:: ════════════════════════════════════════════════════════
echo  [6/7]  Building EXE (this takes 1-3 minutes, please wait)...
echo.

:: base PyInstaller flags (same for both branches)
set BASE_FLAGS=--onefile --windowed --name "FileConverter" ^
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
    --hidden-import docx2pdf ^
    --hidden-import docx ^
    --hidden-import docx.oxml.ns ^
    --hidden-import docx.text.paragraph ^
    --hidden-import docx.table ^
    --hidden-import reportlab ^
    --hidden-import reportlab.platypus ^
    --hidden-import reportlab.lib.styles ^
    --hidden-import reportlab.lib.pagesizes ^
    --hidden-import reportlab.lib.units ^
    --hidden-import reportlab.lib.colors ^
    --hidden-import reportlab.lib.enums ^
    --collect-submodules reportlab ^
    --collect-submodules docx

if "%POPPLER_OK%"=="1" (
    pyinstaller %BASE_FLAGS% --add-data "poppler;poppler" file_converter.py
) else (
    pyinstaller %BASE_FLAGS% file_converter.py
)

if errorlevel 1 (
    echo.
    echo  =====================================================
    echo  [ERROR] Build failed! Check the output above.
    echo.
    echo  Common fixes:
    echo    - Run build.bat as Administrator
    echo    - Upgrade PyInstaller:  pip install pyinstaller --upgrade
    echo    - Delete build\ and dist\ folders manually and retry
    echo  =====================================================
    echo.
    pause & exit /b 1
)

:: ════════════════════════════════════════════════════════
::  STEP 7 — Done
:: ════════════════════════════════════════════════════════
echo.
echo  [7/7]  Finalising...

:: verify the exe was actually created
if not exist "dist\FileConverter.exe" (
    echo  [ERROR] EXE not found in dist\ — build may have failed silently.
    pause & exit /b 1
)

for %%A in ("dist\FileConverter.exe") do set EXE_SIZE=%%~zA
set /a EXE_MB=!EXE_SIZE! / 1048576

echo.
echo  =====================================================
echo    BUILD SUCCESSFUL!
echo.
echo    dist\FileConverter.exe   (!EXE_MB! MB)
echo.
echo    You can share this single file with anyone.
echo    No Python, no pip, no extra setup needed.
echo  =====================================================
echo.

:: open dist folder in Explorer
explorer dist

endlocal
pause

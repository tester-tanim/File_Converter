# 🔄 File Converter

A lightweight desktop app that converts files between formats — no internet required. Just run one Python script and a clean GUI window opens instantly. You can also build it into a standalone `.exe` to share with anyone — no Python needed.

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgrey)

---

## ✨ Features

| Conversion | Input | Output |
|---|---|---|
| Spreadsheet → Image | `.csv`, `.xlsx`, `.xls` | `.png`, `.jpg`, `.jpeg` |
| PDF → Images | `.pdf` | one image per page |
| Images → PDF | `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tiff` | single merged `.pdf` |
| DOCX → PDF | `.docx` | `.pdf` |

- 🖥️ Simple GUI — no command-line knowledge needed
- 📁 Supports multiple files at once (Images → PDF mode)
- 💾 Output saved automatically to the same folder as the input file
- ⚡ Fast conversion with a live progress indicator
- 📦 Can be built into a single `.exe` — no installation needed for end users

---

## 🚀 Option A — Run with Python

### Prerequisites

- Python 3.8 or higher → [Download Python](https://www.python.org/downloads/)

### Installation

**1. Clone the repository**

```bash
git clone https://github.com/your-username/file-converter.git
cd file-converter
```

**2. Install Python dependencies**

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl matplotlib pillow pypdf pdf2image img2pdf
```

**3. Run the app**

```bash
python file_converter.py
```

A GUI window will open. That's it! 🎉

---

## 📦 Option B — Build a Standalone EXE (Windows)

You can package the entire app into a single `.exe` file that works on any Windows machine — no Python, no pip, no setup required for the end user.

### Step 1 — Set up your project folder

```
File_Converter\
├── file_converter.py
├── build.bat
├── requirements.txt
└── poppler\            ← add this manually (see Step 2)
    ├── pdftoppm.exe
    ├── pdfinfo.exe
    └── ...
```

### Step 2 — Add Poppler (for PDF → Images support)

1. Download Poppler from: [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases)
2. Extract the ZIP file
3. Find the folder that contains `pdftoppm.exe`
4. Copy that folder into your project directory and rename it to `poppler`

> **Skip this step** if you don't need PDF → Images. The build will still work, but that feature will require system Poppler to be installed.

### Step 3 — Double-click `build.bat`

The script will automatically:

- ✅ Check Python is installed
- ✅ Install all Python dependencies including PyInstaller
- ✅ Bundle Poppler inside the exe
- ✅ Build `FileConverter.exe`
- ✅ Open the `dist\` folder when done

### Step 4 — Share your EXE

```
dist\FileConverter.exe  ← share this with anyone!
```

No Python, no pip, no Poppler setup needed for the end user. Just double-click and go!

---

## ⚙️ Extra Setup (when running as Python script)

### PDF → Images

This mode requires **Poppler** to be installed on your system.

| OS | Steps |
|---|---|
| **Windows** | Download from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases), extract, and add the `bin/` folder to your system PATH |
| **macOS** | `brew install poppler` |
| **Linux** | `sudo apt install poppler-utils` |

> When running as `.exe`, Poppler is bundled automatically — no extra steps needed.

### DOCX → PDF

This mode requires one of the following:

- **LibreOffice** (free, recommended — works on all platforms) → [libreoffice.org](https://www.libreoffice.org)
- **docx2pdf** (Windows/Mac only, requires Microsoft Word to be installed)

```bash
pip install docx2pdf
```

---

## 🖼️ How to Use

1. Run `python file_converter.py` (or open `FileConverter.exe`)
2. Click a **conversion type** button at the top
3. Click **+ Add files** to select your input file(s)
4. Choose an **output format** if applicable (PNG / JPG / JPEG)
5. Click **Convert**
6. Find the output file in the **same folder as your input**

---

## 📁 Project Structure

```
file-converter/
├── file_converter.py   # Main app — run this
├── build.bat           # Auto-builds FileConverter.exe (Windows)
├── requirements.txt    # Python dependencies
├── poppler\            # Add manually for PDF support in EXE (see Option B)
└── README.md
```

---

## 🛠️ Built With

- [Tkinter](https://docs.python.org/3/library/tkinter.html) — GUI framework (built into Python)
- [Pandas](https://pandas.pydata.org/) — Spreadsheet reading
- [Matplotlib](https://matplotlib.org/) — Table-to-image rendering
- [Pillow](https://python-pillow.org/) — Image processing
- [pypdf](https://pypdf.readthedocs.io/) — PDF handling
- [pdf2image](https://github.com/Belval/pdf2image) — PDF to image conversion
- [img2pdf](https://gitlab.mister-muffin.de/josch/img2pdf) — Images to PDF conversion
- [PyInstaller](https://pyinstaller.org/) — EXE packaging

---

## 🐛 Troubleshooting

**`tkinter` not found (Linux)**
```bash
sudo apt-get install python3-tk
```

**PDF → Images — "Unable to get page count. Is Poppler installed and in PATH?"**
Poppler is not installed or not added to PATH. See the [Extra Setup](#️-extra-setup-when-running-as-python-script) section. If building an EXE, make sure the `poppler\` folder exists in your project before running `build.bat`.

**DOCX → PDF fails**
Install LibreOffice from [libreoffice.org](https://www.libreoffice.org) and make sure it's accessible from your terminal (`libreoffice --version`).

**Images → PDF has color issues**
RGBA or transparent images (e.g. PNGs with transparency) are automatically converted to RGB before merging. No action needed.

**EXE build fails**
Make sure Python and pip are working correctly, then try:
```bash
pip install pyinstaller --upgrade
```
Then run `build.bat` again.

---

## 📄 License

This project is licensed under the MIT License — feel free to use, modify, and distribute it.

---

## 🙌 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m 'Add your feature'`
4. Push to the branch: `git push origin feature/your-feature`
5. Open a Pull Request

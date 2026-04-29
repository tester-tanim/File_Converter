# 🔄 File Converter

A lightweight desktop app that converts files between formats — no internet required. Just run one Python script and a clean GUI window opens instantly.

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

---

## 🚀 Getting Started

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
**3
```bash
pip install docx2pdf
```

**4. Run the app**

```bash
python file_converter.py
```

A GUI window will open. That's it! 🎉

---

## ⚙️ Extra Setup (for some modes)

### PDF → Images

This mode requires **Poppler** to be installed on your system.

| OS | Command |
|---|---|
| **Windows** | Download from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases), extract, and add the `bin/` folder to your system PATH |
| **macOS** | `brew install poppler` |
| **Linux** | `sudo apt install poppler-utils` |

### DOCX → PDF

This mode requires one of the following:

- **LibreOffice** (free, recommended — works on all platforms) → [libreoffice.org](https://www.libreoffice.org)
- **docx2pdf** (Windows/Mac only, requires Microsoft Word to be installed)

```bash
pip install docx2pdf
```

---

## 🖼️ How to Use

1. Run `python file_converter.py`
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
├── requirements.txt    # Python dependencies
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

---

## 🐛 Troubleshooting

**`tkinter` not found (Linux)**
```bash
sudo apt-get install python3-tk
```

**PDF → Images fails**
Make sure Poppler is installed and added to your system PATH. See the [Extra Setup](#️-extra-setup-for-some-modes) section above.

**DOCX → PDF fails**
Install LibreOffice from [libreoffice.org](https://www.libreoffice.org) and make sure it's accessible from your terminal (`libreoffice --version`).

**Images → PDF has color issues**
RGBA or transparent images (e.g., PNGs with transparency) are automatically converted to RGB before merging. No action needed.

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

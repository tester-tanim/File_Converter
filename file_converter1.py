"""
Universal File Converter
Run: python file_converter.py

Conversions supported:
  - CSV / XLSX  → PNG / JPG / JPEG
  - PDF         → PNG / JPG / JPEG  (one image per page)
  - Images      → PDF               (multiple images merged)
  - DOCX        → PDF
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# ═══════════════════════════════════════════════════════════
#  CONVERSION FUNCTIONS
# ═══════════════════════════════════════════════════════════

def spreadsheet_to_image(src, fmt, log):
    ext = Path(src).suffix.lower()
    df  = pd.read_csv(src) if ext == ".csv" else pd.read_excel(src)
    rows, cols = df.shape
    log(f"Loaded {rows} rows × {cols} columns")

    col_width  = max(1.4, min(2.5, 12 / max(cols, 1)))
    fig_w = max(8,  cols * col_width + 1)
    fig_h = max(3, (rows + 1) * 0.45 + 1.2)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    row_fill = [["#F7F9FC" if i % 2 == 0 else "#FFFFFF"] * cols for i in range(rows)]
    tbl = ax.table(
        cellText    = df.astype(str).values.tolist(),
        colLabels   = df.columns.tolist(),
        cellLoc     = "center", loc="center",
        cellColours = row_fill,
        colColours  = ["#2E4057"] * cols,
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(max(7, min(11, int(120 / max(cols, rows, 1)))))
    tbl.auto_set_column_width(col=list(range(cols)))
    for (r, c), cell in tbl.get_celld().items():
        cell.set_edgecolor("#D0D7E3"); cell.set_linewidth(0.6)
        cell.get_text().set_color("#FFFFFF" if r == 0 else "#2C2C2C")
        if r == 0: cell.get_text().set_fontweight("bold")

    title = Path(src).stem.replace("_", " ").title()
    fig.suptitle(title, fontsize=13, fontweight="bold", color="#2E4057", y=0.97)
    fig.tight_layout(pad=0.5)

    save_fmt = "jpeg" if fmt == "jpg" else fmt
    out = str(Path(src).with_suffix(f".{fmt}"))
    fig.savefig(out, format=save_fmt, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    log(f"Saved: {out}")
    return [out]


def pdf_to_images(src, fmt, log):
    try:
        from pdf2image import convert_from_path
    except ImportError:
        raise RuntimeError("pdf2image not installed. Run: pip install pdf2image\nAlso install poppler: https://poppler.freedesktop.org")

    log("Converting PDF pages to images…")
    images = convert_from_path(src, dpi=150)
    log(f"Found {len(images)} page(s)")
    out_dir = Path(src).parent
    stem    = Path(src).stem
    pil_fmt = "JPEG" if fmt in ("jpg", "jpeg") else "PNG"
    ext     = "jpg" if fmt in ("jpg", "jpeg") else "png"
    saved   = []
    for i, img in enumerate(images, 1):
        out = str(out_dir / f"{stem}_page{i}.{ext}")
        img.save(out, pil_fmt)
        log(f"Saved: {Path(out).name}")
        saved.append(out)
    return saved


def images_to_pdf(src_list, log):
    try:
        import img2pdf
    except ImportError:
        raise RuntimeError("img2pdf not installed. Run: pip install img2pdf")

    from PIL import Image

    log(f"Merging {len(src_list)} image(s) into PDF…")
    converted = []
    tmp_files = []
    for p in src_list:
        img = Image.open(p)
        if img.mode in ("RGBA", "P", "LA"):
            rgb = img.convert("RGB")
            tmp = p + "__tmp.jpg"
            rgb.save(tmp, "JPEG")
            tmp_files.append(tmp)
            converted.append(tmp)
        else:
            converted.append(p)

    out = str(Path(src_list[0]).parent / "images_merged.pdf")
    with open(out, "wb") as f:
        f.write(img2pdf.convert(converted))
    for t in tmp_files:
        os.remove(t)
    log(f"Saved: {out}")
    return [out]


def docx_to_pdf(src, log):
    log("Converting DOCX to PDF…")
    import subprocess, shutil

    out = str(Path(src).with_suffix(".pdf"))

    if shutil.which("libreoffice") or shutil.which("soffice"):
        cmd = shutil.which("libreoffice") or shutil.which("soffice")
        result = subprocess.run(
            [cmd, "--headless", "--convert-to", "pdf", "--outdir",
             str(Path(src).parent), src],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr or "LibreOffice conversion failed")
        log(f"Saved: {out}")
        return [out]

    try:
        from docx2pdf import convert
        convert(src, out)
        log(f"Saved: {out}")
        return [out]
    except ImportError:
        pass

    raise RuntimeError(
        "No converter found.\n"
        "Install one of:\n"
        "  • LibreOffice (free): https://www.libreoffice.org\n"
        "  • pip install docx2pdf  (Windows/Mac with Word installed)"
    )


# ═══════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════

MODES = [
    ("CSV / XLSX  →  Image",   "sheet2img"),
    ("PDF  →  Images",         "pdf2img"),
    ("Images  →  PDF",         "img2pdf"),
    ("DOCX  →  PDF",           "docx2pdf"),
]

IMG_FMTS = ["PNG", "JPG", "JPEG"]

BG      = "#F8F9FB"
DARK    = "#2E4057"
GRAY    = "#6B7280"
BORDER  = "#E5E7EB"
GREEN   = "#16A34A"
RED     = "#DC2626"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Converter")
        self.resizable(False, False)
        self.configure(bg=BG)

        self.mode      = tk.StringVar(value="sheet2img")
        self.img_fmt   = tk.StringVar(value="PNG")
        self.files     = []
        self.status_sv = tk.StringVar(value="")

        self._build()
        self._center(560, 560)

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── layout ─────────────────────────────────────────────
    def _build(self):
        # title
        tk.Label(self, text="File Converter", font=("Helvetica", 18, "bold"),
                 bg=BG, fg="#1A1A2E").pack(pady=(26, 4))
        tk.Label(self, text="Choose a conversion, pick your files, and go",
                 font=("Helvetica", 10), bg=BG, fg=GRAY).pack(pady=(0, 18))

        # mode buttons
        tk.Label(self, text="Conversion type", font=("Helvetica", 10, "bold"),
                 bg=BG, fg="#374151", anchor="w").pack(fill="x", padx=28)

        grid = tk.Frame(self, bg=BG)
        grid.pack(fill="x", padx=28, pady=(8, 18))
        self.mode_btns = {}
        for i, (label, key) in enumerate(MODES):
            btn = tk.Button(grid, text=label, font=("Helvetica", 9),
                            relief="flat", cursor="hand2",
                            padx=10, pady=8, wraplength=110,
                            command=lambda k=key: self._set_mode(k))
            btn.grid(row=0, column=i, padx=(0, 8), sticky="ew")
            grid.columnconfigure(i, weight=1)
            self.mode_btns[key] = btn

        # image format (shown only when output is image)
        self.fmt_frame = tk.Frame(self, bg=BG)
        self.fmt_frame.pack(fill="x", padx=28, pady=(0, 14))
        tk.Label(self.fmt_frame, text="Output format",
                 font=("Helvetica", 10, "bold"),
                 bg=BG, fg="#374151", anchor="w").pack(anchor="w", pady=(0, 6))
        btn_row = tk.Frame(self.fmt_frame, bg=BG)
        btn_row.pack(anchor="w")
        self.fmt_btns = {}
        for fmt in IMG_FMTS:
            b = tk.Button(btn_row, text=fmt, font=("Helvetica", 9),
                          relief="flat", cursor="hand2", padx=16, pady=6,
                          command=lambda f=fmt: self._set_fmt(f))
            b.pack(side="left", padx=(0, 8))
            self.fmt_btns[fmt] = b
        self._set_fmt("PNG")

        # file list
        tk.Frame(self, height=1, bg=BORDER).pack(fill="x", padx=28, pady=(0, 14))
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=28)
        self.file_lbl = tk.Label(hdr, text="Selected files",
                                 font=("Helvetica", 10, "bold"),
                                 bg=BG, fg="#374151")
        self.file_lbl.pack(side="left")
        tk.Button(hdr, text="Clear", font=("Helvetica", 9),
                  relief="flat", bg=BORDER, fg=GRAY, padx=8, pady=2,
                  cursor="hand2", command=self._clear_files).pack(side="right")
        tk.Button(hdr, text="+ Add files", font=("Helvetica", 9),
                  relief="flat", bg=DARK, fg="#FFFFFF", padx=10, pady=2,
                  cursor="hand2", command=self._browse).pack(side="right", padx=(0, 8))

        list_frame = tk.Frame(self, bg=BG, bd=0)
        list_frame.pack(fill="both", padx=28, pady=(8, 0), expand=True)
        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")
        self.listbox = tk.Listbox(list_frame, yscrollcommand=sb.set,
                                  font=("Helvetica", 9), relief="flat",
                                  bg="#FFFFFF", fg="#374151",
                                  selectbackground="#E0E7FF",
                                  highlightthickness=1,
                                  highlightbackground=BORDER,
                                  height=7)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.config(command=self.listbox.yview)

        # convert button
        tk.Frame(self, height=1, bg=BORDER).pack(fill="x", padx=28, pady=(14, 0))
        self.convert_btn = tk.Button(self, text="Convert",
                                     font=("Helvetica", 11, "bold"),
                                     relief="flat", bg=DARK, fg="#FFFFFF",
                                     activebackground="#1e2d3d",
                                     activeforeground="#FFFFFF",
                                     cursor="hand2", pady=11,
                                     command=self._start)
        self.convert_btn.pack(fill="x", padx=28, pady=(14, 0))

        # status + progress
        self.status_lbl = tk.Label(self, textvariable=self.status_sv,
                                   font=("Helvetica", 9), bg=BG, fg=GRAY,
                                   wraplength=500, justify="center")
        self.status_lbl.pack(pady=(10, 0))

        style = ttk.Style(); style.theme_use("default")
        style.configure("C.Horizontal.TProgressbar",
                        troughcolor=BORDER, background=DARK,
                        thickness=3, borderwidth=0)
        self.progress = ttk.Progressbar(self, mode="indeterminate",
                                        style="C.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=28, pady=(6, 18))

        # initialise mode AFTER all widgets are built
        self._set_mode("sheet2img")

    # ── mode / format helpers ───────────────────────────────
    def _set_mode(self, key):
        self.mode.set(key)
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.status_sv.set("")
        for k, b in self.mode_btns.items():
            b.config(bg=DARK if k == key else BORDER,
                     fg="#FFFFFF" if k == key else "#374151")
        # show/hide image format selector
        needs_fmt = key in ("sheet2img", "pdf2img")
        if needs_fmt:
            self.fmt_frame.pack(fill="x", padx=28, pady=(0, 14),
                                before=self.convert_btn)
        else:
            self.fmt_frame.pack_forget()

    def _set_fmt(self, fmt):
        self.img_fmt.set(fmt)
        for f, b in self.fmt_btns.items():
            b.config(bg=DARK if f == fmt else BORDER,
                     fg="#FFFFFF" if f == fmt else "#374151")

    # ── file browse ─────────────────────────────────────────
    def _browse(self):
        mode = self.mode.get()
        if mode == "sheet2img":
            types = [("Spreadsheets", "*.csv *.xlsx *.xls")]
            multi = False
        elif mode == "pdf2img":
            types = [("PDF files", "*.pdf")]
            multi = False
        elif mode == "img2pdf":
            types = [("Images", "*.png *.jpg *.jpeg *.bmp *.tiff *.webp")]
            multi = True
        else:
            types = [("Word documents", "*.docx")]
            multi = False

        if multi:
            paths = filedialog.askopenfilenames(title="Choose files", filetypes=types)
            for p in paths:
                if p not in self.files:
                    self.files.append(p)
                    self.listbox.insert(tk.END, Path(p).name)
        else:
            path = filedialog.askopenfilename(title="Choose a file", filetypes=types)
            if path:
                self.files = [path]
                self.listbox.delete(0, tk.END)
                self.listbox.insert(tk.END, Path(path).name)

    def _clear_files(self):
        self.files.clear()
        self.listbox.delete(0, tk.END)
        self.status_sv.set("")

    # ── conversion ──────────────────────────────────────────
    def _log(self, msg):
        self.after(0, lambda: self.status_sv.set(msg))

    def _start(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add at least one file first.")
            return
        self.convert_btn.config(state="disabled")
        self.progress.start(10)
        self.status_sv.set("Working…")
        self.status_lbl.config(fg=GRAY)
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        mode = self.mode.get()
        fmt  = self.img_fmt.get().lower()
        try:
            if mode == "sheet2img":
                for f in self.files:
                    spreadsheet_to_image(f, fmt, self._log)
            elif mode == "pdf2img":
                for f in self.files:
                    pdf_to_images(f, fmt, self._log)
            elif mode == "img2pdf":
                images_to_pdf(self.files, self._log)
            elif mode == "docx2pdf":
                for f in self.files:
                    docx_to_pdf(f, self._log)
            self.after(0, self._done, True, None)
        except Exception as e:
            self.after(0, self._done, False, str(e))

    def _done(self, ok, err):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        if ok:
            self.status_sv.set("✓ Done! Files saved in the same folder as input.")
            self.status_lbl.config(fg=GREEN)
        else:
            self.status_sv.set(f"Error: {err}")
            self.status_lbl.config(fg=RED)


if __name__ == "__main__":
    app = App()
    app.mainloop()

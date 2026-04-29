"""
CSV / XLSX → PNG / JPG / JPEG Converter
Run:  python convert_to_image.py
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# ── conversion logic ─────────────────────────────────────────────────────────

def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path)
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    else:
        raise ValueError(f"Unsupported format: {ext}")


def df_to_image(df, out_path, save_fmt):
    rows, cols = df.shape
    col_width  = max(1.4, min(2.5, 12 / max(cols, 1)))
    row_height = 0.45
    fig_w = max(8,  cols * col_width  + 1)
    fig_h = max(3, (rows + 1) * row_height + 1.2)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis("off")

    cell_text  = df.astype(str).values.tolist()
    header_color = "#2E4057"
    row_colors   = ["#F7F9FC", "#FFFFFF"]
    edge_color   = "#D0D7E3"
    row_fill     = [[row_colors[i % 2]] * cols for i in range(rows)]

    tbl = ax.table(
        cellText    = cell_text,
        colLabels   = df.columns.tolist(),
        cellLoc     = "center",
        loc         = "center",
        cellColours = row_fill,
        colColours  = [header_color] * cols,
    )
    tbl.auto_set_font_size(False)
    font_size = max(7, min(11, int(120 / max(cols, rows, 1))))
    tbl.set_fontsize(font_size)
    tbl.auto_set_column_width(col=list(range(cols)))

    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor(edge_color)
        cell.set_linewidth(0.6)
        if row == 0:
            cell.get_text().set_color("#FFFFFF")
            cell.get_text().set_fontweight("bold")
        else:
            cell.get_text().set_color("#2C2C2C")
        cell.set_height(row_height / fig_h * 0.95)

    title = os.path.basename(out_path).rsplit(".", 1)[0].replace("_", " ").title()
    fig.suptitle(title, fontsize=13, fontweight="bold", color="#2E4057", y=0.97)
    fig.tight_layout(pad=0.5)

    fmt = "jpeg" if save_fmt == "jpg" else save_fmt
    fig.savefig(out_path, format=fmt, dpi=150,
                bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV / XLSX → Image Converter")
        self.resizable(False, False)
        self.configure(bg="#F8F9FB")

        self.file_path = tk.StringVar(value="")
        self.fmt_var   = tk.StringVar(value="png")
        self.status    = tk.StringVar(value="")

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 520, 420
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        pad = dict(padx=28, pady=0)

        # ── title ──────────────────────────────────────────────
        tk.Label(self, text="CSV / XLSX  →  Image",
                 font=("Helvetica", 17, "bold"),
                 bg="#F8F9FB", fg="#1A1A2E").pack(pady=(28, 4))
        tk.Label(self, text="Convert your spreadsheet into a PNG, JPG or JPEG image",
                 font=("Helvetica", 10), bg="#F8F9FB", fg="#6B7280").pack(pady=(0, 22))

        # ── file picker ────────────────────────────────────────
        tk.Label(self, text="Select file", font=("Helvetica", 10, "bold"),
                 bg="#F8F9FB", fg="#374151", anchor="w").pack(fill="x", **pad)

        file_frame = tk.Frame(self, bg="#F8F9FB")
        file_frame.pack(fill="x", padx=28, pady=(6, 18))

        self.file_entry = tk.Entry(file_frame, textvariable=self.file_path,
                                   font=("Helvetica", 10), state="readonly",
                                   relief="flat", bg="#FFFFFF", fg="#374151",
                                   readonlybackground="#FFFFFF",
                                   highlightthickness=1,
                                   highlightbackground="#D1D5DB",
                                   highlightcolor="#6366F1")
        self.file_entry.pack(side="left", fill="x", expand=True,
                             ipady=8, ipadx=8)

        tk.Button(file_frame, text="Browse",
                  font=("Helvetica", 10), relief="flat",
                  bg="#2E4057", fg="#FFFFFF", activebackground="#1e2d3d",
                  activeforeground="#FFFFFF", cursor="hand2",
                  padx=14, pady=8,
                  command=self._browse).pack(side="left", padx=(8, 0))

        # ── format selector ────────────────────────────────────
        tk.Label(self, text="Output format", font=("Helvetica", 10, "bold"),
                 bg="#F8F9FB", fg="#374151", anchor="w").pack(fill="x", **pad)

        fmt_frame = tk.Frame(self, bg="#F8F9FB")
        fmt_frame.pack(fill="x", padx=28, pady=(6, 24))

        self.fmt_buttons = {}
        for fmt in ["png", "jpg", "jpeg"]:
            btn = tk.Button(fmt_frame, text=fmt.upper(),
                            font=("Helvetica", 10),
                            relief="flat", cursor="hand2",
                            padx=20, pady=7,
                            command=lambda f=fmt: self._select_fmt(f))
            btn.pack(side="left", padx=(0, 8))
            self.fmt_buttons[fmt] = btn
        self._select_fmt("png")

        # ── divider ────────────────────────────────────────────
        tk.Frame(self, height=1, bg="#E5E7EB").pack(fill="x", padx=28, pady=(0, 22))

        # ── convert button ─────────────────────────────────────
        self.convert_btn = tk.Button(self, text="Convert",
                                     font=("Helvetica", 11, "bold"),
                                     relief="flat", bg="#2E4057", fg="#FFFFFF",
                                     activebackground="#1e2d3d",
                                     activeforeground="#FFFFFF",
                                     cursor="hand2", pady=12,
                                     command=self._start_convert)
        self.convert_btn.pack(fill="x", padx=28)

        # ── status ─────────────────────────────────────────────
        self.status_lbl = tk.Label(self, textvariable=self.status,
                                   font=("Helvetica", 10),
                                   bg="#F8F9FB", fg="#6B7280",
                                   wraplength=460, justify="center")
        self.status_lbl.pack(pady=(14, 0))

        # ── progress bar ───────────────────────────────────────
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor="#E5E7EB", background="#2E4057",
                        thickness=4, borderwidth=0)
        self.progress = ttk.Progressbar(self, mode="indeterminate",
                                        style="Custom.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=28, pady=(10, 0))

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Choose a file",
            filetypes=[("Spreadsheets", "*.csv *.xlsx *.xls"),
                       ("CSV", "*.csv"),
                       ("Excel", "*.xlsx *.xls")]
        )
        if path:
            self.file_path.set(path)
            self.status.set("")
            self.status_lbl.config(fg="#6B7280")

    def _select_fmt(self, fmt):
        self.fmt_var.set(fmt)
        for f, btn in self.fmt_buttons.items():
            if f == fmt:
                btn.config(bg="#2E4057", fg="#FFFFFF")
            else:
                btn.config(bg="#E5E7EB", fg="#374151")

    def _start_convert(self):
        path = self.file_path.get()
        if not path:
            messagebox.showwarning("No file", "Please select a CSV or XLSX file first.")
            return
        self.convert_btn.config(state="disabled")
        self.progress.start(10)
        self.status.set("Converting…")
        self.status_lbl.config(fg="#6B7280")
        threading.Thread(target=self._convert, daemon=True).start()

    def _convert(self):
        path = self.file_path.get()
        fmt  = self.fmt_var.get()
        base = os.path.splitext(path)[0]
        out  = f"{base}.{fmt}"
        try:
            df = load_file(path)
            df_to_image(df, out, fmt)
            self.after(0, self._done, out, None)
        except Exception as e:
            self.after(0, self._done, None, str(e))

    def _done(self, out_path, error):
        self.progress.stop()
        self.convert_btn.config(state="normal")
        if error:
            self.status.set(f"Error: {error}")
            self.status_lbl.config(fg="#DC2626")
        else:
            self.status.set(f"Saved: {out_path}")
            self.status_lbl.config(fg="#16A34A")


if __name__ == "__main__":
    app = App()
    app.mainloop()

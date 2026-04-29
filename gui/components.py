import os
import tkinter as tk
from tkinter import filedialog

from core.config import THEME, sc, sf, t


class HoverButton(tk.Label):
    def __init__(self, master, text, command=None, style="primary", **kw):
        styles = {
            "primary": (THEME["btn_bg"],  THEME["btn_fg"],  THEME["btn_hover"]),
            "ghost":   (THEME["card"],    THEME["accent"],  THEME["border"]),
            "tool":    (THEME["card"],    THEME["text"],    THEME["bg"]),
        }
        bg, fg, hover = styles.get(style, styles["primary"])
        super().__init__(
            master, text=text, cursor="hand2",
            font=THEME["font_btn"], fg=fg, bg=bg,
            padx=sc(18), pady=sc(10), relief="flat",
            **kw
        )
        self._bg, self._hover, self._cmd = bg, hover, command
        self.bind("<Enter>",    lambda e: self.config(bg=self._hover))
        self.bind("<Leave>",    lambda e: self.config(bg=self._bg))
        self.bind("<Button-1>", lambda e: self._cmd() if self._cmd else None)


def _pick_folder(var: tk.StringVar, preview_label: tk.Label = None,
                 extensions: tuple = (), open_btn=None):
    path = filedialog.askdirectory()
    if path:
        var.set(path)
        if preview_label is not None:
            _update_preview(path, preview_label, extensions)
        if open_btn is not None:
            open_btn.pack(side="right")


def _update_preview(folder: str, label: tk.Label, extensions: tuple):
    try:
        files = sorted([
            f for f in os.listdir(folder)
            if f.lower().endswith(extensions) and not f.startswith('~$')
        ])
        if not files:
            label.config(text="⚠️  Bu klasörde uygun dosya yok.",
                         fg=THEME["log_err"])
        else:
            preview = "  ".join(files[:6])
            if len(files) > 6:
                preview += f"  (+{len(files)-6} daha)"
            label.config(text=f"📂  {len(files)} dosya:  {preview}",
                         fg=THEME["log_ok"])
    except Exception:
        label.config(text="", fg=THEME["subtext"])


def _form_row_with_preview(parent, label_text, var, extensions: tuple):
    frame = tk.Frame(parent, bg=THEME["card"])
    frame.pack(fill="x", pady=(0, sc(12)))
    tk.Label(frame, text=label_text, bg=THEME["card"],
             fg=THEME["subtext"], font=THEME["font_sub"]).pack(anchor="w")
    row = tk.Frame(frame, bg=THEME["card"])
    row.pack(fill="x", pady=(sc(4), 0))
    entry = tk.Entry(row, textvariable=var, font=THEME["font_main"],
                     bg=THEME["bg"], fg=THEME["text"],
                     relief="flat", bd=0,
                     highlightthickness=1,
                     highlightbackground=THEME["border"],
                     highlightcolor=THEME["accent"])
    entry.pack(side="left", fill="x", expand=True,
               ipady=sc(6), padx=(0, sc(8)))

    preview_row = tk.Frame(frame, bg=THEME["card"])
    preview_row.pack(fill="x", pady=(sc(4), 0))

    preview = tk.Label(preview_row, text="", bg=THEME["card"],
                       fg=THEME["subtext"], font=THEME["font_sub"],
                       anchor="w", wraplength=sc(380))
    preview.pack(side="left", fill="x", expand=True)

    def open_folder():
        path = var.get().strip()
        if os.path.isdir(path):
            os.startfile(path)

    open_btn = HoverButton(preview_row, text="🗂 Aç",
                           command=open_folder, style="ghost")
    open_btn.pack_forget()

    browse_cmd = lambda: _pick_folder(var, preview, extensions, open_btn)
    HoverButton(row, text=t("btn_browse"),
                command=browse_cmd, style="ghost").pack(side="right")

    def on_var_change(*_):
        path = var.get().strip()
        if os.path.isdir(path):
            _update_preview(path, preview, extensions)
            open_btn.pack(side="right")
        else:
            open_btn.pack_forget()

    var.trace_add("write", on_var_change)


def _form_text(parent, label_text, var):
    frame = tk.Frame(parent, bg=THEME["card"])
    frame.pack(fill="x", pady=(0, sc(12)))
    tk.Label(frame, text=label_text, bg=THEME["card"],
             fg=THEME["subtext"], font=THEME["font_sub"]).pack(anchor="w")
    entry = tk.Entry(frame, textvariable=var, font=THEME["font_main"],
                     bg=THEME["bg"], fg=THEME["text"],
                     relief="flat", bd=0,
                     highlightthickness=1,
                     highlightbackground=THEME["border"],
                     highlightcolor=THEME["accent"])
    entry.pack(fill="x", pady=(sc(4), 0), ipady=sc(6))


def _form_check(parent, label_text, var):
    frame = tk.Frame(parent, bg=THEME["card"])
    frame.pack(fill="x", pady=(0, sc(12)))
    cb = tk.Checkbutton(frame, text=label_text, variable=var,
                        bg=THEME["card"], fg=THEME["text"],
                        activebackground=THEME["card"],
                        selectcolor=THEME["bg"],
                        font=THEME["font_main"],
                        relief="flat", bd=0)
    cb.pack(anchor="w")

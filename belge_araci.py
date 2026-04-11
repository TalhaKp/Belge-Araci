"""
Belge Aracı - PDF Birleştirici & Word → PDF Dönüştürücü
========================================================
Genişletme notları (geliştirici için):
  - Yeni araç eklemek için: ToolBase sınıfından türet, TOOLS listesine ekle.
  - Dil desteği: STRINGS sözlüğüne yeni dil anahtarı ekle, LanguageManager'ı kullan.
  - Tema desteği: THEME sözlüğünü değiştir veya runtime'da güncelle.

Word → PDF dönüştürücü öncelik sırası:
  1. Microsoft Word (win32com) — en iyi format koruması
  2. LibreOffice (soffice CLI) — Word kurulu değilse otomatik devreye girer
  3. İkisi de yoksa kullanıcıya LibreOffice indirme teklifi yapılır
"""

import os
import re
import sys
import glob
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from abc import ABC, abstractmethod
import webbrowser

# ─────────────────────────────────────────────
#  DPI ÖLÇEKLENDIRME
# ─────────────────────────────────────────────
def _get_dpi_scale() -> float:
    """
    Windows DPI ölçeğini tespit eder (örn. 1.25 = %125, 1.5 = %150).
    Diğer platformlarda 1.0 döner.
    """
    try:
        import ctypes
        awareness = ctypes.c_int()
        ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return dpi / 96.0
    except Exception:
        return 1.0

DPI_SCALE = _get_dpi_scale()

def sc(value: int) -> int:
    """Verilen piksel değerini DPI ölçeğine göre ölçeklendirir."""
    return max(1, round(value * DPI_SCALE))

def sf(size: int) -> int:
    """Font boyutunu DPI ölçeğine göre ölçeklendirir."""
    return max(8, round(size * DPI_SCALE))

# ─────────────────────────────────────────────
#  İKON (base64 gömülü .ico)
# ─────────────────────────────────────────────
import base64, tempfile, atexit

ICON_B64 = (
    (
    "AAABAAEAEBAAAAAAIAC9AAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAIRJ"
    "REFUeJylUzsOgCAMbQ230bDrgkfXBXeC59GpyZNQUvBNNOH9+DAR0bHODw1gv25mJIeYvazPbUkWkQkH"
    "JKGYWaDHWRWwOgtcixhixmS2BIBPFRSrCvR2F7D2BhRHX5qpFcrOOON5qQk0CFlStA5RJSNcbaOFKOhK"
    "MCxQXjHOTPTvO7+4HziJueObpgAAAABJRU5ErkJggg=="
)
)

_icon_tmp_path = None

def _setup_icon(root: tk.Tk):
    """
    Base64 ikonunu geçici bir .ico dosyasına yazar ve pencereye atar.
    Uygulama kapanınca geçici dosya silinir.
    """
    global _icon_tmp_path
    if ICON_B64 == (
    "AAABAAEAEBAAAAAAIAC9AAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAIRJ"
    "REFUeJylUzsOgCAMbQ230bDrgkfXBXeC59GpyZNQUvBNNOH9+DAR0bHODw1gv25mJIeYvazPbUkWkQkH"
    "JKGYWaDHWRWwOgtcixhixmS2BIBPFRSrCvR2F7D2BhRHX5qpFcrOOON5qQk0CFlStA5RJSNcbaOFKOhK"
    "MCxQXjHOTPTvO7+4HziJueObpgAAAABJRU5ErkJggg=="
):
        return  # Geliştirme modunda ikon yoksa sessizce geç
    try:
        data = base64.b64decode(ICON_B64)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".ico")
        tmp.write(data)
        tmp.close()
        _icon_tmp_path = tmp.name
        root.iconbitmap(_icon_tmp_path)
        atexit.register(lambda: os.unlink(_icon_tmp_path) if os.path.exists(_icon_tmp_path) else None)
    except Exception:
        pass  # İkon yüklenemezse uygulama çalışmaya devam eder


# ─────────────────────────────────────────────
#  DİL / STRINGS  (ileride çeviri buradan)
# ─────────────────────────────────────────────
STRINGS = {
    "tr": {
        "app_title":        "Belge Aracı",
        "select_tool":      "Ne yapmak istersiniz?",
        "btn_merge":        "📎  PDF Birleştir",
        "btn_word2pdf":     "📄  Word → PDF",
        "btn_back":         "← Geri",
        "btn_start":        "Başlat",
        "btn_browse":       "Klasör Seç",
        "label_folder":     "Klasör:",
        "label_outname":    "Çıktı dosya adı:",
        "label_delete":     "Orijinal Word dosyaları dönüşüm sonrası silinsin mi?",
        "running":          "İşlem devam ediyor…",
        "done":             "✅  İşlem tamamlandı!",
        "err_no_folder":    "Lütfen bir klasör seçin.",
        "err_no_pdf":       "Seçilen klasörde hiç PDF yok.",
        "err_no_word":      "Seçilen klasörde hiç Word dosyası yok.",
        "err_word_app":     "Microsoft Word başlatılamadı. Word kurulu mu?",
        "merge_title":      "PDF Birleştirici",
        "word2pdf_title":   "Word → PDF Dönüştürücü",
        "merge_desc":       "Seçilen klasördeki tüm PDF'leri tek dosyada birleştirir.",
        "word2pdf_desc":    "Klasördeki tüm .doc/.docx dosyalarını PDF'e çevirir.",
        "log_scanning":     "Klasör taranıyor…",
        "log_merging":      "Birleştiriliyor: {name}",
        "log_converting":   "Dönüştürülüyor: {name}",
        "log_skipped":      "Es geçildi (zaten var): {name}",
        "log_ok":           "✓ {name}",
        "log_err":          "✗ {name} — {err}",
        "log_saved":        "Kaydedildi → {path}",
        "log_summary":      "Toplam: {ok} başarılı, {skip} atlandı.",
        "log_deleted":      "Silindi: {name}",
        "yes":              "Evet",
        "no":               "Hayır",
        "engine_word":      "Microsoft Word kullanılıyor…",
        "engine_libre":     "LibreOffice kullanılıyor…",
        "err_no_engine":    "Dönüştürücü bulunamadı.",
        "libre_offer_title":"LibreOffice Bulunamadı",
        "libre_offer_msg":  (
            "Word → PDF dönüşümü için Microsoft Word veya LibreOffice gereklidir.\n\n"
            "Bilgisayarınızda ikisi de bulunamadı.\n\n"
            "LibreOffice ücretsizdir. İndirme sayfasını açmamı ister misiniz?"
        ),
    }
    # Buraya "en": {...} eklenince LanguageManager otomatik devreye girer.
}

class LanguageManager:
    def __init__(self, lang="tr"):
        self.lang = lang

    def t(self, key, **kwargs):
        text = STRINGS.get(self.lang, STRINGS["tr"]).get(key, key)
        return text.format(**kwargs) if kwargs else text

LM = LanguageManager("tr")
t = LM.t

# ─────────────────────────────────────────────
#  TEMA  (DPI'a duyarlı fontlar)
# ─────────────────────────────────────────────
THEME = {
    "bg":           "#F5F0EB",
    "card":         "#FFFFFF",
    "accent":       "#C0392B",
    "accent_dark":  "#922B21",
    "text":         "#1A1A1A",
    "subtext":      "#6B6B6B",
    "border":       "#E0D8D0",
    "log_bg":       "#1E1E1E",
    "log_fg":       "#D4D4D4",
    "log_ok":       "#6A9F6A",
    "log_err":      "#E06C75",
    "log_info":     "#61AFEF",
    "btn_bg":       "#C0392B",
    "btn_fg":       "#FFFFFF",
    "btn_hover":    "#922B21",
    "font_main":    ("Segoe UI", sf(10)),
    "font_title":   ("Segoe UI", sf(18), "bold"),
    "font_sub":     ("Segoe UI", sf(9)),
    "font_btn":     ("Segoe UI", sf(10), "bold"),
    "font_mono":    ("Consolas", sf(9)),
}

# ─────────────────────────────────────────────
#  YARDIMCI: Özel Buton
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
#  ARAÇ TEMEL SINIFI
# ─────────────────────────────────────────────
class ToolBase(ABC):
    @property
    @abstractmethod
    def title(self) -> str: ...

    @property
    @abstractmethod
    def description(self) -> str: ...

    @property
    @abstractmethod
    def icon(self) -> str: ...

    @abstractmethod
    def build_form(self, parent: tk.Frame, log_fn) -> dict: ...

    @abstractmethod
    def run(self, state: dict, log_fn, done_fn): ...

# ─────────────────────────────────────────────
#  ARAÇ 1: PDF BİRLEŞTİRİCİ
# ─────────────────────────────────────────────
class PdfMergerTool(ToolBase):
    title       = t("merge_title")
    description = t("merge_desc")
    icon        = "📎"

    def build_form(self, parent, log_fn):
        state = {
            "folder":  tk.StringVar(),
            "outname": tk.StringVar(value="Birlestirilmis_Dosya.pdf"),
        }
        _form_row(parent, t("label_folder"),
                  state["folder"], lambda: _pick_folder(state["folder"]))
        _form_text(parent, t("label_outname"), state["outname"])
        return state

    def run(self, state, log_fn, done_fn):
        folder  = state["folder"].get().strip()
        outname = state["outname"].get().strip() or "Birlestirilmis_Dosya.pdf"

        def worker():
            try:
                from pypdf import PdfWriter, PdfReader
            except ImportError:
                log_fn(t("log_err", name="pypdf", err="Kütüphane eksik. pip install pypdf"), "err")
                done_fn(False); return

            if not folder or not os.path.isdir(folder):
                log_fn(t("err_no_folder"), "err"); done_fn(False); return

            log_fn(t("log_scanning"), "info")

            def natural_sort(name):
                return tuple(int(c) if c.isdecimal() else c.lower().strip()
                             for c in re.split(r'(\d+)', name))

            pdfs = sorted(
                [f for f in os.listdir(folder)
                 if f.lower().endswith('.pdf') and f != outname],
                key=natural_sort
            )
            if not pdfs:
                log_fn(t("err_no_pdf"), "err"); done_fn(False); return

            writer = PdfWriter()
            ok = 0
            for fname in pdfs:
                fpath = os.path.join(folder, fname)
                try:
                    reader = PdfReader(fpath)
                    writer.append(reader)
                    log_fn(t("log_ok", name=f"{fname} ({len(reader.pages)} sayfa)"), "ok")
                    ok += 1
                except Exception as e:
                    log_fn(t("log_err", name=fname, err=str(e)), "err")

            out_path = os.path.join(folder, outname)
            try:
                with open(out_path, "wb") as f:
                    writer.write(f)
                log_fn(t("log_saved", path=out_path), "info")
                log_fn(t("log_summary", ok=ok, skip=0), "info")
                done_fn(True)
            except PermissionError:
                log_fn(t("log_err", name=outname, err="Dosya açık olabilir."), "err")
                done_fn(False)

        threading.Thread(target=worker, daemon=True).start()

# ─────────────────────────────────────────────
#  DÖNÜŞTÜRME MOTORU TESPİTİ
# ─────────────────────────────────────────────
def _find_libreoffice() -> str | None:
    try:
        result = subprocess.run(
            ["soffice", "--version"], capture_output=True, timeout=5)
        if result.returncode == 0:
            return "soffice"
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for pattern in [
        r"C:\Program Files\LibreOffice*\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice*\program\soffice.exe",
    ]:
        candidates.extend(glob.glob(pattern))

    for path in candidates:
        if os.path.isfile(path):
            return path
    return None

# ─────────────────────────────────────────────
#  ARAÇ 2: WORD → PDF
# ─────────────────────────────────────────────
class Word2PdfTool(ToolBase):
    title       = t("word2pdf_title")
    description = t("word2pdf_desc")
    icon        = "📄"

    def build_form(self, parent, log_fn):
        state = {
            "folder": tk.StringVar(),
            "delete": tk.BooleanVar(value=False),
        }
        _form_row(parent, t("label_folder"),
                  state["folder"], lambda: _pick_folder(state["folder"]))
        _form_check(parent, t("label_delete"), state["delete"])
        return state

    def run(self, state, log_fn, done_fn):
        from pathlib import Path
        folder    = state["folder"].get().strip()
        do_delete = state["delete"].get()

        def worker():
            if not folder or not os.path.isdir(folder):
                log_fn(t("err_no_folder"), "err"); done_fn(False); return

            klasor = Path(folder).resolve()
            log_fn(t("log_scanning"), "info")

            word_files = [f for f in klasor.glob('*.doc*')
                          if not f.name.startswith('~$')]
            if not word_files:
                log_fn(t("err_no_word"), "err"); done_fn(False); return

            use_word   = False
            libre_path = None

            try:
                import win32com.client as _wc
                _app = _wc.DispatchEx("Word.Application")
                _app.Quit()
                use_word = True
            except Exception:
                pass

            if not use_word:
                libre_path = _find_libreoffice()

            if not use_word and libre_path is None:
                def _ask():
                    ans = messagebox.askyesno(
                        t("libre_offer_title"), t("libre_offer_msg"))
                    if ans:
                        webbrowser.open(
                            "https://www.libreoffice.org/download/download-libreoffice/")
                    log_fn(t("err_no_engine"), "err")
                    done_fn(False)
                tk._default_root.after(0, _ask)
                return

            if use_word:
                log_fn(t("engine_word"), "info")
                self._run_with_word(word_files, do_delete, log_fn, done_fn)
            else:
                log_fn(t("engine_libre"), "info")
                self._run_with_libreoffice(libre_path, word_files, do_delete, log_fn, done_fn)

        threading.Thread(target=worker, daemon=True).start()

    def _run_with_word(self, word_files, do_delete, log_fn, done_fn):
        import win32com.client
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
        except Exception as e:
            log_fn(t("err_word_app") + f" ({e})", "err"); done_fn(False); return

        ok, skip, converted = 0, 0, []
        try:
            for dosya in word_files:
                pdf_yolu = dosya.with_suffix('.pdf')
                if pdf_yolu.exists():
                    log_fn(t("log_skipped", name=dosya.name), "info"); skip += 1; continue
                log_fn(t("log_converting", name=dosya.name), "info")
                try:
                    doc = word.Documents.Open(str(dosya.resolve()))
                    doc.SaveAs(str(pdf_yolu.resolve()), FileFormat=17)
                    doc.Close(0)
                    log_fn(t("log_ok", name=pdf_yolu.name), "ok")
                    ok += 1; converted.append(dosya)
                except Exception as e:
                    log_fn(t("log_err", name=dosya.name, err=str(e)), "err")
        finally:
            word.Quit()

        self._finish(ok, skip, converted, do_delete, log_fn, done_fn)

    def _run_with_libreoffice(self, soffice, word_files, do_delete, log_fn, done_fn):
        ok, skip, converted = 0, 0, []
        for dosya in word_files:
            pdf_yolu = dosya.with_suffix('.pdf')
            if pdf_yolu.exists():
                log_fn(t("log_skipped", name=dosya.name), "info"); skip += 1; continue
            log_fn(t("log_converting", name=dosya.name), "info")
            try:
                result = subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf",
                     "--outdir", str(dosya.parent), str(dosya)],
                    capture_output=True, timeout=60
                )
                if result.returncode == 0 and pdf_yolu.exists():
                    log_fn(t("log_ok", name=pdf_yolu.name), "ok")
                    ok += 1; converted.append(dosya)
                else:
                    err_msg = result.stderr.decode(errors="replace").strip()
                    log_fn(t("log_err", name=dosya.name,
                             err=err_msg or "bilinmeyen hata"), "err")
            except subprocess.TimeoutExpired:
                log_fn(t("log_err", name=dosya.name, err="zaman aşımı"), "err")
            except Exception as e:
                log_fn(t("log_err", name=dosya.name, err=str(e)), "err")

        self._finish(ok, skip, converted, do_delete, log_fn, done_fn)

    def _finish(self, ok, skip, converted, do_delete, log_fn, done_fn):
        log_fn(t("log_summary", ok=ok, skip=skip), "info")
        if do_delete and converted:
            for f in converted:
                try:
                    f.unlink()
                    log_fn(t("log_deleted", name=f.name), "info")
                except Exception as e:
                    log_fn(t("log_err", name=f.name, err=str(e)), "err")
        done_fn(True)

# ─────────────────────────────────────────────
#  ARAÇ LİSTESİ  ← buraya yeni araç ekle
# ─────────────────────────────────────────────
TOOLS: list[ToolBase] = [
    PdfMergerTool(),
    Word2PdfTool(),
]

# ─────────────────────────────────────────────
#  FORM YARDIMCILARI
# ─────────────────────────────────────────────
def _pick_folder(var: tk.StringVar):
    path = filedialog.askdirectory()
    if path:
        var.set(path)

def _form_row(parent, label_text, var, browse_cmd):
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
    HoverButton(row, text=t("btn_browse"),
                command=browse_cmd, style="ghost").pack(side="right")

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

# ─────────────────────────────────────────────
#  ANA UYGULAMA
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(t("app_title"))
        self.resizable(True, True)
        self.minsize(sc(480), sc(540))
        self.configure(bg=THEME["bg"])
        self._center(sc(520), sc(640))

        # İkonu ayarla
        _setup_icon(self)

        self._container = tk.Frame(self, bg=THEME["bg"])
        self._container.pack(fill="both", expand=True)

        self._show_home()

    def _center(self, w, h):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _clear(self):
        for w in self._container.winfo_children():
            w.destroy()

    def _show_home(self):
        self._clear()
        p = self._container

        hdr = tk.Frame(p, bg=THEME["accent"], height=sc(6))
        hdr.pack(fill="x")

        tk.Label(p, text=t("app_title"), font=THEME["font_title"],
                 bg=THEME["bg"], fg=THEME["text"],
                 pady=sc(24)).pack()

        tk.Label(p, text=t("select_tool"), font=THEME["font_main"],
                 bg=THEME["bg"], fg=THEME["subtext"]).pack(pady=(0, sc(16)))

        for tool in TOOLS:
            self._tool_card(p, tool)

    def _tool_card(self, parent, tool: ToolBase):
        card = tk.Frame(parent, bg=THEME["card"],
                        highlightthickness=1,
                        highlightbackground=THEME["border"],
                        cursor="hand2")
        card.pack(fill="x", padx=sc(32), pady=sc(8), ipady=sc(4))

        inner = tk.Frame(card, bg=THEME["card"])
        inner.pack(fill="x", padx=sc(20), pady=sc(14))

        tk.Label(inner, text=f"{tool.icon}  {tool.title}",
                 font=("Segoe UI", sf(12), "bold"),
                 bg=THEME["card"], fg=THEME["text"],
                 anchor="w").pack(fill="x")

        tk.Label(inner, text=tool.description,
                 font=THEME["font_sub"],
                 bg=THEME["card"], fg=THEME["subtext"],
                 anchor="w").pack(fill="x", pady=(sc(4), 0))

        all_widgets = [card, inner] + inner.winfo_children()
        for w in all_widgets:
            w.bind("<Button-1>", lambda e, tl=tool: self._show_tool(tl))
            w.bind("<Enter>", lambda e, c=card: c.config(
                bg=THEME["bg"], highlightbackground=THEME["accent"]))
            w.bind("<Leave>", lambda e, c=card: c.config(
                bg=THEME["card"], highlightbackground=THEME["border"]))

    def _show_tool(self, tool: ToolBase):
        self._clear()
        p = self._container

        bar = tk.Frame(p, bg=THEME["accent"], height=sc(6))
        bar.pack(fill="x")

        top = tk.Frame(p, bg=THEME["bg"])
        top.pack(fill="x", padx=sc(24), pady=(sc(16), 0))

        HoverButton(top, text=t("btn_back"),
                    command=self._show_home,
                    style="ghost").pack(side="left")

        tk.Label(top, text=f"{tool.icon}  {tool.title}",
                 font=("Segoe UI", sf(13), "bold"),
                 bg=THEME["bg"], fg=THEME["text"]).pack(side="left", padx=sc(12))

        card = tk.Frame(p, bg=THEME["card"],
                        highlightthickness=1,
                        highlightbackground=THEME["border"])
        card.pack(fill="x", padx=sc(24), pady=sc(16))

        form_inner = tk.Frame(card, bg=THEME["card"])
        form_inner.pack(fill="x", padx=sc(20), pady=sc(16))

        state = tool.build_form(form_inner, self._log)

        self._start_btn = HoverButton(
            p, text=t("btn_start"),
            command=lambda: self._run_tool(tool, state),
            style="primary")
        self._start_btn.pack(pady=(0, sc(12)))

        log_frame = tk.Frame(p, bg=THEME["log_bg"],
                             highlightthickness=1,
                             highlightbackground=THEME["border"])
        log_frame.pack(fill="both", expand=True,
                       padx=sc(24), pady=(0, sc(16)))

        self._log_text = tk.Text(
            log_frame, bg=THEME["log_bg"], fg=THEME["log_fg"],
            font=THEME["font_mono"],
            relief="flat", bd=0,
            state="disabled", wrap="word",
            padx=sc(10), pady=sc(10))
        self._log_text.pack(fill="both", expand=True)

        self._log_text.tag_config("ok",   foreground=THEME["log_ok"])
        self._log_text.tag_config("err",  foreground=THEME["log_err"])
        self._log_text.tag_config("info", foreground=THEME["log_info"])

    def _log(self, msg: str, tag: str = ""):
        def _write():
            self._log_text.config(state="normal")
            self._log_text.insert("end", msg + "\n", tag)
            self._log_text.see("end")
            self._log_text.config(state="disabled")
        self.after(0, _write)

    def _run_tool(self, tool: ToolBase, state: dict):
        self._start_btn.config(text=t("running"))
        self._start_btn.unbind("<Button-1>")

        def done(success: bool):
            def _ui():
                self._start_btn.config(text=t("btn_start"))
                self._start_btn.bind("<Button-1>",
                    lambda e: self._run_tool(tool, state))
                if success:
                    self._log(t("done"), "ok")
            self.after(0, _ui)

        tool.run(state, self._log, done)


# ─────────────────────────────────────────────
#  GİRİŞ NOKTASI
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
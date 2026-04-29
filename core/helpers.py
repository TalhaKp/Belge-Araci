import os
import re
import sys


def get_resource_path(relative_path):
    """ PyInstaller ile paketlendiğinde veya normal çalışırken doğru dosya yolunu bulur. """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def _get_dpi_scale() -> float:
    try:
        import ctypes
        awareness = ctypes.c_int()
        ctypes.windll.shcore.GetProcessDpiAwareness(0, ctypes.byref(awareness))
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        dpi = ctypes.windll.user32.GetDpiForSystem()
        return dpi / 96.0
    except Exception:
        return 1.0


DPI_SCALE = _get_dpi_scale()


def _font_scale() -> float:
    try:
        import tkinter as _tk
        _r = _tk.Tk()
        sh = _r.winfo_screenheight()
        _r.destroy()
        screen_factor = min(1.0, sh / 900)
        return min(1.5, DPI_SCALE * screen_factor)
    except Exception:
        return min(1.5, DPI_SCALE)


FONT_SCALE = _font_scale()


def sanitize_filename(name: str, fallback: str = "Birlestirilmis_Dosya.pdf") -> str:
    if not name or not name.strip():
        return fallback

    name = re.sub(r'[\x00-\x1f\x7f]', '', name)
    name = re.sub(r'[/\\:*?"<>|]', '', name)
    name = name.replace('..', '')
    name = name.strip('. ')

    if not name:
        return fallback

    stem = os.path.splitext(name)[0].strip('. ')
    if not stem:
        return fallback

    stem = stem[:200]
    return stem + ".pdf"

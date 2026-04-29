import os
import re
import threading
import tkinter as tk

from core.config import t
from core.helpers import sanitize_filename
from gui.components import _form_row_with_preview, _form_text
from tools.tool_base import ToolBase


class PdfMergerTool(ToolBase):
    title       = t("merge_title")
    description = t("merge_desc")
    icon        = "📎"

    def build_form(self, parent, log_fn):
        state = {
            "folder":  tk.StringVar(),
            "outname": tk.StringVar(value="Birlestirilmis_Dosya.pdf"),
        }
        _form_row_with_preview(parent, t("label_folder"),
                               state["folder"], ('.pdf',))
        _form_text(parent, t("label_outname"), state["outname"])
        return state

    def run(self, state, log_fn, done_fn):
        folder  = state["folder"].get().strip()
        outname = sanitize_filename(state["outname"].get())

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

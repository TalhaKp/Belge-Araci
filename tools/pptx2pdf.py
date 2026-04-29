import os
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox
import webbrowser

from core.config import t
from gui.components import _form_row_with_preview, _form_check
from tools.tool_base import ToolBase, _find_libreoffice


class Pptx2PdfTool(ToolBase):
    title       = t("pptx2pdf_title")
    description = t("pptx2pdf_desc")
    icon        = "📊"

    def build_form(self, parent, log_fn):
        state = {
            "folder": tk.StringVar(),
            "delete": tk.BooleanVar(value=False),
        }
        _form_row_with_preview(parent, t("label_folder"),
                               state["folder"], ('.ppt', '.pptx'))
        _form_check(parent, t("label_delete_pptx"), state["delete"])
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

            pptx_files = [f for f in klasor.glob('*.ppt*')
                          if not f.name.startswith('~$')]
            if not pptx_files:
                log_fn(t("err_no_pptx"), "err"); done_fn(False); return

            use_ppt    = False
            libre_path = None

            try:
                import win32com.client as _wc
                _app = _wc.DispatchEx("PowerPoint.Application")
                _app.Quit()
                use_ppt = True
            except Exception:
                pass

            if not use_ppt:
                libre_path = _find_libreoffice()

            if not use_ppt and libre_path is None:
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

            if use_ppt:
                log_fn(t("engine_ppt"), "info")
                self._run_with_ppt(pptx_files, do_delete, log_fn, done_fn)
            else:
                log_fn(t("engine_libre"), "info")
                self._run_with_libreoffice(libre_path, pptx_files, do_delete, log_fn, done_fn)

        threading.Thread(target=worker, daemon=True).start()

    def _run_with_ppt(self, pptx_files, do_delete, log_fn, done_fn):
        import win32com.client
        try:
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.Visible = True
        except Exception as e:
            log_fn(t("err_ppt_app") + f" ({e})", "err"); done_fn(False); return

        ok, skip, converted = 0, 0, []
        try:
            for dosya in pptx_files:
                pdf_yolu = dosya.with_suffix('.pdf')
                if pdf_yolu.exists():
                    log_fn(t("log_skipped", name=dosya.name), "info"); skip += 1; continue
                log_fn(t("log_converting", name=dosya.name), "info")
                try:
                    prs = ppt.Presentations.Open(str(dosya.resolve()), WithWindow=False)
                    prs.SaveAs(str(pdf_yolu.resolve()), FileFormat=32)
                    prs.Close()
                    log_fn(t("log_ok", name=pdf_yolu.name), "ok")
                    ok += 1; converted.append(dosya)
                except Exception as e:
                    log_fn(t("log_err", name=dosya.name, err=str(e)), "err")
        finally:
            ppt.Quit()

        self._finish(ok, skip, converted, do_delete, log_fn, done_fn)

    def _run_with_libreoffice(self, soffice, pptx_files, do_delete, log_fn, done_fn):
        ok, skip, converted = 0, 0, []
        for dosya in pptx_files:
            pdf_yolu = dosya.with_suffix('.pdf')
            if pdf_yolu.exists():
                log_fn(t("log_skipped", name=dosya.name), "info"); skip += 1; continue
            log_fn(t("log_converting", name=dosya.name), "info")
            try:
                result = subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf",
                     "--outdir", str(dosya.parent), str(dosya)],
                    capture_output=True, timeout=120
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

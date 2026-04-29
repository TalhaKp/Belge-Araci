import os
import re
import threading
import tkinter as tk
from tkinter import messagebox
import webbrowser
from pathlib import Path

from core.config import t
from gui.components import _form_row_with_preview, _form_text
from tools.tool_base import ToolBase, _find_libreoffice


class PptxMergerTool(ToolBase):
    title       = t("pptx_merge_title")
    description = t("pptx_merge_desc")
    icon        = "📑"

    def build_form(self, parent, log_fn):
        state = {
            "folder":  tk.StringVar(),
            "outname": tk.StringVar(value="Birlestirilmis_Sunum.pptx"),
        }
        _form_row_with_preview(parent, t("label_folder"),
                               state["folder"], ('.pptx',))
        _form_text(parent, t("label_outname_pptx"), state["outname"])
        return state

    def run(self, state, log_fn, done_fn):
        folder  = state["folder"].get().strip()
        outname = self._sanitize_pptx_name(state["outname"].get())

        def worker():
            if not folder or not os.path.isdir(folder):
                log_fn(t("err_no_folder"), "err"); done_fn(False); return

            log_fn(t("log_scanning"), "info")

            def natural_sort(name):
                return tuple(int(c) if c.isdecimal() else c.lower().strip()
                             for c in re.split(r'(\d+)', name))

            pptx_files = sorted(
                [f for f in os.listdir(folder)
                 if f.lower().endswith('.pptx') and f != outname
                 and not f.startswith('~$')],
                key=natural_sort
            )

            if not pptx_files:
                log_fn(t("err_no_pptx_merge"), "err"); done_fn(False); return

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
                self._run_with_ppt(folder, pptx_files, outname, log_fn, done_fn)
            else:
                log_fn(t("engine_libre"), "info")
                self._run_with_libreoffice(libre_path, folder, pptx_files, outname, log_fn, done_fn)

        threading.Thread(target=worker, daemon=True).start()

    def _run_with_ppt(self, folder, pptx_files, outname, log_fn, done_fn):
        import win32com.client

        try:
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.Visible = True
        except Exception as e:
            log_fn(t("err_ppt_app") + f" ({e})", "err"); done_fn(False); return

        try:
            base_path = str(Path(folder) / pptx_files[0])
            merged = ppt.Presentations.Open(base_path, WithWindow=False)
            slide_count = merged.Slides.Count
            log_fn(t("log_ok", name=f"{pptx_files[0]} ({slide_count} slayt)"), "ok")

            for fname in pptx_files[1:]:
                fpath = str(Path(folder) / fname)
                try:
                    src_prs   = ppt.Presentations.Open(fpath, WithWindow=False)
                    src_count = src_prs.Slides.Count
                    src_prs.Close()

                    merged.Slides.InsertFromFile(
                        fpath,
                        merged.Slides.Count,
                        1,
                        src_count
                    )
                    log_fn(t("log_ok", name=f"{fname} ({src_count} slayt)"), "ok")
                except Exception as e:
                    log_fn(t("log_err", name=fname, err=str(e)), "err")

            out_path = str(Path(folder) / outname)
            merged.SaveAs(out_path, FileFormat=1)
            merged.Close()
            log_fn(t("log_saved", path=out_path), "info")
            log_fn(t("log_summary", ok=len(pptx_files), skip=0), "info")
            done_fn(True)

        except Exception as e:
            log_fn(t("log_err", name=outname, err=str(e)), "err")
            done_fn(False)
        finally:
            try:
                ppt.Quit()
            except Exception:
                pass

    def _run_with_libreoffice(self, soffice, folder, pptx_files, outname, log_fn, done_fn):
        log_fn(
            "⚠️  PPTX birleştirme yalnızca Microsoft PowerPoint ile desteklenmektedir. "
            "LibreOffice ile bu işlem yapılamıyor. Lütfen PowerPoint kurun.",
            "err"
        )
        done_fn(False)

    @staticmethod
    def _sanitize_pptx_name(name: str,
                             fallback: str = "Birlestirilmis_Sunum.pptx") -> str:
        if not name or not name.strip():
            return fallback
        name = re.sub(r'[\x00-\x1f\x7f]', '', name)
        name = re.sub(r'[/\\:*?"<>|]', '', name)
        name = name.replace('..', '').strip('. ')
        if not name:
            return fallback
        stem = os.path.splitext(name)[0].strip('. ')
        if not stem:
            return fallback
        return stem[:200] + ".pptx"

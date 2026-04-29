import os
import tkinter as tk

from core.config import THEME, sc, sf, t
from core.helpers import get_resource_path
from gui.components import HoverButton
from tools import TOOLS
from tools.tool_base import ToolBase


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(t("app_title"))
        self.resizable(True, True)
        self.minsize(sc(520), sc(480))
        self.configure(bg=THEME["bg"])
        self._center(sc(620), sc(680))

        icon_path = get_resource_path("belge_araci.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

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

import glob
import os
import subprocess
import tkinter as tk
from abc import ABC, abstractmethod


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

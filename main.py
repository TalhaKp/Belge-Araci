import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.app import App

if __name__ == "__main__":
    try:
        import ctypes
        myappid = 'talhakp.belgearaci.v1.4.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = App()
    app.mainloop()

"""Entry point for CAT-RG."""

import sys

# Before Tk loads: Windows groups this process with its own taskbar / shortcut icon.
if sys.platform == "win32":
    try:
        import ctypes

        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "PatrickKoebbe.CAT-RG.Application.2.2"
        )
    except Exception:
        pass

import tkinter as tk

from catrg.gui.main_window import CyberTipAnalyzer, _DND_AVAILABLE

try:
    import tkinterdnd2
except ImportError:
    tkinterdnd2 = None


def main() -> None:
    root = None
    if _DND_AVAILABLE and tkinterdnd2 is not None:
        try:
            root = tkinterdnd2.TkinterDnD.Tk()
        except RuntimeError:
            root = None

    if root is None:
        root = tk.Tk()

    app = CyberTipAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()

"""Entry point for CAT-RG."""

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

"""Keep the bundled Qt DLLs ahead of any system/Python Qt installation."""

import os
import sys


if sys.platform.startswith("win"):
    root = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    qt_bin = os.path.join(root, "PyQt6", "Qt6", "bin")
    os.environ["PATH"] = os.pathsep.join(p for p in (qt_bin, root, os.environ.get("PATH", "")) if p)
    _qt_dll_directory = os.add_dll_directory(qt_bin)

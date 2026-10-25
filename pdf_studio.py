"""
PDF Studio - PDF reader/editor for Windows
  * tabs (like a browser): every opened PDF gets its own tab, in order
  * continuous view (scroll through all pages like one document), page x / n in the status bar
  * drag pages from the preview to export as separate PDF or JPG files
  * edit, move and resize text, images and shapes (Edit mode), delete / add / move / rotate pages, insert PDFs, undo
  * built-in print dialog: printer, A0-A5 + custom size, orientation, scale, page range, duplex,
    paper source (auto / bypass tray), paper type,
    save-toner and grayscale modes, flip on left/right or top/bottom edge, and live preview
  * Explorer right-click: combine multiple PDFs/images, or open a single selection directly
  * single window: opening another PDF while the app is running adds a new tab
  * Acrobat-style look, dark grey by default (start with --light for the light theme): document tabs above
    the toolbar, thumbnail panel on the left, grey page canvas, blue pill buttons, red app mark

Requirements:   pip install pymupdf PyQt6
Run:            python pdf_studio.py [file.pdf ...]
Register menu:  python pdf_studio.py --install      (remove: --uninstall)
                adds Explorer entries, the PDF Studio icon for .pdf files, a Start-menu shortcut,
                and lists the app under Settings > Default apps
Icon file:      python pdf_studio.py --export-icon pdfstudio.ico   (used for the .exe build)
"""
import sys, os, re, time, glob, tempfile, math, bisect, string, html as _html, shutil

# PyMuPDF loads its optional page-layout analyser (onnxruntime + numpy) on import when it is installed. PDF Studio
# never uses it, and it costs about half a second on every start, so keep it out. Must run before `import pymupdf`.
os.environ.setdefault("PYMUPDF_SUGGEST_LAYOUT_ANALYZER", "0")
sys.modules.setdefault("pymupdf.layout", None)
sys.modules.setdefault("pymupdf4llm", None)         # markdown export helper (pulls in numpy); also unused

SERVER_NAME = "PDFStudio_single_instance_" + re.sub(r"\W", "_", os.environ.get("USERNAME", os.environ.get("USER", "u")))
APP_USER_MODEL_ID = "PDFStudio.PDFStudio.1"     # Windows taskbar identity (own icon + own button group)


# ============================================================================
#  FAST MERGE MODE  (runs before any Qt import so it starts quickly)
# ============================================================================
IMG_EXT = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".gif")
ALL_EXT = (".pdf",) + IMG_EXT


def _natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", os.path.basename(s))]


def _page_export_root():
    """A stable location for files handed to Explorer or another drop target."""
    return os.path.join(tempfile.gettempdir(), "PDFStudio_page_exports")


def _cleanup_old_page_exports():
    """Keep drag files available after the OLE drag loop; remove only old app-owned exports."""
    root = _page_export_root()
    try:
        os.makedirs(root, exist_ok=True)
        cutoff = time.time() - 7 * 24 * 60 * 60
        for entry in os.scandir(root):
            try:
                if (entry.name.startswith(("drag_", "single_open_"))
                        and entry.is_dir(follow_symlinks=False)
                        and entry.stat(follow_symlinks=False).st_mtime < cutoff):
                    shutil.rmtree(entry.path, ignore_errors=True)
            except OSError:
                pass
    except OSError:
        pass


def _msg(text, icon=0x10):
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, text, "PDF Studio", icon)
    except Exception:
        print(text)


def merge_main(paths):
    """Explorer starts one process per selected file. The first process to grab the
    lock becomes the 'leader': it waits until all the other processes have dropped
    their path into the queue folder, then merges everything once."""
    qdir = os.path.join(tempfile.gettempdir(), "pdfstudio_merge")
    os.makedirs(qdir, exist_ok=True)
    with open(os.path.join(qdir, f"{time.time_ns()}_{os.getpid()}.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(paths))

    lock = os.path.join(qdir, "leader.lock")
    try:
        if time.time() - os.path.getmtime(lock) > 60:      # stale lock
            os.remove(lock)
    except OSError:
        pass
    try:
        os.close(os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except FileExistsError:
        return                                              # another process is the leader

    try:
        last, stable = -1, time.time()
        while True:
            files = glob.glob(os.path.join(qdir, "*.txt"))
            if len(files) != last:
                last, stable = len(files), time.time()
            if time.time() - stable >= 0.7:
                break
            time.sleep(0.1)
        items = []
        for fp in files:
            try:
                with open(fp, encoding="utf-8") as f:
                    items += [l for l in f.read().splitlines() if l.strip()]
                os.remove(fp)
            except OSError:
                pass
    finally:
        try:
            os.remove(lock)
        except OSError:
            pass

    items = sorted({os.path.abspath(p) for p in items if p.lower().endswith(ALL_EXT) and os.path.exists(p)},
                   key=_natural_key)
    if not items:
        _msg("فایل معتبری برای باز کردن پیدا نشد.\nNo supported file was found.", 0x30)
        return

    if len(items) == 1:
        only = items[0]
        if only.lower().endswith(".pdf"):
            return only

        # The viewer opens PDFs, not raw image paths. Wrap a single image in a
        # temporary one-page PDF so Explorer's Combine command opens it directly.
        import pymupdf as fitz
        temp_dir = None
        out = None
        try:
            os.makedirs(_page_export_root(), exist_ok=True)
            temp_dir = tempfile.mkdtemp(prefix="single_open_", dir=_page_export_root())
            base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_",
                          os.path.splitext(os.path.basename(only))[0]).strip(" .")[:80] or "image"
            target = os.path.join(temp_dir, f"{base}.pdf")
            out = fitz.open()
            pm = fitz.Pixmap(only)
            w, h = (842, 595) if pm.width > pm.height else (595, 842)
            page = out.new_page(width=w, height=h)
            k = min((w - 40) / pm.width, (h - 40) / pm.height)
            iw, ih = pm.width * k, pm.height * k
            page.insert_image(
                fitz.Rect(w / 2 - iw / 2, h / 2 - ih / 2, w / 2 + iw / 2, h / 2 + ih / 2),
                filename=only)
            out.save(target, garbage=3, deflate=True)
            return target
        except Exception as e:
            if temp_dir:
                shutil.rmtree(temp_dir, ignore_errors=True)
            _msg(f"فایل باز نشد:\n{e}\nCould not open file.", 0x30)
            return
        finally:
            if out is not None:
                out.close()

    import pymupdf as fitz
    out = fitz.open()
    skipped = []
    for p in items:
        try:
            if p.lower().endswith(IMG_EXT):                 # image -> one fitted page
                pm = fitz.Pixmap(p)
                w, h = (842, 595) if pm.width > pm.height else (595, 842)
                page = out.new_page(width=w, height=h)
                k = min((w - 40) / pm.width, (h - 40) / pm.height)
                iw, ih = pm.width * k, pm.height * k
                page.insert_image(fitz.Rect(w / 2 - iw / 2, h / 2 - ih / 2, w / 2 + iw / 2, h / 2 + ih / 2),
                                  filename=p)
                continue
            with fitz.open(p) as d:
                if d.needs_pass and not d.authenticate(""):
                    skipped.append(os.path.basename(p) + " (password)")
                    continue
                out.insert_pdf(d)
        except Exception as e:
            skipped.append(f"{os.path.basename(p)} ({e})")
    if len(out) == 0:
        _msg("Nothing could be merged.")
        return
    folder = os.path.dirname(items[0])
    target, n = os.path.join(folder, "Combined.pdf"), 1
    while os.path.exists(target):
        n += 1
        target = os.path.join(folder, f"Combined ({n}).pdf")
    out.save(target, garbage=3, deflate=True)
    out.close()
    if skipped:
        _msg("Skipped:\n" + "\n".join(skipped), 0x30)
    return target                                           # opened in the app below


_open_after = None
if len(sys.argv) > 1 and sys.argv[1] == "--merge":
    _open_after = merge_main(sys.argv[2:])
    if not _open_after:                                     # follower process, or nothing to merge
        sys.exit(0)


def _fast_forward(paths):
    """If PDF Studio is already running, give it the files and return True.
    Standard library only, so a second launch (double-click on another PDF) costs a few milliseconds:
    no Qt and no PyMuPDF are imported. The running window opens a new tab and comes to the front."""
    data = "\n".join(paths).encode("utf-8")
    try:
        if sys.platform == "win32":
            import ctypes
            try:
                ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)   # ASFW_ANY: the running window may take focus
            except Exception:
                pass
            pipe = "\\\\.\\pipe\\" + SERVER_NAME                          # where Qt's QLocalServer listens
            for _ in range(20):
                try:
                    with open(pipe, "r+b", buffering=0) as f:        # r+b = OPEN_EXISTING (no create/truncate)
                        f.write(data)
                    return True
                except FileNotFoundError:
                    return False                            # nobody is listening: start normally
                except OSError:
                    time.sleep(0.05)                        # pipe busy: another launch is being handed over
            return False
        import socket
        s = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        s.settimeout(0.5)
        try:
            s.connect(os.path.join(tempfile.gettempdir(), SERVER_NAME))
            s.sendall(data)
            return True
        finally:
            s.close()
    except Exception:
        return False                                        # anything unexpected: fall back to the normal path


if not (len(sys.argv) > 1 and sys.argv[1] in ("--install", "--uninstall", "--export-icon")):
    if _fast_forward([_open_after] if _open_after else
                     [os.path.abspath(a) for a in sys.argv[1:] if not a.startswith("--")]):
        sys.exit(0)


# ============================================================================
#  EXPLORER CONTEXT MENU (registry, current user only - no admin needed)
# ============================================================================
_APP_EXE = None          # path of the installed copy of the exe, set by install_context_menu()


def _app_exe():
    return _APP_EXE or sys.executable


def _install_dir():
    return os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Programs", "PDF Studio")


def _registered_exe():
    """The exe path that Explorer currently launches for PDFs (first quoted token of the open command)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            f"Software\\Classes\\{_APP_ID}\\shell\\open\\command") as k:
            cmd = winreg.QueryValueEx(k, "")[0]
    except OSError:
        return None
    m = re.match(r'\s*"([^"]+)"', cmd)
    return m.group(1) if m else None


def _deploy_exe():
    """Frozen exe only: replace any previous version with this one. Stops running copies, removes the old
    registration, start-menu link and settings, deletes the old exe and the install folder, then copies this exe
    to a fixed per-user folder (no admin needed). Returns the installed exe path, or None when not frozen."""
    if not getattr(sys, "frozen", False):
        return None
    me = os.path.abspath(sys.executable)
    target = os.path.join(_install_dir(), "PDFStudio.exe")
    if os.path.normcase(me) == os.path.normcase(target):
        return target                                       # already running from the installed copy
    import subprocess
    old = _registered_exe()
    try:                                                    # onefile = a launcher process + the real one: spare both
        subprocess.run(["taskkill", "/F", "/IM", "PDFStudio.exe", "/FI", f"PID ne {os.getpid()}",
                        "/FI", f"PID ne {os.getppid()}"], creationflags=0x08000000, timeout=15,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass
    time.sleep(0.5)                                         # let Windows release the old exe
    uninstall_context_menu()
    for stale in (old, target):
        if (stale and os.path.basename(stale).lower() == "pdfstudio.exe"
                and os.path.normcase(os.path.abspath(stale)) != os.path.normcase(me)):
            try:
                os.remove(stale)
            except OSError:
                pass
    shutil.rmtree(_install_dir(), ignore_errors=True)
    os.makedirs(_install_dir(), exist_ok=True)
    shutil.copy2(me, target)
    return target


def _base_cmd():
    if getattr(sys, "frozen", False):
        return f'"{_app_exe()}"'
    py = sys.executable
    if py.lower().endswith("python.exe"):
        py = py[:-10] + "pythonw.exe"                       # no console window
    return f'"{py}" "{os.path.abspath(__file__)}"'


_REG_ROOT = "Software\\Classes\\SystemFileAssociations"
_APP_ID = "PDFStudio.pdf"


def _reg_set(path, name, value):
    import winreg
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
    winreg.SetValueEx(k, name, 0, winreg.REG_SZ, value)


def _reg_delete_tree(path):
    import winreg
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_ALL_ACCESS)
    except OSError:
        return
    while True:
        try:
            sub = winreg.EnumKey(k, 0)
        except OSError:
            break
        _reg_delete_tree(path + "\\" + sub)
    winreg.CloseKey(k)
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except OSError:
        pass


def _notify_assoc_changed():
    """Tell Explorer that file associations/icons changed so it redraws them right away."""
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)      # SHCNE_ASSOCCHANGED
    except Exception:
        pass


def _start_menu_link():
    return os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs",
                        "PDF Studio.lnk")


def _create_start_menu_shortcut(icon_path):
    """Start-menu entry with the PDF Studio icon (uses PowerShell's WScript.Shell; silent on failure)."""
    import subprocess
    if getattr(sys, "frozen", False):
        target, args = _app_exe(), ""
    else:
        py = sys.executable
        if py.lower().endswith("python.exe"):
            py = py[:-10] + "pythonw.exe"
        target, args = py, f'"{os.path.abspath(__file__)}"'
    q = lambda s: s.replace("'", "''")
    ps = (f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{q(_start_menu_link())}');"
          f"$s.TargetPath='{q(target)}';$s.Arguments='{q(args)}';$s.IconLocation='{q(icon_path)},0';"
          f"$s.Description='PDF Studio';$s.WorkingDirectory='{q(os.path.dirname(target))}';$s.Save()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", ps],
                       creationflags=0x08000000, timeout=30)
    except Exception:
        pass


def install_context_menu():
    global _APP_EXE
    _APP_EXE = _deploy_exe() or _APP_EXE                   # new install replaces the old version completely
    cmd = _base_cmd()
    icon = _icon_ref()                                      # exe resource when frozen, else a generated .ico
    # "Combine" in the right-click menu of PDFs and images (works for mixed selections)
    for ext in ALL_EXT:
        base = f"{_REG_ROOT}\\{ext}\\shell\\PDFStudioCombine"
        _reg_set(base, "MUIVerb", "Combine (PDF Studio)")
        _reg_set(base, "Icon", icon)
        _reg_set(base, "MultiSelectModel", "Player")
        _reg_set(base + "\\command", "", f'{cmd} --merge "%1"')
    # "Open with PDF Studio" for single PDFs
    base = f"{_REG_ROOT}\\.pdf\\shell\\PDFStudioEdit"
    _reg_set(base, "MUIVerb", "Open with PDF Studio")
    _reg_set(base, "Icon", icon)
    _reg_set(base + "\\command", "", f'{cmd} "%1"')
    # Register as an installed app so it shows up in Settings > Default apps
    _reg_set(f"Software\\Classes\\{_APP_ID}", "", "PDF Document (PDF Studio)")
    _reg_set(f"Software\\Classes\\{_APP_ID}\\DefaultIcon", "", icon)           # icon of .pdf files in Explorer
    _reg_set(f"Software\\Classes\\{_APP_ID}\\shell\\open\\command", "", f'{cmd} "%1"')
    _reg_set("Software\\PDFStudio\\Capabilities", "ApplicationName", "PDF Studio")
    _reg_set("Software\\PDFStudio\\Capabilities", "ApplicationDescription", "PDF reader and editor")
    _reg_set("Software\\PDFStudio\\Capabilities", "ApplicationIcon", icon)
    _reg_set("Software\\PDFStudio\\Capabilities\\FileAssociations", ".pdf", _APP_ID)
    _reg_set("Software\\RegisteredApplications", "PDF Studio", "Software\\PDFStudio\\Capabilities")
    _reg_set(f"Software\\Classes\\.pdf\\OpenWithProgids", _APP_ID, "")
    _create_start_menu_shortcut(icon.rsplit(",", 1)[0].strip('"'))
    _notify_assoc_changed()


def refresh_registration_if_stale():
    """Explorer launches PDF Studio with the exe path that was written to the registry at install time. If that
    file is gone (moved, deleted or replaced), double-clicking a PDF starts nothing. When an earlier install
    exists (never after --uninstall) and its exe no longer exists, install again from the exe that is running now.
    Runs in the background."""
    if sys.platform != "win32":
        return
    registered = _registered_exe()
    if registered is None or os.path.exists(registered):
        return                                              # never installed / uninstalled, or still valid
    import threading
    threading.Thread(target=install_context_menu, daemon=True).start()


def uninstall_context_menu():
    import winreg
    for ext in ALL_EXT:
        _reg_delete_tree(f"{_REG_ROOT}\\{ext}\\shell\\PDFStudioCombine")
    for old in ("PDFStudioMerge", "PDFStudioEdit"):
        _reg_delete_tree(f"{_REG_ROOT}\\.pdf\\shell\\{old}")
    _reg_delete_tree(f"Software\\Classes\\{_APP_ID}")
    _reg_delete_tree("Software\\PDFStudio")
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\RegisteredApplications", 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(k, "PDF Studio")
    except OSError:
        pass
    try:
        os.remove(_start_menu_link())
    except OSError:
        pass
    if (getattr(sys, "frozen", False) and os.path.isdir(_install_dir())
            and not os.path.normcase(os.path.abspath(sys.executable)).startswith(os.path.normcase(_install_dir()))):
        shutil.rmtree(_install_dir(), ignore_errors=True)   # the installed copy, unless it is the one running
    _notify_assoc_changed()


# ============================================================================
#  GUI
# ============================================================================
import pymupdf as fitz
from PyQt6.QtCore import (Qt, QSize, QSizeF, QPoint, QPointF, QRectF, QMarginsF, QTimer, QEvent,
                          QMimeData, QObject, QSettings, QUrl, pyqtSignal)
from PyQt6.QtGui import (QAction, QBrush, QColor, QCursor, QDrag, QFont, QIcon, QImage, QKeySequence, QLinearGradient,
                         QPalette, QPixmap, QPainter, QPen, QPainterPath, QPageLayout, QPageSize, QRadialGradient,
                         QTextCursor, QTransform, QIntValidator)
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
from PyQt6.QtWidgets import (QApplication, QButtonGroup, QComboBox, QDialog, QDialogButtonBox,
                             QCheckBox, QDoubleSpinBox, QFileDialog, QFormLayout, QFrame, QGraphicsDropShadowEffect,
                             QGridLayout, QGroupBox, QHBoxLayout,
                             QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem,
                             QMainWindow, QMenu, QMessageBox, QPlainTextEdit, QProgressDialog, QPushButton, QRadioButton,
                             QScrollArea, QAbstractItemView, QSpinBox, QSplitter,
                             QStackedWidget, QTabBar, QTabWidget, QToolBar, QToolButton, QVBoxLayout, QWidget)

PAGE_GAP = 12      # px between pages in the continuous view
PAGE_MARGIN = 12   # px around the whole document
PRINT_MARGIN_MM = 5
CELL_GAP_PT = 10   # gap between tiles when printing several pages per sheet
PAPERS = {"A0": QPageSize.PageSizeId.A0, "A1": QPageSize.PageSizeId.A1, "A2": QPageSize.PageSizeId.A2,
          "A3": QPageSize.PageSizeId.A3, "A4": QPageSize.PageSizeId.A4, "A5": QPageSize.PageSizeId.A5}
CUSTOM_PAPER = "custom"
RENDER_DPI_DEFAULT = 300
MAX_BAND_PIXELS = 25_000_000      # big sheets (A0, plotter rolls) are rendered in strips of at most this size


def page_size_for(paper, custom_mm=(210.0, 297.0)):
    """QPageSize for a PAPERS key, or for the manual width x height (mm) when paper == CUSTOM_PAPER."""
    if paper == CUSTOM_PAPER:
        return QPageSize(QSizeF(float(custom_mm[0]), float(custom_mm[1])), QPageSize.Unit.Millimeter,
                         "Custom", QPageSize.SizeMatchPolicy.ExactMatch)
    return QPageSize(PAPERS[paper])


def pix_to_qimage(pix):
    fmt = QImage.Format.Format_RGBA8888 if pix.alpha else QImage.Format.Format_RGB888
    return QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()


def icon_pixmap(n=256):
    """PDF Studio mark at n x n pixels: a red rounded square with a white folded page (drawn in code)."""
    S = 256
    pix = QPixmap(n, n)
    pix.fill(Qt.GlobalColor.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
    p.scale(n / S, n / S)
    p.setPen(Qt.PenStyle.NoPen)
    r = QRectF(24, 20, S - 48, S - 48)
    rad = 46
    for k, a in ((9, 14), (5, 24), (2, 36)):                        # soft shadow under the square
        p.setBrush(QColor(0, 0, 0, a))
        p.drawRoundedRect(r.adjusted(-k + 2, -k + 8, k - 2, k + 8), rad + k, rad + k)
    body = QLinearGradient(r.topLeft(), r.bottomLeft())
    body.setColorAt(0.0, QColor("#F2392D"))
    body.setColorAt(1.0, QColor("#D3170C"))
    p.setBrush(body)
    p.drawRoundedRect(r, rad, rad)
    # page with a folded corner
    pg = QRectF(r.center().x() - 46, r.center().y() - 60, 92, 120)
    fold = 28
    path = QPainterPath()
    path.moveTo(pg.left() + 10, pg.top())
    path.lineTo(pg.right() - fold, pg.top())
    path.lineTo(pg.right(), pg.top() + fold)
    path.lineTo(pg.right(), pg.bottom() - 10)
    path.quadTo(pg.right(), pg.bottom(), pg.right() - 10, pg.bottom())
    path.lineTo(pg.left() + 10, pg.bottom())
    path.quadTo(pg.left(), pg.bottom(), pg.left(), pg.bottom() - 10)
    path.lineTo(pg.left(), pg.top() + 10)
    path.quadTo(pg.left(), pg.top(), pg.left() + 10, pg.top())
    p.setBrush(QColor(255, 255, 255))
    p.drawPath(path)
    p.setBrush(QColor(214, 214, 214))
    tri = QPainterPath()
    tri.moveTo(pg.right() - fold, pg.top())
    tri.lineTo(pg.right() - fold, pg.top() + fold)
    tri.lineTo(pg.right(), pg.top() + fold)
    tri.closeSubpath()
    p.drawPath(tri)
    p.setBrush(QColor("#E0261A"))
    for i, wd in enumerate((56, 56, 36)):
        p.drawRoundedRect(QRectF(pg.left() + 18, pg.top() + 52 + i * 18, wd, 8), 4, 4)
    p.end()
    return pix


_ICON = None


def app_icon():
    """The application icon with every size Windows asks for (title bar, taskbar, Alt-Tab, dialogs)."""
    global _ICON
    if _ICON is None:
        _ICON = QIcon()
        for n in (16, 20, 24, 32, 40, 48, 64, 96, 128, 256):
            _ICON.addPixmap(icon_pixmap(n))
    return _ICON


def save_ico(path, sizes=(16, 24, 32, 48, 64, 128, 256)):
    """Write a multi-size .ico (PNG-compressed entries) for the exe, Explorer, shortcuts and file types."""
    import struct
    from PyQt6.QtCore import QBuffer, QIODevice
    imgs = []
    for n in sizes:
        buf = QBuffer()
        buf.open(QIODevice.OpenModeFlag.WriteOnly)
        icon_pixmap(n).save(buf, "PNG")
        imgs.append((n, bytes(buf.data())))
    head = struct.pack("<HHH", 0, 1, len(imgs))
    entries, blobs, off = b"", b"", 6 + 16 * len(imgs)
    for n, data in imgs:
        entries += struct.pack("<BBBBHHII", n % 256, n % 256, 0, 0, 1, 32, len(data), off + len(blobs))
        blobs += data
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "wb") as f:
        f.write(head + entries + blobs)


def _icon_ref():
    """Icon location string for the registry: the exe itself when frozen, otherwise a generated .ico file."""
    if getattr(sys, "frozen", False):
        return f'"{_app_exe()}",0'
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "PDFStudio", "pdfstudio.ico")
    save_ico(path)
    return f'"{path}",0'


# ============================================================================
#  ACROBAT-STYLE THEME  (light chrome, white toolbar under a tab strip, grey page canvas,
#  pill-shaped blue primary buttons, red app mark)
# ============================================================================
_THEMES = {
    "dark": dict(
        BG="#3C3C3C", SURFACE="#444444", STRIP="#303030", CANVAS="#353535", PANEL="#3C3C3C",
        LINE="#565656", LINE2="#6A6A6A", HOVER="#505050", PRESS="#5C5C5C", SEL="#24456E", SEL_TEXT="#9CC7FF",
        TEXT="#E6E6E6", TEXT2="#BDBDBD", MUTED="#9A9A9A", DISABLED="#6A6A6A", TITLE="#FFFFFF",
        FIELD="#333333", FIELD_BORDER="#7A7A7A", FIELD_BORDER_H="#B0B0B0",
        BTN_BG="#444444", BTN_BORDER="#8E8E8E", BTN_BORDER_H="#E6E6E6", BTN_HOVER="#505050", BTN_PRESS="#5C5C5C",
        DIS_BG="#3C3C3C", DIS_BORDER="#565656", PRIM_DIS_BG="#4C4C4C", TAB_HOVER="#3A3A3A",
        SB_TRACK="#3C3C3C", SB_HANDLE="#6A6A6A", SB_HANDLE_H="#8A8A8A", CLOSE_HOVER="#5C5C5C", CHIP="#4C4C4C",
        SELECTION="#2F5C9E", TRACK="#565656", TIP_BG="#4C4C4C", TIP_BORDER="#6A6A6A", MENU_SEL_TEXT="#FFFFFF",
        CHECK_COL="#E6E6E6", ARROW="#CFCFCF",
        ACC="#2680EB", ACC_HOVER="#4B9CF5", ACC_PRESS="#1473E6", DRAG="#378EF0",
        BODY="#4A4A4A", BODY_PANEL="#3B3B3B", BODY_TEXT="#BDBDBD", BODY_HOVER="#505050",
        BODY_SB_TRACK="#3B3B3B", BODY_SB_HANDLE="#6E6E6E", BODY_SB_HANDLE_H="#8C8C8C",
        RIM=(255, 255, 255, 38), THUMB_RIM=(255, 255, 255, 70), DARK=True),
    "light": dict(
        BG="#F4F4F4", SURFACE="#FFFFFF", STRIP="#EBEBEB", CANVAS="#E4E4E4", PANEL="#F4F4F4",
        LINE="#D8D8D8", LINE2="#BDBDBD", HOVER="#EDEDED", PRESS="#DCDCDC", SEL="#D6E6FB", SEL_TEXT="#0D66D0",
        TEXT="#2C2C2C", TEXT2="#4B4B4B", MUTED="#6E6E6E", DISABLED="#A6A6A6", TITLE="#1F1F1F",
        FIELD="#FFFFFF", FIELD_BORDER="#909090", FIELD_BORDER_H="#505050",
        BTN_BG="#FFFFFF", BTN_BORDER="#6E6E6E", BTN_BORDER_H="#2C2C2C", BTN_HOVER="#F0F0F0", BTN_PRESS="#E1E1E1",
        DIS_BG="#F4F4F4", DIS_BORDER="#D0D0D0", PRIM_DIS_BG="#E1E1E1", TAB_HOVER="#F7F7F7",
        SB_TRACK="#F0F0F0", SB_HANDLE="#C1C1C1", SB_HANDLE_H="#9E9E9E", CLOSE_HOVER="#E1E1E1", CHIP="#F0F0F0",
        SELECTION="#B7D4F9", TRACK="#E1E1E1", TIP_BG="#FFFFFF", TIP_BORDER="#8E8E8E", MENU_SEL_TEXT="#000000",
        CHECK_COL="#2C2C2C", ARROW="#4B4B4B",
        ACC="#1473E6", ACC_HOVER="#0D66D0", ACC_PRESS="#095ABA", DRAG="#1473E6",
        BODY="#4A4A4A", BODY_PANEL="#3B3B3B", BODY_TEXT="#D4D4D4", BODY_HOVER="#4A4A4A",   # dark-grey body (viewer + thumbnails)
        BODY_SB_TRACK="#3B3B3B", BODY_SB_HANDLE="#6E6E6E", BODY_SB_HANDLE_H="#8C8C8C",
        RIM=(0, 0, 0, 60), THUMB_RIM=(0, 0, 0, 70), DARK=False),
}
# dark grey is the default; start with  python pdf_studio.py --light  for the light theme
T = _THEMES["light" if "--light" in sys.argv else "dark"]
ACCENT = T["DRAG"]             # drop lines / auto-scroll mark


def _theme_assets():
    """Tiny PNGs (check mark, radio dot, chevrons) that the stylesheet points at."""
    d = os.path.join(tempfile.gettempdir(), "pdfstudio_theme_" + ("dark" if T["DARK"] else "light") + "_v3")
    os.makedirs(d, exist_ok=True)
    out = {}

    def make(name, draw):
        path = os.path.join(d, name + ".png")
        pm = QPixmap(32, 32)
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        draw(p)
        p.end()
        pm.save(path, "PNG")
        out[name] = path.replace("\\", "/")

    def pen(p, color, w=3.6):
        q = QPen(QColor(color), w)
        q.setCapStyle(Qt.PenCapStyle.RoundCap)
        q.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        p.setPen(q)
        p.setBrush(Qt.BrushStyle.NoBrush)

    def check(p):                      # white tick on the blue check box
        pen(p, "#FFFFFF", 4.2)
        p.drawPolyline([QPointF(8, 17), QPointF(14, 23), QPointF(24.5, 9.5)])

    def checkd(p):                     # tick for checked menu items
        pen(p, T["CHECK_COL"], 4.2)
        p.drawPolyline([QPointF(8, 17), QPointF(14, 23), QPointF(24.5, 9.5)])

    def dot(p):
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(255, 255, 255))
        p.drawEllipse(QPointF(16, 16), 6, 6)

    def down(p):
        pen(p, T["ARROW"])
        p.drawPolyline([QPointF(9, 12.5), QPointF(16, 19.5), QPointF(23, 12.5)])

    def up(p):
        pen(p, T["ARROW"])
        p.drawPolyline([QPointF(9, 19.5), QPointF(16, 12.5), QPointF(23, 19.5)])

    for n, fn in (("check", check), ("checkd", checkd), ("dot", dot), ("down", down), ("up", up)):
        make(n, fn)
    return out


_THEME_QSS = r"""
QToolTip { color:$TEXT; background:$TIP_BG; border:1px solid $TIP_BORDER; padding:4px 8px; }

QDialog, QMessageBox, QInputDialog, QProgressDialog { background:$BG; color:$TEXT; }
QMainWindow#AppWindow, QDialog#AppWindow { background:$BG; }
QLabel { color:$TEXT; }

/* ---------- menu bar + menus ---------- */
QMenuBar { background:$SURFACE; border-bottom:1px solid $LINE; padding:0 6px; spacing:0; }
QMenuBar::item { background:transparent; padding:6px 11px; border-radius:0; color:$TEXT; }
QMenuBar::item:selected { background:$HOVER; }
QMenuBar::item:pressed { background:$PRESS; }
QMenu { background:$SURFACE; border:1px solid $LINE2; padding:4px 0; color:$TEXT; }
QMenu::item { background:transparent; padding:6px 34px 6px 28px; margin:0; }
QMenu::item:selected { background:$SEL; color:$MENU_SEL_TEXT; }
QMenu::item:disabled { background:transparent; color:$DISABLED; }
QMenu::separator { height:1px; background:$LINE; margin:4px 0; }
QMenu::indicator { width:14px; height:14px; left:7px; }
QMenu::indicator:checked { image:url($CHECKD); }

/* ---------- toolbar (strip under the tabs) ---------- */
QToolBar { background:$SURFACE; border:none; border-top:1px solid $LINE; border-bottom:1px solid $LINE;
           padding:3px 8px; spacing:2px; }
QToolBar QToolButton { background:transparent; border:1px solid transparent; border-radius:4px;
                       padding:5px 10px; margin:0 1px; color:$TEXT; }
QToolBar QToolButton:hover { background:$HOVER; }
QToolBar QToolButton:pressed { background:$PRESS; }
QToolBar QToolButton:checked { background:$SEL; border-color:$ACC; color:$SEL_TEXT; }
QToolBar QToolButton#qt_toolbar_ext_button { padding:0; margin:0 0 0 4px; min-width:24px; }
QToolBar::separator { background:$LINE; width:1px; margin:6px 6px; }

/* ---------- document tabs ---------- */
QTabBar { background:$STRIP; qproperty-drawBase:0; }
QTabBar::tab { background:transparent; border:1px solid transparent; border-bottom:none;
               border-top-left-radius:4px; border-top-right-radius:4px;
               padding:7px 6px 7px 14px; margin:6px 1px 0 1px; min-width:90px; max-width:240px; color:$TEXT2; }
QTabBar::tab:hover { background:$TAB_HOVER; color:$TEXT; }
QTabBar::tab:selected { background:$SURFACE; border:1px solid $LINE; border-bottom:none;
                        border-top:2px solid $ACC; color:$TITLE; }
QToolButton#TabClose { background:transparent; border:none; border-radius:3px; padding:2px 5px;
                       color:$MUTED; font-size:8pt; }
QToolButton#TabClose:hover { background:$CLOSE_HOVER; color:$TITLE; }

/* ---------- status bar ---------- */
QStatusBar { background:$SURFACE; border-top:1px solid $LINE; color:$TEXT2; padding:0 8px; }
QStatusBar::item { border:none; }
QLabel#StatusPill { background:transparent; border:none; padding:3px 12px; color:$TEXT; }
QWidget#PageBar { background:$SURFACE; }
QLabel#PageText { background:transparent; border:none; color:$TEXT2; font-size:15px; font-weight:600; }
QLineEdit#PageBox { font-size:16px; font-weight:700; padding:1px 6px; min-height:24px; border-radius:6px; }

/* ---------- panels ---------- */
QListWidget#Thumbs { background:$BODY_PANEL; border:none; padding:6px 4px; color:$BODY_TEXT; outline:0; }
QListWidget#Thumbs::item { border-radius:3px; color:$BODY_TEXT; }
QListWidget#Thumbs::item:hover { background:$BODY_HOVER; }
QListWidget#Thumbs::item:selected { background:$SEL; color:$SEL_TEXT; }
QListWidget#Thumbs::item:selected:!active { background:$SEL; color:$SEL_TEXT; }
QScrollArea#PageScroll { background:$BODY; border:none; }
QListWidget#Thumbs QScrollBar:vertical, QScrollArea#PageScroll QScrollBar:vertical,
QScrollArea#PageScroll QScrollBar:horizontal { background:$BODY_SB_TRACK; }
QListWidget#Thumbs QScrollBar::handle:vertical, QScrollArea#PageScroll QScrollBar::handle:vertical,
QScrollArea#PageScroll QScrollBar::handle:horizontal { background:$BODY_SB_HANDLE; }
QListWidget#Thumbs QScrollBar::handle:vertical:hover, QScrollArea#PageScroll QScrollBar::handle:vertical:hover,
QScrollArea#PageScroll QScrollBar::handle:horizontal:hover { background:$BODY_SB_HANDLE_H; }
QScrollArea#PageScroll > QWidget > QWidget { background:transparent; }
QSplitter::handle { background:$LINE; }
QSplitter::handle:hover { background:$LINE2; }

QFrame#EmptyCard { background:$SURFACE; border:1px solid $LINE; border-radius:8px; }
QLabel#EmptyTitle { font-size:22pt; font-weight:700; color:$TITLE; background:transparent; }
QLabel#EmptySub { color:$MUTED; background:transparent; }
QLabel#KeyHint { background:$CHIP; border:1px solid $LINE; border-radius:4px; padding:3px 10px; color:$MUTED; }

/* ---------- scrollbars ---------- */
QScrollBar:vertical { background:$SB_TRACK; width:13px; margin:0; }
QScrollBar::handle:vertical { background:$SB_HANDLE; border-radius:3px; min-height:40px; margin:2px 2px; }
QScrollBar::handle:vertical:hover { background:$SB_HANDLE_H; }
QScrollBar:horizontal { background:$SB_TRACK; height:13px; margin:0; }
QScrollBar::handle:horizontal { background:$SB_HANDLE; border-radius:3px; min-width:40px; margin:2px 2px; }
QScrollBar::handle:horizontal:hover { background:$SB_HANDLE_H; }
QScrollBar::add-line, QScrollBar::sub-line { width:0; height:0; background:none; border:none; }
QScrollBar::add-page, QScrollBar::sub-page { background:transparent; }

/* ---------- controls ---------- */
QPushButton { background:$BTN_BG; border:1px solid $BTN_BORDER; border-radius:14px; padding:5px 18px; color:$TEXT; }
QPushButton:hover { background:$BTN_HOVER; border-color:$BTN_BORDER_H; }
QPushButton:pressed { background:$BTN_PRESS; border-color:$BTN_BORDER_H; }
QPushButton:disabled { color:$DISABLED; background:$DIS_BG; border-color:$DIS_BORDER; }
QPushButton#NavBtn { padding:0; border-radius:4px; min-height:24px; }
QPushButton#RotBtn { padding:0; font-size:22px; font-weight:bold; }
QPushButton#Primary { background:$ACC; border:1px solid $ACC; color:#FFFFFF; font-weight:600; }
QPushButton#Primary:hover { background:$ACC_HOVER; border-color:$ACC_HOVER; }
QPushButton#Primary:pressed { background:$ACC_PRESS; border-color:$ACC_PRESS; }
QPushButton#Primary:disabled { background:$PRIM_DIS_BG; border-color:$PRIM_DIS_BG; color:$DISABLED; }

QComboBox, QSpinBox, QDoubleSpinBox, QLineEdit { background:$FIELD; border:1px solid $FIELD_BORDER; border-radius:4px;
    padding:4px 8px; min-height:20px; color:$TEXT; selection-background-color:$SELECTION; selection-color:#FFFFFF; }
QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover, QLineEdit:hover { border-color:$FIELD_BORDER_H; }
QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QLineEdit:focus { border:1px solid $ACC; }
QComboBox:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QLineEdit:disabled { color:$DISABLED; background:$DIS_BG;
    border-color:$DIS_BORDER; }
QComboBox { combobox-popup:0; padding-right:26px; }
QComboBox::drop-down { subcontrol-origin:padding; subcontrol-position:center right; width:26px; border:none; }
QComboBox::down-arrow { image:url($DOWN); width:14px; height:14px; }
QComboBox QAbstractItemView { background:$SURFACE; border:1px solid $LINE2; border-radius:0; padding:2px; outline:0;
    color:$TEXT; selection-background-color:$SEL; selection-color:$MENU_SEL_TEXT; }
QSpinBox::up-button, QDoubleSpinBox::up-button { subcontrol-origin:border; subcontrol-position:top right;
    width:20px; border:none; background:transparent; }
QSpinBox::down-button, QDoubleSpinBox::down-button { subcontrol-origin:border; subcontrol-position:bottom right;
    width:20px; border:none; background:transparent; }
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow { image:url($UP); width:10px; height:10px; }
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow { image:url($DOWN); width:10px; height:10px; }

QGroupBox { background:$SURFACE; border:1px solid $LINE; border-radius:4px; margin:0;
            padding:34px 12px 12px 12px; font-weight:600; }
QGroupBox::title { subcontrol-origin:padding; subcontrol-position:top left; left:14px; top:9px;
                   padding:0; color:$TITLE; background:transparent; }

QCheckBox, QRadioButton { spacing:8px; background:transparent; color:$TEXT; }
QCheckBox:disabled, QRadioButton:disabled { color:$DISABLED; }
QCheckBox::indicator { width:16px; height:16px; border-radius:2px; border:1px solid $BTN_BORDER; background:$FIELD; }
QCheckBox::indicator:hover { border-color:$BTN_BORDER_H; }
QCheckBox::indicator:checked { background:$ACC; border-color:$ACC; image:url($CHECK); }
QRadioButton::indicator { width:16px; height:16px; border-radius:9px; border:1px solid $BTN_BORDER; background:$FIELD; }
QRadioButton::indicator:hover { border-color:$BTN_BORDER_H; }
QRadioButton::indicator:checked { background:$ACC; border-color:$ACC; image:url($DOT); }

QProgressBar { background:$TRACK; border:none; border-radius:4px; text-align:center; min-height:8px; color:$TEXT; }
QProgressBar::chunk { background:$ACC; border-radius:4px; }
"""


def _theme_palette():
    C = QColor
    pal = QPalette()
    for role, col in ((QPalette.ColorRole.Window, T["BG"]), (QPalette.ColorRole.WindowText, T["TEXT"]),
                      (QPalette.ColorRole.Base, T["FIELD"]), (QPalette.ColorRole.AlternateBase, T["PANEL"]),
                      (QPalette.ColorRole.ToolTipBase, T["TIP_BG"]), (QPalette.ColorRole.ToolTipText, T["TEXT"]),
                      (QPalette.ColorRole.Text, T["TEXT"]), (QPalette.ColorRole.Button, T["BTN_BG"]),
                      (QPalette.ColorRole.ButtonText, T["TEXT"]),
                      (QPalette.ColorRole.BrightText, "#FFFFFF" if T["DARK"] else "#000000"),
                      (QPalette.ColorRole.Highlight, T["ACC"]), (QPalette.ColorRole.HighlightedText, "#FFFFFF"),
                      (QPalette.ColorRole.Link, T["DRAG"])):
        pal.setColor(role, C(col))
    ph = C(T["TEXT"])
    ph.setAlpha(110)
    pal.setColor(QPalette.ColorRole.PlaceholderText, ph)
    for role in (QPalette.ColorRole.Text, QPalette.ColorRole.ButtonText, QPalette.ColorRole.WindowText):
        pal.setColor(QPalette.ColorGroup.Disabled, role, C(T["DISABLED"]))
    return pal


def apply_theme(app):
    """Fusion style + palette + Acrobat-style stylesheet (dark grey by default) for the whole application."""
    app.setStyle("Fusion")
    app.setPalette(_theme_palette())
    f = QFont()
    f.setFamilies(["Adobe Clean", "Segoe UI Variable Text", "Segoe UI", "Noto Sans", "Helvetica Neue", "Arial"])
    f.setPointSizeF(9.5)
    f.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
    app.setFont(f)
    a = _theme_assets()
    tokens = {k: v for k, v in T.items() if isinstance(v, str)}
    qss = string.Template(_THEME_QSS).safe_substitute(
        **tokens, CHECK=a["check"], CHECKD=a["checkd"], DOT=a["dot"], DOWN=a["down"], UP=a["up"])
    app.setStyleSheet(qss)


class DarkTitleBarMixin:
    """Dark theme: ask Windows 10/11 for a dark title bar on this top-level window (no-op elsewhere)."""
    def showEvent(self, e):
        super().showEvent(e)
        if T["DARK"] and sys.platform == "win32":
            try:
                import ctypes
                from ctypes import wintypes
                dwm = ctypes.windll.dwmapi
                dwm.DwmSetWindowAttribute.argtypes = [wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
                v = ctypes.c_int(1)
                h = wintypes.HWND(int(self.winId()))
                for attr in (20, 19):                        # DWMWA_USE_IMMERSIVE_DARK_MODE (new / old id)
                    if dwm.DwmSetWindowAttribute(h, attr, ctypes.byref(v), ctypes.sizeof(v)) == 0:
                        break
            except Exception:
                pass


class EmptyState(QWidget):
    """What you see before a PDF is open: one centred card with an Open button."""
    def __init__(self, on_open):
        super().__init__()
        card = QFrame()
        card.setObjectName("EmptyCard")
        card.setFixedWidth(440)
        sh = QGraphicsDropShadowEffect(card)
        sh.setBlurRadius(36)
        sh.setOffset(0, 8)
        sh.setColor(QColor(0, 0, 0, 120 if T["DARK"] else 40))
        card.setGraphicsEffect(sh)
        icon = QLabel()
        icon.setPixmap(app_icon().pixmap(96, 96))
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title = QLabel("PDF Studio")
        title.setObjectName("EmptyTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub = QLabel("Open a PDF, or drop PDF files anywhere in this window")
        sub.setObjectName("EmptySub")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sub.setWordWrap(True)
        btn = QPushButton("Open PDF…")
        btn.setObjectName("Primary")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setMinimumHeight(34)
        btn.clicked.connect(lambda: on_open())
        hint = QLabel("Ctrl + O")
        hint.setObjectName("KeyHint")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(36, 34, 36, 30)
        lay.setSpacing(10)
        lay.addWidget(icon)
        lay.addWidget(title)
        lay.addWidget(sub)
        lay.addSpacing(10)
        lay.addWidget(btn)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(hint)
        row.addStretch(1)
        lay.addLayout(row)
        outer = QVBoxLayout(self)
        outer.addStretch(1)
        outer.addWidget(card, 0, Qt.AlignmentFlag.AlignHCenter)
        outer.addStretch(1)


# ---- Windows printer-driver access: paper source, paper type, print quality ----------
import ctypes

_DM_IN_PROMPT, _DM_OUT_BUFFER, _DM_IN_BUFFER = 4, 2, 8
_DM_ORIENTATION, _DM_PAPERSIZE, _DM_PAPERLENGTH, _DM_PAPERWIDTH = 0x1, 0x2, 0x4, 0x8
_DM_COPIES, _DM_DEFAULTSOURCE, _DM_PRINTQUALITY, _DM_DUPLEX, _DM_YRESOLUTION = (
    0x100, 0x200, 0x400, 0x1000, 0x2000)
_DM_COLOR = 0x800
_DMCOLOR_MONOCHROME, _DMCOLOR_COLOR = 1, 2
_DM_MEDIATYPE = 0x04000000
_DC_BINS, _DC_BINNAMES, _DC_ENUMRESOLUTIONS, _DC_MEDIATYPENAMES, _DC_MEDIATYPES = 6, 12, 13, 34, 35
DMRES_DRAFT, DMRES_MEDIUM, DMRES_HIGH = -1, -3, -4
# (label, DEVMODE dmPrintQuality, dpi used to render the page)
QUALITY_PRESETS = (("Draft (fast, saves ink)", DMRES_DRAFT, 150),
                   ("Normal", DMRES_MEDIUM, 300),
                   ("High", DMRES_HIGH, 600))


class _DEVMODEW(ctypes.Structure):          # printer half of the Windows DEVMODEW (220 bytes)
    _fields_ = [("dmDeviceName", ctypes.c_uint16 * 32), ("dmSpecVersion", ctypes.c_uint16),
                ("dmDriverVersion", ctypes.c_uint16), ("dmSize", ctypes.c_uint16),
                ("dmDriverExtra", ctypes.c_uint16), ("dmFields", ctypes.c_uint32),
                ("dmOrientation", ctypes.c_int16), ("dmPaperSize", ctypes.c_int16),
                ("dmPaperLength", ctypes.c_int16), ("dmPaperWidth", ctypes.c_int16),
                ("dmScale", ctypes.c_int16), ("dmCopies", ctypes.c_int16),
                ("dmDefaultSource", ctypes.c_int16), ("dmPrintQuality", ctypes.c_int16),
                ("dmColor", ctypes.c_int16), ("dmDuplex", ctypes.c_int16),
                ("dmYResolution", ctypes.c_int16), ("dmTTOption", ctypes.c_int16),
                ("dmCollate", ctypes.c_int16), ("dmFormName", ctypes.c_uint16 * 32),
                ("dmLogPixels", ctypes.c_uint16), ("dmBitsPerPel", ctypes.c_uint32),
                ("dmPelsWidth", ctypes.c_uint32), ("dmPelsHeight", ctypes.c_uint32),
                ("dmDisplayFlags", ctypes.c_uint32), ("dmDisplayFrequency", ctypes.c_uint32),
                ("dmICMMethod", ctypes.c_uint32), ("dmICMIntent", ctypes.c_uint32),
                ("dmMediaType", ctypes.c_uint32), ("dmDitherType", ctypes.c_uint32),
                ("dmReserved1", ctypes.c_uint32), ("dmReserved2", ctypes.c_uint32),
                ("dmPanningWidth", ctypes.c_uint32), ("dmPanningHeight", ctypes.c_uint32)]


_WIN_PAPER_IDS = {"A3": 8, "A4": 9, "A5": 11, "A2": 66}
_WIN_PAPER_CUSTOM_MM = {"A0": (841.0, 1189.0), "A1": (594.0, 841.0)}


def _configure_devmode_page(dm, paper, custom_mm, landscape):
    """Write the selected paper and orientation to Windows' native printer settings."""
    if paper == CUSTOM_PAPER or paper in _WIN_PAPER_CUSTOM_MM:
        try:
            dims = _WIN_PAPER_CUSTOM_MM[paper] if paper in _WIN_PAPER_CUSTOM_MM else custom_mm
            width, height = (float(dims[0]), float(dims[1]))
            # DEVMODE stores custom dimensions as signed tenths of a millimetre.
            if 0 < width <= 3276.7 and 0 < height <= 3276.7:
                dm.dmPaperSize = 0
                dm.dmPaperWidth = round(width * 10)
                dm.dmPaperLength = round(height * 10)
                dm.dmFields |= _DM_PAPERSIZE | _DM_PAPERWIDTH | _DM_PAPERLENGTH
            else:
                return False
        except (TypeError, ValueError, IndexError):
            return False
    elif paper in _WIN_PAPER_IDS:
        dm.dmPaperSize = _WIN_PAPER_IDS[paper]
        dm.dmPaperWidth = dm.dmPaperLength = 0
        dm.dmFields = (dm.dmFields & ~(_DM_PAPERWIDTH | _DM_PAPERLENGTH)) | _DM_PAPERSIZE
    else:
        return False

    if landscape is not None:
        dm.dmOrientation = 2 if landscape else 1
        dm.dmFields |= _DM_ORIENTATION
    return True


_DM_MEDIATYPE_END = _DEVMODEW.dmMediaType.offset + 4
_DM_COLOR_END = _DEVMODEW.dmColor.offset + 2


class _WinSpool:
    """Small ctypes wrapper over winspool.drv. `_WinSpool.get()` is None when not on Windows."""
    _inst = None

    @classmethod
    def get(cls):
        if cls._inst is None:
            try:
                cls._inst = cls()
            except (ImportError, AttributeError, OSError):
                cls._inst = False
        return cls._inst or None

    def __init__(self):
        from ctypes import wintypes as w
        self.w = w
        lib = self.lib = ctypes.WinDLL("winspool.drv")
        lib.OpenPrinterW.argtypes = [w.LPWSTR, ctypes.POINTER(w.HANDLE), w.LPVOID]
        lib.OpenPrinterW.restype = w.BOOL
        lib.ClosePrinter.argtypes = [w.HANDLE]
        lib.ClosePrinter.restype = w.BOOL
        lib.DocumentPropertiesW.argtypes = [w.HWND, w.HANDLE, w.LPWSTR, w.LPVOID, w.LPVOID, w.DWORD]
        lib.DocumentPropertiesW.restype = ctypes.c_long
        lib.DeviceCapabilitiesW.argtypes = [w.LPCWSTR, w.LPCWSTR, w.WORD, w.LPVOID, w.LPVOID]
        lib.DeviceCapabilitiesW.restype = ctypes.c_int
        lib.GetPrinterW.argtypes = [w.HANDLE, w.DWORD, w.LPVOID, w.DWORD, ctypes.POINTER(w.DWORD)]
        lib.GetPrinterW.restype = w.BOOL
        lib.SetPrinterW.argtypes = [w.HANDLE, w.DWORD, w.LPVOID, w.DWORD]
        lib.SetPrinterW.restype = w.BOOL

    # -- handles --
    def open(self, name):
        h = self.w.HANDLE()
        return h if self.lib.OpenPrinterW(name, ctypes.byref(h), None) else None

    def close(self, h):
        if h:
            self.lib.ClosePrinter(h)

    # -- what the driver offers --
    def caps(self, name):
        """{'bins': [(name, id)], 'media': [(name, id)], 'dpis': [(x, y)]} as reported by the driver."""
        lib = self.lib

        def count(cap):
            return max(0, lib.DeviceCapabilitiesW(name, None, cap, None, None))

        def numbers(cap, ctype, per=1):
            n = count(cap)
            if not n:
                return []
            buf = (ctype * (n * per))()
            got = lib.DeviceCapabilitiesW(name, None, cap, buf, None)
            return list(buf)[:max(0, got) * per]

        def strings(cap, width):
            n = count(cap)
            if not n:
                return []
            buf = ctypes.create_unicode_buffer(n * width)
            got = lib.DeviceCapabilitiesW(name, None, cap, buf, None)
            return [buf[i * width:(i + 1) * width].split("\0")[0].strip() for i in range(max(0, got))]

        out = {"bins": [], "media": [], "dpis": []}
        try:
            out["bins"] = [(n, i) for n, i in zip(strings(_DC_BINNAMES, 24), numbers(_DC_BINS, ctypes.c_uint16)) if n]
            out["media"] = [(n, i) for n, i in zip(strings(_DC_MEDIATYPENAMES, 64),
                                                   numbers(_DC_MEDIATYPES, ctypes.c_uint32)) if n]
            flat = numbers(_DC_ENUMRESOLUTIONS, ctypes.c_long, 2)
            out["dpis"] = sorted(set(zip(flat[0::2], flat[1::2])))
        except Exception:
            pass
        return out

    # -- DEVMODE blobs --
    def default_devmode(self, h, name):
        size = self.lib.DocumentPropertiesW(None, h, name, None, None, 0)
        if size <= 0:
            return None
        buf = ctypes.create_string_buffer(size)
        return buf.raw if self.lib.DocumentPropertiesW(None, h, name, buf, None, _DM_OUT_BUFFER) >= 0 else None

    def validated(self, h, name, blob):
        """Let the driver merge/validate a DEVMODE (keeps its private data consistent)."""
        size = self.lib.DocumentPropertiesW(None, h, name, None, None, 0)
        if size <= 0 or len(blob) != size:
            return None
        src, out = ctypes.create_string_buffer(blob, size), ctypes.create_string_buffer(size)
        r = self.lib.DocumentPropertiesW(None, h, name, out, src, _DM_IN_BUFFER | _DM_OUT_BUFFER)
        return out.raw if r >= 0 else None

    def build_devmode(self, name, base=None, source=None, media=None, quality=None,
                      save_toner=False, grayscale=None, paper=None, paper_custom_mm=None,
                      landscape=None):
        """Merge the selected options into the driver's DEVMODE and validate it."""
        h = self.open(name)
        if not h:
            return None
        try:
            blob = base or self.default_devmode(h, name)
            if not blob:
                return None
            buf = ctypes.create_string_buffer(blob, len(blob))
            dm = _DEVMODEW.from_buffer(buf)
            if paper is not None:
                _configure_devmode_page(dm, paper, paper_custom_mm, landscape)
            if source is not None:
                dm.dmDefaultSource = int(source)
                dm.dmFields |= _DM_DEFAULTSOURCE
            if media is not None and dm.dmSize >= _DM_MEDIATYPE_END:
                dm.dmMediaType = int(media)
                dm.dmFields |= _DM_MEDIATYPE
            if quality:                                   # (kind, value, render_dpi, y_dpi)
                dm.dmPrintQuality = int(quality[1])
                dm.dmFields |= _DM_PRINTQUALITY
                if quality[3]:
                    dm.dmYResolution = int(quality[3])
                    dm.dmFields |= _DM_YRESOLUTION
            if save_toner:
                # DMRES_DRAFT asks capable drivers for their lower-consumption mode.
                dm.dmPrintQuality = DMRES_DRAFT
                dm.dmFields |= _DM_PRINTQUALITY
                dm.dmFields &= ~_DM_YRESOLUTION
            if grayscale is not None and dm.dmSize >= _DM_COLOR_END:
                dm.dmColor = _DMCOLOR_MONOCHROME if grayscale else _DMCOLOR_COLOR
                dm.dmFields |= _DM_COLOR
            return self.validated(h, name, buf.raw) or buf.raw
        finally:
            self.close(h)

    # -- per-user default DEVMODE (what Windows applications print with) --
    def get_user_devmode(self, h):
        need = self.w.DWORD()
        self.lib.GetPrinterW(h, 9, None, 0, ctypes.byref(need))
        if not need.value:
            return None
        buf = ctypes.create_string_buffer(need.value)
        if not self.lib.GetPrinterW(h, 9, buf, need.value, ctypes.byref(need)):
            return None
        ptr = ctypes.c_void_p.from_buffer(buf).value
        if not ptr:
            return None
        dm = _DEVMODEW.from_address(ptr)
        return ctypes.string_at(ptr, dm.dmSize + dm.dmDriverExtra)

    def set_user_devmode(self, h, blob):
        buf = ctypes.create_string_buffer(blob, len(blob)) if blob else None
        info = ctypes.c_void_p(ctypes.addressof(buf) if buf else None)
        return bool(self.lib.SetPrinterW(h, 9, ctypes.byref(info), 0))


class _TempDefaultDevMode:
    """Makes `blob` the printer's per-user default DEVMODE while printing, then puts the old one back.
    (Qt's QPrinter starts from that default, which is the only way to reach the driver's
    paper source / paper type / quality from Qt.)  `.active` tells whether it took effect."""

    def __init__(self, name, blob):
        self.name, self.blob, self.active, self.h, self.orig = name, blob, False, None, None

    def __enter__(self):
        ws = _WinSpool.get()
        if ws and self.blob:
            self.h = ws.open(self.name)
            if self.h:
                self.orig = ws.get_user_devmode(self.h)
                self.active = ws.set_user_devmode(self.h, self.blob)
        return self

    def __exit__(self, *exc):
        ws = _WinSpool.get()
        if ws and self.h:
            if self.active:
                ws.set_user_devmode(self.h, self.orig)
            ws.close(self.h)
        return False


def open_native_printer_properties(widget, printer_name, devmode=None):
    """Opens the printer driver's own settings window (paper source, paper type, quality, …).
    Returns (ok, error_text, devmode_bytes) — devmode_bytes is what the user chose when they
    pressed OK, so it can be used for the actual print job."""
    ws = _WinSpool.get()
    if not ws:
        return False, "Native printer settings are only available on Windows.", None
    h = ws.open(printer_name)
    if not h:
        return False, "Could not open that printer.", None
    try:
        hwnd = int(widget.winId())
        size = ws.lib.DocumentPropertiesW(hwnd, h, printer_name, None, None, 0)
        if size <= 0:
            return False, "This printer doesn't expose a settings window.", None
        out, mode, src = ctypes.create_string_buffer(size), _DM_IN_PROMPT | _DM_OUT_BUFFER, None
        if devmode and len(devmode) == size:
            src, mode = ctypes.create_string_buffer(devmode, size), mode | _DM_IN_BUFFER
        r = ws.lib.DocumentPropertiesW(hwnd, h, printer_name, out, src, mode)
        return r >= 0, None, (out.raw if r == 1 else None)      # 1 == IDOK
    finally:
        ws.close(h)


def _grid_for(n):
    """Rows x cols tile grid for n pages on one sheet. Known layouts use the usual
    office-printer arrangement; anything else gets the most square-ish grid."""
    presets = {1: (1, 1), 2: (1, 2), 4: (2, 2), 6: (2, 3), 8: (2, 4), 9: (3, 3), 10: (2, 5)}
    if n in presets:
        return presets[n]
    cols = max(1, math.ceil(math.sqrt(n)))
    rows = math.ceil(n / cols)
    return rows, cols


def _sheet_landscape(orient, per_sheet, rows, cols, first_page_rect):
    if orient == "l":
        return True
    if orient == "p":
        return False
    if per_sheet > 1:
        return cols > rows
    return first_page_rect.width > first_page_rect.height


def rot_bbox(w, h, rotation):
    """Width/height of the bounding box of a w x h page turned by `rotation` degrees (any angle)."""
    t = math.radians(rotation)
    c, sn = abs(math.cos(t)), abs(math.sin(t))
    return w * c + h * sn, w * sn + h * c


def layout_sheet(doc, group, rows, cols, aw, ah, scale_mode, scale_percent, offset_pt, rotation=0):
    """Pure geometry for one printed sheet, in PDF points, (0, 0) = top-left of the
    printable area. Returns [(page_index, clip_rect, target_rect), ...] where both
    rects are (x, y, w, h) tuples; clip_rect is each tile's own slot (so an enlarged
    or panned page never bleeds into its neighbour), target_rect is where the page
    is actually drawn."""
    gap = CELL_GAP_PT if rows * cols > 1 else 0
    cw = (aw - gap * (cols - 1)) / cols
    ch = (ah - gap * (rows - 1)) / rows
    ox, oy = offset_pt
    out = []
    for idx, pno in enumerate(group):
        r, c = divmod(idx, cols)
        cx, cy = c * (cw + gap), r * (ch + gap)
        pr = doc[pno].rect
        page_w, page_h = rot_bbox(pr.width, pr.height, rotation)       # bounding box of the turned page
        s_fit = min(cw / page_w, ch / page_h) if page_w and page_h else 1.0
        if scale_mode == "actual":
            s = 1.0
        elif scale_mode == "custom":
            s = s_fit * (scale_percent / 100.0)
        else:
            s = s_fit
        dw, dh = page_w * s, page_h * s
        tx, ty = cx + (cw - dw) / 2 + ox, cy + (ch - dh) / 2 + oy
        out.append((pno, (cx, cy, cw, ch), (tx, ty, dw, dh)))
    return out


# ============================================================================
#  PRINTING
# ============================================================================
def rows_to_range(rows):
    """[0,1,2,4] -> '1-3, 5'"""
    rows, out, i = sorted(set(rows)), [], 0
    while i < len(rows):
        j = i
        while j + 1 < len(rows) and rows[j + 1] == rows[j] + 1:
            j += 1
        out.append(str(rows[i] + 1) if i == j else f"{rows[i] + 1}-{rows[j] + 1}")
        i = j + 1
    return ", ".join(out)


def parse_range(text, n):
    """'1-3, 5, 8-'  ->  [0,1,2,4,7,...]   (None when invalid)"""
    out = []
    for part in re.split(r"[,;،]", text):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", part)
        if m and (m.group(1) or m.group(2)):
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else n
            if a < 1 or b > n or a > b:
                return None
            out += range(a - 1, b)
        elif part.isdigit():
            v = int(part)
            if not 1 <= v <= n:
                return None
            out.append(v - 1)
        else:
            return None
    return out or None


def _layout(psize, landscape):
    return QPageLayout(psize,
                       QPageLayout.Orientation.Landscape if landscape else QPageLayout.Orientation.Portrait,
                       QMarginsF(PRINT_MARGIN_MM, PRINT_MARGIN_MM, PRINT_MARGIN_MM, PRINT_MARGIN_MM),
                       QPageLayout.Unit.Millimeter)


def _apply_layout(printer, psize, landscape):
    if not printer.setPageLayout(_layout(psize, landscape)):    # printer refuses the margins -> use its own
        printer.setPageSize(psize)
        printer.setPageOrientation(QPageLayout.Orientation.Landscape if landscape
                                   else QPageLayout.Orientation.Portrait)


def _is_landscape(orient, rect):
    return orient == "l" or (orient == "auto" and rect.width > rect.height)


def _content_box(source, pr):
    """Smallest rectangle (page points) holding everything visible on the page, found from a small probe
    render and padded so nothing can be cut off. None means the page is completely blank."""
    k = min(100 / 72.0, math.sqrt(4_000_000 / max(1.0, pr.width * pr.height)))
    pix = source.get_pixmap(matrix=fitz.Matrix(k, k), alpha=False)
    w, h, n, stride, s = pix.width, pix.height, pix.n, pix.stride, pix.samples
    top = bottom = None
    left, right = w * n, 0                                    # in bytes
    for y in range(h):
        row = s[y * stride:y * stride + w * n]
        t = row.lstrip(b"\xff")
        if not t:
            continue
        if top is None:
            top = y
        bottom = y
        left = min(left, len(row) - len(t))
        right = max(right, len(row.rstrip(b"\xff")))
    if top is None:
        return None
    pad = 4
    x0, x1 = max(0, left // n - pad), min(w, -(-right // n) + pad)
    y0, y1 = max(0, top - pad), min(h, bottom + 1 + pad)
    return fitz.Rect(pr.x0 + x0 / k, pr.y0 + y0 / k, pr.x0 + x1 / k, pr.y0 + y1 / k) & pr


def _draw_page_banded(painter, page, dest, r, rotation=0, grayscale=False, compress=False):
    """Draw `page` rasterised at r px/pt into dest=(x, y, w, h) device pixels, in horizontal strips so a
    huge sheet (A0, a plotter roll) never needs one giant bitmap. Rotation is applied to the page
    content around its center (any angle); dest is the page's own, unrotated rectangle.
    compress=True sends far less data to the printer without touching quality: only the area that
    holds content is rendered and sent, strips that are pure white are skipped (the paper is already
    white), and grayscale pages are rendered straight to 8-bit gray."""
    pr = page.rect
    dx, dy, dw, dh = dest
    rotation %= 360
    source, box = page, pr
    if compress:
        source = page.get_displaylist()                       # parse the page once, reuse it for every strip
        box = _content_box(source, pr)
        if box is None:
            return                                            # nothing visible on this page
    if rotation:
        cx, cy = dx + dw / 2, dy + dh / 2
        painter.save()
        painter.translate(cx, cy)
        painter.rotate(rotation)
        painter.translate(-cx, -cy)
    full_w, full_h = max(1, round(pr.width * r)), max(1, round(pr.height * r))
    rows = max(1, min(full_h, int(MAX_BAND_PIXELS // full_w)))
    sx, sy = dw / full_w, dh / full_h
    if compress:
        rows = min(rows, 384)
        x0, x1 = max(0, int((box.x0 - pr.x0) * r)), min(full_w, math.ceil((box.x1 - pr.x0) * r))
        y, y_end = max(0, int((box.y0 - pr.y0) * r)), min(full_h, math.ceil((box.y1 - pr.y0) * r))
        cs = fitz.csGRAY if grayscale else fitz.csRGB
    else:
        x0, x1, y, y_end = 0, full_w, 0, full_h
    while y < y_end:
        n = min(rows, y_end - y)
        if compress:
            clip = fitz.Rect(pr.x0 + x0 / r, pr.y0 + y / r, pr.x0 + x1 / r, pr.y0 + min(pr.height, (y + n) / r))
            pix = source.get_pixmap(matrix=fitz.Matrix(r, r), clip=clip, colorspace=cs, alpha=False)
            samples = pix.samples
            if samples.count(b"\xff") == len(samples):         # pure white strip: print nothing
                y += n
                continue
            if grayscale:
                img = QImage(samples, pix.width, pix.height, pix.stride, QImage.Format.Format_Grayscale8).copy()
            else:
                img = pix_to_qimage(pix)
            left, width = dx + x0 * sx, (x1 - x0) * sx
        else:
            clip = fitz.Rect(pr.x0, pr.y0 + y / r, pr.x1, pr.y0 + min(pr.height, (y + n) / r))
            img = pix_to_qimage(page.get_pixmap(matrix=fitz.Matrix(r, r), clip=clip, alpha=False))
            if grayscale:
                img = img.convertToFormat(QImage.Format.Format_Grayscale8).convertToFormat(QImage.Format.Format_RGB32)
            left, width = dx, dw
        top, bottom = round(dy + y * sy), round(dy + (y + n) * sy)
        painter.drawImage(QRectF(left, top, width, max(1, bottom - top)), img)
        y += n
    if rotation:
        painter.restore()


def print_pdf(parent, doc, st, docname):
    """Send the pages in st['pages'] to the printer chosen in the dialog, st['per_sheet']
    pages at a time per physical sheet.  Paper source / paper type / quality from the
    dialog are applied through the driver's DEVMODE (Windows) for the duration of the job."""
    req = st.get("driver") or {}
    ws = _WinSpool.get()
    blob = None
    per_sheet = max(1, st.get("per_sheet", 1))
    rows, cols = _grid_for(per_sheet)
    landscape = _sheet_landscape(st["orient"], per_sheet, rows, cols, doc[st["pages"][0]].rect)
    driver_override = (req.get("blob")
                      or any(req.get(k) is not None for k in ("source", "media", "quality"))
                      or req.get("save_toner") or req.get("options_touched")
                      or st.get("paper") is not None)
    if ws and driver_override:
        blob = (ws.build_devmode(st["printer"], req.get("blob"), req.get("source"),
                                 req.get("media"), req.get("quality"),
                                 save_toner=req.get("save_toner", False),
                                 grayscale=req.get("grayscale"),
                                 paper=st.get("paper"),
                                 paper_custom_mm=st.get("paper_custom_mm"),
                                 landscape=landscape) or req.get("blob"))
    with _TempDefaultDevMode(st["printer"], blob) as tdm:
        if blob and not tdm.active:
            st["warning"] = ("Windows refused to apply one or more printer settings, including the selected "
                             "paper size, so the driver's own defaults may be used.")
        return _print_pdf(parent, doc, st, docname, driver_mode=bool(blob))


def _print_pdf(parent, doc, st, docname, driver_mode=False):
    pages = st["pages"]
    per_sheet = max(1, st.get("per_sheet", 1))
    groups = [pages[i:i + per_sheet] for i in range(0, len(pages), per_sheet)]
    rows, cols = _grid_for(per_sheet)
    offset_pt = (st["offset_mm"][0] / 25.4 * 72, st["offset_mm"][1] / 25.4 * 72)

    psize = page_size_for(st["paper"], st.get("paper_custom_mm", (210.0, 297.0)))
    render_dpi = st.get("render_dpi", RENDER_DPI_DEFAULT)
    if st.get("save_toner"):
        render_dpi = min(render_dpi, 150)
    # with driver overrides keep the driver's own resolution; otherwise ask for the maximum like before
    printer = QPrinter(QPrinterInfo.printerInfo(st["printer"]),
                       QPrinter.PrinterMode.PrinterResolution if driver_mode else QPrinter.PrinterMode.HighResolution)
    printer.setDocName(docname)
    printer.setCopyCount(st["copies"])
    printer.setCollateCopies(True)
    printer.setDuplex(st["duplex"])
    printer.setColorMode(QPrinter.ColorMode.GrayScale if st.get("grayscale")
                         else QPrinter.ColorMode.Color)
    if st.get("save_toner"):
        printer.setResolution(150)
    _apply_layout(printer, psize,
                  _sheet_landscape(st["orient"], per_sheet, rows, cols, doc[groups[0][0]].rect))

    prog = QProgressDialog("Printing…", "Cancel", 0, len(groups), parent)
    prog.setWindowTitle("PDF Studio")
    prog.setWindowModality(Qt.WindowModality.WindowModal)
    prog.setMinimumDuration(0)
    painter = QPainter()
    if not painter.begin(printer):
        raise RuntimeError("Could not start the printer.")
    canceled = False
    try:
        for n, group in enumerate(groups):
            prog.setValue(n)
            QApplication.processEvents()
            if prog.wasCanceled():
                canceled = True
                break
            ls = _sheet_landscape(st["orient"], per_sheet, rows, cols, doc[group[0]].rect)
            if n > 0:
                _apply_layout(printer, psize, ls)
                printer.newPage()
            area = printer.pageRect(QPrinter.Unit.DevicePixel)
            dpi = printer.resolution()
            px_per_pt = dpi / 72.0
            aw, ah = area.width() / px_per_pt, area.height() / px_per_pt         # printable area, in points
            rotation = st.get("image_rotation", 0) % 360
            cells = layout_sheet(doc, group, rows, cols, aw, ah, st["scale_mode"],
                                 st["scale_percent"], offset_pt, rotation)
            for pno, clip, target in cells:
                cx, cy, cw, ch = clip
                tx, ty, dw, dh = target
                pr = doc[pno].rect
                bw, _bh = rot_bbox(pr.width, pr.height, rotation)
                s = dw / bw if bw else 1.0
                uw, uh = pr.width * s, pr.height * s                              # the page's own size on paper
                ucx, ucy = tx + dw / 2, ty + dh / 2
                r = min(px_per_pt, (render_dpi / 72.0) * max(1.0, s))             # render at the chosen quality
                painter.setClipRect(QRectF(cx * px_per_pt, cy * px_per_pt, cw * px_per_pt, ch * px_per_pt))
                _draw_page_banded(painter, doc[pno],
                                  ((ucx - uw / 2) * px_per_pt, (ucy - uh / 2) * px_per_pt,
                                   uw * px_per_pt, uh * px_per_pt),
                                  r, rotation, st.get("grayscale", False), st.get("compress", False))
    finally:
        painter.end()
    if canceled:
        printer.abort()
    prog.setValue(len(groups))
    return not canceled


class PreviewLabel(QLabel):
    """The preview label itself: click-drag anywhere on it to nudge the printed
    content around the sheet (emits the drag delta, in label pixels).
    Hold Shift and a rotate frame appears over the page: drag with Shift held to turn the
    page by any angle around the frame's centre (hold Ctrl as well to snap to 5 degrees)."""
    dragged = pyqtSignal(QPointF)
    zoomed = pyqtSignal(int)                    # +1 / -1 per wheel notch
    zoom_reset = pyqtSignal()
    rotated = pyqtSignal(float)                 # new absolute rotation, degrees

    def __init__(self, parent=None):
        super().__init__(parent)
        self._last = None
        self._frame = None                      # (cx, cy, w, h, angle) in label pixels
        self._shift = False
        self._rot_start = None                  # (pointer angle, rotation) while rotating
        self.get_rotation = lambda: 0
        self.setCursor(Qt.CursorShape.OpenHandCursor)

    def set_frame(self, cx, cy=None, w=None, h=None, angle=0):
        self._frame = None if cx is None else (cx, cy, w, h, angle)
        self.update()

    def set_shift(self, down):
        if down != self._shift:
            self._shift = down
            if self._rot_start is None:
                self.setCursor(Qt.CursorShape.CrossCursor if down else Qt.CursorShape.OpenHandCursor)
            self.update()

    def _pointer_angle(self, pos):
        cx, cy = self._frame[0], self._frame[1]
        return math.degrees(math.atan2(pos.y() - cy, pos.x() - cx))

    def paintEvent(self, e):
        super().paintEvent(e)
        if not (self._shift and self._frame):
            return
        cx, cy, w, h, ang = self._frame
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.translate(cx, cy)
        p.rotate(ang)
        col = QColor(ACCENT)
        pen = QPen(col, 1.6, Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(QColor(col.red(), col.green(), col.blue(), 28))
        p.drawRect(QRectF(-w / 2, -h / 2, w, h))
        p.setPen(QPen(col, 1.6))
        p.drawLine(QPointF(0, -h / 2), QPointF(0, -h / 2 - 16))              # stem to the handle
        p.setBrush(col)
        p.drawEllipse(QPointF(0, -h / 2 - 16), 5, 5)
        p.drawLine(QPointF(-6, 0), QPointF(6, 0))                           # centre cross
        p.drawLine(QPointF(0, -6), QPointF(0, 6))
        p.resetTransform()
        a = (self.get_rotation() if self._rot_start is None else self._frame[4]) % 360
        txt = f"{a:.1f}\u00b0".replace(".0\u00b0", "\u00b0")
        p.setPen(QColor("#ffffff"))
        fm = p.fontMetrics()
        tw = fm.horizontalAdvance(txt) + 12
        box = QRectF(cx - tw / 2, cy + 8, tw, fm.height() + 6)
        p.setBrush(QColor(0, 0, 0, 170))
        p.setPen(Qt.PenStyle.NoPen)
        p.drawRoundedRect(box, 5, 5)
        p.setPen(QColor("#ffffff"))
        p.drawText(box, Qt.AlignmentFlag.AlignCenter, txt)
        p.end()

    def wheelEvent(self, e):
        dy = e.angleDelta().y()
        if dy:
            self.zoomed.emit(1 if dy > 0 else -1)
        e.accept()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.zoom_reset.emit()

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            if (e.modifiers() & Qt.KeyboardModifier.ShiftModifier) and self._frame:
                self._rot_start = (self._pointer_angle(e.position()), self.get_rotation())
                self.set_shift(True)
                self.setCursor(Qt.CursorShape.ClosedHandCursor)
                return
            self._last = e.position()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)

    def mouseMoveEvent(self, e):
        self.set_shift(bool(e.modifiers() & Qt.KeyboardModifier.ShiftModifier))
        if self._rot_start is not None and self._frame:
            a0, r0 = self._rot_start
            new = r0 + self._pointer_angle(e.position()) - a0
            if e.modifiers() & Qt.KeyboardModifier.ControlModifier:
                new = round(new / 5.0) * 5.0
            else:
                for snap in (0, 90, 180, 270, 360):                         # magnetic at the right angles
                    if abs((new % 360) - snap) < 1.5:
                        new = snap
                        break
            self.rotated.emit(new % 360)
        elif self._last is not None:
            d = e.position() - self._last
            self._last = e.position()
            self.dragged.emit(d)

    def mouseReleaseEvent(self, e):
        self._last = None
        self._rot_start = None
        self.setCursor(Qt.CursorShape.CrossCursor if self._shift else Qt.CursorShape.OpenHandCursor)
        self.update()


class RadioChoice(QObject):
    """A column of radio buttons that answers like a QComboBox (currentData / findData /
    setCurrentIndex / currentIndexChanged), so Adobe-style radio groups can replace dropdowns."""
    currentIndexChanged = pyqtSignal(int)

    def __init__(self, parent=None, horizontal=False, per_row=0):
        super().__init__(parent)
        self._group = QButtonGroup(self)
        self._buttons, self._data = [], []
        self._horizontal = horizontal
        self._per_row = per_row                      # >0: lay the radios out in a grid, this many per row
        self.box = QWidget()
        if per_row:
            self.layout_ = QVBoxLayout(self.box)
            self.layout_.setSpacing(6)
        else:
            self.layout_ = QHBoxLayout(self.box) if horizontal else QVBoxLayout(self.box)
            self.layout_.setSpacing(10 if horizontal else 6)
        self.layout_.setContentsMargins(0, 0, 0, 0)
        self._group.idToggled.connect(self._toggled)

    def addItem(self, text, data, extra=None):
        btn = QRadioButton(text)
        idx = len(self._buttons)
        self._group.addButton(btn, idx)
        self._buttons.append(btn)
        self._data.append(data)
        if self._per_row:
            row, col = divmod(idx, self._per_row)
            while self.layout_.count() <= row:                            # a new row: radios packed to the left
                rl = QHBoxLayout()
                rl.setContentsMargins(0, 0, 0, 0)
                rl.setSpacing(6)
                rl.addStretch(1)
                self.layout_.addLayout(rl)
            rl = self.layout_.itemAt(row).layout()
            btn.setMinimumWidth(66)                                       # equal columns without a grid
            rl.insertWidget(rl.count() - 1, btn)
            if extra is not None:
                rl.insertWidget(rl.count() - 1, extra)
        elif self._horizontal:
            self.layout_.addWidget(btn)
            if extra is not None:
                self.layout_.addWidget(extra)
        elif extra is None:
            self.layout_.addWidget(btn)
        else:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.addWidget(btn)
            row.addWidget(extra)
            row.addStretch(1)
            self.layout_.addLayout(row)
        if idx == 0:
            btn.setChecked(True)

    def _toggled(self, idx, checked):
        if checked:
            self.currentIndexChanged.emit(idx)

    def currentIndex(self):
        return self._group.checkedId()

    def currentData(self):
        i = self.currentIndex()
        return self._data[i] if 0 <= i < len(self._data) else None

    def findData(self, value):
        return self._data.index(value) if value in self._data else -1

    def setCurrentIndex(self, idx):
        if 0 <= idx < len(self._buttons):
            self._buttons[idx].setChecked(True)


class PrintDialog(DarkTitleBarMixin, QDialog):
    PW, PH = 300, 330                                      # preview size (compact, landscape dialog)

    def __init__(self, parent, doc, cur, preset=None, initial_preferences=None):
        super().__init__(parent)
        self.setObjectName("AppWindow")
        self.setWindowTitle("Print")
        self.doc, self.cur_page, self.pv = doc, cur, 0
        self._initial_preferences = initial_preferences or {}
        self.ox_mm, self.oy_mm, self._k = 0.0, 0.0, 1.0     # manual pan offset + last preview scale
        self.image_rotation = 0                               # image rotation, independent of paper orientation
        self.zoom = 1.0                                       # preview zoom (mouse wheel), 1.0 = fit
        self._fw, self._fh = self.PW, self.PH                 # preview frame = the sheet's own proportions
        self._printer_options_touched = False

        names = QPrinterInfo.availablePrinterNames()
        self.cb_printer = QComboBox()
        self.cb_printer.addItems(names)
        dflt = QPrinterInfo.defaultPrinterName()
        preferred_printer = self._initial_preferences.get("printer")
        if preferred_printer in names:
            self.cb_printer.setCurrentIndex(names.index(preferred_printer))
        elif dflt in names:
            self.cb_printer.setCurrentIndex(names.index(dflt))
        self.btn_printer_props = QPushButton("⚙ Settings…")
        self.btn_printer_props.setToolTip("Open this printer's own driver settings "
                                          "(paper source, quality, color…)")
        self.btn_printer_props.clicked.connect(self.open_printer_props)
        self.ck_save_toner = QCheckBox("SAVE TONER")
        self.ck_save_toner.setToolTip("Requests the printer's Draft mode and 150 dpi output; "
                                      "the driver must support Draft mode.")
        self.ck_compress = QCheckBox("COMPRESS")
        self.ck_compress.setToolTip("Sends much less data to the printer so the job starts printing sooner,\n"
                                    "with no change in print quality: only the part of each page that has\n"
                                    "content is sent, blank areas are skipped and gray pages go as 8-bit gray.")
        self.ck_grayscale = QCheckBox("GRAY SCALE")
        self.ck_grayscale.setToolTip("Print in grayscale (black and white); also changes the preview.")
        self.sp_copies = QSpinBox()
        self.sp_copies.setRange(1, 999)

        self.cb_paper = QComboBox()
        for k in PAPERS:
            self.cb_paper.addItem(k, k)
        self.cb_paper.addItem("Custom size…", CUSTOM_PAPER)
        self.cb_paper.setCurrentIndex(self.cb_paper.findData("A4"))
        self.sp_pw, self.sp_ph = QDoubleSpinBox(), QDoubleSpinBox()
        for sp, lo, hi, v in ((self.sp_pw, 50, 2000, 210), (self.sp_ph, 50, 20000, 297)):
            sp.setRange(lo, hi)
            sp.setDecimals(1)
            sp.setSuffix(" mm")
            sp.setValue(v)
        self.lbl_custom = QLabel("Width × Height:")
        self.w_custom = QWidget()
        rc = QHBoxLayout(self.w_custom)
        rc.setContentsMargins(0, 0, 0, 0)
        rc.addWidget(self.sp_pw)
        rc.addWidget(QLabel("×"))
        rc.addWidget(self.sp_ph)
        self.lbl_custom.setVisible(False)
        self.w_custom.setVisible(False)
        self.cb_orient = RadioChoice(self, horizontal=True)    # Adobe: radio buttons, not a dropdown
        for t, v in (("Auto portrait/landscape", "auto"), ("Portrait", "p"), ("Landscape", "l")):
            self.cb_orient.addItem(t, v)
        self.cb_orient.layout_.addStretch(1)

        self.sp_scale = QSpinBox()
        self.sp_scale.setRange(10, 400)
        self.sp_scale.setSingleStep(5)
        self.sp_scale.setValue(100)
        self.sp_scale.setSuffix(" %")
        self.sp_scale.setMaximumWidth(86)
        self.sp_scale.setToolTip("Percentage relative to \"Fit\" — changing it updates "
                                 "the preview immediately and selects Custom Scale.")
        self.cb_scale = RadioChoice(self, horizontal=True)       # Fit | Actual size | Custom Scale in one line
        self.cb_scale.addItem("Fit", "fit")
        self.cb_scale.addItem("Actual size", "actual")
        self.cb_scale.addItem("Custom Scale:", "custom", extra=self.sp_scale)
        self.cb_scale.layout_.addStretch(1)

        self.sp_persheet = QSpinBox()
        self.sp_persheet.setRange(1, 100)
        self.sp_persheet.setValue(9)
        self.sp_persheet.setSuffix("")
        self.sp_persheet.setVisible(False)
        self.cb_persheet = RadioChoice(self, horizontal=True, per_row=4)    # 1in1 | 2in1 | 4in1 ... radios
        for t, v in (("1in1", 1), ("2in1", 2), ("4in1", 4), ("6in1", 6),
                     ("8in1", 8), ("10in1", 10)):
            self.cb_persheet.addItem(t, v)
        self.cb_persheet.addItem("Custom", "custom", extra=self.sp_persheet)

        self.rb_all = QRadioButton(f"All pages (1 – {len(doc)})")
        self.rb_cur = QRadioButton(f"Current page ({cur + 1})")
        self.rb_range = QRadioButton("Pages:")
        self.rb_all.setChecked(True)
        grp = QButtonGroup(self)
        for b in (self.rb_all, self.rb_cur, self.rb_range):
            grp.addButton(b)
        self.ed_range = QLineEdit()
        self.ed_range.setPlaceholderText("e.g. 1-3, 5, 8-")

        # Two-sided printing: one tick turns it on, two more ticks pick the flip edge
        self.ck_duplex = QCheckBox("Print on both sides of paper")
        self.ck_duplex.setToolTip("Both sides of the sheet. Needs a printer with automatic duplex.")
        self.ck_flip_lr = QCheckBox("RIGHT")
        self.ck_flip_lr.setToolTip("Turn the sheet over on its left / right edge.")
        self.ck_flip_tb = QCheckBox("UP")
        self.ck_flip_tb.setToolTip("Turn the sheet over on its top / bottom edge.")
        self.grp_flip = QButtonGroup(self)
        self.grp_flip.setExclusive(True)                    # exactly one flip edge stays ticked
        self.grp_flip.addButton(self.ck_flip_lr)
        self.grp_flip.addButton(self.ck_flip_tb)
        self.ck_flip_lr.setChecked(True)
        self.w_flip = QWidget()
        flip_row = QHBoxLayout(self.w_flip)
        flip_row.setContentsMargins(26, 0, 0, 0)
        flip_row.setSpacing(18)
        flip_row.addWidget(self.ck_flip_lr)
        flip_row.addWidget(self.ck_flip_tb)
        flip_row.addStretch(1)
        self.w_flip.setEnabled(False)

        self.sp_ox = QDoubleSpinBox()
        self.sp_oy = QDoubleSpinBox()
        for sp in (self.sp_ox, self.sp_oy):
            sp.setRange(-200, 200)
            sp.setDecimals(1)
            sp.setSuffix(" mm")
        self.bt_reset_pos = QPushButton("Reset view")
        self.bt_reset_pos.clicked.connect(self.reset_offset)

        def box(title, form):
            g = QGroupBox(title)
            g.setLayout(form)
            return g

        # paper source (AUTO / BYPASS TRAY) and paper type, filled from what the chosen driver reports;
        # quality has no control any more, it only follows the driver's own settings window
        self.devmode_blob = None                           # settings captured from the driver's own window
        self._caps = {"bins": [], "media": [], "dpis": []}
        self.cb_source, self.cb_media, self.cb_quality = QComboBox(), QComboBox(), QComboBox()
        self.cb_quality.setVisible(False)
        for cb in (self.cb_source, self.cb_media):
            cb.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            cb.setMinimumContentsLength(4)
            cb.setFixedWidth(116)

        row_printer = QHBoxLayout()
        self.cb_printer.setMinimumWidth(220)                    # wide printer drop-down; Size / Source / Type sit to its right
        self.cb_printer.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.cb_printer.setMinimumContentsLength(8)
        row_printer.addWidget(self.cb_printer, 3)
        row_printer.addSpacing(6)
        row_printer.addWidget(QLabel("Size:"))                  # size / source / type sit between printer and Settings
        self.cb_paper.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.cb_paper.setMinimumContentsLength(7)
        self.cb_paper.setFixedWidth(84)
        row_printer.addWidget(self.cb_paper)
        row_printer.addSpacing(4)
        row_printer.addWidget(QLabel("Source:"))
        row_printer.addWidget(self.cb_source)
        row_printer.addSpacing(8)
        row_printer.addWidget(QLabel("Type:"))
        row_printer.addWidget(self.cb_media)
        row_printer.addSpacing(8)
        row_printer.addWidget(self.btn_printer_props)
        w_printer = QWidget()
        w_printer.setLayout(row_printer)

        # Top strip, full width: Printer, size, source, type and Settings in one line
        f1 = QFormLayout()
        f1.addRow(w_printer)
        copies_row = QHBoxLayout()
        copies_row.addWidget(QLabel("Copies:"))
        copies_row.addWidget(self.sp_copies)
        copies_row.addSpacing(12)
        copies_row.addWidget(self.ck_grayscale)
        copies_row.addWidget(self.ck_save_toner)
        copies_row.addWidget(self.ck_compress)
        copies_row.addStretch(1)

        # Page Sizing & Handling: paper, size (scale) radios, pages per sheet, both-sides ticks
        f2 = QFormLayout()
        f2.addRow(self.lbl_custom, self.w_custom)               # only shown for a custom paper size
        row_per = QHBoxLayout()
        row_per.addWidget(self.cb_persheet.box, 1)
        w_per = QWidget()
        w_per.setLayout(row_per)
        v2 = QVBoxLayout()
        v2.setSpacing(8)
        v2.addLayout(f2)
        v2.addWidget(QLabel("Orientation:"))
        v2.addWidget(self.cb_orient.box)
        v2.addWidget(QLabel("Size:"))
        v2.addWidget(self.cb_scale.box)
        f_per = QFormLayout()
        f_per.addRow("Pages per sheet:", w_per)
        v2.addLayout(f_per)
        v_duplex = QVBoxLayout()                               # two-sided printing, in the old Orientation spot
        v_duplex.addWidget(self.ck_duplex)
        v_duplex.addWidget(self.w_flip)
        v_duplex.addStretch(1)
        f3 = QVBoxLayout()
        f3.addWidget(self.rb_all)
        f3.addWidget(self.rb_cur)
        row = QHBoxLayout()
        row.addWidget(self.rb_range)
        row.addWidget(self.ed_range)
        f3.addLayout(row)

        row_pg = QHBoxLayout()                                   # Pages to Print | Two-sided, side by side
        row_pg.setSpacing(10)
        row_pg.addWidget(box("Pages to Print", f3), 1)
        row_pg.addWidget(box("Two-sided", v_duplex), 1)
        left = QVBoxLayout()
        left.setSpacing(10)
        left.addLayout(copies_row)
        left.addLayout(row_pg)
        left.addWidget(box("Page Sizing && Handling", v2))
        left.addStretch(1)

        self.lbl_pv = PreviewLabel()
        self.lbl_pv.setFixedSize(self.PW + 2, self.PH + 2)
        self.lbl_pv.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_pv.setStyleSheet(f"background:{T['CANVAS']};border:1px solid {T['LINE2']};"
                                  f"border-radius:2px;color:{T['TEXT']}")
        self.bt_prev, self.bt_next = QPushButton("◀"), QPushButton("▶")
        for b in (self.bt_prev, self.bt_next):
            b.setFixedWidth(40)
            b.setObjectName("NavBtn")
        self.bt_rotate_left, self.bt_rotate_right = QPushButton("↶"), QPushButton("↷")
        for b in (self.bt_rotate_left, self.bt_rotate_right):
            b.setFixedSize(42, 34)
            b.setObjectName("RotBtn")
        self.bt_rotate_left.setToolTip("Rotate image left 90°")
        self.bt_rotate_left.setAccessibleName("Rotate image left")
        self.bt_rotate_right.setToolTip("Rotate image right 90°")
        self.bt_rotate_right.setAccessibleName("Rotate image right")
        rotate_row = QHBoxLayout()                            # [Reset view ..... ↶ ↷ ..... 100%], rotate buttons centred
        rotate_row.setSpacing(6)
        row_left, row_right = QHBoxLayout(), QHBoxLayout()
        row_left.setContentsMargins(0, 0, 0, 0)
        row_right.setContentsMargins(0, 0, 0, 0)
        self.lbl_info = QLabel()
        self.lbl_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        nav = QHBoxLayout()
        nav.addWidget(self.bt_prev)
        nav.addWidget(self.lbl_info, 1)
        nav.addWidget(self.bt_next)
        self.lbl_zoom = QLabel("100%")
        self.lbl_zoom.setStyleSheet(f"color:{T['MUTED']};font-size:11px")
        row_left.addWidget(self.bt_reset_pos)                 # reset zoom + position
        row_left.addStretch(1)
        row_right.addStretch(1)
        row_right.addWidget(self.lbl_zoom)
        rotate_row.addLayout(row_left, 1)
        rotate_row.addWidget(self.bt_rotate_left)
        rotate_row.addWidget(self.bt_rotate_right)
        rotate_row.addLayout(row_right, 1)
        self.lbl_info.setWordWrap(True)
        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(6)
        pv_box = QWidget()                                       # fixed area; the frame inside follows the sheet's shape
        pv_box.setFixedSize(self.PW + 2, self.PH + 2)
        pv_lay = QVBoxLayout(pv_box)
        pv_lay.setContentsMargins(0, 0, 0, 0)
        pv_lay.addWidget(self.lbl_pv, 0, Qt.AlignmentFlag.AlignCenter)
        right.addWidget(pv_box, 0, Qt.AlignmentFlag.AlignHCenter)
        right.addLayout(rotate_row)
        right.addLayout(nav)
        right.addStretch(1)
        w_right = QWidget()                                      # as wide as the preview, so it sits centred
        w_right.setFixedWidth(self.PW + 40)
        w_right.setLayout(right)

        self.btn_print = QPushButton("🖨 Print")
        self.btn_print.setDefault(True)
        self.btn_print.setObjectName("Primary")
        bb = QDialogButtonBox()
        bb.addButton(self.btn_print, QDialogButtonBox.ButtonRole.AcceptRole)
        bb.addButton(QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)

        top = QHBoxLayout()
        top.setSpacing(14)
        top.addLayout(left, 1)
        top.addWidget(w_right, 0, Qt.AlignmentFlag.AlignTop)
        row_btn = QHBoxLayout()                                  # Print / Cancel in the bottom-right corner
        row_btn.addStretch(1)
        row_btn.addWidget(bb)
        main = QVBoxLayout(self)
        main.setContentsMargins(16, 12, 16, 12)
        main.setSpacing(8)
        main.addLayout(f1)
        main.addLayout(top, 1)
        main.addLayout(row_btn)
        # tighter group boxes than the main window uses, so the whole dialog stays small
        self.setStyleSheet("QGroupBox{padding:28px 10px 8px 10px;}QGroupBox::title{top:7px;}")

        self.cb_orient.currentIndexChanged.connect(self.update_preview)
        self.ck_duplex.toggled.connect(self._duplex_toggled)
        self.ck_flip_lr.toggled.connect(self.update_preview)
        self.ck_flip_tb.toggled.connect(self.update_preview)
        self.cb_paper.currentIndexChanged.connect(self._paper_changed)
        self.sp_pw.valueChanged.connect(self.update_preview)
        self.sp_ph.valueChanged.connect(self.update_preview)
        self.cb_printer.currentIndexChanged.connect(self._printer_changed)
        self.ck_save_toner.toggled.connect(self._printer_options_changed)
        self.ck_grayscale.toggled.connect(self._printer_options_changed)
        self.cb_scale.currentIndexChanged.connect(self._scale_mode_changed)
        self.sp_scale.valueChanged.connect(self._custom_scale_changed)
        self.cb_persheet.currentIndexChanged.connect(self._persheet_mode_changed)
        self.sp_persheet.valueChanged.connect(self.update_preview)
        for r in (self.rb_all, self.rb_cur, self.rb_range):
            r.toggled.connect(self.update_preview)
        self.ed_range.textChanged.connect(self.update_preview)
        self.ed_range.textEdited.connect(lambda _: self.rb_range.setChecked(True))
        self.bt_prev.clicked.connect(lambda: self.step(-1))
        self.bt_next.clicked.connect(lambda: self.step(1))
        self.sp_ox.valueChanged.connect(self._ox_changed)
        self.sp_oy.valueChanged.connect(self._oy_changed)
        self.lbl_pv.dragged.connect(self.on_drag)
        self.lbl_pv.rotated.connect(self._rotate_free)
        self.lbl_pv.get_rotation = lambda: self.image_rotation
        self.lbl_pv.zoomed.connect(self._wheel_zoom)
        self.lbl_pv.zoom_reset.connect(self._reset_zoom)
        self.bt_rotate_left.clicked.connect(lambda: self._rotate_image(-90))
        self.bt_rotate_right.clicked.connect(lambda: self._rotate_image(90))
        self._load_caps()
        self._restore_preferences(preset)
        self.update_preview()

    # ---- helpers ----
    def custom_mm(self):
        return (self.sp_pw.value(), self.sp_ph.value())

    def duplex_value(self):
        """'none' (one-sided), 'lr' (flip RIGHT) or 'tb' (flip UP), from the three ticks."""
        if not self.ck_duplex.isChecked():
            return "none"
        return "lr" if self.ck_flip_lr.isChecked() else "tb"

    def set_duplex_value(self, value):
        self.ck_duplex.setChecked(value != "none")
        if value == "tb":
            self.ck_flip_tb.setChecked(True)
        elif value == "lr":
            self.ck_flip_lr.setChecked(True)

    def _duplex_toggled(self, on):
        self.w_flip.setEnabled(on)
        self.update_preview()

    def _paper_changed(self, *_):
        custom = self.cb_paper.currentData() == CUSTOM_PAPER
        self.lbl_custom.setVisible(custom)
        self.w_custom.setVisible(custom)
        self.update_preview()

    def _printer_changed(self, *_):
        self.devmode_blob = None
        self._printer_options_touched = False
        for checkbox in (self.ck_save_toner, self.ck_grayscale, self.ck_compress):
            old = checkbox.blockSignals(True)
            checkbox.setChecked(False)
            checkbox.blockSignals(old)
        self.cb_paper.setCurrentIndex(self.cb_paper.findData("A4"))
        self.cb_orient.setCurrentIndex(self.cb_orient.findData("auto"))
        self.set_duplex_value("none")
        self.sp_copies.setValue(1)
        self.sp_pw.setValue(210.0)
        self.sp_ph.setValue(297.0)
        self._load_caps()

    def _load_caps(self):
        ws, name = _WinSpool.get(), self.cb_printer.currentText()
        self._caps = (ws.caps(name) if ws and name else None) or {"bins": [], "media": [], "dpis": []}
        self.cb_source.clear()
        self.cb_source.addItem("AUTO", None)
        self.cb_source.addItem("BYPASS TRAY", self._bypass_bin())
        self.cb_media.clear()
        self.cb_media.addItem("Default", None)
        for n, i in self._caps["media"]:
            self.cb_media.addItem(n, i)
        self.cb_quality.clear()
        self.cb_quality.addItem("Printer default", None)
        for t, v, dpi in QUALITY_PRESETS:
            self.cb_quality.addItem(t, ("q", v, dpi, None))
        for x, y in self._caps["dpis"][-6:]:
            self.cb_quality.addItem(f"{x} × {y} dpi" if x != y else f"{x} dpi", ("dpi", x, min(x, 600), y))
        tip = "" if ws else "Paper source / paper type need the Windows printer driver."
        for cb in (self.cb_source, self.cb_media):
            cb.setEnabled(bool(ws))
            cb.setToolTip(tip)
        default_blob = None
        if ws and name:
            h = ws.open(name)
            if h:
                try:
                    default_blob = ws.default_devmode(h, name)
                finally:
                    ws.close(h)
        self._sync_driver_settings(default_blob, include_advanced=False)

    def _sync_driver_settings(self, blob, include_advanced=True):
        """Reflect the selected printer's DEVMODE in the page controls and live preview."""
        if not blob or len(blob) < ctypes.sizeof(_DEVMODEW):
            return
        try:
            dm = _DEVMODEW.from_buffer_copy(blob)
        except (ValueError, TypeError):
            return
        fields = dm.dmFields

        # Paper: use explicit dimensions when the driver exposes them, otherwise known
        # Windows paper IDs. The UI keeps custom dimensions in portrait order.
        paper_dims = {
            1: (215.9, 279.4), 2: (215.9, 279.4), 3: (279.4, 431.8),
            4: (279.4, 431.8), 5: (215.9, 355.6), 6: (139.7, 215.9),
            7: (184.15, 266.7), 8: (297.0, 420.0), 9: (210.0, 297.0),
            10: (210.0, 297.0), 11: (148.0, 210.0),
        }
        dimensions = None
        if ((fields & _DM_PAPERWIDTH) and (fields & _DM_PAPERLENGTH)
                and dm.dmPaperWidth > 0 and dm.dmPaperLength > 0):
            dimensions = tuple(sorted((dm.dmPaperWidth / 10.0, dm.dmPaperLength / 10.0)))
        elif (fields & _DM_PAPERSIZE) and dm.dmPaperSize in paper_dims:
            dimensions = paper_dims[dm.dmPaperSize]
        if not dimensions:
            form_name = "".join(chr(c) for c in dm.dmFormName).split("\0", 1)[0].strip().upper()
            form_sizes = {
                "A0": (841.0, 1189.0), "A1": (594.0, 841.0), "A2": (420.0, 594.0),
                "A3": (297.0, 420.0), "A4": (210.0, 297.0), "A5": (148.0, 210.0),
                "LETTER": paper_dims[1], "LEGAL": paper_dims[5], "TABLOID": paper_dims[3],
                "LEDGER": paper_dims[4], "EXECUTIVE": paper_dims[7],
            }
            dimensions = form_sizes.get(form_name)
        if dimensions:
            known = {
                "A0": (841.0, 1189.0), "A1": (594.0, 841.0), "A2": (420.0, 594.0),
                "A3": (297.0, 420.0), "A4": (210.0, 297.0), "A5": (148.0, 210.0),
            }
            matched = next((key for key, size in known.items()
                            if max(abs(dimensions[0] - size[0]), abs(dimensions[1] - size[1])) <= 2.0), None)
            if matched:
                idx = self.cb_paper.findData(matched)
                if idx >= 0:
                    self.cb_paper.setCurrentIndex(idx)
            else:
                self.sp_pw.setValue(dimensions[0])
                self.sp_ph.setValue(dimensions[1])
                idx = self.cb_paper.findData(CUSTOM_PAPER)
                if idx >= 0:
                    self.cb_paper.setCurrentIndex(idx)

        if include_advanced and fields & _DM_ORIENTATION and dm.dmOrientation in (1, 2):
            self.cb_orient.setCurrentIndex(self.cb_orient.findData("p" if dm.dmOrientation == 1 else "l"))
        if fields & _DM_COPIES and dm.dmCopies > 0:
            self.sp_copies.setValue(dm.dmCopies)
        if fields & _DM_DUPLEX:
            duplex = "none"
            if dm.dmDuplex in (2, 3):
                landscape = dm.dmOrientation == 2 if fields & _DM_ORIENTATION else False
                long_edge = (dm.dmDuplex == 2) != landscape
                duplex = "lr" if long_edge else "tb"
            self.set_duplex_value(duplex)
        if fields & _DM_COLOR and dm.dmSize >= _DM_COLOR_END:
            old = self.ck_grayscale.blockSignals(True)
            self.ck_grayscale.setChecked(dm.dmColor == _DMCOLOR_MONOCHROME)
            self.ck_grayscale.blockSignals(old)

        if include_advanced:
            if fields & _DM_DEFAULTSOURCE:
                idx = self.cb_source.findData(dm.dmDefaultSource)
                if idx >= 0:
                    self.cb_source.setCurrentIndex(idx)
            if fields & _DM_MEDIATYPE and dm.dmSize >= _DM_MEDIATYPE_END:
                media_idx = self.cb_media.findData(dm.dmMediaType)
                if media_idx >= 0:
                    self.cb_media.setCurrentIndex(media_idx)
            if fields & _DM_PRINTQUALITY:
                quality_idx = -1
                for i in range(self.cb_quality.count()):
                    data = self.cb_quality.itemData(i)
                    if not data:
                        continue
                    if dm.dmPrintQuality < 0 and data[0] == "q" and data[1] == dm.dmPrintQuality:
                        quality_idx = i
                        break
                    if dm.dmPrintQuality > 0 and data[0] == "dpi" and data[1] == dm.dmPrintQuality:
                        if not (fields & _DM_YRESOLUTION) or data[3] == dm.dmYResolution:
                            quality_idx = i
                            break
                if quality_idx >= 0:
                    self.cb_quality.setCurrentIndex(quality_idx)
        self.update_preview()

    def _bypass_bin(self):
        """Driver id of the bypass / manual-feed tray: the tray whose name says so, else the standard
        Windows 'manual feed' bin."""
        pat = re.compile(r"by-?\s*pass|manual|multi|\bmp\b|\bmpt\b", re.I)
        for name, bid in self._caps["bins"]:
            if pat.search(name):
                return bid
        return 4                                           # DMBIN_MANUAL

    def driver_request(self):
        return {"blob": self.devmode_blob, "source": self.cb_source.currentData(),
                "media": self.cb_media.currentData(), "quality": self.cb_quality.currentData(),
                "save_toner": self.ck_save_toner.isChecked(),
                "grayscale": self.ck_grayscale.isChecked(),
                "options_touched": self._printer_options_touched}

    def pages(self):
        if self.rb_all.isChecked():
            return list(range(len(self.doc)))
        if self.rb_cur.isChecked():
            return [self.cur_page]
        return parse_range(self.ed_range.text(), len(self.doc))

    def per_sheet(self):
        d = self.cb_persheet.currentData()
        return max(1, self.sp_persheet.value()) if d == "custom" else d

    def groups(self):
        pg = self.pages()
        if not pg:
            return []
        n = self.per_sheet()
        return [pg[i:i + n] for i in range(0, len(pg), n)]

    def step(self, d):
        self.pv += d
        self.update_preview()

    def duplex_mode(self, first_landscape):
        d = self.duplex_value()
        if d == "none":
            return QPrinter.DuplexMode.DuplexNone
        # "left/right edge" is the long edge of a portrait sheet and the short edge of a landscape one
        long_edge = (d == "lr") != first_landscape
        return QPrinter.DuplexMode.DuplexLongSide if long_edge else QPrinter.DuplexMode.DuplexShortSide

    @staticmethod
    def _select_data(combo, value):
        idx = combo.findData(value)
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def _restore_preferences(self, preset=None):
        """Restore the last print-dialog choices held in memory by the main window."""
        prefs = self._initial_preferences
        if prefs:
            self.sp_copies.setValue(prefs.get("copies", self.sp_copies.value()))
            self._select_data(self.cb_paper, prefs.get("paper"))
            custom_mm = prefs.get("paper_custom_mm")
            if custom_mm and len(custom_mm) == 2:
                self.sp_pw.setValue(custom_mm[0])
                self.sp_ph.setValue(custom_mm[1])
            self._select_data(self.cb_orient, prefs.get("orient", "auto"))
            self.sp_scale.setValue(prefs.get("scale_percent", self.sp_scale.value()))
            self._select_data(self.cb_scale, prefs.get("scale_mode", "fit"))
            self._select_data(self.cb_persheet, prefs.get("per_sheet_mode", 1))
            self.sp_persheet.setValue(prefs.get("per_sheet_custom", self.sp_persheet.value()))
            self.set_duplex_value(prefs.get("duplex", "none"))
            self.sp_ox.setValue(prefs.get("offset_mm", (0.0, 0.0))[0])
            self.sp_oy.setValue(prefs.get("offset_mm", (0.0, 0.0))[1])
            self.image_rotation = prefs.get("image_rotation", 0)

            mode = prefs.get("pages_mode", "all")
            if mode == "range":
                self.ed_range.setText(prefs.get("range_text", ""))
                self.rb_range.setChecked(True)
            elif mode == "current":
                self.rb_cur.setChecked(True)
            else:
                self.rb_all.setChecked(True)

            self.ck_save_toner.setChecked(prefs.get("save_toner", False))
            self.ck_compress.setChecked(prefs.get("compress", False))
            self.ck_grayscale.setChecked(prefs.get("grayscale", False))

            # These options and the native driver DEVMODE are specific to a printer.
            if prefs.get("printer") == self.cb_printer.currentText():
                self._select_data(self.cb_source, prefs.get("source"))
                self._select_data(self.cb_media, prefs.get("media"))
                self._select_data(self.cb_quality, prefs.get("quality"))
                self.devmode_blob = prefs.get("devmode_blob")
                self._printer_options_touched = prefs.get("printer_options_touched", False)

        # A contextual "print selected pages" action takes priority for this job.
        if preset:
            if len(preset) == 1:
                self.rb_cur.setChecked(True)
            else:
                self.ed_range.setText(rows_to_range(preset))
                self.rb_range.setChecked(True)
        self._persheet_mode_changed()
        self._paper_changed()

    def preferences(self):
        """Snapshot every user-facing print option for reuse until this app closes."""
        pages_mode = "range" if self.rb_range.isChecked() else ("current" if self.rb_cur.isChecked() else "all")
        return {
            "printer": self.cb_printer.currentText(),
            "copies": self.sp_copies.value(),
            "paper": self.cb_paper.currentData(),
            "paper_custom_mm": self.custom_mm(),
            "orient": self.cb_orient.currentData(),
            "scale_mode": self.cb_scale.currentData(),
            "scale_percent": self.sp_scale.value(),
            "per_sheet_mode": self.cb_persheet.currentData(),
            "per_sheet_custom": self.sp_persheet.value(),
            "pages_mode": pages_mode,
            "range_text": self.ed_range.text(),
            "duplex": self.duplex_value(),
            "offset_mm": (self.sp_ox.value(), self.sp_oy.value()),
            "image_rotation": self.image_rotation,
            "save_toner": self.ck_save_toner.isChecked(),
            "compress": self.ck_compress.isChecked(),
            "grayscale": self.ck_grayscale.isChecked(),
            "source": self.cb_source.currentData(),
            "media": self.cb_media.currentData(),
            "quality": self.cb_quality.currentData(),
            "devmode_blob": self.devmode_blob,
            "printer_options_touched": self._printer_options_touched,
        }

    def settings(self):
        pg = self.pages()
        per_sheet = self.per_sheet()
        rows, cols = _grid_for(per_sheet)
        ls0 = _sheet_landscape(self.cb_orient.currentData(), per_sheet, rows, cols, self.doc[pg[0]].rect)
        save_toner = self.ck_save_toner.isChecked()
        return {"printer": self.cb_printer.currentText(), "copies": self.sp_copies.value(),
                "paper": self.cb_paper.currentData(), "paper_custom_mm": self.custom_mm(),
                "driver": self.driver_request(),
                "render_dpi": min((self.cb_quality.currentData()
                                   or (None, None, RENDER_DPI_DEFAULT))[2], 150) if save_toner
                else (self.cb_quality.currentData() or (None, None, RENDER_DPI_DEFAULT))[2],
                "orient": self.cb_orient.currentData(),
                "scale_mode": self.cb_scale.currentData(), "scale_percent": self.sp_scale.value(),
                "image_rotation": self.image_rotation,
                "save_toner": save_toner, "grayscale": self.ck_grayscale.isChecked(),
                "compress": self.ck_compress.isChecked(),
                "per_sheet": per_sheet, "offset_mm": (self.ox_mm, self.oy_mm),
                "pages": pg, "duplex": self.duplex_mode(ls0)}

    # ---- UI reactions ----
    def _printer_options_changed(self, *_):
        self._printer_options_touched = True
        self.update_preview()

    def _scale_mode_changed(self, *_):
        self.update_preview()

    def _custom_scale_changed(self, *_):
        if self.cb_scale.currentData() != "custom":
            idx = self.cb_scale.findData("custom")
            if idx >= 0:
                self.cb_scale.setCurrentIndex(idx)
                return
        self.update_preview()

    def eventFilter(self, obj, ev):
        # Shift down/up anywhere in the dialog shows / hides the rotate frame on the preview
        t = ev.type()
        if t in (QEvent.Type.KeyPress, QEvent.Type.KeyRelease) and ev.key() == Qt.Key.Key_Shift:
            self.lbl_pv.set_shift(t == QEvent.Type.KeyPress)
        elif t == QEvent.Type.WindowDeactivate and obj is self:
            self.lbl_pv.set_shift(False)
        return super().eventFilter(obj, ev)

    def showEvent(self, e):
        super().showEvent(e)
        QApplication.instance().installEventFilter(self)

    def hideEvent(self, e):
        QApplication.instance().removeEventFilter(self)
        self.lbl_pv.set_shift(False)
        super().hideEvent(e)

    def _rotate_image(self, degrees):
        self.image_rotation = round(self.image_rotation + degrees, 2) % 360
        self.update_preview()

    def _rotate_free(self, angle):
        """Free rotation from the Shift+drag frame on the preview."""
        self.image_rotation = round(angle, 1) % 360
        self.update_preview()

    def _persheet_mode_changed(self, *_):
        self.sp_persheet.setVisible(self.cb_persheet.currentData() == "custom")
        self.pv = 0
        self.update_preview()

    def _ox_changed(self, v):
        self.ox_mm = v
        self.update_preview()

    def _oy_changed(self, v):
        self.oy_mm = v
        self.update_preview()

    def _wheel_zoom(self, step):
        z = self.zoom * (1.15 ** step)
        self.zoom = max(0.5, min(4.0, round(z, 3)))
        self.update_preview()

    def _reset_zoom(self):
        self.zoom = 1.0
        self.update_preview()

    def reset_offset(self):
        self.zoom = 1.0
        self.ox_mm = self.oy_mm = 0.0
        for sp, v in ((self.sp_ox, self.ox_mm), (self.sp_oy, self.oy_mm)):
            sp.blockSignals(True)
            sp.setValue(v)
            sp.blockSignals(False)
        self.update_preview()

    def on_drag(self, d):
        if not self._k:
            return
        px_per_mm = self._k * 72.0 / 25.4
        self.ox_mm += d.x() / px_per_mm
        self.oy_mm += d.y() / px_per_mm
        for sp, v in ((self.sp_ox, self.ox_mm), (self.sp_oy, self.oy_mm)):
            sp.blockSignals(True)
            sp.setValue(v)
            sp.blockSignals(False)
        self.update_preview()

    def open_printer_props(self):
        name = self.cb_printer.currentText()
        if not name:
            return
        ok, err, blob = open_native_printer_properties(self, name, self.devmode_blob)
        if not ok:
            QMessageBox.information(self, "Printer settings",
                                    err or "Could not open this printer's settings window.")
        elif blob:
            self.devmode_blob = blob
            self._sync_driver_settings(blob)

    # ---- preview ----
    def _draw_flipped(self, p, paper, ox, oy, mode, back):
        """Draw the sheet as if it were being turned over: the hinge is the flip edge
        (left edge for 'left / right', top edge for 'top / bottom'); the opposite, free
        edge swings up towards the viewer so the paper looks curved in that direction.
        The back side is drawn the same way but tinted and badged."""
        W, H = paper.width(), paper.height()
        vertical_hinge = (mode == "lr")                  # hinge on the left, paper bends towards the right
        LIFT, SQUEEZE = 0.12, 0.08                       # how big the free edge gets / how much it foreshortens
        n = W if vertical_hinge else H
        cx_, cy_ = ox + W / 2, oy + H / 2

        def curve(t):                                    # 0 at the hinge .. 1 at the free edge
            return t * t * (3 - 2 * t) * 0.35 + t ** 3 * 0.65

        # the paper itself, strip by strip, with a perspective bend
        step = 1.0
        i = 0.0
        while i < n:
            t0, t1 = i / n, min(n, i + step) / n
            c0, c1 = curve(t0), curve(t1)
            sc = 1 + LIFT * c0
            if vertical_hinge:
                x0 = ox + (i * (1 - SQUEEZE * c0))
                x1 = ox + (min(n, i + step) * (1 - SQUEEZE * c1))
                hh = H * sc
                dst = QRectF(x0, cy_ - hh / 2, max(1.0, x1 - x0) + 0.6, hh)
                src = QRectF(i, 0, step, H)
            else:
                y0 = oy + (i * (1 - SQUEEZE * c0))
                y1 = oy + (min(n, i + step) * (1 - SQUEEZE * c1))
                ww = W * sc
                dst = QRectF(cx_ - ww / 2, y0, ww, max(1.0, y1 - y0) + 0.6)
                src = QRectF(0, i, W, step)
            p.drawImage(dst, paper, src)
            i += step

        # smooth shading: darker towards the lifted edge (and a bit more on the back side)
        from PyQt6.QtGui import QLinearGradient
        if vertical_hinge:
            g = QLinearGradient(ox, 0, ox + W * (1 - SQUEEZE), 0)
            area = QRectF(ox, oy - H * LIFT / 2, W * (1 - SQUEEZE), H * (1 + LIFT))
        else:
            g = QLinearGradient(0, oy, 0, oy + H * (1 - SQUEEZE))
            area = QRectF(ox - W * LIFT / 2, oy, W * (1 + LIFT), H * (1 - SQUEEZE))
        base = 22 if back else 0
        g.setColorAt(0.0, QColor(20, 30, 60, base))
        g.setColorAt(0.55, QColor(20, 30, 60, base + 14))
        g.setColorAt(1.0, QColor(20, 30, 60, base + 85))
        # clip the gradient to the bent paper outline so the dark background is not tinted
        outline = QPainterPath()
        N = 60
        pts_a, pts_b = [], []
        for j in range(N + 1):
            t = j / N
            c = curve(t)
            sc = 1 + LIFT * c
            if vertical_hinge:
                x = ox + t * n * (1 - SQUEEZE * c)
                pts_a.append(QPointF(x, cy_ - H * sc / 2))
                pts_b.append(QPointF(x, cy_ + H * sc / 2))
            else:
                y = oy + t * n * (1 - SQUEEZE * c)
                pts_a.append(QPointF(cx_ - W * sc / 2, y))
                pts_b.append(QPointF(cx_ + W * sc / 2, y))
        outline.moveTo(pts_a[0])
        for q in pts_a[1:] + pts_b[::-1]:
            outline.lineTo(q)
        outline.closeSubpath()
        p.save()
        p.setClipPath(outline)
        p.fillRect(area, g)
        p.restore()

        # hinge line (where the sheet turns) + curved arrow showing the turning direction
        accent = QColor("#ff9f1c")
        pen = QPen(accent, 3, Qt.PenStyle.DashLine)
        p.save()
        p.setPen(pen)
        if vertical_hinge:
            p.drawLine(QPointF(ox - 3, oy - 6), QPointF(ox - 3, oy + H + 6))
        else:
            p.drawLine(QPointF(ox - 6, oy - 3), QPointF(ox + W + 6, oy - 3))
        # arrow: arc starting at the hinge side, curving over to the free side
        path = QPainterPath()
        if vertical_hinge:
            a, b = QPointF(ox + W * 0.18, oy - 14), QPointF(ox + W * 0.82, oy - 14)
            path.moveTo(a)
            path.quadTo(QPointF((a.x() + b.x()) / 2, oy - 40), b)
            tip, d1, d2 = b, QPointF(-9, -2), QPointF(-4, 8)
        else:
            a, b = QPointF(ox + W + 14, oy + H * 0.18), QPointF(ox + W + 14, oy + H * 0.82)
            path.moveTo(a)
            path.quadTo(QPointF(ox + W + 40, (a.y() + b.y()) / 2), b)
            tip, d1, d2 = b, QPointF(-2, -9), QPointF(8, -4)
        pen2 = QPen(accent, 2.5)
        pen2.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen2)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawPath(path)
        head = QPainterPath()
        head.moveTo(tip)
        head.lineTo(tip + d1)
        head.lineTo(tip + d2)
        head.closeSubpath()
        p.setBrush(accent)
        p.drawPath(head)
        p.restore()

        # front / back badge
        txt = "BACK" if back else "FRONT"
        p.save()
        f = p.font()
        f.setBold(True)
        f.setPointSize(8)
        p.setFont(f)
        bw, bh = 52, 18
        br = QRectF(self._fw - bw - 8, self._fh - bh - 8, bw, bh)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#ff9f1c") if back else QColor("#1473E6"))
        p.drawRoundedRect(br, 9, 9)
        p.setPen(QColor("white"))
        p.drawText(br, Qt.AlignmentFlag.AlignCenter, txt)
        p.restore()

    def update_preview(self, *_):
        pg = self.pages()
        self.btn_print.setEnabled(bool(pg))
        bad = not pg
        self.ed_range.setStyleSheet("border:1px solid #D7373F" if bad and self.rb_range.isChecked() else "")
        if bad:
            self.lbl_pv.setPixmap(QPixmap())
            self.lbl_pv.setText("Invalid page range")
            self.lbl_info.setText("")
            self.bt_prev.setEnabled(False)
            self.bt_next.setEnabled(False)
            return

        grp = self.groups()
        self.pv = max(0, min(self.pv, len(grp) - 1))
        self.bt_prev.setEnabled(self.pv > 0)
        self.bt_next.setEnabled(self.pv < len(grp) - 1)

        sheet = grp[self.pv]
        per_sheet = self.per_sheet()
        rows, cols = _grid_for(per_sheet)
        ls = _sheet_landscape(self.cb_orient.currentData(), per_sheet, rows, cols, self.doc[sheet[0]].rect)
        paper_key = self.cb_paper.currentData()
        ps = page_size_for(paper_key, self.custom_mm()).size(QPageSize.Unit.Point)
        paper_label = "{:g} × {:g} mm".format(*self.custom_mm()) if paper_key == CUSTOM_PAPER else paper_key
        pw, ph = (ps.height(), ps.width()) if ls else (ps.width(), ps.height())
        duplex = self.duplex_value()
        back = duplex != "none" and self.pv % 2 == 1
        # the frame always has the printed sheet's proportions (A4 portrait, landscape, custom size ...),
        # so what you drag and zoom matches the page exactly
        fit = min(self.PW / pw, self.PH / ph)
        self._fw, self._fh = max(1, round(pw * fit)), max(1, round(ph * fit))
        self.lbl_pv.setFixedSize(self._fw + 2, self._fh + 2)
        pad = 0 if duplex == "none" else 64                                # extra room for the curl
        k = min((self._fw - pad) / pw, (self._fh - pad) / ph) * self.zoom  # preview px per point
        self._k = k
        ox, oy = (self._fw - pw * k) / 2, (self._fh - ph * k) / 2          # centred; overflow is clipped
        self.lbl_zoom.setText(f"{round(self.zoom * 100)}%")
        m = PRINT_MARGIN_MM / 25.4 * 72
        aw, ah = pw - 2 * m, ph - 2 * m
        base_x, base_y = m * k, m * k                                      # relative to the paper's top-left

        offset_pt = (self.ox_mm / 25.4 * 72, self.oy_mm / 25.4 * 72)
        rotation = self.image_rotation % 360
        cells = layout_sheet(self.doc, sheet, rows, cols, aw, ah,
                             self.cb_scale.currentData(), self.sp_scale.value(), offset_pt, rotation)

        # 1) render the flat sheet into its own image
        paper = QImage(max(1, round(pw * k)), max(1, round(ph * k)), QImage.Format.Format_ARGB32_Premultiplied)
        paper.fill(QColor("white"))
        pp = QPainter(paper)
        first_center = None
        for pno, clip, target in cells:
            cx, cy, cw, ch = clip
            tx, ty, dw, dh = target
            clip_rect = QRectF(base_x + cx * k, base_y + cy * k, cw * k, ch * k)
            pp.save()
            pp.setClipRect(clip_rect)
            if len(cells) > 1:
                pp.setPen(QColor("#cccccc"))
                pp.drawRect(clip_rect.adjusted(0, 0, -1, -1))
            pr = self.doc[pno].rect
            bw, _bh = rot_bbox(pr.width, pr.height, rotation)
            s = dw / bw if bw else 1.0
            uw, uh = pr.width * s, pr.height * s                           # the page's own size on paper
            mscale = max(0.05, min(s * k, 2.5))                            # cap preview render cost
            img = pix_to_qimage(self.doc[pno].get_pixmap(matrix=fitz.Matrix(mscale, mscale), alpha=False))
            if self.ck_grayscale.isChecked():
                img = img.convertToFormat(QImage.Format.Format_Grayscale8).convertToFormat(QImage.Format.Format_RGB32)
            ccx, ccy = base_x + (tx + dw / 2) * k, base_y + (ty + dh / 2) * k
            if first_center is None:
                first_center = (ccx, ccy, uw * k, uh * k)
            pp.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
            pp.translate(ccx, ccy)
            pp.rotate(rotation)
            pp.drawImage(QRectF(-uw * k / 2, -uh * k / 2, uw * k, uh * k), img)
            pp.restore()
        pp.end()

        # 2) put it on the preview canvas (flat, or bent towards the flip edge for two-sided)
        pm = QPixmap(self._fw, self._fh)
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        if duplex == "none":
            p.setPen(Qt.PenStyle.NoPen)
            for kk, aa in ((14, 10), (9, 16), (5, 24), (2, 34)):             # soft shadow under the sheet
                p.setBrush(QColor(0, 0, 0, aa))
                p.drawRoundedRect(QRectF(ox - kk, oy - kk + 4, paper.width() + 2 * kk, paper.height() + 2 * kk),
                                  3 + kk, 3 + kk)
            p.drawImage(QPointF(ox, oy), paper)
        else:
            self._draw_flipped(p, paper, ox, oy, duplex, back)
        p.end()
        self.lbl_pv.setText("")
        self.lbl_pv.setPixmap(pm)
        if first_center is not None:                                       # for the Shift rotate frame (label coords)
            fcx, fcy, fw_, fh_ = first_center
            self.lbl_pv.set_frame(ox + fcx + (self.lbl_pv.width() - pm.width()) / 2,
                                  oy + fcy + (self.lbl_pv.height() - pm.height()) / 2,
                                  fw_, fh_, rotation)
        else:
            self.lbl_pv.set_frame(None)

        rng = f"{sheet[0] + 1}" + (f"–{sheet[-1] + 1}" if len(sheet) > 1 else "")
        info = f"Sheet {self.pv + 1} / {len(grp)}  •  page(s) {rng}\n{paper_label} {'landscape' if ls else 'portrait'}"
        if duplex != "none":
            info += f"\nsheet {self.pv // 2 + 1}, {'front' if self.pv % 2 == 0 else 'back'}"
        self.lbl_info.setText(info)



# ============================================================================
#  OBJECT EDITING HELPERS  (Edit mode: text, images and vector shapes)
# ============================================================================
def _rgb_from_int(c):
    return (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255)


def _redact_flags(images=0, graphics=0, text=1):
    """Keyword args for Page.apply_redactions that touch only the chosen kind of content."""
    return dict(images=images, graphics=graphics, text=text)


def _apply_redactions(page, images=None, graphics=None, text=None):
    """Remove only the requested content kinds (older PyMuPDF versions: fall back to the plain call)."""
    kw = {}
    kw["images"] = getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0) if images is None else images
    kw["graphics"] = getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0) if graphics is None else graphics
    kw["text"] = getattr(fitz, "PDF_REDACT_TEXT_NONE", 1) if text is None else text
    try:
        page.apply_redactions(**kw)
    except TypeError:
        page.apply_redactions(images=kw["images"])


def _emit_text(page, text, origin, size, flags, color, box=None):
    """Write one line of text onto the page at `origin` (baseline-left). Latin text uses the standard PDF fonts
    (matching bold / italic / serif of the original); other scripts (Persian, Arabic ...) are shaped by the HTML engine."""
    if not text:
        return
    size = max(1.0, float(size))
    if all(ord(ch) < 256 for ch in text):
        pre = "ti" if flags & 4 else "he"
        fn = {"ti": {"ro": "tiro", "bo": "tibo", "it": "tiit", "bi": "tibi"},
              "he": {"lv": "helv", "bo": "hebo", "it": "heit", "bi": "hebi"}}[pre]
        key = {(0, 0): "ro" if pre == "ti" else "lv", (1, 0): "bo", (0, 1): "it", (1, 1): "bi"}[
            (1 if flags & 16 else 0, 1 if flags & 2 else 0)]
        page.insert_text(fitz.Point(origin), text, fontsize=size, fontname=fn[key], color=color)
        return
    rtl = bool(re.search(r"[\u0590-\u08FF\uFB50-\uFEFF]", text))
    fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    css = (f'@font-face{{font-family:ft;src:url("tahoma.ttf");}}'
           f'*{{font-family:ft;font-size:{size}pt;'
           f'color:rgb({int(color[0]*255)},{int(color[1]*255)},{int(color[2]*255)});}}p{{margin:0}}')
    if box is None:
        box = fitz.Rect(origin.x, origin.y - size, origin.x + size * max(2, len(text)), origin.y + size * 0.4)
    box = fitz.Rect(box.x0 - (60 if rtl else 0), box.y0 - 2, box.x1 + (0 if rtl else 60), box.y1 + 2)
    try:
        page.insert_htmlbox(box, f'<p dir="{"rtl" if rtl else "ltr"}" style="text-align:'
                                 f'{"right" if rtl else "left"}">{_html.escape(text)}</p>',
                            css=css, archive=fitz.Archive(fonts))
    except Exception:                                       # no Windows fonts (other OS): keep what Latin-1 can show
        page.insert_text(fitz.Point(origin), text.encode("latin-1", "replace").decode("latin-1"),
                         fontsize=size, fontname="helv", color=color)


def _redraw_path(page, path, T, avg):
    """Draw a vector path again with every point passed through T (move / scale)."""
    if not path.get("items"):
        return
    sh = page.new_shape()
    for it in path["items"]:
        op = it[0]
        if op == "l":
            sh.draw_line(T(it[1]), T(it[2]))
        elif op == "c":
            sh.draw_bezier(T(it[1]), T(it[2]), T(it[3]), T(it[4]))
        elif op == "re":
            a, b = T(it[1].tl), T(it[1].br)
            sh.draw_rect(fitz.Rect(min(a.x, b.x), min(a.y, b.y), max(a.x, b.x), max(a.y, b.y)))
        elif op == "qu":
            q = it[1]
            sh.draw_quad(fitz.Quad(T(q.ul), T(q.ur), T(q.ll), T(q.lr)))
    fill, color = path.get("fill"), path.get("color")
    if fill is None and color is None:
        return
    cap = path.get("lineCap") or (0,)
    dashes = path.get("dashes")
    try:
        sh.finish(fill=fill, color=color, width=max(0.05, (path.get("width") or 1.0) * avg),
                  dashes=dashes if dashes and dashes.strip() not in ("[] 0", "") else None,
                  even_odd=bool(path.get("even_odd")), closePath=bool(path.get("closePath")),
                  lineCap=int(cap[0]) if isinstance(cap, (tuple, list)) else int(cap),
                  fill_opacity=path.get("fill_opacity") if path.get("fill_opacity") is not None else 1,
                  stroke_opacity=path.get("stroke_opacity") if path.get("stroke_opacity") is not None else 1)
        sh.commit()
    except Exception:
        pass


class InlineTextEditor(QPlainTextEdit):
    """Small text box that appears on the page while a text object is typed into."""
    committed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, parent, text, px_size, width, height):
        super().__init__(parent)
        self._done = False
        f = QFont("Arial")
        f.setPixelSize(max(8, int(px_size)))
        self.setFont(f)
        self.setPlainText(text)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setStyleSheet(f"QPlainTextEdit{{background:rgba(255,255,255,235);color:#111;"
                           f"border:2px solid {ACCENT};selection-background-color:{ACCENT};}}")
        self.document().setDocumentMargin(2)
        self.resize(int(max(width, 60)), int(max(height, px_size * 1.6)))
        self.moveCursor(QTextCursor.MoveOperation.End)

    def _finish(self, ok):
        if self._done:
            return
        self._done = True
        if ok:
            self.committed.emit(self.toPlainText())
        else:
            self.cancelled.emit()

    def keyPressEvent(self, e):
        if e.key() == Qt.Key.Key_Escape:
            self._finish(False)
            return
        if e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and (e.modifiers() & Qt.KeyboardModifier.ControlModifier):
            self._finish(True)
            return
        super().keyPressEvent(e)

    def focusOutEvent(self, e):
        super().focusOutEvent(e)
        self._finish(True)


# ============================================================================
#  DOCUMENT TAB
# ============================================================================
class PageLabel(QLabel):
    clicked = pyqtSignal(int, QPointF)
    context_requested = pyqtSignal(int, QPoint)             # page index, global position of the right-click
    export_drag_requested = pyqtSignal(int)

    def contextMenuEvent(self, e):
        self.context_requested.emit(self.index, e.globalPos())
        e.accept()

    def __init__(self, index, parent):
        super().__init__(parent)
        self.index = index
        self.setScaledContents(True)
        self.setStyleSheet("background:white")
        self.export_drag_enabled = False
        self._press_pos = None
        self.ctl = None                                     # the DocTab, which owns Edit-mode behaviour
        self.setMouseTracking(True)

    def _editing(self):
        return self.ctl is not None and self.ctl.edit_mode

    def paintEvent(self, e):
        super().paintEvent(e)
        if self._editing():
            p = QPainter(self)
            p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
            self.ctl.paint_overlay(self, p)
            p.end()

    def mousePressEvent(self, e):
        if self._editing():
            if e.button() == Qt.MouseButton.LeftButton:
                self.setFocus()
                self.ctl.edit_press(self, e.position(), e.modifiers())
                e.accept()
            else:
                e.ignore()
            return
        if e.button() == Qt.MouseButton.LeftButton:
            self._press_pos = e.position()
            self.clicked.emit(self.index, e.position())
            e.accept()
        else:
            e.ignore()                                      # e.g. middle button -> handled by the scroll area

    def mouseDoubleClickEvent(self, e):
        if self._editing() and e.button() == Qt.MouseButton.LeftButton:
            self.ctl.edit_double_click(self, e.position())
            e.accept()
            return
        super().mouseDoubleClickEvent(e)

    def keyPressEvent(self, e):
        if self._editing() and self.ctl.edit_key(self, e):
            e.accept()
            return
        super().keyPressEvent(e)

    def mouseMoveEvent(self, e):
        if self._editing():
            self.ctl.edit_move(self, e.position(), e.buttons(), e.modifiers())
            e.accept()
            return
        if (self.export_drag_enabled and self._press_pos is not None
                and e.buttons() & Qt.MouseButton.LeftButton):
            distance = (e.position().toPoint() - self._press_pos.toPoint()).manhattanLength()
            if distance >= QApplication.startDragDistance():
                self._press_pos = None
                self.export_drag_requested.emit(self.index)
            e.accept()
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self._editing():
            if e.button() == Qt.MouseButton.LeftButton:
                self.ctl.edit_release(self, e.position(), e.modifiers())
            e.accept()
            return
        self._press_pos = None
        super().mouseReleaseEvent(e)

    def leaveEvent(self, e):
        if self._editing():
            self.ctl.edit_leave(self)
        super().leaveEvent(e)


class PageCanvas(QWidget):
    """Container of the continuous page view: transparent, with a soft shadow under each page."""
    def __init__(self):
        super().__init__()
        self.setAutoFillBackground(False)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setPen(Qt.PenStyle.NoPen)
        r = e.rect()
        for w in self.children():
            if not isinstance(w, PageLabel) or w.isHidden():
                continue
            g = QRectF(w.geometry())
            if not g.adjusted(-32, -32, 32, 40).intersects(QRectF(r)):
                continue
            for k, a in ((6, 10), (3, 18), (1, 30)):
                p.setBrush(QColor(0, 0, 0, a))
                p.drawRoundedRect(g.adjusted(-k, -k + 2, k, k + 2), 1 + k / 2, 1 + k / 2)
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.setPen(QPen(QColor(*T["RIM"]), 1))
            p.drawRect(g.adjusted(-0.5, -0.5, 0.5, 0.5))
            p.setPen(Qt.PenStyle.NoPen)


def _drop_paths(e):
    """Supported files (PDF / images) carried by a drag-and-drop event, in natural order."""
    md = e.mimeData()
    if md.hasFormat("application/x-pdfstudio-page-export"):
        return []                                          # page-export drags are for external targets, not reinsertion
    if not md.hasUrls():
        return []
    ps = [u.toLocalFile() for u in md.urls()]
    return sorted((x for x in ps if x.lower().endswith(ALL_EXT) and os.path.isfile(x)), key=_natural_key)


class ZoomScroll(QScrollArea):
    """Page view scroll area. Ctrl+wheel zooms; PDF/image files dragged over it are
    reported (with the y position inside the page column) so they can be inserted."""
    zoomed = pyqtSignal(int)
    files_hover = pyqtSignal(int)                    # y in container coordinates
    files_drop = pyqtSignal(list, int)
    files_leave = pyqtSignal()

    AS_DEAD = 10                                     # px around the anchor with no scrolling

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        # middle-button autoscroll (like a web browser)
        self._as_on, self._as_anchor, self._as_rem = False, QPoint(), 0.0
        self._as_timer = QTimer(self)
        self._as_timer.setInterval(16)
        self._as_timer.timeout.connect(self._as_tick)
        self._as_mark = QLabel("⬍", self.viewport())
        self._as_mark.setFixedSize(30, 30)
        self._as_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._as_mark.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._as_mark.setStyleSheet(f"background:{T['SURFACE']};color:{T['TEXT']};border:2px solid {ACCENT};"
                                    "border-radius:15px;font-size:16px")
        self._as_mark.hide()

    # ---- middle-click autoscroll ----
    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.MiddleButton and not self._as_on:
            self._as_start(e.globalPosition().toPoint())
            e.accept()
            return
        super().mousePressEvent(e)

    def _as_start(self, gpos):
        self._as_on, self._as_anchor, self._as_rem, self._as_moved = True, gpos, 0.0, False
        c = self.viewport().mapFromGlobal(gpos)
        self._as_mark.move(c.x() - 15, c.y() - 15)
        self._as_mark.show()
        self._as_mark.raise_()
        QApplication.setOverrideCursor(Qt.CursorShape.SizeVerCursor)
        QApplication.instance().installEventFilter(self)
        self._as_timer.start()

    def _as_stop(self):
        if not self._as_on:
            return
        self._as_on = False
        self._as_timer.stop()
        self._as_mark.hide()
        QApplication.instance().removeEventFilter(self)
        QApplication.restoreOverrideCursor()

    def _as_tick(self):
        dy = QCursor.pos().y() - self._as_anchor.y()
        v = abs(dy) - self.AS_DEAD
        if v <= 0:
            return
        if v > 8:
            self._as_moved = True
        speed = min(v * 0.35 + v * v * 0.004, 320.0)        # px per tick, grows with distance
        self._as_rem += speed if dy > 0 else -speed
        step = int(self._as_rem)
        self._as_rem -= step
        if step:
            sb = self.verticalScrollBar()
            sb.setValue(sb.value() + step)

    def eventFilter(self, obj, ev):
        if self._as_on:
            t = ev.type()
            if t == QEvent.Type.MouseButtonPress:           # any click ends autoscroll
                self._as_stop()
                return True
            if t == QEvent.Type.MouseButtonRelease and ev.button() == Qt.MouseButton.MiddleButton:
                if self._as_moved:                          # press-drag-release mode
                    self._as_stop()
                return True
            if t == QEvent.Type.KeyPress and ev.key() == Qt.Key.Key_Escape:
                self._as_stop()
                return True
            if t in (QEvent.Type.WindowDeactivate, QEvent.Type.Hide) and obj is self.window():
                self._as_stop()
        return super().eventFilter(obj, ev)

    def hideEvent(self, e):
        self._as_stop()
        super().hideEvent(e)

    def wheelEvent(self, e):
        if e.modifiers() & Qt.KeyboardModifier.ControlModifier:
            self.zoomed.emit(e.angleDelta().y())
            e.accept()
        else:
            super().wheelEvent(e)

    def _y(self, e):
        return int(e.position().y()) + self.verticalScrollBar().value()

    def dragEnterEvent(self, e):
        if _drop_paths(e):
            e.acceptProposedAction()
            self.files_hover.emit(self._y(e))
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if not _drop_paths(e):
            e.ignore()
            return
        e.acceptProposedAction()
        y = int(e.position().y())                    # auto-scroll when hovering near the top / bottom edge
        vh, sb = self.viewport().height(), self.verticalScrollBar()
        if y < 40:
            sb.setValue(sb.value() - 18)
        elif y > vh - 40:
            sb.setValue(sb.value() + 18)
        self.files_hover.emit(self._y(e))

    def dragLeaveEvent(self, e):
        self.files_leave.emit()

    def dropEvent(self, e):
        paths = _drop_paths(e)
        self.files_leave.emit()
        if paths:
            e.acceptProposedAction()
            self.files_drop.emit(paths, self._y(e))
        else:
            e.ignore()


class ThumbList(QListWidget):
    """Thumbnail strip; accepts PDF/image files dropped between thumbnails."""
    files_drop = pyqtSignal(list, int)               # paths, insertion index
    layout_changed = pyqtSignal()                    # scrolled / resized -> reposition delete buttons
    pages_moved = pyqtSignal(list, int)              # dragged rows, destination slot (0..n)
    context_requested = pyqtSignal(int, QPoint)      # row, global position of the right-click
    export_drag_requested = pyqtSignal(int)
    export_selection_drag_requested = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDragEnabled(True)                    # drag thumbnails to reorder pages
        self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)
        self.setDefaultDropAction(Qt.DropAction.CopyAction)
        self.setObjectName("Thumbs")
        self.setSpacing(4)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.export_drag_index = None
        self._line = QFrame(self.viewport())
        self._line.setStyleSheet(f"background:{ACCENT};border-radius:2px")
        self._line.hide()

    def scrollContentsBy(self, dx, dy):
        super().scrollContentsBy(dx, dy)
        self.layout_changed.emit()

    def updateGeometries(self):
        super().updateGeometries()
        self.layout_changed.emit()

    def contextMenuEvent(self, e):
        it = self.itemAt(e.pos())
        if it is not None:
            self.context_requested.emit(self.row(it), e.globalPos())
        e.accept()

    def _slot(self, e):
        y, n = e.position().y(), self.count()
        for i in range(n):
            if y < self.visualItemRect(self.item(i)).center().y():
                return i
        return n

    def _show_line(self, idx):
        n = self.count()
        if n == 0:
            return
        if idx < n:
            ly = self.visualItemRect(self.item(idx)).top() - 1
        else:
            ly = self.visualItemRect(self.item(n - 1)).bottom() + 1
        self._line.setGeometry(0, max(0, ly - 1), self.viewport().width(), 4)
        self._line.show()
        self._line.raise_()

    def _internal(self, e):
        return (e.source() is self
                and e.mimeData().hasFormat("application/x-pdfstudio-page-reorder"))

    def startDrag(self, supportedActions):
        rows = sorted({self.row(item) for item in self.selectedItems()})
        if rows:
            # DocTab supplies real file URLs so Explorer can receive the pages.
            # The same MIME payload also lets this list recognize an internal reorder.
            self.export_selection_drag_requested.emit(rows)
            return
        super().startDrag(supportedActions)

    def _autoscroll(self, e):
        y, vh, sb = int(e.position().y()), self.viewport().height(), self.verticalScrollBar()
        if y < 40:
            sb.setValue(sb.value() - 18)
        elif y > vh - 40:
            sb.setValue(sb.value() + 18)

    def dragEnterEvent(self, e):
        if self._internal(e):
            e.setDropAction(Qt.DropAction.MoveAction)
            e.accept()
            self._show_line(self._slot(e))
        elif _drop_paths(e):
            e.acceptProposedAction()
            self._show_line(self._slot(e))
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if self._internal(e):
            e.setDropAction(Qt.DropAction.MoveAction)
            e.accept()
            self._autoscroll(e)
            self._show_line(self._slot(e))
        elif _drop_paths(e):
            e.acceptProposedAction()
            self._autoscroll(e)
            self._show_line(self._slot(e))
        else:
            e.ignore()

    def dragLeaveEvent(self, e):
        self._line.hide()

    def dropEvent(self, e):
        self._line.hide()
        if self._internal(e):
            try:
                rows = sorted({int(x) for x in bytes(
                    e.mimeData().data("application/x-pdfstudio-page-reorder")).decode("ascii").split(",")
                    if x.strip()})
            except (ValueError, UnicodeDecodeError):
                rows = sorted({self.row(i) for i in self.selectedItems()})
            slot = self._slot(e)
            e.setDropAction(Qt.DropAction.MoveAction)   # page order changes only after our model update
            e.accept()
            if rows:
                self.pages_moved.emit(rows, slot)
            return
        paths = _drop_paths(e)
        if paths:
            e.acceptProposedAction()
            self.files_drop.emit(paths, self._slot(e))
        else:
            e.ignore()


class DelButton(QPushButton):
    """Small trash-can button in the top-left corner of a page thumbnail: deletes exactly that page."""
    def __init__(self, index, parent):
        super().__init__("", parent)
        self.index = index
        self.setFixedSize(24, 24)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setToolTip(f"Delete page {index + 1}  (Ctrl+Z to undo)")
        self.setStyleSheet("QPushButton{border:none;background:transparent}")

    def paintEvent(self, e):
        hot = self.underMouse()
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setPen(QPen(QColor("#D7373F" if hot else T["LINE2"]), 1))
        p.setBrush(QColor("#D7373F") if hot else QColor(T["SURFACE"]))
        p.drawRoundedRect(QRectF(0.5, 0.5, self.width() - 1, self.height() - 1), 6, 6)
        col = QColor("white") if hot else QColor(T["TEXT2"])
        p.setPen(QPen(col, 1.4, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.translate(self.width() / 2, self.height() / 2)
        p.drawLine(QPointF(-5.5, -4.5), QPointF(5.5, -4.5))                  # lid
        p.drawLine(QPointF(-1.8, -4.5), QPointF(-1.8, -6.2))                 # handle
        p.drawLine(QPointF(1.8, -4.5), QPointF(1.8, -6.2))
        p.drawLine(QPointF(-1.8, -6.2), QPointF(1.8, -6.2))
        body = QPainterPath()                                                # bin body, slightly tapered
        body.moveTo(-4.2, -2.6)
        body.lineTo(-3.5, 5.6)
        body.lineTo(3.5, 5.6)
        body.lineTo(4.2, -2.6)
        p.drawPath(body)
        p.drawLine(QPointF(-1.5, -0.4), QPointF(-1.5, 3.4))                  # ribs
        p.drawLine(QPointF(1.5, -0.4), QPointF(1.5, 3.4))
        p.end()

    def enterEvent(self, e):
        self.update()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self.update()
        super().leaveEvent(e)


class DocTab(QWidget):
    """One open PDF: thumbnail strip + continuous page view + its own undo history."""
    changed = pyqtSignal()

    def __init__(self, main):
        super().__init__()
        self.main = main
        self.doc, self.path, self.cur = None, None, 0
        self.zoom, self.undo_stack, self.dirty = 1.4, [], False
        self.labels, self.ys, self.hs, self.imgs = [], [], [], {}
        self._armed_export_format = None
        self.edit_mode = False
        self._objs = {}                                     # page index -> editable objects (text / image / drawing)
        self.edit_sel, self.edit_page = [], None            # selected objects (all on one page)
        self._drag, self._ghost, self._hover, self._editor = None, None, None, None
        self._nav, self._pending, self._frac, self._screen_off = False, False, None, None
        self._center_h = False                              # centre the page horizontally after (re)loading

        self.del_btns = []
        self.thumbs = ThumbList()
        self.thumbs.setIconSize(QSize(118, 160))
        self.thumbs.setFixedWidth(172)
        self.thumbs.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.thumbs.currentRowChanged.connect(self.goto_page)
        self.thumbs.layout_changed.connect(self.place_del_buttons)
        self.thumbs.verticalScrollBar().valueChanged.connect(lambda _: self.place_del_buttons())
        self.thumbs.export_drag_requested.connect(
            lambda idx: self.export_page_drag(idx, source=self.thumbs))
        self.thumbs.export_selection_drag_requested.connect(
            lambda rows: self.export_pages_drag(rows, source=self.thumbs))

        # continuous page view: one container, pages laid out vertically inside it
        self.container = PageCanvas()
        self.scroll = ZoomScroll()
        self.scroll.setObjectName("PageScroll")
        self.scroll.setWidget(self.container)
        self.container.setAutoFillBackground(False)         # QScrollArea turns this on; the canvas colour comes from the stylesheet
        self.scroll.viewport().setAutoFillBackground(False)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.scroll.verticalScrollBar().valueChanged.connect(self.on_scroll)
        self.scroll.zoomed.connect(lambda d: self.set_zoom(self.zoom * (1.1 if d > 0 else 1 / 1.1)))
        self.scroll.files_hover.connect(self.show_drop_line)
        self.scroll.files_leave.connect(self.hide_drop_line)
        self.scroll.files_drop.connect(self.drop_files_at_y)
        self.thumbs.files_drop.connect(self.insert_files)
        self.thumbs.context_requested.connect(self.page_menu)
        self.thumbs.pages_moved.connect(lambda rows, slot: QTimer.singleShot(0, lambda: self.move_pages(rows, slot)))
        self.drop_line = QFrame(self.container)
        self.drop_line.setStyleSheet(f"background:{ACCENT};border-radius:2px")
        self.drop_line.hide()

        split = QSplitter()
        split.setHandleWidth(4)
        split.setChildrenCollapsible(False)
        split.addWidget(self.thumbs)
        split.addWidget(self.scroll)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([self.thumbs.width(), 1100])
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(split)

    # ---------- state ----------
    def set_dirty(self, v):
        self.dirty = v
        self.changed.emit()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.update_visible()

    def showEvent(self, e):
        super().showEvent(e)
        if self._pending:
            QTimer.singleShot(0, self._after_build)

    # ---------- file ----------
    def load(self, p):
        try:
            d = fitz.open(p)
            if d.needs_pass:
                pw, ok = QInputDialog.getText(self.main, "Password", "Password:", QLineEdit.EchoMode.Password)
                if not ok or not d.authenticate(pw):
                    return False
        except Exception as e:
            QMessageBox.warning(self.main, "Error", str(e))
            return False
        self.doc, self.path, self.cur, self.undo_stack, self.dirty = d, p, 0, [], False
        self._center_h = True
        self.refresh_thumbs()
        return True

    def save(self):
        if not self.doc:
            return False
        if not self.path:
            return self.save_as()
        return self._write(self.path)

    def save_as(self):
        if not self.doc:
            return False
        p, _ = QFileDialog.getSaveFileName(self.main, "Save PDF", self.path or "", "PDF (*.pdf)")
        return self._write(p) if p else False

    def _write(self, p):
        try:
            data = self.doc.tobytes(garbage=3, deflate=True)
            with open(p, "wb") as f:
                f.write(data)
        except Exception as e:
            QMessageBox.warning(self.main, "Save failed", str(e))
            return False
        self.path = p
        self.set_dirty(False)
        return True

    # ---------- rendering (continuous view) ----------
    def thumb_icon(self, i):
        pix = self.doc[i].get_pixmap(matrix=fitz.Matrix(0.32, 0.32), alpha=False)
        img = QPixmap.fromImage(pix_to_qimage(pix))
        out = QPixmap(img.size())                           # square page with a thin grey outline
        out.fill(Qt.GlobalColor.transparent)
        p = QPainter(out)
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        clip = QPainterPath()
        clip.addRect(QRectF(out.rect()))
        p.setClipPath(clip)
        p.drawPixmap(0, 0, img)
        p.setClipping(False)
        p.setPen(QPen(QColor(*T["THUMB_RIM"]), 1))
        p.drawRect(QRectF(0.5, 0.5, out.width() - 1, out.height() - 1))
        p.end()
        ic = QIcon()
        for mode in (QIcon.Mode.Normal, QIcon.Mode.Active, QIcon.Mode.Selected):
            ic.addPixmap(out, mode)                          # same pixmap for every state -> no blue selection tint
        return ic

    def refresh_thumbs(self, keep=None, screen_off=None):
        keep = self.cur if keep is None else keep
        self.thumbs.blockSignals(True)
        self.thumbs.clear()
        for i in range(len(self.doc)):
            item = QListWidgetItem(self.thumb_icon(i), str(i + 1))
            item.setSizeHint(QSize(140, 180))                 # room for the padding around each thumbnail
            item.setToolTip(f"Drag page {i + 1} to a folder or the Desktop to export it. "
                            "Select several pages to export them together.")
            self.thumbs.addItem(item)
        self.cur = max(0, min(keep, len(self.doc) - 1))
        self.thumbs.setCurrentRow(self.cur)
        self.thumbs.blockSignals(False)
        for b in self.del_btns:
            b.hide()
            b.deleteLater()
        self.del_btns = []
        for i in range(len(self.doc)):
            btn = DelButton(i, self.thumbs.viewport())
            btn.clicked.connect(lambda _=False, b=btn: self.delete_page_at(b.index))
            self.del_btns.append(btn)
        self.place_del_buttons()
        QTimer.singleShot(0, self.place_del_buttons)
        self.build_pages(screen_off=screen_off)

    def build_pages(self, frac=None, screen_off=None):
        """Lay out every page (as an empty white box) in one tall column. Page images
        are rendered lazily, only for pages near the visible area."""
        self._clear_armed_export()
        for lab in self.labels:
            lab.hide()
            lab.deleteLater()
        self.labels, self.ys, self.hs, self.imgs = [], [], [], {}
        self._frac, self._screen_off = frac, screen_off
        self._reset_edit_state()
        self.drop_line.hide()
        if not self.doc:
            return
        sizes = []
        for p in self.doc:
            r = p.rect
            sizes.append((math.ceil(r.width * self.zoom), math.ceil(r.height * self.zoom)))
        maxw = max(w for w, _ in sizes)
        total_w = maxw + 2 * PAGE_MARGIN
        cur_cursor = Qt.CursorShape.ArrowCursor
        y = PAGE_MARGIN
        for i, (w, h) in enumerate(sizes):
            lab = PageLabel(i, self.container)
            lab.setGeometry((total_w - w) // 2, y, w, h)
            lab.setCursor(cur_cursor)
            lab.ctl = self
            lab.setFocusPolicy(Qt.FocusPolicy.StrongFocus if self.edit_mode else Qt.FocusPolicy.NoFocus)
            lab.export_drag_enabled = not self.edit_mode
            lab.setToolTip(f"Drag page {i + 1} to the Desktop to create a separate PDF; "
                           "right-click to choose PDF or JPG. Files remain available for 7 days.")
            lab.clicked.connect(self.on_click)
            lab.context_requested.connect(self.page_menu)
            lab.export_drag_requested.connect(self.export_page_drag)
            lab.show()
            self.labels.append(lab)
            self.ys.append(y)
            self.hs.append(h)
            y += h + PAGE_GAP
        self.container.resize(total_w, y - PAGE_GAP + PAGE_MARGIN)
        self._pending = True
        QTimer.singleShot(0, self._after_build)
        self.changed.emit()

    def place_del_buttons(self):
        """Each trash button sits in the top-left corner of its page thumbnail."""
        vp = self.thumbs.viewport()
        for i, b in enumerate(self.del_btns):
            it = self.thumbs.item(i)
            if it is None:
                b.hide()
                continue
            r = self.thumbs.visualItemRect(it)
            if r.bottom() < 0 or r.top() > vp.height():
                b.hide()
                continue
            b.move(r.left() + 4, r.top() + 11)                 # inside the thumbnail, top-left corner
            b.show()
            b.raise_()

    def _after_build(self):
        if not self.doc or not self.ys or not self.isVisible():
            return                                          # showEvent will call us again
        self._pending = False
        self._nav = True
        base = self.ys[self.cur] - PAGE_MARGIN
        off = int(self._frac * self.hs[self.cur]) if self._frac is not None else 0
        target = base + off
        if self._screen_off is not None:                    # keep the page at the same spot on screen
            target = self.ys[self.cur] - self._screen_off
        self.scroll.verticalScrollBar().setValue(max(0, target))
        if self._center_h:
            self._center_horizontally()
            QTimer.singleShot(0, self._center_horizontally)   # again once the scrollbar range is final
            QTimer.singleShot(80, self._finish_center)
        self._nav = False
        self._frac = None
        self._screen_off = None
        self.update_visible()
        self.changed.emit()

    def _center_horizontally(self):
        hb = self.scroll.horizontalScrollBar()
        hb.setValue((hb.minimum() + hb.maximum()) // 2)

    def _finish_center(self):
        self._center_horizontally()
        self._center_h = False

    def render_page(self, i, force=False):
        if i >= len(self.labels) or (i in self.imgs and not force):
            return
        pix = self.doc[i].get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
        img = pix_to_qimage(pix)
        self.imgs[i] = img
        self.labels[i].setPixmap(QPixmap.fromImage(img))

    def update_visible(self):
        if not self.doc or not self.labels:
            return
        sb = self.scroll.verticalScrollBar().value()
        vh = max(self.scroll.viewport().height(), 400)
        lo, hi = sb - vh, sb + 2 * vh
        for i, (y, h) in enumerate(zip(self.ys, self.hs)):
            if y + h >= lo and y <= hi:
                self.render_page(i)
            elif i in self.imgs:                            # far away: free memory
                del self.imgs[i]
                self.labels[i].clear()

    def on_scroll(self, value):
        self.update_visible()
        if self._nav or not self.ys:
            return
        probe = value + self.scroll.viewport().height() // 3
        i = max(0, bisect.bisect_right(self.ys, probe) - 1)
        if i != self.cur:
            self.cur = i
            self.thumbs.blockSignals(True)
            self.thumbs.setCurrentRow(i)
            self.thumbs.blockSignals(False)
            self.place_del_buttons()                        # the list may have auto-scrolled while signals were blocked
            self.changed.emit()

    def goto_page(self, i):
        if not self.doc or i < 0 or i >= len(self.ys):
            return
        self.cur = i
        self._nav = True
        self.scroll.verticalScrollBar().setValue(max(0, self.ys[i] - PAGE_MARGIN))
        self._nav = False
        self.update_visible()
        self.place_del_buttons()
        self.changed.emit()

    def set_zoom(self, z):
        if not self.doc:
            return
        z = max(0.3, min(6.0, z))
        if abs(z - self.zoom) < 1e-6:
            return
        frac = None
        if self.ys:
            base = self.ys[self.cur] - PAGE_MARGIN
            frac = (self.scroll.verticalScrollBar().value() - base) / max(1, self.hs[self.cur])
        self.zoom = z
        self.build_pages(frac)

    # ---------- history ----------
    def snapshot(self):
        self.undo_stack.append(self.doc.tobytes())
        del self.undo_stack[:-15]
        self.set_dirty(True)

    def undo(self):
        if self.undo_stack:
            self.doc = fitz.open("pdf", self.undo_stack.pop())
            self.dirty = True
            self.refresh_thumbs()

    # ---------- page operations ----------
    def delete_pages(self):
        if not self.doc:
            return
        rows = sorted({self.thumbs.row(i) for i in self.thumbs.selectedItems()}) or [self.cur]
        if len(rows) >= len(self.doc):
            return QMessageBox.information(self.main, "PDF Studio", "Can't delete every page.")
        self.snapshot()
        self.doc.delete_pages(rows)
        self.refresh_thumbs(min(rows[0], len(self.doc) - 1))

    def delete_page_at(self, i):
        """Delete one specific page (the x beside it) without having to select it."""
        if not self.doc or not (0 <= i < len(self.doc)):
            return
        if len(self.doc) <= 1:
            return QMessageBox.information(self.main, "PDF Studio", "Can't delete every page.")
        screen_off = self.ys[i] - self.scroll.verticalScrollBar().value() if i < len(self.ys) else None
        self.snapshot()
        self.doc.delete_page(i)
        self.refresh_thumbs(min(i, len(self.doc) - 1), screen_off=screen_off)

    # ---------- dropping files between pages ----------
    def slot_at_y(self, y):
        """Insertion index (0..n) for a y position inside the page column."""
        n = len(self.ys)
        for i in range(n):
            if y < self.ys[i] + self.hs[i] / 2:
                return i
        return n

    def show_drop_line(self, y):
        if not self.ys:
            return
        i, n = self.slot_at_y(y), len(self.ys)
        if i == 0:
            ly = self.ys[0] - PAGE_GAP // 2 - 2
        elif i == n:
            ly = self.ys[-1] + self.hs[-1] + PAGE_GAP // 2 + 2
        else:
            ly = (self.ys[i - 1] + self.hs[i - 1] + self.ys[i]) // 2
        self.drop_line.setGeometry(0, max(0, ly - 2), self.container.width(), 5)
        self.drop_line.show()
        self.drop_line.raise_()

    def hide_drop_line(self):
        self.drop_line.hide()

    def drop_files_at_y(self, paths, y):
        self.insert_files(paths, self.slot_at_y(y))

    def _image_page(self, path, at):
        pm = fitz.Pixmap(path)
        w, h = (842, 595) if pm.width > pm.height else (595, 842)
        page = self.doc.new_page(pno=at, width=w, height=h)
        k = min((w - 40) / pm.width, (h - 40) / pm.height)
        iw, ih = pm.width * k, pm.height * k
        page.insert_image(fitz.Rect(w / 2 - iw / 2, h / 2 - ih / 2, w / 2 + iw / 2, h / 2 + ih / 2),
                          filename=path)

    def insert_files(self, paths, at):
        """Insert PDFs and/or images as pages at index `at` (between existing pages)."""
        if not self.doc or not paths:
            return
        at = max(0, min(at, len(self.doc)))
        first, skipped, was_dirty = at, [], self.dirty
        self.snapshot()
        for p in paths:
            try:
                if p.lower().endswith(IMG_EXT):
                    self._image_page(p, at)
                    at += 1
                else:
                    with fitz.open(p) as src:
                        if src.needs_pass and not src.authenticate(""):
                            skipped.append(os.path.basename(p) + " (password)")
                            continue
                        self.doc.insert_pdf(src, start_at=at)
                        at += len(src)
            except Exception as e:
                skipped.append(f"{os.path.basename(p)} ({e})")
        if at == first:                                     # nothing was added -> drop the useless snapshot
            self.undo_stack.pop()
            self.set_dirty(was_dirty)
        else:
            self.refresh_thumbs(first)
        if skipped:
            QMessageBox.warning(self.main, "PDF Studio", "Skipped:\n" + "\n".join(skipped))

    def add_blank(self):
        if not self.doc:
            return
        self.snapshot()
        r = self.doc[self.cur].rect
        self.doc.new_page(pno=self.cur + 1, width=r.width, height=r.height)
        self.refresh_thumbs(self.cur + 1)

    def insert_pdf(self):
        if not self.doc:
            return
        p, _ = QFileDialog.getOpenFileName(self.main, "Insert PDF after current page", "", "PDF (*.pdf)")
        if p:
            self._insert([p], self.cur + 1)

    def merge_end(self):
        if not self.doc:
            return
        ps, _ = QFileDialog.getOpenFileNames(self.main, "Merge PDFs", "", "PDF (*.pdf)")
        if ps:
            self._insert(ps, len(self.doc))

    def _insert(self, paths, at):
        self.snapshot()
        try:
            for p in paths:
                with fitz.open(p) as src:
                    self.doc.insert_pdf(src, start_at=at)
                    at += len(src)
        except Exception as e:
            QMessageBox.warning(self.main, "Error", str(e))
        self.refresh_thumbs(at - 1)

    def move(self, d):
        if not self.doc:
            return
        c, n = self.cur, len(self.doc)
        if (d < 0 and c == 0) or (d > 0 and c == n - 1):
            return
        self.snapshot()
        self.doc.move_page(c, c - 1) if d < 0 else self.doc.move_page(c, c + 2 if c + 2 < n else -1)
        self.refresh_thumbs(c + d)

    def move_pages(self, rows, slot):
        """Drag-and-drop reorder from the thumbnail strip: move the selected pages to before `slot`."""
        if not self.doc or not rows:
            return
        n = len(self.doc)
        rows = [r for r in rows if 0 <= r < n]
        rest = [i for i in range(n) if i not in rows]
        pos = sum(1 for i in rest if i < slot)
        order = rest[:pos] + rows + rest[pos:]
        if not rows or order == list(range(n)):
            return
        self.snapshot()
        self.doc.select(order)
        self.refresh_thumbs(pos)
        for r in range(pos, pos + len(rows)):
            it = self.thumbs.item(r)
            if it:
                it.setSelected(True)

    # ---------- right-click menu ----------
    def _reselect(self, rows):
        self.thumbs.blockSignals(True)
        for r in rows:
            it = self.thumbs.item(r)
            if it:
                it.setSelected(True)
        self.thumbs.blockSignals(False)

    def delete_rows(self, rows):
        if not self.doc or not rows:
            return
        if len(rows) >= len(self.doc):
            return QMessageBox.information(self.main, "PDF Studio", "Can't delete every page.")
        self.snapshot()
        self.doc.delete_pages(rows)
        self.refresh_thumbs(min(rows[0], len(self.doc) - 1))

    def rotate_rows(self, rows, delta):
        self.snapshot()
        for r in rows:
            pg = self.doc[r]
            pg.set_rotation((pg.rotation + delta) % 360)
        self.refresh_thumbs(rows[0])
        self._reselect(rows)

    def duplicate_rows(self, rows):
        self.snapshot()
        for r in sorted(rows, reverse=True):
            self.doc.copy_page(r, r + 1 if r + 1 < len(self.doc) else -1)
        self.refresh_thumbs(rows[0] + 1)

    def blank_after(self, row):
        self.snapshot()
        r = self.doc[row].rect
        self.doc.new_page(pno=row + 1, width=r.width, height=r.height)
        self.refresh_thumbs(row + 1)

    def insert_pdf_after(self, row):
        p, _ = QFileDialog.getOpenFileName(self.main, "Insert PDF here", "", "PDF (*.pdf)")
        if p:
            self._insert([p], row + 1)

    def extract_rows(self, rows):
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "pages"
        start = os.path.join(os.path.dirname(self.path), f"{base}_extract.pdf") if self.path else f"{base}_extract.pdf"
        p, _ = QFileDialog.getSaveFileName(self.main, "Save selected pages as new PDF", start, "PDF (*.pdf)")
        if not p:
            return
        try:
            out = fitz.open()
            for r in rows:
                out.insert_pdf(self.doc, from_page=r, to_page=r)
            out.save(p, garbage=3, deflate=True)
            out.close()
            self.main.statusBar().showMessage(f"Saved {len(rows)} page(s) to {p}", 6000)
        except Exception as e:
            QMessageBox.warning(self.main, "Save failed", str(e))

    def arm_page_export(self, index, file_format, indexes=None):
        """Choose the export format for the next drag of the selected page(s)."""
        if not self.doc or not (0 <= index < len(self.doc)):
            return
        indexes = sorted({int(i) for i in (indexes or [index]) if 0 <= int(i) < len(self.doc)})
        if index not in indexes:
            indexes.append(index)
            indexes.sort()
        self._clear_armed_export()
        self._armed_export_format = (indexes, file_format)
        self.thumbs.export_drag_index = index
        what = "page" if len(indexes) == 1 else f"{len(indexes)} selected pages"
        self.main.statusBar().showMessage(
            f"Now drag the {what} to the Desktop to export as {file_format.upper()}.", 8000)
        for row in indexes:
            self.labels[row].setToolTip(
                f"Drag the selected page(s) to the Desktop now to export as {file_format.upper()}")

    def _clear_armed_export(self):
        pending = self._armed_export_format
        if pending:
            for idx in pending[0]:
                if 0 <= idx < len(self.labels):
                    self.labels[idx].setToolTip(
                        f"Drag page {idx + 1} to the Desktop to create a separate PDF; "
                        "right-click to choose PDF or JPG. Files remain available for 7 days.")
        self._armed_export_format = None
        self.thumbs.export_drag_index = None

    def export_page_drag(self, index, source=None):
        """Drag one unmodified page out as a PDF or JPG file."""
        pending = self._armed_export_format
        rows = pending[0] if pending and index in pending[0] else [index]
        self.export_pages_drag(rows, source=source)

    def export_pages_drag(self, indexes, source=None):
        """Drag selected pages as real files; keep them available for external drop targets."""
        if not self.doc:
            return
        indexes = sorted({int(i) for i in indexes if 0 <= int(i) < len(self.doc)})
        if not indexes:
            return
        pending = self._armed_export_format
        file_format = pending[1] if pending and indexes == pending[0] else "pdf"
        self._clear_armed_export()
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "PDF_Studio"
        base = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base).strip(" .")[:80] or "PDF_Studio"
        try:
            os.makedirs(_page_export_root(), exist_ok=True)
            temp_dir = tempfile.mkdtemp(prefix="drag_", dir=_page_export_root())
        except Exception as e:
            QMessageBox.warning(self.main, "Export page failed", str(e))
            return
        extension = "jpg" if file_format == "jpg" else "pdf"
        paths = []
        try:
            for index in indexes:
                temp_path = os.path.join(temp_dir, f"{base}_page_{index + 1}.{extension}")
                if file_format == "jpg":
                    rect = self.doc[index].rect
                    scale = min(200 / 72, 6000 / max(rect.width, rect.height))
                    pix = self.doc[index].get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                    pix.save(temp_path)
                else:
                    out = fitz.open()
                    try:
                        out.insert_pdf(self.doc, from_page=index, to_page=index)
                        out.save(temp_path, garbage=3, deflate=True)
                    finally:
                        out.close()
                paths.append(temp_path)

            mime = QMimeData()
            mime.setUrls([QUrl.fromLocalFile(path) for path in paths])
            mime.setData("application/x-pdfstudio-page-export", b"1")
            if source is self.thumbs:
                mime.setData("application/x-pdfstudio-page-reorder",
                             ",".join(str(i) for i in indexes).encode("ascii"))
            drag = QDrag(source or self.labels[indexes[0]])
            drag.setMimeData(mime)
            image = self.imgs.get(indexes[0])
            if image is not None:
                drag.setPixmap(QPixmap.fromImage(image).scaled(
                    QSize(160, 220), Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation))
                drag.setHotSpot(QPoint(16, 16))
            actions = Qt.DropAction.CopyAction
            default_action = Qt.DropAction.CopyAction
            if source is self.thumbs:
                actions |= Qt.DropAction.MoveAction
            result = drag.exec(actions, default_action)
            if result != Qt.DropAction.IgnoreAction:
                names = ", ".join(os.path.basename(path) for path in paths)
                self.main.statusBar().showMessage(
                    f"Drag accepted: {names}. Export files are kept for 7 days.", 7000)
        except Exception as e:
            # Failed exports are not useful to a drop target, so clean them immediately.
            shutil.rmtree(temp_dir, ignore_errors=True)
            QMessageBox.warning(self.main, "Export page failed", str(e))

    def page_menu(self, i, gpos):
        """Right-click menu for page `i` (works on the page itself and on its thumbnail)."""
        if not self.doc or not (0 <= i < len(self.doc)):
            return
        sel = sorted({self.thumbs.row(x) for x in self.thumbs.selectedItems()})
        if i not in sel:                                    # right-click on an unselected page -> act on that page
            sel = [i]
            self.thumbs.blockSignals(True)
            self.thumbs.clearSelection()
            self.thumbs.setCurrentRow(i)
            item = self.thumbs.item(i)
            if item is not None:
                item.setSelected(True)
            self.thumbs.blockSignals(False)
            self.cur = i
            self.place_del_buttons()
            self.changed.emit()
        n, total = len(sel), len(self.doc)
        what = f"page {sel[0] + 1}" if n == 1 else f"{n} pages"
        m = QMenu(self)
        m.addAction(f"🖨 Print {what}…", lambda: self.main.print_doc(sel))
        if n < total:
            m.addAction("🖨 Print whole document…", lambda: self.main.print_doc())
        m.addSeparator()
        export_pdf = m.addAction(f"↗ Drag {what} to Desktop as PDF",
                                 lambda page=i, rows=tuple(sel): self.arm_page_export(page, "pdf", rows))
        export_jpg = m.addAction(f"↗ Drag {what} to Desktop as JPG",
                                 lambda page=i, rows=tuple(sel): self.arm_page_export(page, "jpg", rows))
        export_pdf.setEnabled(not self.edit_mode)
        export_jpg.setEnabled(not self.edit_mode)
        m.addSeparator()
        m.addAction("⟳ Rotate right", lambda: self.rotate_rows(sel, 90))
        m.addAction("⟲ Rotate left", lambda: self.rotate_rows(sel, -90))
        m.addAction("⧉ Duplicate", lambda: self.duplicate_rows(sel))
        m.addSeparator()
        mv = m.addMenu("Move")
        a = mv.addAction("↑ Up", lambda: self.move_pages(sel, sel[0] - 1))
        a.setEnabled(sel[0] > 0)
        a = mv.addAction("↓ Down", lambda: self.move_pages(sel, sel[-1] + 2))
        a.setEnabled(sel[-1] < total - 1)
        a = mv.addAction("⤒ To the beginning", lambda: self.move_pages(sel, 0))
        a.setEnabled(sel[0] > 0)
        a = mv.addAction("⤓ To the end", lambda: self.move_pages(sel, total))
        a.setEnabled(sel[-1] < total - 1)
        m.addAction("＋ Add blank page after", lambda: self.blank_after(sel[-1]))
        m.addAction("⇪ Insert PDF after…", lambda: self.insert_pdf_after(sel[-1]))
        m.addAction(f"📤 Save {what} as new PDF…", lambda: self.extract_rows(sel))
        m.addSeparator()
        a = m.addAction(f"🗑 Delete {what}", lambda: self.delete_rows(sel))
        a.setEnabled(n < total)
        a = m.addAction("↶ Undo", self.undo)
        a.setEnabled(bool(self.undo_stack))
        m.exec(gpos)

    def rotate(self):
        if not self.doc:
            return
        self.snapshot()
        p = self.doc[self.cur]
        p.set_rotation((p.rotation + 90) % 360)
        self.refresh_thumbs()

    # ---------- object editing (Edit mode) ----------
    HANDLE_PX = 8

    def _reset_edit_state(self):
        self._objs = {}
        self.edit_sel, self.edit_page = [], None
        self._drag = self._ghost = self._hover = None
        if self._editor is not None:
            ed, self._editor = self._editor, None
            ed._done = True
            ed.hide()
            ed.deleteLater()

    def set_edit(self, on):
        self.edit_mode = on
        if on:
            self._clear_armed_export()
        else:
            self._commit_editor()
            self.edit_sel, self.edit_page, self._drag, self._ghost, self._hover = [], None, None, None, None
        for lab in self.labels:
            lab.setCursor(Qt.CursorShape.ArrowCursor)
            lab.export_drag_enabled = not on
            lab.setFocusPolicy(Qt.FocusPolicy.StrongFocus if on else Qt.FocusPolicy.NoFocus)
            lab.update()

    def on_click(self, idx, pos):                           # old click-to-edit-a-line path; Edit mode now handles input itself
        return

    # -- finding the objects on a page --
    def page_objects(self, idx):
        if idx in self._objs:
            return self._objs[idx]
        page = self.doc[idx]
        objs = []
        try:
            for b in page.get_text("dict")["blocks"]:
                if b.get("type") != 0:
                    continue
                lines = [ln for ln in b["lines"] if "".join(sp["text"] for sp in ln["spans"]).strip()]
                if not lines:
                    continue
                r = fitz.Rect(lines[0]["bbox"])
                for ln in lines[1:]:
                    r |= fitz.Rect(ln["bbox"])
                objs.append({"kind": "text", "rect": r, "lines": lines})
        except Exception:
            pass
        try:
            for im in page.get_image_info(xrefs=True):
                if im.get("xref") and not (im.get("width", 2) <= 1 and im.get("height", 2) <= 1):   # 1x1 = a deleted image
                    objs.append({"kind": "image", "rect": fitz.Rect(im["bbox"]), "xref": im["xref"]})
        except Exception:
            pass
        try:
            paths = [p for p in page.get_drawings() if p.get("type") in ("f", "s", "fs") and p.get("items")]
            if paths:
                clusters = page.cluster_drawings(drawings=paths)
                pr = page.rect
                for c in clusters:
                    if c.width * c.height >= 0.85 * pr.width * pr.height:
                        continue                            # the page background, not an object
                    mine = [p for p in paths if (fitz.Rect(c) + (-1, -1, 1, 1)).contains(p["rect"])]
                    if mine:
                        objs.append({"kind": "draw", "rect": fitz.Rect(c), "paths": mine})
        except Exception:
            pass
        self._objs[idx] = objs
        return objs

    def _hit(self, idx, pt):
        tol = 3.0 / self.zoom
        best, best_area = None, None
        for o in self.page_objects(idx):
            if (o["rect"] + (-tol, -tol, tol, tol)).contains(pt):
                area = max(1.0, o["rect"].width * o["rect"].height) * (0.5 if o["kind"] == "text" else 1.0)
                if best is None or area < best_area:
                    best, best_area = o, area
        return best

    def _sel_rect(self):
        if not self.edit_sel:
            return None
        r = fitz.Rect(self.edit_sel[0]["rect"])
        for o in self.edit_sel[1:]:
            r |= o["rect"]
        return r

    def _handle_points(self, r):
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        return {"nw": (r.x0, r.y0), "n": (cx, r.y0), "ne": (r.x1, r.y0), "e": (r.x1, cy),
                "se": (r.x1, r.y1), "s": (cx, r.y1), "sw": (r.x0, r.y1), "w": (r.x0, cy)}

    def _handle_at(self, idx, pos):
        r = self._sel_rect()
        if r is None or idx != self.edit_page:
            return None
        for name, (hx, hy) in self._handle_points(r).items():
            if abs(pos.x() - hx * self.zoom) <= self.HANDLE_PX and abs(pos.y() - hy * self.zoom) <= self.HANDLE_PX:
                return name
        return None

    _HANDLE_CURSORS = {"nw": Qt.CursorShape.SizeFDiagCursor, "se": Qt.CursorShape.SizeFDiagCursor,
                       "ne": Qt.CursorShape.SizeBDiagCursor, "sw": Qt.CursorShape.SizeBDiagCursor,
                       "n": Qt.CursorShape.SizeVerCursor, "s": Qt.CursorShape.SizeVerCursor,
                       "e": Qt.CursorShape.SizeHorCursor, "w": Qt.CursorShape.SizeHorCursor}

    def _page_ok(self, idx):
        if self.doc[idx].rotation:
            QMessageBox.information(self.main, "PDF Studio", "Set the page rotation back to 0° to edit this page.")
            return False
        return True

    # -- mouse --
    def edit_press(self, lab, pos, mods):
        idx = lab.index
        self._commit_editor()
        if not self._page_ok(idx):
            return
        pt = fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom)
        h = self._handle_at(idx, pos)
        if h:
            r = self._sel_rect()
            hp = self._handle_points(r)
            opp = {"nw": "se", "n": "s", "ne": "sw", "e": "w", "se": "nw", "s": "n", "sw": "ne", "w": "e"}[h]
            ax, ay = hp[opp]
            self._drag = {"mode": "resize", "handle": h, "anchor": (ax, ay), "hpos": hp[h], "rect": fitz.Rect(r), "idx": idx}
            self._ghost = (1.0, 1.0, 0.0, 0.0)
            return
        o = self._hit(idx, pt)
        shift = bool(mods & Qt.KeyboardModifier.ShiftModifier)
        if o is not None:
            if idx != self.edit_page:
                self.edit_sel = []
            self.edit_page = idx
            if shift:
                self.edit_sel = [x for x in self.edit_sel if x is not o] if o in self.edit_sel else self.edit_sel + [o]
            elif o not in self.edit_sel:
                self.edit_sel = [o]
            self._drag = {"mode": "move", "start": pos, "idx": idx, "moved": False}
            self._ghost = (1.0, 1.0, 0.0, 0.0)
        else:
            if not shift:
                self.edit_sel = []
            self.edit_page = idx
            self._drag = {"mode": "band", "start": pos, "idx": idx}
            self._band = QRectF(pos, pos)
        for l in self.labels:
            l.update()

    def edit_move(self, lab, pos, buttons, mods):
        idx = lab.index
        d = self._drag
        if d is None or not (buttons & Qt.MouseButton.LeftButton) or d["idx"] != idx:
            if d is None:                                   # hover feedback
                h = self._handle_at(idx, pos)
                if h:
                    lab.setCursor(self._HANDLE_CURSORS[h])
                    return
                o = self._hit(idx, fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom))
                if o is not self._hover:
                    self._hover = o
                    lab.update()
                if o is not None:
                    lab.setCursor(Qt.CursorShape.SizeAllCursor if o in self.edit_sel
                                  else Qt.CursorShape.PointingHandCursor)
                else:
                    lab.setCursor(Qt.CursorShape.ArrowCursor)
            return
        if d["mode"] == "move":
            dx, dy = (pos.x() - d["start"].x()) / self.zoom, (pos.y() - d["start"].y()) / self.zoom
            if not d["moved"] and math.hypot(dx, dy) * self.zoom < 3:
                return
            d["moved"] = True
            if mods & Qt.KeyboardModifier.ShiftModifier:    # Shift: keep to one axis
                if abs(dx) > abs(dy):
                    dy = 0
                else:
                    dx = 0
            self._ghost = (1.0, 1.0, dx, dy)
        elif d["mode"] == "resize":
            h, r = d["handle"], d["rect"]
            ax, ay = d["anchor"]
            mx, my = pos.x() / self.zoom, pos.y() / self.zoom
            w0, h0 = max(1e-3, r.width), max(1e-3, r.height)
            has_text = any(o["kind"] == "text" for o in self.edit_sel)
            hx, hy = d["hpos"]
            sx = (mx - ax) / (hx - ax) if ("e" in h or "w" in h) and abs(hx - ax) > 1e-6 else 1.0
            sy = (my - ay) / (hy - ay) if ("n" in h or "s" in h) and abs(hy - ay) > 1e-6 else 1.0
            if len(h) == 2:                                 # corners keep the proportions
                sx = sy = max(0.05, max(sx, sy))
            elif has_text:                                  # text only scales evenly
                sx = sy = max(0.05, sx if h in ("e", "w") else sy)
            else:
                sx, sy = max(0.05, sx), max(0.05, sy)
            self._ghost = (sx, sy, 0.0, 0.0)
        elif d["mode"] == "band":
            self._band = QRectF(d["start"], pos).normalized()
        lab.update()

    def edit_release(self, lab, pos, mods):
        d, self._drag = self._drag, None
        g, self._ghost = self._ghost, None
        if d is None:
            return
        idx = d["idx"]
        if d["mode"] == "band":
            b = self._band
            self._band = None
            if b.width() > 3 or b.height() > 3:
                rect = fitz.Rect(b.left() / self.zoom, b.top() / self.zoom, b.right() / self.zoom, b.bottom() / self.zoom)
                add = [o for o in self.page_objects(idx) if rect.contains(o["rect"]) and o not in self.edit_sel]
                self.edit_sel = self.edit_sel + add if (mods & Qt.KeyboardModifier.ShiftModifier) else add
        elif g is not None and (abs(g[0] - 1) > 1e-4 or abs(g[1] - 1) > 1e-4 or abs(g[2]) > 1e-4 or abs(g[3]) > 1e-4):
            if d["mode"] == "resize":
                self._apply_transform(idx, d["anchor"], g[0], g[1], 0.0, 0.0)
            else:
                self._apply_transform(idx, (0.0, 0.0), 1.0, 1.0, g[2], g[3])
        lab.update()

    def edit_leave(self, lab):
        if self._hover is not None:
            self._hover = None
            lab.update()

    def edit_double_click(self, lab, pos):
        idx = lab.index
        if not self._page_ok(idx):
            return
        pt = fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom)
        o = self._hit(idx, pt)
        if o is not None and o["kind"] == "text":
            self.edit_page, self.edit_sel = idx, [o]
            self._open_editor(lab, o)
        elif o is None:
            self.edit_page, self.edit_sel = idx, []
            self._open_editor(lab, None, pt)
        lab.update()

    def edit_key(self, lab, e):
        k = e.key()
        if self._editor is not None:
            return False
        if k == Qt.Key.Key_Escape:
            self.edit_sel = []
            lab.update()
            return True
        if not self.edit_sel or lab.index != self.edit_page:
            return False
        if k in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self._apply_delete(lab.index)
            return True
        if k in (Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_F2):
            o = self.edit_sel[0]
            if len(self.edit_sel) == 1 and o["kind"] == "text":
                self._open_editor(lab, o)
            return True
        step = 10.0 if e.modifiers() & Qt.KeyboardModifier.ShiftModifier else 1.0
        dxy = {Qt.Key.Key_Left: (-step, 0), Qt.Key.Key_Right: (step, 0),
               Qt.Key.Key_Up: (0, -step), Qt.Key.Key_Down: (0, step)}.get(k)
        if dxy:
            self._apply_transform(lab.index, (0.0, 0.0), 1.0, 1.0, dxy[0], dxy[1])
            return True
        return False

    # -- drawing the selection --
    def _tx_rect(self, r, g, anchor):
        sx, sy, dx, dy = g
        ax, ay = anchor
        return fitz.Rect(ax + (r.x0 - ax) * sx + dx, ay + (r.y0 - ay) * sy + dy,
                         ax + (r.x1 - ax) * sx + dx, ay + (r.y1 - ay) * sy + dy)

    def paint_overlay(self, lab, p):
        idx, z = lab.index, self.zoom
        accent = QColor(ACCENT)
        if self._hover is not None and self._hover in self.page_objects(idx) and self._hover not in self.edit_sel and self._drag is None:
            r = self._hover["rect"]
            p.setPen(QPen(QColor(accent.red(), accent.green(), accent.blue(), 150), 1))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(QRectF(r.x0 * z, r.y0 * z, r.width * z, r.height * z))
        if idx == self.edit_page and self._band is not None and self._drag and self._drag["mode"] == "band":
            p.setPen(QPen(accent, 1, Qt.PenStyle.DashLine))
            p.setBrush(QColor(accent.red(), accent.green(), accent.blue(), 30))
            p.drawRect(self._band)
        if idx != self.edit_page or not self.edit_sel:
            return
        g = self._ghost or (1.0, 1.0, 0.0, 0.0)
        anchor = self._drag["anchor"] if (self._drag and self._drag["mode"] == "resize") else (0.0, 0.0)
        img = self.imgs.get(idx)
        moving = self._ghost is not None and g != (1.0, 1.0, 0.0, 0.0)
        for o in self.edit_sel:
            r0 = o["rect"]
            r = self._tx_rect(r0, g, anchor)
            tgt = QRectF(r.x0 * z, r.y0 * z, r.width * z, r.height * z)
            if moving and img is not None:                  # live ghost of the content being moved / scaled
                p.setOpacity(0.8)
                p.drawImage(tgt, img, QRectF(r0.x0 * z, r0.y0 * z, r0.width * z, r0.height * z))
                p.setOpacity(1.0)
            p.setPen(QPen(accent, 1, Qt.PenStyle.DashLine))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(tgt)
        r = self._tx_rect(self._sel_rect(), g, anchor)
        p.setPen(QPen(accent, 1.4))
        p.setBrush(QColor("white"))
        hs = self.HANDLE_PX - 2
        for hx, hy in self._handle_points(r).values():
            p.drawRect(QRectF(hx * z - hs / 2, hy * z - hs / 2, hs, hs))

    _band = None

    # -- inline text editor --
    def _open_editor(self, lab, obj, pt=None):
        self._commit_editor()
        z = self.zoom
        if obj is not None:
            lines = obj["lines"]
            text = "\n".join("".join(sp["text"] for sp in ln["spans"]) for ln in lines)
            size = lines[0]["spans"][0]["size"]
            r = obj["rect"]
            x, y, w, h = r.x0 * z - 2, r.y0 * z - 2, r.width * z + 60, r.height * z + 6
        else:
            text, size = "", 12.0
            x, y, w, h = pt.x * z, (pt.y - size) * z, 220, size * z * 1.8
        ed = InlineTextEditor(lab, text, size * z, w, h)
        ed.move(int(x), int(y))
        ed.committed.connect(lambda t, o=obj, p0=pt, i=lab.index: self._editor_done(i, o, p0, t))
        ed.cancelled.connect(lambda: self._editor_closed())
        self._editor = ed
        ed.show()
        ed.setFocus()
        if text:
            ed.selectAll()

    def _commit_editor(self):
        ed = self._editor
        if ed is not None:
            ed._finish(True)

    def _editor_closed(self):
        ed, self._editor = self._editor, None
        if ed is not None:
            ed.hide()
            ed.deleteLater()

    def _editor_done(self, idx, obj, pt, text):
        self._editor_closed()
        try:
            if obj is None:
                if text.strip():
                    self._apply_new_text(idx, pt, text)
            else:
                old = "\n".join("".join(sp["text"] for sp in ln["spans"]) for ln in obj["lines"])
                if text != old:
                    self._apply_text_edit(idx, obj, text)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))

    # -- applying changes to the PDF --
    def _finish_edit(self, idx, select_rects=None):
        """Redraw the changed page + its thumbnail (keeps the scroll position) and re-select what was edited."""
        self._objs.pop(idx, None)
        self.render_page(idx, force=True)
        self.thumbs.item(idx).setIcon(self.thumb_icon(idx))
        self.edit_page = idx
        self.edit_sel = []
        if select_rects:
            for o in self.page_objects(idx):
                for rr in select_rects:
                    if (abs(o["rect"].x0 - rr.x0) < 4 and abs(o["rect"].y0 - rr.y0) < 4
                            and abs(o["rect"].x1 - rr.x1) < 6 and abs(o["rect"].y1 - rr.y1) < 6):
                        self.edit_sel.append(o)
                        break
        self.labels[idx].update()

    def _remove_objects(self, page, objs):
        """Take the objects off the page (text, images and vector shapes each in their own pass)."""
        texts = [o for o in objs if o["kind"] == "text"]
        for o in texts:
            page.add_redact_annot(o["rect"] + (0.3, 0.3, -0.3, -0.3), fill=False)
        if texts:
            _apply_redactions(page, text=getattr(fitz, "PDF_REDACT_TEXT_REMOVE", 0))
        for o in objs:
            if o["kind"] == "image":
                try:
                    page.delete_image(o["xref"])
                except Exception:
                    page.add_redact_annot(o["rect"], fill=False)
                    _apply_redactions(page, images=getattr(fitz, "PDF_REDACT_IMAGE_REMOVE", 1))
        draws = [o for o in objs if o["kind"] == "draw"]
        for o in draws:
            page.add_redact_annot(o["rect"] + (0.4, 0.4, -0.4, -0.4), fill=False)
        if draws:                                           # "if touched": curves (circles ...) are not removed by "if covered"
            _apply_redactions(page, graphics=getattr(fitz, "PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED", 2))

    def _image_payload(self, o):
        info = self.doc.extract_image(o["xref"])
        if not info:
            return None
        mask = None
        if info.get("smask"):
            try:
                mask = fitz.Pixmap(self.doc, info["smask"]).tobytes("png")
            except Exception:
                mask = None
        return info["image"], mask

    def _apply_transform(self, idx, anchor, sx, sy, dx, dy):
        objs = list(self.edit_sel)
        if not objs:
            return
        ax, ay = anchor
        T = lambda p: fitz.Point(ax + (p.x - ax) * sx + dx, ay + (p.y - ay) * sy + dy)
        if abs(sx - 1) < 1e-6:
            s_text = sy
        elif abs(sy - 1) < 1e-6:
            s_text = sx
        else:
            s_text = (sx + sy) / 2
        avg = (abs(sx) + abs(sy)) / 2
        page = self.doc[idx]
        self.snapshot()
        new_rects = [self._tx_rect(o["rect"], (sx, sy, dx, dy), anchor) for o in objs]
        try:
            payloads = {o["xref"]: self._image_payload(o) for o in objs if o["kind"] == "image"}
            self._remove_objects(page, objs)
            for o, nr in zip(objs, new_rects):
                if o["kind"] == "text":
                    for ln in o["lines"]:
                        for sp in ln["spans"]:
                            if not sp["text"].strip():
                                continue
                            bb = fitz.Rect(sp["bbox"])
                            nb = fitz.Rect(T(bb.tl), T(bb.br))
                            _emit_text(page, sp["text"], T(fitz.Point(sp["origin"])), sp["size"] * s_text,
                                       sp["flags"], _rgb_from_int(sp["color"]), nb)
                elif o["kind"] == "image":
                    pl = payloads.get(o["xref"])
                    if pl:
                        page.insert_image(nr, stream=pl[0], mask=pl[1], keep_proportion=False)
                else:
                    for path in o["paths"]:
                        _redraw_path(page, path, T, avg)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))
        self._finish_edit(idx, new_rects)

    def _apply_delete(self, idx):
        objs = list(self.edit_sel)
        if not objs:
            return
        self.snapshot()
        try:
            self._remove_objects(self.doc[idx], objs)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))
        self._finish_edit(idx)

    def _apply_text_edit(self, idx, obj, text):
        page = self.doc[idx]
        self.snapshot()
        lines = obj["lines"]
        sp0 = lines[0]["spans"][0]
        size, flags, color = sp0["size"], sp0["flags"], _rgb_from_int(sp0["color"])
        ox, oy = sp0["origin"]
        step = (lines[1]["spans"][0]["origin"][1] - oy) if len(lines) > 1 else size * 1.2
        if step <= 0:
            step = size * 1.2
        r = obj["rect"]
        self._remove_objects(page, [obj])
        new_rects = []
        for i, t in enumerate(text.split("\n")):
            if t.strip():
                y = oy + i * step
                _emit_text(page, t, fitz.Point(ox, y), size, flags, color,
                           fitz.Rect(r.x0, y - size, r.x1, y + size * 0.4))
        self._finish_edit(idx)

    def _apply_new_text(self, idx, pt, text):
        page = self.doc[idx]
        self.snapshot()
        size = 12.0
        for i, t in enumerate(text.split("\n")):
            if t.strip():
                y = pt.y + i * size * 1.2
                _emit_text(page, t, fitz.Point(pt.x, y), size, 0, (0, 0, 0))
        self._finish_edit(idx)


# ============================================================================
#  MAIN WINDOW (browser-style tabs)
# ============================================================================
class TabBar(QTabBar):
    def tabInserted(self, i):
        super().tabInserted(i)
        b = QToolButton(self)                               # small close button (replaces the default one)
        b.setObjectName("TabClose")
        b.setText("✕")
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        b.clicked.connect(lambda _=False, btn=b: self._close_button_clicked(btn))
        self.setTabButton(i, QTabBar.ButtonPosition.RightSide, b)

    def _close_button_clicked(self, btn):
        for j in range(self.count()):
            if self.tabButton(j, QTabBar.ButtonPosition.RightSide) is btn:
                self.tabCloseRequested.emit(j)
                return

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.MouseButton.MiddleButton:       # middle click closes a tab
            i = self.tabAt(e.position().toPoint())
            if i >= 0:
                self.tabCloseRequested.emit(i)
                return
        super().mouseReleaseEvent(e)


class Tabs(QObject):
    """Tab strip (`bar`) + stacked pages (`pages`) with the same API the window used on QTabWidget.
    They are separate widgets so the window can put the toolbar BETWEEN them, like Acrobat does."""
    tabCloseRequested = pyqtSignal(int)
    currentChanged = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.bar = TabBar()
        self.bar.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.bar.setTabsClosable(False)                      # TabBar adds its own close buttons
        self.bar.setMovable(True)
        self.bar.setExpanding(False)
        self.bar.setDocumentMode(True)
        self.bar.setElideMode(Qt.TextElideMode.ElideRight)
        self.bar.setUsesScrollButtons(True)
        self.bar.hide()                                      # shown while at least one PDF is open
        self.pages = QStackedWidget()
        self._w = []                                         # page widgets, in tab order
        self.bar.currentChanged.connect(self._on_current)
        self.bar.tabCloseRequested.connect(self.tabCloseRequested)
        self.bar.tabMoved.connect(lambda a, b: self._w.insert(b, self._w.pop(a)))

    def _on_current(self, i):
        if 0 <= i < len(self._w):
            self.pages.setCurrentWidget(self._w[i])
        self.currentChanged.emit(i)

    def count(self):
        return len(self._w)

    def widget(self, i):
        return self._w[i] if 0 <= i < len(self._w) else None

    def indexOf(self, w):
        return self._w.index(w) if w in self._w else -1

    def currentIndex(self):
        return self.bar.currentIndex()

    def currentWidget(self):
        return self.widget(self.bar.currentIndex())

    def setCurrentIndex(self, i):
        self.bar.setCurrentIndex(i)

    def setCurrentWidget(self, w):
        i = self.indexOf(w)
        if i >= 0:
            self.bar.setCurrentIndex(i)

    def addTab(self, w, text):
        self._w.append(w)
        self.pages.addWidget(w)
        i = self.bar.addTab(text)
        self.bar.show()
        return i

    def removeTab(self, i):
        if not 0 <= i < len(self._w):
            return
        w = self._w.pop(i)
        self.bar.removeTab(i)
        self.pages.removeWidget(w)
        if not self._w:
            self.bar.hide()

    def setTabText(self, i, text):
        self.bar.setTabText(i, text)

    def setTabToolTip(self, i, text):
        self.bar.setTabToolTip(i, text)


class Main(DarkTitleBarMixin, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setObjectName("AppWindow")
        self.setWindowIcon(app_icon())
        self.server = None
        self._print_dialog_preferences = None    # session-only; cleared when this app process exits
        self._settings = QSettings("PDFStudio", "PDFStudio")     # remembers window size/position between runs
        self.resize(1200, 850)
        g = self._settings.value("window/geometry")
        if g is not None:
            try:
                self.restoreGeometry(g)
            except Exception:
                pass

        self.tabs = Tabs()
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(lambda _: self.update_status())
        self.hint = EmptyState(self.open_dialog)
        self.stack = QStackedWidget()
        self.stack.addWidget(self.hint)
        self.stack.addWidget(self.tabs.pages)
        central = QWidget()                                  # tab strip / toolbar / document area, top to bottom
        self._central_lay = QVBoxLayout(central)
        self._central_lay.setContentsMargins(0, 0, 0, 0)
        self._central_lay.setSpacing(0)
        self._central_lay.addWidget(self.tabs.bar)
        self._central_lay.addWidget(self.stack, 1)
        self.setCentralWidget(central)

        # status bar: zoom + "Page x / n" at the bottom right
        self.zoom_lbl = QLabel()
        self.zoom_lbl.setObjectName("StatusPill")
        self.statusBar().addPermanentWidget(self.zoom_lbl)
        self.statusBar().setMinimumHeight(38)
        # page counter in the middle of the status bar: "Page [ 3 ] / 8"; type a number + Enter to jump there
        self.page_lbl = QWidget(self.statusBar())
        self.page_lbl.setObjectName("PageBar")
        _pl = QHBoxLayout(self.page_lbl)
        _pl.setContentsMargins(10, 0, 10, 0)
        _pl.setSpacing(8)
        _t1, self.page_total = QLabel("Page"), QLabel("/ 0")
        for _l in (_t1, self.page_total):
            _l.setObjectName("PageText")
        self.page_box = QLineEdit()
        self.page_box.setObjectName("PageBox")
        self.page_box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_box.setFixedWidth(68)
        self.page_box.setValidator(QIntValidator(1, 999999, self))
        self.page_box.setToolTip("Type a page number and press Enter to go to that page")
        self.page_box.returnPressed.connect(self.page_box_go)
        self.page_box.installEventFilter(self)
        _pl.addWidget(_t1)
        _pl.addWidget(self.page_box)
        _pl.addWidget(self.page_total)
        self.statusBar().installEventFilter(self)
        self.page_lbl.hide()

        self._build_actions()
        self._central_lay.insertWidget(1, self.toolbar)       # right under the tab strip
        self.setAcceptDrops(True)
        self.update_status()

    # ---------- tabs ----------
    def cur_tab(self):
        w = self.tabs.currentWidget()
        return w if isinstance(w, DocTab) else None

    @staticmethod
    def _is_pdf_file(p):
        """A .pdf name, or any file that starts with the PDF signature (Explorer 'Open with' can hand over
        files whose name has no .pdf, e.g. downloads)."""
        if p.lower().endswith(".pdf"):
            return True
        try:
            with open(p, "rb") as f:
                return f.read(1024).lstrip(b"\xef\xbb\xbf \r\n\t\x00").startswith(b"%PDF-")
        except OSError:
            return False

    def open_paths(self, paths):
        """Open every PDF in `paths`. A path that cannot be opened is reported with its exact text instead of
        being skipped silently, so a failing 'Open with' shows what Windows actually handed over."""
        problems = []
        for p in paths:
            if not os.path.exists(p):
                problems.append(f"File not found:\n{p!r}")
            elif not os.path.isfile(p) or not self._is_pdf_file(p):
                problems.append(f"Not a PDF file:\n{p!r}")
            else:
                self.open_path(p)
        if problems:
            QMessageBox.warning(self, "PDF Studio", "\n\n".join(problems))

    def open_path(self, p):
        ap = os.path.normcase(os.path.abspath(p))
        for i in range(self.tabs.count()):                  # already open -> just switch to it
            t = self.tabs.widget(i)
            if t.path and os.path.normcase(os.path.abspath(t.path)) == ap:
                self.tabs.setCurrentIndex(i)
                return
        tab = DocTab(self)
        if not tab.load(p):
            tab.deleteLater()
            return
        tab.set_edit(self.a_edit.isChecked())
        tab.changed.connect(lambda t=tab: self.on_tab_changed(t))
        i = self.tabs.addTab(tab, "")                       # new tabs go to the end, in order
        self.stack.setCurrentIndex(1)
        self.tabs.setCurrentIndex(i)
        self.on_tab_changed(tab)

    def on_tab_changed(self, tab):
        i = self.tabs.indexOf(tab)
        if i < 0:
            return
        name = os.path.basename(tab.path) if tab.path else "Untitled"
        self.tabs.setTabText(i, ("● " if tab.dirty else "") + name)
        self.tabs.setTabToolTip(i, tab.path or name)
        if tab is self.cur_tab():
            self.update_status()

    def place_page_bar(self):
        """Keep the page counter centred in the status bar."""
        sb = self.statusBar()
        self.page_lbl.adjustSize()
        w, h = self.page_lbl.sizeHint().width(), sb.height()
        self.page_lbl.setGeometry((sb.width() - w) // 2, 0, w, h)
        self.page_lbl.raise_()

    def page_box_go(self):
        t = self.cur_tab()
        if not (t and t.doc):
            return
        txt = self.page_box.text().strip()
        if txt.isdigit():
            t.goto_page(max(1, min(int(txt), len(t.doc))) - 1)
        t.thumbs.blockSignals(True)                          # keep the thumbnail highlight on the page we jumped to
        t.thumbs.setCurrentRow(t.cur)
        t.thumbs.blockSignals(False)
        t.place_del_buttons()
        self.page_box.setText(str(t.cur + 1))
        self.page_box.clearFocus()
        t.scroll.setFocus()

    def eventFilter(self, obj, ev):
        if obj is self.statusBar() and ev.type() in (QEvent.Type.Resize, QEvent.Type.Show):
            if self.page_lbl.isVisible():
                self.place_page_bar()
        elif obj is self.page_box:
            if ev.type() == QEvent.Type.FocusIn:
                QTimer.singleShot(0, self.page_box.selectAll)
            elif ev.type() == QEvent.Type.FocusOut:
                t = self.cur_tab()
                if t and t.doc:
                    self.page_box.setText(str(t.cur + 1))
            elif ev.type() == QEvent.Type.KeyPress and ev.key() == Qt.Key.Key_Escape:
                self.page_box.clearFocus()
                t = self.cur_tab()
                if t:
                    t.scroll.setFocus()
                return True
        return super().eventFilter(obj, ev)

    def update_status(self):
        t = self.cur_tab()
        if t and t.doc:
            if not self.page_box.hasFocus():                 # don't overwrite a number the user is typing
                self.page_box.setText(str(t.cur + 1))
            self.page_total.setText(f"/ {len(t.doc)}")
            self.zoom_lbl.setText(f"{round(t.zoom * 100)}%")
            self.page_lbl.show()
            self.place_page_bar()
            self.zoom_lbl.show()
            name = os.path.basename(t.path) if t.path else "Untitled"
            self.setWindowTitle(f"{'*' if t.dirty else ''}{name} - PDF Studio")
        else:
            self.page_box.clear()
            self.zoom_lbl.setText("")
            self.page_lbl.hide()
            self.zoom_lbl.hide()
            self.setWindowTitle("PDF Studio")

    def maybe_save(self, tab):
        if not tab.dirty:
            return True
        self.tabs.setCurrentWidget(tab)
        r = QMessageBox.question(self, "PDF Studio",
                                 f"Save changes to {os.path.basename(tab.path) if tab.path else 'this file'}?",
                                 QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard |
                                 QMessageBox.StandardButton.Cancel)
        if r == QMessageBox.StandardButton.Cancel:
            return False
        if r == QMessageBox.StandardButton.Save:
            return tab.save()
        return True

    def close_tab(self, i=None):
        if i is None or isinstance(i, bool):
            i = self.tabs.currentIndex()
        if i < 0:
            return
        t = self.tabs.widget(i)
        if not self.maybe_save(t):
            return
        self.tabs.removeTab(i)
        t.deleteLater()
        if self.tabs.count() == 0:
            self.stack.setCurrentIndex(0)
        self.update_status()

    def step_tab(self, d):
        n = self.tabs.count()
        if n:
            self.tabs.setCurrentIndex((self.tabs.currentIndex() + d) % n)

    # ---------- single instance ----------
    def bring_to_front(self):
        """Restore (if minimized) and put this window on top of every other window and give it focus, as
        Acrobat does when you open another PDF. Windows only lets the foreground process take focus, so we
        join the foreground thread's input queue and use the topmost toggle as a fallback."""
        self.setWindowState((self.windowState() & ~Qt.WindowState.WindowMinimized) | Qt.WindowState.WindowActive)
        self.show()
        self.raise_()
        self.activateWindow()
        if sys.platform != "win32":
            return
        try:
            from ctypes import wintypes
            user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
            user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int,
                                            ctypes.c_int, ctypes.c_int, wintypes.UINT]
            hwnd = int(self.winId())
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, 9)                              # SW_RESTORE
            this_thread = kernel32.GetCurrentThreadId()
            fg = user32.GetForegroundWindow()
            fg_thread = user32.GetWindowThreadProcessId(fg, None) if fg else 0
            attached = bool(fg_thread and fg_thread != this_thread
                            and user32.AttachThreadInput(this_thread, fg_thread, True))
            try:
                flags = 0x0001 | 0x0002 | 0x0040                         # NOSIZE | NOMOVE | SHOWWINDOW
                user32.SetWindowPos(hwnd, wintypes.HWND(-1), 0, 0, 0, 0, flags)         # HWND_TOPMOST: above every window now
                user32.SetWindowPos(hwnd, wintypes.HWND(-2), 0, 0, 0, 0, flags)         # HWND_NOTOPMOST: back to a normal window
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
                user32.SetFocus(hwnd)
            finally:
                if attached:
                    user32.AttachThreadInput(this_thread, fg_thread, False)
        except Exception:
            pass

    def attach_server(self, server):
        self.server = server
        server.newConnection.connect(self._on_connection)

    def _on_connection(self):
        sock = self.server.nextPendingConnection()
        buf = bytearray()
        finished = []
        sock.readyRead.connect(lambda: buf.extend(bytes(sock.readAll())))

        def done():
            if finished:                                    # disconnected() and the state check can both fire
                return
            finished.append(True)
            buf.extend(bytes(sock.readAll()))
            paths = [l for l in buf.decode("utf-8", "ignore").splitlines() if l.strip()]
            self.open_paths(paths)
            self.bring_to_front()
            sock.deleteLater()
        sock.disconnected.connect(done)
        if sock.state() == QLocalSocket.LocalSocketState.UnconnectedState:
            QTimer.singleShot(0, done)                      # the sender was already gone: handle what it wrote

    # ---------- UI ----------
    def _act(self, text, slot, key=None, checkable=False):
        a = QAction(text, self)
        a.triggered.connect(slot)
        if key:
            a.setShortcut(QKeySequence(key))
        a.setCheckable(checkable)
        return a

    def _tab_call(self, name, *args):
        def run():
            t = self.cur_tab()
            if t:
                getattr(t, name)(*args)
        return run

    def _build_actions(self):
        m = self.menuBar()
        tb = QToolBar("main")
        tb.setMovable(False)
        tb.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        self.toolbar = tb                                    # placed under the tab strip (see __init__)
        T = self._tab_call
        self.a_edit = self._act("✎ Edit text", self.toggle_edit, "Ctrl+E", True)
        # Close tab / Delete page(s) / Move up / Move down / Merge PDFs are no longer shown in the menus or the
        # toolbar (tabs have their x, thumbnails have the trash button and drag-to-reorder). The keyboard
        # shortcuts keep working through hidden window-level actions.
        for _a in (self._act("Close tab", self.close_tab, "Ctrl+W"),
                   self._act("Delete page(s)", T("delete_pages"), "Delete"),
                   self._act("Move up", T("move", -1), "Alt+Up"),
                   self._act("Move down", T("move", 1), "Alt+Down")):
            self.addAction(_a)
        acts = {
            "File": [self._act("Open…", self.open_dialog, "Ctrl+O"),
                     self._act("Save", T("save"), "Ctrl+S"),
                     self._act("Save As…", T("save_as"), "Ctrl+Shift+S"),
                     self._act("🖨 Print…", self.print_doc, "Ctrl+P")],
            "Pages": [self._act("＋ Add blank page", T("add_blank"), "Ctrl+Shift+N"),
                      self._act("⇪ Insert PDF…", T("insert_pdf"), "Ctrl+I"),
                      self._act("⟳ Rotate", T("rotate"), "Ctrl+R")],
            "Edit": [self._act("↶ Undo", T("undo"), "Ctrl+Z"), self.a_edit],
            "View": [self._act("Zoom in", lambda: self._zoom(1.2), "Ctrl+="),
                     self._act("Zoom out", lambda: self._zoom(1 / 1.2), "Ctrl+-"),
                     self._act("Next tab", lambda: self.step_tab(1), "Ctrl+PgDown"),
                     self._act("Previous tab", lambda: self.step_tab(-1), "Ctrl+PgUp")],
            "Tools": [self._act("Add right-click menu in Explorer", self.do_install),
                      self._act("Remove right-click menu", self.do_uninstall)],
        }
        for title, lst in acts.items():
            menu = QMenu(title, self)
            m.addMenu(menu)
            for a in lst:
                menu.addAction(a)
        for title in ("File", "Pages", "Edit"):
            for a in acts[title]:
                tb.addAction(a)
            tb.addSeparator()

    def _zoom(self, f):
        t = self.cur_tab()
        if t:
            t.set_zoom(t.zoom * f)

    def toggle_edit(self):
        on = self.a_edit.isChecked()
        for i in range(self.tabs.count()):
            self.tabs.widget(i).set_edit(on)

    def open_dialog(self):
        ps, _ = QFileDialog.getOpenFileNames(self, "Open PDF", "", "PDF (*.pdf)")
        self.open_paths(ps)

    def print_doc(self, pages=None):
        t = self.cur_tab()
        if not (t and t.doc):
            return
        if not QPrinterInfo.availablePrinterNames():
            return QMessageBox.warning(self, "Print", "No printer is installed.")
        pages = pages if isinstance(pages, (list, tuple)) and pages else None   # (menu actions pass a bool)
        dlg = PrintDialog(self, t.doc, pages[0] if pages else t.cur, preset=pages,
                          initial_preferences=self._print_dialog_preferences)
        result = dlg.exec()
        # Preserve the most recent choices even if the print dialog was canceled.
        self._print_dialog_preferences = dlg.preferences()
        if result != QDialog.DialogCode.Accepted:
            return
        st = dlg.settings()
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            ok = print_pdf(self, t.doc, st, os.path.basename(t.path) if t.path else "PDF Studio")
        except Exception as e:
            QApplication.restoreOverrideCursor()
            return QMessageBox.warning(self, "Print failed", str(e))
        QApplication.restoreOverrideCursor()
        self.statusBar().showMessage(
            f"Sent {len(st['pages'])} page(s) to {st['printer']}" if ok else "Printing canceled", 5000)
        if st.get("warning"):
            QMessageBox.information(self, "Print", st["warning"])

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.open_paths([u.toLocalFile() for u in e.mimeData().urls()])

    def closeEvent(self, e):
        for i in range(self.tabs.count()):
            if not self.maybe_save(self.tabs.widget(i)):
                return e.ignore()
        self._settings.setValue("window/geometry", self.saveGeometry())
        e.accept()

    # ---------- explorer menu ----------
    def offer_install_on_first_run(self):
        """The exe registers itself only through 'Add right-click menu in Explorer' / --install. A freshly
        downloaded exe has done neither, so ask once (not again after 'No') instead of leaving the user with
        no 'Open with' entry and no right-click menu."""
        if sys.platform != "win32" or not getattr(sys, "frozen", False):
            return
        if _registered_exe() is not None or self._settings.value("install/declined") == "1":
            return
        answer = QMessageBox.question(
            self, "PDF Studio",
            "Add PDF Studio to Explorer?\n\nIt will appear in 'Open with' and in the right-click menu of PDFs, "
            "and any older version is replaced.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes)
        if answer == QMessageBox.StandardButton.Yes:
            self.do_install()
        else:
            self._settings.setValue("install/declined", "1")

    def do_install(self):
        try:
            install_context_menu()
            QMessageBox.information(self, "PDF Studio",
                                    "Done. Select PDFs/images in Explorer → right-click → Combine (PDF Studio).\n"
                                    "To make it the default PDF app: Settings > Apps > Default apps > PDF Studio.\n"
                                    "(Windows 11: under 'Show more options' / Shift+F10)")
        except Exception as e:
            QMessageBox.warning(self, "Error", str(e))

    def do_uninstall(self):
        uninstall_context_menu()
        QMessageBox.information(self, "PDF Studio", "Removed.")


# ============================================================================
#  SINGLE INSTANCE: a second launch hands its files to the running window (-> new tabs)
# ============================================================================
def _forward(paths):
    if sys.platform == "win32":
        try:
            ctypes.windll.user32.AllowSetForegroundWindow(0xFFFFFFFF)   # ASFW_ANY: the running window may take focus
        except Exception:
            pass
    sock = QLocalSocket()
    sock.connectToServer(SERVER_NAME)
    if not sock.waitForConnected(300):
        return False
    sock.write("\n".join(paths).encode("utf-8"))
    sock.waitForBytesWritten(1500)
    sock.disconnectFromServer()
    if sock.state() != QLocalSocket.LocalSocketState.UnconnectedState:
        sock.waitForDisconnected(500)
    return True


def acquire_instance(paths):
    for _ in range(6):
        if _forward(paths):
            return "forwarded", None
        srv = QLocalServer()
        if srv.listen(SERVER_NAME):
            return "server", srv
        time.sleep(0.15)                                    # someone else is starting up right now
    QLocalServer.removeServer(SERVER_NAME)                  # stale socket
    srv = QLocalServer()
    if srv.listen(SERVER_NAME):
        return "server", srv
    return "alone", None


def _set_app_user_model_id():
    """Windows picks the taskbar icon and groups taskbar buttons by AppUserModelID. Without our own id a
    script run shows the Python icon in the taskbar; with it, PDF Studio gets its own icon and button."""
    if sys.platform == "win32":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_USER_MODEL_ID)
        except Exception:
            pass


def _report_exception(exc_type, exc, tb):
    """A windowed exe has no console, so an unexpected error would silently end or freeze the app. Show it."""
    import traceback
    text = "".join(traceback.format_exception(exc_type, exc, tb))
    try:
        QMessageBox.critical(None, "PDF Studio - unexpected error", text[-3000:])
    except Exception:
        _msg(text[-1500:])


if __name__ == "__main__":
    sys.excepthook = _report_exception
    _set_app_user_model_id()
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Studio")
    app.setWindowIcon(app_icon())
    if len(sys.argv) > 1 and sys.argv[1] in ("--install", "--uninstall"):
        (install_context_menu if sys.argv[1] == "--install" else uninstall_context_menu)()
        if sys.argv[1] == "--install":
            _msg(f"Installed (previous version removed):\n{_app_exe()}\n\nSettings > Default apps will open now: "
                 "search 'PDF Studio' (or .pdf) and set it as default.", 0x40)
            try:
                os.startfile("ms-settings:defaultapps")
            except OSError:
                pass
        else:
            _msg("Removed.", 0x40)
        sys.exit(0)
    if len(sys.argv) > 2 and sys.argv[1] == "--export-icon":
        save_ico(sys.argv[2])
        sys.exit(0)
    apply_theme(app)
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    paths = [_open_after] if _open_after else [os.path.abspath(a) for a in args]
    mode, server = acquire_instance(paths)
    if mode == "forwarded":
        sys.exit(0)
    w = Main()
    if server:
        w.attach_server(server)
    w.show()
    w.open_paths(paths)
    QTimer.singleShot(4000, _cleanup_old_page_exports)      # housekeeping after the window is already up
    QTimer.singleShot(2500, refresh_registration_if_stale)  # fix a dead exe path in the Explorer registration
    if not paths:                                           # started on its own (not to open a file): offer to install
        QTimer.singleShot(700, w.offer_install_on_first_run)
    sys.exit(app.exec())

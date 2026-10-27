"""
PDF Studio - PDF reader/editor for Windows
  * tabs (like a browser): every opened PDF gets its own tab, in order
  * continuous view (scroll through all pages like one document), page x / n in the status bar
  * drag pages from the preview to export as separate PDF or JPG files
  * edit, move and resize text, images and shapes (Edit mode), delete / add / move / rotate pages, insert PDFs, undo
  * built-in print dialog: printer, A0-A5 + custom size, orientation, scale, page range, duplex,
    paper source (auto / bypass tray), paper type,
    save-toner and grayscale modes, flip on left/right or top/bottom edge, and live preview
  * Explorer right-click: combine multiple PDFs/images (kept in memory, no file is written until you save), or open a single selection directly
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
MERGE_TAG = "::merge::"                                    # first line of a hand-over: "combine the files below"


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
            if time.time() - stable >= 0.35:
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

    if len(items) == 1 and items[0].lower().endswith(".pdf"):
        return [items[0]]                                   # a single PDF is simply opened
    return [MERGE_TAG] + items                              # everything else is combined inside the app, in memory


def combine_to_doc(items):
    """Combine PDFs and images into ONE document held in memory. Nothing is written to disk: the result opens as an
    unsaved tab and becomes a file only when the user saves it. Returns (document, [skipped files])."""
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
    return out, skipped


def _unique_target(folder, base="Combined"):
    target, n = os.path.join(folder, base + ".pdf"), 1
    while os.path.exists(target):
        n += 1
        target = os.path.join(folder, f"{base} ({n}).pdf")
    return target


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
    if "--background" in sys.argv and _fast_forward(["::noop::"]):
        sys.exit(0)                                         # already running: nothing to do
    if _fast_forward(_open_after or
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
    src_dir = os.path.dirname(me)
    if os.path.isdir(os.path.join(src_dir, "_internal")):   # folder build (starts much faster than a single-file exe)
        shutil.rmtree(_install_dir(), ignore_errors=True)
        shutil.copytree(src_dir, _install_dir())
    else:
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
                         QTextCursor, QTextBlockFormat, QTransform, QIntValidator)
from PyQt6.QtNetwork import QLocalServer, QLocalSocket
from PyQt6.QtPrintSupport import QPrinter, QPrinterInfo
from PyQt6.QtWidgets import (QSystemTrayIcon, QApplication, QButtonGroup, QColorDialog, QComboBox, QDialog, QDialogButtonBox,
                             QCheckBox, QDoubleSpinBox, QFileDialog, QFormLayout, QFrame, QGraphicsDropShadowEffect,
                             QGridLayout, QGroupBox, QHBoxLayout,
                             QInputDialog, QLabel, QLineEdit, QListWidget, QListWidgetItem,
                             QMainWindow, QMenu, QMessageBox, QPlainTextEdit, QProgressDialog, QPushButton, QRadioButton,
                             QScrollArea, QAbstractItemView, QSpinBox, QSplitter, QStyledItemDelegate,
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
    img = QImage(pix.samples_mv, pix.width, pix.height, pix.stride, fmt)
    img._keep = pix                                         # the image reads straight from the pixmap: no copy
    return img


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
QLabel#PageText { background:transparent; border:none; color:$TEXT2; font-size:22px; font-weight:700; }
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


def _autostart_enabled():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as k:
            winreg.QueryValueEx(k, "PDFStudio")
            return True
    except OSError:
        return False


def _set_autostart(on):
    """Start PDF Studio hidden when Windows starts, so double-clicking a PDF opens it instantly."""
    try:
        import winreg
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run")
        if on:
            winreg.SetValueEx(k, "PDFStudio", 0, winreg.REG_SZ, f'{_base_cmd()} --background')
        else:
            try:
                winreg.DeleteValue(k, "PDFStudio")
            except OSError:
                pass
    except OSError:
        pass


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

    def _set_print_scale(self, pct):
        """Make `pct` (relative to Fit) the real print scale: switches to Custom Scale and sets the
        spin box, so the preview, the saved settings and the printed output all use the same value."""
        pct = int(max(self.sp_scale.minimum(), min(self.sp_scale.maximum(), round(pct))))
        self.sp_scale.blockSignals(True)
        self.sp_scale.setValue(pct)
        self.sp_scale.blockSignals(False)
        idx = self.cb_scale.findData("custom")
        if idx >= 0 and self.cb_scale.currentIndex() != idx:
            self.cb_scale.setCurrentIndex(idx)          # fires update_preview through _scale_mode_changed
        else:
            self.update_preview()

    def _wheel_zoom(self, step):
        # the wheel changes the content's real scale (what gets printed), not just the preview camera
        cur = getattr(self, "_eff_pct", 100.0)
        new = cur * (1.15 ** step)
        if round(new) == round(cur):
            new = cur + (1 if step > 0 else -1)
        self._set_print_scale(new)

    def _reset_zoom(self):
        self.cb_scale.setCurrentIndex(max(0, self.cb_scale.findData("fit")))
        self.sp_scale.blockSignals(True)
        self.sp_scale.setValue(100)
        self.sp_scale.blockSignals(False)
        self.update_preview()

    def reset_offset(self):
        self.zoom = 1.0
        self.image_rotation = 0                       # undo button / Shift+drag rotation too
        self.cb_scale.blockSignals(True)
        self.cb_scale.setCurrentIndex(max(0, self.cb_scale.findData("fit")))
        self.cb_scale.blockSignals(False)
        self.sp_scale.blockSignals(True)
        self.sp_scale.setValue(100)
        self.sp_scale.blockSignals(False)
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
        m = PRINT_MARGIN_MM / 25.4 * 72
        aw, ah = pw - 2 * m, ph - 2 * m
        base_x, base_y = m * k, m * k                                      # relative to the paper's top-left

        offset_pt = (self.ox_mm / 25.4 * 72, self.oy_mm / 25.4 * 72)
        rotation = self.image_rotation % 360
        cells = layout_sheet(self.doc, sheet, rows, cols, aw, ah,
                             self.cb_scale.currentData(), self.sp_scale.value(), offset_pt, rotation)
        # effective scale relative to Fit (also right for "Actual size"), shown next to the preview
        fit_cells = layout_sheet(self.doc, sheet, rows, cols, aw, ah, "fit", 100, offset_pt, rotation)
        try:
            self._eff_pct = cells[0][2][2] / fit_cells[0][2][2] * 100.0
        except (ZeroDivisionError, IndexError):
            self._eff_pct = 100.0
        self.lbl_zoom.setText(f"{round(self._eff_pct)}%")

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


# ---- fonts: keep the original font of edited text, or use the one picked in the Edit bar ----------------------------
_RTL_RE = re.compile(r"[\u0590-\u08FF\uFB50-\uFEFF]")
_STYLE_WORDS = {"bold", "italic", "oblique", "regular", "normal", "bolditalic", "boldoblique"}
_SYS_FONTS = None                       # normalised family -> {(bold, italic): file}
_SYS_FAMILY_NAMES = {}                  # normalised family -> display name


def _fam_key(s):
    return re.sub(r"[^0-9a-z\u0080-\uffff]", "", (s or "").lower())


def _pdf_font_parts(name, flags=0):
    """'ABCDEF+Arial-BoldMT' -> ('arial', bold, italic)."""
    n = re.sub(r"^[A-Z]{6}\+", "", name or "")
    low = n.lower()
    bold = bool(flags & 16) or "bold" in low or "black" in low or "heavy" in low
    italic = bool(flags & 2) or "italic" in low or "oblique" in low
    low = re.sub(r"(bold|italic|oblique|regular|roman|normal|black|heavy|semibold|light|medium)", "", low)
    low = _fam_key(low)
    low = re.sub(r"(psmt|mt|ps)+$", "", low)
    return low, bold, italic


def _split_style(display):
    name = re.sub(r"\s*\((TrueType|OpenType|All res|Type 1)[^)]*\)\s*$", "", display).strip()
    words, bold, italic = name.split(), False, False
    while words and words[-1].lower() in _STYLE_WORDS:
        w = words.pop().lower()
        bold = bold or "bold" in w
        italic = italic or "italic" in w or "oblique" in w
    return " ".join(words), bold, italic


def _system_fonts():
    global _SYS_FONTS
    if _SYS_FONTS is not None:
        return _SYS_FONTS
    idx = {}

    def add(family, bold, italic, path):
        k = _fam_key(family)
        if k and path and os.path.isfile(path) and path.lower().endswith((".ttf", ".otf")):
            idx.setdefault(k, {}).setdefault((bold, italic), path)
            _SYS_FAMILY_NAMES.setdefault(k, family)
    try:
        import winreg
        wdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for hive, base in ((winreg.HKEY_LOCAL_MACHINE, wdir),
                           (winreg.HKEY_CURRENT_USER, os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                                                   r"Microsoft\Windows\Fonts"))):
            try:
                key = winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            except OSError:
                continue
            i = 0
            while True:
                try:
                    disp, val, _t = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                if "&" in disp:
                    continue
                fam, b, it = _split_style(disp)
                add(fam, b, it, val if os.path.isabs(val) else os.path.join(base, val))
    except ImportError:                                     # not Windows (tests / other systems): ask fontconfig
        try:
            import subprocess
            out = subprocess.run(["fc-list", "--format", "%{file}|%{family[0]}|%{style[0]}\\n"],
                                 capture_output=True, text=True, timeout=20).stdout
            for ln in out.splitlines():
                parts = ln.split("|")
                if len(parts) == 3:
                    st = parts[2].lower()
                    add(parts[1], "bold" in st, "italic" in st or "oblique" in st, parts[0])
        except Exception:
            pass
    _SYS_FONTS = idx
    return idx


def font_families():
    """Display names of the fonts that can be applied to text (installed system fonts)."""
    _system_fonts()
    return sorted(_SYS_FAMILY_NAMES.values(), key=str.lower)


def _find_system_font(fam_key, bold, italic):
    d = _system_fonts().get(fam_key)
    if not d:
        return None
    for k in ((bold, italic), (bold, False), (False, italic), (False, False)):
        if k in d:
            return d[k]
    return next(iter(d.values()))


class _FontSpec:
    """A font ready to write with: embedded / system file (Unicode, TrueType) or one of the 14 standard fonts."""
    def __init__(self, key, buf=None, path=None, base14=None):
        self.key, self.buf, self.path, self.base14 = key, buf, path, base14
        self.font = None
        if base14:
            self.font = fitz.Font(base14)
        else:
            self.font = fitz.Font(fontbuffer=buf) if buf else fitz.Font(fontfile=path)

    def covers(self, text):
        if self.base14:
            return all(ord(c) < 256 for c in text)
        try:
            return all(self.font.has_glyph(ord(c)) for c in text if not c.isspace())
        except Exception:
            return False

    def length(self, text, size):
        try:
            return self.font.text_length(text, fontsize=size)
        except Exception:
            return len(text) * size * 0.5

    def data(self):
        if self.buf:
            return self.buf
        with open(self.path, "rb") as fh:
            return fh.read()


def _base14_name(flags, bold, italic):
    if flags & 8:                                           # monospaced
        return {(0, 0): "cour", (1, 0): "cobo", (0, 1): "coit", (1, 1): "cobi"}[(int(bold), int(italic))]
    if flags & 4:                                           # serif
        return {(0, 0): "tiro", (1, 0): "tibo", (0, 1): "tiit", (1, 1): "tibi"}[(int(bold), int(italic))]
    return {(0, 0): "helv", (1, 0): "hebo", (0, 1): "heit", (1, 1): "hebi"}[(int(bold), int(italic))]


_FALLBACK_FAMILIES = ("tahoma", "arial", "segoeui", "calibri", "dejavusans", "notosansarabic", "notosans",
                      "freesans", "liberationsans")


def _embedded_font_buf(doc, page, pdf_font, cache):
    k = ("emb", id(doc), pdf_font)
    if k in cache:
        return cache[k]
    buf = None
    try:
        want = _fam_key(re.sub(r"^[A-Z]{6}\+", "", pdf_font or ""))
        for f in page.get_fonts(full=True):
            if _fam_key(re.sub(r"^[A-Z]{6}\+", "", f[3] or "")) == want:
                _n, ext, _t, content = doc.extract_font(f[0])
                if content and ext in ("ttf", "otf", "ttc"):
                    buf = content
                break
    except Exception:
        buf = None
    cache[k] = buf
    return buf


def _spec_for(cache, key, **kw):
    if key not in cache:
        try:
            cache[key] = _FontSpec(key, **kw)
        except Exception:
            cache[key] = None
    return cache[key]


def _resolve_font(page, style, text, cache):
    """Pick the font for `text`: the user's choice, else the PDF's own font (embedded or the same family on this
    PC), else a stand-in that has the needed letters."""
    flags = style.get("flags", 0)
    fam_key, bold, italic = _pdf_font_parts(style.get("font"), flags)
    if style.get("family"):                                 # picked in the Edit bar
        fam_key = _fam_key(style["family"])
        bold, italic = bool(flags & 16), bool(flags & 2)
        user = True
    else:
        user = False
    tried = []
    if not user and style.get("font"):
        buf = _embedded_font_buf(page.parent, page, style["font"], cache)
        if buf:
            tried.append(_spec_for(cache, ("buf", id(page.parent), style["font"]), buf=buf))
    p = _find_system_font(fam_key, bold, italic)
    if p:
        tried.append(_spec_for(cache, ("file", p), path=p))
    for t in tried:
        if t is not None and t.covers(text):
            return t
    rtl_or_wide = any(ord(c) >= 256 for c in text)
    if not rtl_or_wide:
        return _spec_for(cache, ("b14", _base14_name(flags, bold, italic)), base14=_base14_name(flags, bold, italic))
    for fk in _FALLBACK_FAMILIES:
        p = _find_system_font(fk, bold, italic)
        if p:
            s = _spec_for(cache, ("file", p), path=p)
            if s is not None and s.covers(text):
                return s
    for fk, d in _system_fonts().items():                   # any installed font that has these letters
        p = d.get((bold, italic)) or next(iter(d.values()))
        s = _spec_for(cache, ("file", p), path=p)
        if s is not None and s.covers(text):
            return s
    return _spec_for(cache, ("b14", "helv"), base14="helv")


def _text_width(page, text, style, cache):
    return _resolve_font(page, style, text, cache).length(text, style["size"])


def _remove_draws(page, draws):
    """Take vector objects off the page. Shapes are first removed only if they lie completely inside their own
    (slightly widened) area, so a shape standing close by is left alone; whatever is still there afterwards is
    removed with the 'touched' rule. Empty frames are removed along their four edges."""
    FULL = getattr(fitz, "PDF_REDACT_LINE_ART_REMOVE_IF_COVERED", 1)
    TOUCH = getattr(fitz, "PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED", 2)
    NOIMG = getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0)
    NOTXT = getattr(fitz, "PDF_REDACT_TEXT_NONE", 1)

    def run(rects, mode):
        for rc in rects:
            page.add_redact_annot(rc, fill=False)
        try:
            page.apply_redactions(images=NOIMG, graphics=mode, text=NOTXT)
        except TypeError:
            page.apply_redactions(images=NOIMG)

    solid = [o for o in draws if not o.get("hollow")]
    hollow = [o for o in draws if o.get("hollow")]
    if solid:
        rects = []
        for o in solid:
            m = max([(p.get("width") or 1.0) for p in o["paths"]] + [1.0]) * 6 + 2     # room for the stroke's mitre reach
            rects.append(fitz.Rect(o["rect"]) + (-m, -m, m, m))
        run(rects, FULL)
        try:
            left = page.get_drawings()
        except Exception:
            left = []
        rest = []
        for o in solid:
            for p in o["paths"]:
                pr = p["rect"]
                if any(abs(d["rect"].x0 - pr.x0) < 0.6 and abs(d["rect"].y0 - pr.y0) < 0.6
                       and abs(d["rect"].x1 - pr.x1) < 0.6 and abs(d["rect"].y1 - pr.y1) < 0.6 for d in left):
                    rest.append(o)
                    break
        if rest:
            run([fitz.Rect(o["rect"]) + (0.4, 0.4, -0.4, -0.4) for o in rest], TOUCH)
    if hollow:
        rects = []
        for o in hollow:
            r = fitz.Rect(o["rect"])
            t = max(1.0, (o["paths"][0].get("width") or 1.0)) / 2 + 1.5
            rects += [fitz.Rect(r.x0 - t, r.y0 - t, r.x1 + t, r.y0 + t), fitz.Rect(r.x0 - t, r.y1 - t, r.x1 + t, r.y1 + t),
                      fitz.Rect(r.x0 - t, r.y0 - t, r.x0 + t, r.y1 + t), fitz.Rect(r.x1 - t, r.y0 - t, r.x1 + t, r.y1 + t)]
        run(rects, TOUCH)


def _emit_text(page, text, baseline_y, style, align, x0, x1, cache=None):
    """Write one line of text with its baseline at `baseline_y`. `align` ('left' / 'center' / 'right') says which of the
    container edges x0 / x1 (or its middle) the text sticks to. Keeps the original font; Persian / Arabic are shaped
    by the HTML engine."""
    if not text or not text.strip():
        return
    cache = {} if cache is None else cache
    size = max(1.0, float(style["size"]))
    color = style.get("color", (0, 0, 0))
    spec = _resolve_font(page, style, text, cache)
    w = spec.length(text, size)
    rtl = bool(_RTL_RE.search(text))
    if align == "right":
        left = x1 - w
    elif align == "center":
        left = (x0 + x1) / 2 - w / 2
    else:
        left = x0
    if not rtl:
        if spec.base14:
            page.insert_text(fitz.Point(left, baseline_y), text, fontsize=size, fontname=spec.base14, color=color)
        else:
            page.insert_font(fontname=spec.key if isinstance(spec.key, str) else "F" + str(abs(hash(spec.key)) % 10**8),
                             fontbuffer=spec.data())
            page.insert_text(fitz.Point(left, baseline_y), text, fontsize=size,
                             fontname=spec.key if isinstance(spec.key, str) else "F" + str(abs(hash(spec.key)) % 10**8),
                             color=color)
        return
    # right-to-left / shaped text: HTML engine with the chosen font
    try:
        asc = abs(getattr(spec.font, "ascender", 0.9)) or 0.9
        desc = abs(getattr(spec.font, "descender", 0.25)) or 0.25
        lh = size * max(1.2, asc + desc)
        top = baseline_y - ((lh - (asc + desc) * size) / 2 + asc * size)
        bw = w * 1.35 + 24
        if align == "right":
            bx0, bx1 = x1 - bw, x1
        elif align == "center":
            bx0, bx1 = (x0 + x1) / 2 - bw / 2, (x0 + x1) / 2 + bw / 2
        else:
            bx0, bx1 = x0, x0 + bw
        box = fitz.Rect(bx0, top - 2, bx1, top + lh + 4)
        arch = fitz.Archive()
        arch.add(spec.data() if not spec.base14 else b"", "ft.ttf")
        face = '@font-face{font-family:ft;src:url("ft.ttf");}' if not spec.base14 else ""
        css = (f'{face}*{{font-family:{"ft" if not spec.base14 else "sans-serif"};font-size:{size}pt;'
               f'color:rgb({int(color[0]*255)},{int(color[1]*255)},{int(color[2]*255)});white-space:nowrap}}'
               f'p{{margin:0;line-height:{lh}pt}}')
        ta = {"right": "left", "left": "right", "center": "center"}[align if align in ("left", "right", "center") else "left"]   # MuPDF's rtl text-align means start / end
        page.insert_htmlbox(box, f'<p dir="rtl" style="text-align:{ta}">{_html.escape(text)}</p>', css=css, archive=arch)
    except Exception:
        page.insert_text(fitz.Point(left, baseline_y), text.encode("latin-1", "replace").decode("latin-1"),
                         fontsize=size, fontname="helv", color=color)


def _redraw_path(page, path, T, avg, ov=None):
    """Draw a vector path again with every point passed through T (move / scale). `ov` = new look picked in the
    Edit bar: keys color (outline), fill, width, dash."""
    if not path.get("items"):
        return
    ov = ov or {}
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
    fill = ov["fill"] if "fill" in ov else path.get("fill")
    color = ov["color"] if "color" in ov else path.get("color")
    if fill is None and color is None:
        fill, color = path.get("fill"), path.get("color")
        if fill is None and color is None:
            return
    base_w = ov["width"] if "width" in ov else (path.get("width") or 1.0)
    if color is None and "color" not in ov and fill is not None:
        base_w = path.get("width") or 1.0
    cap = path.get("lineCap") or (0,)
    dashes = path.get("dashes")
    dashes = dashes if dashes and dashes.strip() not in ("[] 0", "") else None
    if "dash" in ov:
        dashes = _dash_string(ov["dash"], base_w)
    try:
        sh.finish(fill=fill, color=color, width=max(0.05, base_w * (1 if "width" in ov else avg)),
                  dashes=dashes,
                  even_odd=bool(path.get("even_odd")), closePath=bool(path.get("closePath")),
                  lineCap=int(cap[0]) if isinstance(cap, (tuple, list)) else int(cap),
                  fill_opacity=path.get("fill_opacity") if path.get("fill_opacity") is not None else 1,
                  stroke_opacity=path.get("stroke_opacity") if path.get("stroke_opacity") is not None else 1)
        sh.commit()
    except Exception:
        pass


# ---- ready-made shapes and page borders ---------------------------------------------------------------------------
SHAPES = [("line", "Line"), ("arrow", "Arrow"), ("rect", "Rectangle"), ("square", "Square"),
          ("rrect", "Rounded rectangle"), ("ellipse", "Ellipse"), ("circle", "Circle"),
          ("triangle", "Triangle"), ("rtriangle", "Right triangle"), ("diamond", "Diamond"),
          ("pentagon", "Pentagon"), ("hexagon", "Hexagon"), ("star", "Star"), ("plus", "Plus"),
          ("rarrow", "Block arrow")]
SQUARE_SHAPES = {"square", "circle"}                    # always drawn 1:1
LINE_SHAPES = {"line", "arrow"}
DASH_STYLES = [("solid", "Solid"), ("dashed", "Dashed"), ("dotted", "Dotted")]


def _regular_poly(n, star=False):
    pts = []
    for k in range(n * (2 if star else 1)):
        a = -math.pi / 2 + k * (math.pi / n if star else 2 * math.pi / n)
        r = 1.0 if (not star or k % 2 == 0) else 0.382
        pts.append((r * math.cos(a), r * math.sin(a)))
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
    return [((x - x0) / (x1 - x0), (y - y0) / (y1 - y0)) for x, y in pts]


_UNIT_POLYS = {
    "triangle": [(.5, 0), (1, 1), (0, 1)],
    "rtriangle": [(0, 0), (0, 1), (1, 1)],
    "diamond": [(.5, 0), (1, .5), (.5, 1), (0, .5)],
    "hexagon": [(.25, 0), (.75, 0), (1, .5), (.75, 1), (.25, 1), (0, .5)],
    "plus": [(1/3, 0), (2/3, 0), (2/3, 1/3), (1, 1/3), (1, 2/3), (2/3, 2/3), (2/3, 1), (1/3, 1), (1/3, 2/3),
             (0, 2/3), (0, 1/3), (1/3, 1/3)],
    "rarrow": [(0, .3), (.6, .3), (.6, 0), (1, .5), (.6, 1), (.6, .7), (0, .7)],
}


def _shape_geom(kind, x0, y0, x1, y1, width=1.0):
    """Drawing primitives of a shape inside the box (x0,y0)-(x1,y1) (page points)."""
    l, r, t, b = min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)
    w, h = r - l, b - t
    if kind in ("line", "arrow"):
        if kind == "line":
            return [{"t": "poly", "pts": [(x0, y0), (x1, y1)], "closed": False}]
        ln = math.hypot(x1 - x0, y1 - y0) or 1.0
        ux, uy = (x1 - x0) / ln, (y1 - y0) / ln
        head = min(max(10.0, width * 4.5), ln * 0.45)
        bx, by = x1 - ux * head, y1 - uy * head
        nx, ny = -uy, ux
        return [{"t": "poly", "pts": [(x0, y0), (bx, by)], "closed": False},
                {"t": "poly", "pts": [(x1, y1), (bx + nx * head * .42, by + ny * head * .42),
                                      (bx - nx * head * .42, by - ny * head * .42)], "closed": True, "solid": True}]
    if kind in ("rect", "square"):
        return [{"t": "rect", "r": (l, t, r, b)}]
    if kind == "rrect":
        return [{"t": "rrect", "r": (l, t, r, b), "k": 0.18}]
    if kind in ("ellipse", "circle"):
        return [{"t": "oval", "r": (l, t, r, b)}]
    if kind == "pentagon":
        unit = _regular_poly(5)
    elif kind == "star":
        unit = _regular_poly(5, star=True)
    else:
        unit = _UNIT_POLYS.get(kind, [(0, 0), (1, 0), (1, 1), (0, 1)])
    return [{"t": "poly", "pts": [(l + fx * w, t + fy * h) for fx, fy in unit], "closed": True}]


def _dash_string(style, width):
    w = max(0.5, float(width))
    if style == "dashed":
        return f"[{3 * w:.2f} {2 * w:.2f}] 0"
    if style == "dotted":
        return f"[0 {2 * w:.2f}] 0"
    return None


def _draw_shape_pdf(page, kind, x0, y0, x1, y1, stroke, fill, width, dash):
    """Draw a ready-made shape onto the page. stroke / fill are RGB tuples in 0..1 or None."""
    if stroke is None and fill is None:
        stroke = (0, 0, 0)
    width = max(0.1, float(width))
    dashes = _dash_string(dash, width)
    cap = 1 if dash == "dotted" else 0
    sh = page.new_shape()
    for pr in _shape_geom(kind, x0, y0, x1, y1, width):
        t = pr["t"]
        if t == "poly":
            pts = [fitz.Point(*p) for p in pr["pts"]]
            sh.draw_polyline(pts)
            if pr.get("solid"):
                sh.finish(fill=stroke or fill, color=stroke or fill, width=0.2, closePath=True, lineJoin=1)
            elif pr["closed"]:
                sh.finish(fill=fill, color=stroke, width=width, dashes=dashes, closePath=True, lineJoin=1, lineCap=cap)
            else:
                sh.finish(fill=None, color=stroke or fill, width=width, dashes=dashes, closePath=False, lineCap=1, lineJoin=1)
        else:
            rr = fitz.Rect(*pr["r"])
            if t == "oval":
                sh.draw_oval(rr)
            elif t == "rrect":
                sh.draw_rect(rr, radius=pr["k"])
            else:
                sh.draw_rect(rr)
            sh.finish(fill=fill, color=stroke, width=width, dashes=dashes, closePath=True, lineJoin=1, lineCap=cap)
    sh.commit()


def _draw_border(page, color, width, margin, style, radius):
    """Frame around the page: style 'solid' / 'dashed' / 'dotted' / 'double'."""
    r = page.rect
    box = fitz.Rect(r.x0 + margin, r.y0 + margin, r.x1 - margin, r.y1 - margin)
    if box.width < 10 or box.height < 10:
        return
    rects, w, dash = [(box, width)], width, style if style in ("dashed", "dotted") else None
    if style == "double":
        w = max(0.5, width / 3.0)
        gap = w * 1.6
        inner = fitz.Rect(box.x0 + w + gap, box.y0 + w + gap, box.x1 - w - gap, box.y1 - w - gap)
        rects = [(box, w)] + ([(inner, w)] if inner.width > 4 and inner.height > 4 else [])
    for rc, ww in rects:
        sh = page.new_shape()
        k = min(0.5, max(0.0, radius / max(1.0, min(rc.width, rc.height)))) if radius > 0 else None
        sh.draw_rect(rc, radius=k) if k else sh.draw_rect(rc)
        sh.finish(fill=None, color=color, width=ww, dashes=_dash_string(dash, ww), closePath=True, lineJoin=1,
                  lineCap=1 if dash == "dotted" else 0)
        sh.commit()


def _shape_qpath(kind, x0, y0, x1, y1, width=1.0):
    """The same shape as a QPainterPath (for the live preview and the toolbar icons). Returns (outline, solid parts)."""
    outline, solid = QPainterPath(), QPainterPath()
    for pr in _shape_geom(kind, x0, y0, x1, y1, width):
        t = pr["t"]
        tgt = solid if pr.get("solid") else outline
        if t == "poly":
            pts = [QPointF(*p) for p in pr["pts"]]
            tgt.moveTo(pts[0])
            for p in pts[1:]:
                tgt.lineTo(p)
            if pr["closed"]:
                tgt.closeSubpath()
        else:
            rr = QRectF(QPointF(pr["r"][0], pr["r"][1]), QPointF(pr["r"][2], pr["r"][3]))
            if t == "oval":
                tgt.addEllipse(rr)
            elif t == "rrect":
                rad = min(rr.width(), rr.height()) * pr["k"]
                tgt.addRoundedRect(rr, rad, rad)
            else:
                tgt.addRect(rr)
    return outline, solid


def _shape_icon(kind, color="#BBBBBB", size=22):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    pad = 4
    if kind in LINE_SHAPES:
        out, solid = _shape_qpath(kind, pad, size - pad, size - pad, pad, 1.0)
    else:
        out, solid = _shape_qpath(kind, pad, pad + 1, size - pad, size - pad - 1, 1.0)
    pen = QPen(QColor(color), 1.6)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    p.setPen(pen)
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawPath(out)
    p.setBrush(QColor(color))
    p.drawPath(solid)
    p.end()
    return QIcon(pm)


def _align_icon(kind, color="#BBBBBB", size=20):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setPen(QPen(QColor(color), 1.7, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap))
    widths = [14, 9, 12, 7]
    for i, wd in enumerate(widths):
        y = 4.5 + i * 3.6
        x = 3 if kind == "left" else (size - 3 - wd if kind == "right" else (size - wd) / 2)
        p.drawLine(QPointF(x, y), QPointF(x + wd, y))
    p.end()
    return QIcon(pm)


class InlineTextEditor(QPlainTextEdit):
    """Small text box that appears on the page while a text object is typed into. It shows the chosen font, size,
    colour, alignment and line spacing live; `fmt` collects what the user changed in the Edit bar meanwhile."""
    committed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, parent, text, width, height, look, keep=None):
        super().__init__(parent)
        self._done = False
        self.keep = keep                                    # the Edit bar: clicking it must not close the editor
        self.fmt = {}
        self.look = dict(look)
        self.setPlainText(text)
        self.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.document().setDocumentMargin(2)
        self.resize(int(max(width, 60)), int(max(height, look["px"] * 1.6)))
        self.apply_look()
        self.moveCursor(QTextCursor.MoveOperation.End)

    def set_look(self, **kw):
        self.look.update(kw)
        self.apply_look()

    def apply_look(self):
        lk = self.look
        f = QFont(lk.get("family") or "Arial")
        f.setPixelSize(max(8, int(lk["px"])))
        f.setBold(bool(lk.get("bold")))
        f.setItalic(bool(lk.get("italic")))
        self.setFont(f)
        col = QColor(*[int(c * 255) for c in lk.get("color", (0, 0, 0))])
        lum = 0.299 * col.red() + 0.587 * col.green() + 0.114 * col.blue()
        bg = "rgba(255,255,255,235)" if lum < 180 else "rgba(40,40,40,235)"
        self.setStyleSheet(f"QPlainTextEdit{{background:{bg};color:{col.name()};"
                           f"border:2px solid {ACCENT};selection-background-color:{ACCENT};}}")
        opt = self.document().defaultTextOption()
        opt.setAlignment({"center": Qt.AlignmentFlag.AlignHCenter, "right": Qt.AlignmentFlag.AlignRight}.get(
            lk.get("align"), Qt.AlignmentFlag.AlignLeft))
        self.document().setDefaultTextOption(opt)
        sp = lk.get("spacing")
        cur = QTextCursor(self.document())
        cur.select(QTextCursor.SelectionType.Document)
        bf = QTextBlockFormat()
        bf.setLineHeight(float(sp) / 1.2 * 100 if sp else 100.0, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        cur.mergeBlockFormat(bf)

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
        if e.reason() in (Qt.FocusReason.PopupFocusReason, Qt.FocusReason.ActiveWindowFocusReason):
            return                                          # a drop-down / colour dialog opened, or the window lost focus
        fw = QApplication.focusWidget()
        if self.keep is not None and fw is not None and (fw is self.keep or self.keep.isAncestorOf(fw)):
            return                                          # the user went to the Edit bar to change the look
        self._finish(True)


def _swatch_icon(color, size=18):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    if color is None:
        p.setPen(QPen(QColor("#888"), 1))
        p.setBrush(QColor("#fff"))
        p.drawRoundedRect(QRectF(1.5, 1.5, size - 3, size - 3), 3, 3)
        p.setPen(QPen(QColor("#d33"), 1.6))
        p.drawLine(QPointF(3, size - 3), QPointF(size - 3, 3))
    else:
        c = QColor(*[int(v * 255) for v in color])
        p.setPen(QPen(QColor(0, 0, 0, 90), 1))
        p.setBrush(c)
        p.drawRoundedRect(QRectF(1.5, 1.5, size - 3, size - 3), 3, 3)
    p.end()
    return QIcon(pm)


class LazyFontCombo(QComboBox):
    """Font list that is only filled when the drop-down is opened (reading every installed font up front made Edit slow)."""
    _filled = False

    def fill(self):
        if self._filled:
            return
        self._filled = True
        cur = self.currentText()
        self.blockSignals(True)
        self.addItems(font_families())
        i = self.findText(cur, Qt.MatchFlag.MatchFixedString) if cur else -1
        self.setCurrentIndex(i)
        if i < 0:
            self.setEditText(cur)
        self.blockSignals(False)

    def showPopup(self):
        self.fill()
        super().showPopup()


class EditPanel(QFrame):
    """The bar under the toolbar while Edit is on: text font / size / colour / alignment / line spacing,
    ready-made shapes with their line and fill, and page borders."""
    fmt_changed = pyqtSignal(str, object)               # family / size / bold / italic / color / align / spacing
    shape_style_changed = pyqtSignal(str, object)       # stroke / fill / width / dash
    tool_changed = pyqtSignal(object)                   # shape kind being drawn, or None
    border_requested = pyqtSignal()
    arrange = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setObjectName("EditPanel")
        self._text_color = (0.0, 0.0, 0.0)
        self._stroke, self._fill = (0.0, 0.0, 0.0), None
        self._align = None
        self._kind = "rect"
        self.setStyleSheet(
            f"#EditPanel{{background:{T['PANEL']};border-bottom:1px solid {T['LINE']};}}"
            f"#EditPanel QLabel{{color:{T['MUTED']};}}"
            f"#EditPanel QToolButton{{background:transparent;border:1px solid transparent;border-radius:4px;"
            f"padding:2px 6px;color:{T['TEXT']};min-height:22px;}}"
            f"#EditPanel QToolButton:hover{{background:{T['HOVER']};}}"
            f"#EditPanel QToolButton:checked{{background:{T['SEL']};border-color:{T['ACC']};color:{T['SEL_TEXT']};}}"
            f"#EditPanel QToolButton::menu-indicator{{image:none;width:0px;}}")
        col = QVBoxLayout(self)
        col.setContentsMargins(8, 4, 8, 4)
        col.setSpacing(3)

        # ---- row 1: text ----
        r1 = QHBoxLayout()
        r1.setSpacing(4)
        r1.addWidget(QLabel("Text"))
        self.font_box = LazyFontCombo()
        self.font_box.setEditable(True)
        self.font_box.setMinimumWidth(190)
        self.font_box.setMaximumWidth(240)
        self.font_box.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.font_box.setCurrentIndex(-1)
        self.font_box.setToolTip("Font of the selected text (blank = the text's own font)")
        self.font_box.activated.connect(lambda _i: self._emit_family())
        self.font_box.lineEdit().editingFinished.connect(self._emit_family_typed)
        r1.addWidget(self.font_box)
        self.size_box = QDoubleSpinBox()
        self.size_box.setRange(1, 500)
        self.size_box.setDecimals(1)
        self.size_box.setSingleStep(0.5)
        self.size_box.setValue(12)
        self.size_box.setSuffix(" pt")
        self.size_box.setToolTip("Font size")
        self.size_box.valueChanged.connect(lambda v: self.fmt_changed.emit("size", float(v)))
        r1.addWidget(self.size_box)
        self.b_bold, self.b_italic = self._tool("B", "Bold", True), self._tool("I", "Italic", True)
        f = self.b_bold.font(); f.setBold(True); self.b_bold.setFont(f)
        f = self.b_italic.font(); f.setItalic(True); self.b_italic.setFont(f)
        self.b_bold.toggled.connect(lambda v: self.fmt_changed.emit("bold", bool(v)))
        self.b_italic.toggled.connect(lambda v: self.fmt_changed.emit("italic", bool(v)))
        self.b_color = self._tool("", "Text colour", False)
        self.b_color.setIcon(_swatch_icon(self._text_color))
        self.b_color.clicked.connect(self._pick_text_color)
        for w in (self.b_bold, self.b_italic, self.b_color):
            r1.addWidget(w)
        r1.addWidget(self._sep())
        self.b_align = {}
        for kind, tip in (("left", "Align left"), ("center", "Align centre"), ("right", "Align right")):
            b = self._tool("", tip + " (click again for the text's own alignment)", True)
            b.setIcon(_align_icon(kind, T["TEXT"]))
            b.clicked.connect(lambda _c=False, k=kind: self._align_clicked(k))
            self.b_align[kind] = b
            r1.addWidget(b)
        r1.addWidget(self._sep())
        r1.addWidget(QLabel("Line spacing"))
        self.sp_box = QDoubleSpinBox()
        self.sp_box.setRange(0.5, 4.0)
        self.sp_box.setSingleStep(0.05)
        self.sp_box.setDecimals(2)
        self.sp_box.setValue(1.2)
        self.sp_box.setToolTip("Space between the lines, as a multiple of the font size")
        self.sp_box.valueChanged.connect(lambda v: self.fmt_changed.emit("spacing", float(v)))
        r1.addWidget(self.sp_box)
        r1.addWidget(self._sep())
        self.b_arr = self._tool("Align lines ▾", "Line up the selected objects with each other", False)
        m = QMenu(self.b_arr)
        for key, label in (("left", "Left edges"), ("hcenter", "Centres (horizontal)"), ("right", "Right edges"),
                           ("top", "Top edges"), ("vcenter", "Centres (vertical)"), ("bottom", "Bottom edges")):
            m.addAction(label, lambda k=key: self.arrange.emit(k))
        self.b_arr.setMenu(m)
        self.b_arr.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        r1.addWidget(self.b_arr)
        r1.addStretch(1)
        col.addLayout(r1)

        # ---- row 2: shapes and borders ----
        r2 = QHBoxLayout()
        r2.setSpacing(4)
        r2.addWidget(QLabel("Shapes"))
        self.b_shape = self._tool("Rectangle", "Draw a shape: click here, then drag on the page (Esc to cancel)", True)
        self.b_shape.setIcon(_shape_icon("rect", T["TEXT"]))
        self.b_shape.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.b_shape.toggled.connect(lambda on: self.tool_changed.emit(self._kind if on else None))
        r2.addWidget(self.b_shape)
        self.b_shape_menu = self._tool("▾", "Choose a shape", False)
        sm = QMenu(self.b_shape_menu)
        for kind, label in SHAPES:
            a = sm.addAction(_shape_icon(kind, T["TEXT"]), label)
            a.triggered.connect(lambda _c=False, k=kind, lb=label: self.choose_shape(k, lb))
        self.b_shape_menu.setMenu(sm)
        self.b_shape_menu.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        r2.addWidget(self.b_shape_menu)
        r2.addWidget(self._sep())
        r2.addWidget(QLabel("Line"))
        self.b_stroke = self._tool("", "Outline colour", False)
        self.b_stroke.setIcon(_swatch_icon(self._stroke))
        self.b_stroke.clicked.connect(self._pick_stroke)
        r2.addWidget(self.b_stroke)
        self.w_box = QDoubleSpinBox()
        self.w_box.setRange(0.25, 60)
        self.w_box.setSingleStep(0.25)
        self.w_box.setValue(2.0)
        self.w_box.setSuffix(" pt")
        self.w_box.setToolTip("Outline thickness")
        self.w_box.valueChanged.connect(lambda v: self.shape_style_changed.emit("width", float(v)))
        r2.addWidget(self.w_box)
        self.dash_box = QComboBox()
        for key, label in DASH_STYLES:
            self.dash_box.addItem(label, key)
        self.dash_box.setToolTip("Outline style")
        self.dash_box.activated.connect(lambda _i: self.shape_style_changed.emit("dash", self.dash_box.currentData()))
        r2.addWidget(self.dash_box)
        r2.addWidget(QLabel("Fill"))
        self.b_fill = self._tool("", "Fill colour (click the ✕ button for no fill)", False)
        self.b_fill.setIcon(_swatch_icon(self._fill))
        self.b_fill.clicked.connect(self._pick_fill)
        r2.addWidget(self.b_fill)
        self.b_nofill = self._tool("✕", "No fill", False)
        self.b_nofill.clicked.connect(self._clear_fill)
        r2.addWidget(self.b_nofill)
        r2.addWidget(self._sep())
        self.b_border = self._tool("▭ Page border…", "Add a frame around one page or all pages", False)
        self.b_border.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        self.b_border.clicked.connect(self.border_requested.emit)
        r2.addWidget(self.b_border)
        r2.addStretch(1)
        col.addLayout(r2)

    # ---- small helpers ----
    def _tool(self, text, tip, checkable):
        b = QToolButton()
        b.setText(text)
        b.setToolTip(tip)
        b.setCheckable(checkable)
        b.setFocusPolicy(Qt.FocusPolicy.NoFocus)            # keep the keyboard focus (and the open text box) where it is
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        return b

    def _sep(self):
        s = QFrame()
        s.setFrameShape(QFrame.Shape.VLine)
        s.setStyleSheet(f"color:{T['LINE2']};")
        s.setFixedHeight(20)
        return s

    def _ask_color(self, cur, title):
        c0 = QColor(*[int(v * 255) for v in cur]) if cur else QColor("#ffcc00")
        c = QColorDialog.getColor(c0, self.window(), title)
        return (c.redF(), c.greenF(), c.blueF()) if c.isValid() else None

    # ---- text ----
    def _emit_family(self):
        t = self.font_box.currentText().strip()
        if t:
            self.fmt_changed.emit("family", t)

    def _emit_family_typed(self):
        t = self.font_box.currentText().strip()
        if t and any(t.lower() == f.lower() for f in font_families()):
            self.fmt_changed.emit("family", t)

    def _pick_text_color(self):
        c = self._ask_color(self._text_color, "Text colour")
        if c:
            self._text_color = c
            self.b_color.setIcon(_swatch_icon(c))
            self.fmt_changed.emit("color", c)

    def _align_clicked(self, kind):
        on = self.b_align[kind].isChecked()
        for k, b in self.b_align.items():
            if k != kind:
                b.blockSignals(True)
                b.setChecked(False)
                b.blockSignals(False)
        self._align = kind if on else None
        self.fmt_changed.emit("align", self._align)

    def set_text_values(self, family, size, bold, italic, color, align, spacing):
        """Show the properties of the selected text without announcing them as changes."""
        for w in (self.font_box, self.size_box, self.b_bold, self.b_italic, self.sp_box, *self.b_align.values()):
            w.blockSignals(True)
        i = self.font_box.findText(family or "", Qt.MatchFlag.MatchFixedString) if (family and self.font_box._filled) else -1
        self.font_box.setCurrentIndex(i)
        if i < 0:
            self.font_box.setEditText(family or "")
        self.size_box.setValue(float(size))
        self.b_bold.setChecked(bool(bold))
        self.b_italic.setChecked(bool(italic))
        self._text_color = color
        self.b_color.setIcon(_swatch_icon(color))
        self._align = align
        for k, b in self.b_align.items():
            b.setChecked(k == align)
        if spacing:
            self.sp_box.setValue(float(spacing))
        for w in (self.font_box, self.size_box, self.b_bold, self.b_italic, self.sp_box, *self.b_align.values()):
            w.blockSignals(False)

    def current_text_fmt(self):
        """Everything shown in the text row (used for brand-new text)."""
        fam = self.font_box.currentText().strip()
        fmt = {"size": float(self.size_box.value()), "bold": self.b_bold.isChecked(),
               "italic": self.b_italic.isChecked(), "color": self._text_color, "spacing": float(self.sp_box.value())}
        if fam and any(fam.lower() == f.lower() for f in font_families()):
            fmt["family"] = fam
        if self._align:
            fmt["align"] = self._align
        return fmt

    # ---- shapes ----
    def choose_shape(self, kind, label):
        self._kind = kind
        self.b_shape.setText(label)
        self.b_shape.setIcon(_shape_icon(kind, T["TEXT"]))
        if self.b_shape.isChecked():
            self.tool_changed.emit(kind)
        else:
            self.b_shape.setChecked(True)

    def set_tool(self, kind):
        self.b_shape.blockSignals(True)
        self.b_shape.setChecked(kind is not None)
        self.b_shape.blockSignals(False)

    def _pick_stroke(self):
        c = self._ask_color(self._stroke, "Outline colour")
        if c:
            self._stroke = c
            self.b_stroke.setIcon(_swatch_icon(c))
            self.shape_style_changed.emit("stroke", c)

    def _pick_fill(self):
        c = self._ask_color(self._fill, "Fill colour")
        if c:
            self._fill = c
            self.b_fill.setIcon(_swatch_icon(c))
            self.shape_style_changed.emit("fill", c)

    def _clear_fill(self):
        self._fill = None
        self.b_fill.setIcon(_swatch_icon(None))
        self.shape_style_changed.emit("fill", None)

    def set_shape_values(self, stroke, fill, width, dash):
        for w in (self.w_box, self.dash_box):
            w.blockSignals(True)
        self._stroke, self._fill = stroke, fill
        self.b_stroke.setIcon(_swatch_icon(stroke))
        self.b_fill.setIcon(_swatch_icon(fill))
        if width:
            self.w_box.setValue(float(width))
        i = self.dash_box.findData(dash or "solid")
        self.dash_box.setCurrentIndex(max(0, i))
        for w in (self.w_box, self.dash_box):
            w.blockSignals(False)

    def shape_style(self):
        return {"stroke": self._stroke, "fill": self._fill, "width": float(self.w_box.value()),
                "dash": self.dash_box.currentData()}


class BorderDialog(DarkTitleBarMixin, QDialog):
    """Frame around one page, a range of pages, or every page."""
    def __init__(self, parent, n_pages, cur_page):
        super().__init__(parent)
        self.setWindowTitle("Page border")
        self.setModal(True)
        self.n, self.cur = n_pages, cur_page
        self.color = (0.0, 0.0, 0.0)
        lay = QVBoxLayout(self)
        g = QGroupBox("Apply to")
        gl = QVBoxLayout(g)
        self.r_cur = QRadioButton(f"This page (page {cur_page + 1})")
        self.r_all = QRadioButton(f"All pages ({n_pages})")
        self.r_rng = QRadioButton("Pages:")
        self.r_cur.setChecked(True)
        self.rng = QLineEdit()
        self.rng.setPlaceholderText("for example 1-3, 5, 8-")
        self.rng.textEdited.connect(lambda _t: self.r_rng.setChecked(True))
        rr = QHBoxLayout()
        rr.addWidget(self.r_rng)
        rr.addWidget(self.rng, 1)
        gl.addWidget(self.r_cur)
        gl.addWidget(self.r_all)
        gl.addLayout(rr)
        lay.addWidget(g)
        form = QFormLayout()
        self.b_col = QPushButton()
        self.b_col.setIcon(_swatch_icon(self.color, 16))
        self.b_col.setText("Colour…")
        self.b_col.clicked.connect(self._pick)
        form.addRow("Colour", self.b_col)
        self.w = QDoubleSpinBox(); self.w.setRange(0.25, 40); self.w.setValue(2.0); self.w.setSingleStep(0.25); self.w.setSuffix(" pt")
        self.m = QDoubleSpinBox(); self.m.setRange(0, 200); self.m.setValue(24); self.m.setSuffix(" pt")
        self.rad = QDoubleSpinBox(); self.rad.setRange(0, 80); self.rad.setValue(0); self.rad.setSuffix(" pt")
        self.st = QComboBox()
        for key, label in (("solid", "Solid"), ("dashed", "Dashed"), ("dotted", "Dotted"), ("double", "Double")):
            self.st.addItem(label, key)
        form.addRow("Thickness", self.w)
        form.addRow("Distance from the page edge", self.m)
        form.addRow("Rounded corners", self.rad)
        form.addRow("Style", self.st)
        lay.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.pages = []

    def _pick(self):
        c0 = QColor(*[int(v * 255) for v in self.color])
        c = QColorDialog.getColor(c0, self, "Border colour")
        if c.isValid():
            self.color = (c.redF(), c.greenF(), c.blueF())
            self.b_col.setIcon(_swatch_icon(self.color, 16))

    def _ok(self):
        if self.r_all.isChecked():
            self.pages = list(range(self.n))
        elif self.r_rng.isChecked():
            self.pages = parse_range(self.rng.text(), self.n) or []
            if not self.pages:
                QMessageBox.information(self, "Page border", "Type pages like 1-3, 5, 8-")
                return
        else:
            self.pages = [self.cur]
        self.accept()

    def values(self):
        return dict(pages=self.pages, color=self.color, width=self.w.value(), margin=self.m.value(),
                    style=self.st.currentData(), radius=self.rad.value())


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


THUMB_STRIP = 44            # px of every page thumbnail that stays visible in the stacked strip (its top edge)
THUMB_PAD = 5               # px above each thumbnail inside its strip slot
THUMB_BOX = QSize(118, 160) # largest size of a thumbnail


class _NoPaintDelegate(QStyledItemDelegate):
    """ThumbList paints its own overlapping pages (ThumbList.paintEvent); items only supply geometry."""
    def paint(self, painter, option, index):
        pass

    def sizeHint(self, option, index):
        s = index.data(Qt.ItemDataRole.SizeHintRole)
        return s if isinstance(s, QSize) else QSize(140, THUMB_STRIP)


def _ease(p):
    return p * p * (3.0 - 2.0 * p)


class ThumbList(QListWidget):
    """Thumbnail strip: pages are stacked on top of each other so only the top edge of each shows; the page
    under the mouse slides open smoothly. Accepts PDF/image files dropped between thumbnails."""
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
        self.setSpacing(0)
        self.setFrameShape(QFrame.Shape.NoFrame)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setItemDelegate(_NoPaintDelegate(self))
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.verticalScrollBar().setSingleStep(24)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.itemSelectionChanged.connect(self.viewport().update)
        self.export_drag_index = None
        self._pm_cache = {}                          # QIcon cache key -> scaled page pixmap
        self._prog = {}                              # row -> how far its page has slid open (0..1)
        self._hover = None                           # row the mouse is on
        self._pinned = None                          # page being viewed in the main area: stays slid open
        self._show_sel = False                       # blue frame only after the user clicked a page
        self._anim = QTimer(self)
        self._anim.setInterval(16)
        self._anim.timeout.connect(self._tick)
        self._line = QFrame(self.viewport())
        self._line.setStyleSheet(f"background:{ACCENT};border-radius:2px")
        self._line.hide()

    def scrollContentsBy(self, dx, dy):
        super().scrollContentsBy(dx, dy)
        self._hover_from_cursor()                    # the wheel moves pages under a still mouse
        self.layout_changed.emit()

    # ---------- stacked pages: geometry, hover animation, painting ----------
    def _pm(self, row):
        it = self.item(row)
        if it is None:
            return None
        ic = it.icon()
        key = ic.cacheKey()
        pm = self._pm_cache.get(key)
        if pm is None:
            if len(self._pm_cache) > 600:
                self._pm_cache.clear()
            pm = ic.pixmap(THUMB_BOX)
            self._pm_cache[key] = pm
        return pm

    def thumb_rect(self, row):
        """Full rectangle of the page thumbnail `row` (viewport coordinates); most of it is covered by the next page."""
        it = self.item(row)
        pm = self._pm(row)
        if it is None or pm is None or pm.isNull():
            return None
        sz = pm.deviceIndependentSize()
        vr = self.visualItemRect(it)
        return QRectF((self.viewport().width() - sz.width()) / 2.0, vr.top() + THUMB_PAD, sz.width(), sz.height())

    def _reveal_rect(self, row, prog):
        r = self.thumb_rect(row)
        if r is None:
            return None
        base = min(r.height(), THUMB_STRIP - THUMB_PAD)
        return QRectF(r.x(), r.y(), r.width(), base + (r.height() - base) * _ease(prog))

    def covered_by_reveal(self, row, rect):
        """True if a page that is sliding open lies over `rect` (a QRectF) of the page `row`."""
        for h, prog in self._prog.items():
            if h < row and prog > 0:
                rr = self._reveal_rect(h, prog)
                if rr is not None and rr.intersects(rect):
                    return True
        return False

    def indexAt(self, pos):
        h = getattr(self, "_hover", None)               # Qt may call this while the widget is still being built
        if h is not None and self._prog.get(h, 0) > 0:
            rr = self._reveal_rect(h, self._prog[h])
            if rr is not None and rr.contains(QPointF(pos)):
                return self.model().index(h, 0)         # the opened page owns the area it covers
        return super().indexAt(pos)

    def _target(self):
        """The page that should be open: the one under the mouse, else the one being viewed."""
        return self._hover if self._hover is not None else self._pinned

    def _retarget(self):
        t = self._target()
        if t is not None:
            self._prog.setdefault(t, 0.0)
        if not self._anim.isActive():
            self._anim.start()
        self.layout_changed.emit()                   # show / hide the trash icon right away

    def _set_hover(self, row):
        if row == self._hover:
            return
        self._hover = row
        self._retarget()

    def set_pinned(self, row):
        """Keep page `row` slid open (the page shown in the main view) and scroll the strip so all of it is visible."""
        if row is not None and not (0 <= row < self.count()):
            row = None
        self._pinned = row
        self._retarget()
        if row is None:
            return
        self.scrollTo(self.model().index(row, 0), QAbstractItemView.ScrollHint.EnsureVisible)
        r = self.thumb_rect(row)
        if r is not None:
            over = r.bottom() + 6 - self.viewport().height()
            if over > 0:
                sb = self.verticalScrollBar()
                sb.setValue(sb.value() + int(over + 0.999))

    def show_selection(self, on=True):
        self._show_sel = on
        self.viewport().update()

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if self.indexAt(e.position().toPoint()).isValid():
            self.show_selection(True)

    def _hover_from(self, pos):
        row = self.indexAt(pos).row()
        self._set_hover(row if row >= 0 else None)

    def _hover_from_cursor(self):
        vp = self.viewport()
        if vp.underMouse():
            self._hover_from(vp.mapFromGlobal(QCursor.pos()))

    def _tick(self):
        target = self._target()
        for row, p in list(self._prog.items()):
            tgt = 1.0 if row == target else 0.0
            p = min(tgt, p + 0.14) if p < tgt else max(tgt, p - 0.14)
            if p <= 0.0 and tgt == 0.0:
                del self._prog[row]
            else:
                self._prog[row] = p
        if all(row == target and p >= 1.0 for row, p in self._prog.items()):
            self._anim.stop()
        self.viewport().update()
        self.layout_changed.emit()

    def mouseMoveEvent(self, e):
        super().mouseMoveEvent(e)
        if not (e.buttons() & Qt.MouseButton.LeftButton):
            self._hover_from(e.position().toPoint())

    def leaveEvent(self, e):
        self._set_hover(None)
        super().leaveEvent(e)

    def _draw_thumb(self, p, row, shown_h, prog):
        full = self.thumb_rect(row)
        pm = self._pm(row)
        if full is None:
            return
        shown = QRectF(full.x(), full.y(), full.width(), min(shown_h, full.height()))
        a_top = 70 if T["DARK"] else 45
        g = QLinearGradient(0, shown.top() - 7, 0, shown.top())          # this page casts a shadow on the one below it
        g.setColorAt(0, QColor(0, 0, 0, 0))
        g.setColorAt(1, QColor(0, 0, 0, a_top))
        p.fillRect(QRectF(shown.x() - 1, shown.top() - 7, shown.width() + 2, 7), QBrush(g))
        if prog > 0:
            g2 = QLinearGradient(0, shown.bottom(), 0, shown.bottom() + 12)
            g2.setColorAt(0, QColor(0, 0, 0, int(90 * prog)))
            g2.setColorAt(1, QColor(0, 0, 0, 0))
            p.fillRect(QRectF(shown.x() - 1, shown.bottom(), shown.width() + 2, 12), QBrush(g2))
        p.save()
        p.setClipRect(shown, Qt.ClipOperation.IntersectClip)
        p.drawPixmap(full.topLeft(), pm)
        selected = self._show_sel and self.selectionModel().isSelected(self.model().index(row, 0))
        if selected:
            p.setPen(QPen(QColor(ACCENT), 2))
            p.setBrush(Qt.BrushStyle.NoBrush)
            p.drawRect(full.adjusted(1, 1, -1, -1))
        f = p.font()
        f.setPixelSize(10)
        f.setBold(True)
        p.setFont(f)
        txt = str(row + 1)
        w = max(18, p.fontMetrics().horizontalAdvance(txt) + 10)
        badge = QRectF(full.right() - w - 4, full.top() + 4, w, 16)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(ACCENT) if selected else QColor(0, 0, 0, 160))
        p.drawRoundedRect(badge, 8, 8)
        p.setPen(QColor("white"))
        p.drawText(badge, Qt.AlignmentFlag.AlignCenter, txt)
        p.restore()

    def paintEvent(self, e):
        super().paintEvent(e)                        # background only: the delegate paints nothing
        n = self.count()
        if not n:
            return
        p = QPainter(self.viewport())
        p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
        p.setClipRegion(e.region())
        vh = self.viewport().height()
        first = max(0, (self.verticalScrollBar().value() - 6) // THUMB_STRIP - 5)
        for row in range(first, n):                  # stacked: every page lies over the one before it
            r = self.thumb_rect(row)
            if r is None:
                continue
            if r.top() - 8 > vh:
                break
            self._draw_thumb(p, row, r.height(), 0.0)
        for row, prog in sorted(self._prog.items(), key=lambda kv: (kv[0] == self._hover, kv[0] == self._pinned, kv[1])):
            rr = self._reveal_rect(row, prog)        # pages sliding open are drawn over the stack
            if rr is not None and rr.bottom() > 0 and rr.top() < vh:
                self._draw_thumb(p, row, rr.height(), prog)
        p.end()

    def updateGeometries(self):
        super().updateGeometries()
        self.layout_changed.emit()

    def contextMenuEvent(self, e):
        idx = self.indexAt(e.pos())
        if idx.isValid():
            self.context_requested.emit(idx.row(), e.globalPos())
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
        self._set_hover(None)
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
        self.title, self.suggest = None, None               # unsaved in-memory documents (Combine): tab name, proposed file name
        self.zoom, self.undo_stack, self.dirty = 1.4, [], False
        self.labels, self.ys, self.hs, self.imgs = [], [], [], {}
        self._armed_export_format = None
        self.edit_mode = False
        self._objs = {}                                     # page index -> editable objects (text / image / drawing)
        self.edit_sel, self.edit_page = [], None            # selected objects (all on one page)
        self._drag, self._ghost, self._hover, self._editor = None, None, None, None
        self.edit_tool = None                               # ready-made shape being drawn (None = select / move)
        self._fonts_added = False                           # fonts were written: shrink them when saving
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

        self._panel = None                                  # the Edit bar is built the first time Edit is switched on

        split = QSplitter()
        split.setHandleWidth(4)
        split.setChildrenCollapsible(False)
        split.addWidget(self.thumbs)
        split.addWidget(self.scroll)
        split.setStretchFactor(0, 0)
        split.setStretchFactor(1, 1)
        split.setSizes([self.thumbs.width(), 1100])
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self._lay = lay
        lay.addWidget(split, 1)

    # ---------- state ----------
    def set_dirty(self, v):
        self.dirty = v
        self.changed.emit()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.update_visible()

    def showEvent(self, e):
        super().showEvent(e)
        if self.edit_mode and self._panel is None:
            self.panel.setVisible(True)                     # built when the tab is actually shown, not for every open tab
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

    def load_doc(self, d, title, suggest):
        """Show a document that only exists in memory. It is dirty, so closing asks first, and Save asks for a file."""
        self.doc, self.path, self.cur, self.undo_stack, self.dirty = d, None, 0, [], True
        self.title, self.suggest = title, suggest
        self._center_h = True
        self.refresh_thumbs()

    def save(self):
        if not self.doc:
            return False
        if not self.path:
            return self.save_as()
        return self._write(self.path)

    def save_as(self):
        if not self.doc:
            return False
        p, _ = QFileDialog.getSaveFileName(self.main, "Save PDF", self.path or self.suggest or "", "PDF (*.pdf)")
        return self._write(p) if p else False

    def _pdf_bytes(self):
        """The PDF as bytes. Fonts written by Edit are whole font files, so when fontTools is installed they are
        cut down to the letters actually used (a copy is made first; on any problem the plain file is saved)."""
        if self._fonts_added:
            try:
                import fontTools                                    # noqa: F401  (optional)
                tmp = fitz.open("pdf", self.doc.tobytes())
                tmp.subset_fonts()
                data = tmp.tobytes(garbage=4, deflate=True)
                chk = fitz.open("pdf", data)
                if len(chk) == len(self.doc):
                    return data
            except Exception:
                pass
        return self.doc.tobytes(garbage=3, deflate=True)

    def _write(self, p):
        try:
            data = self._pdf_bytes()
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

    def _thumb_placeholder(self, i):
        r = self.doc[i].rect
        out = QPixmap(max(1, math.ceil(r.width * 0.32)), max(1, math.ceil(r.height * 0.32)))
        out.fill(QColor("white"))
        p = QPainter(out)
        p.setPen(QPen(QColor(*T["THUMB_RIM"]), 1))
        p.drawRect(QRectF(0.5, 0.5, out.width() - 1, out.height() - 1))
        p.end()
        ic = QIcon()
        for mode in (QIcon.Mode.Normal, QIcon.Mode.Active, QIcon.Mode.Selected):
            ic.addPixmap(out, mode)
        return ic

    def _thumb_tick(self, gen):
        """Draw the missing thumbnails a few at a time so a big PDF opens at once and the window stays responsive."""
        if gen != self._thumb_gen or not self.doc:
            return
        for _ in range(3):
            if not self._thumb_todo:
                return
            i = self._thumb_todo.pop()
            it = self.thumbs.item(i)
            if it is not None and i < len(self.doc):
                it.setIcon(self.thumb_icon(i))
        self.thumbs.viewport().update()
        if self._thumb_todo:
            QTimer.singleShot(1, lambda g=gen: self._thumb_tick(g))

    def refresh_thumbs(self, keep=None, screen_off=None):
        keep = self.cur if keep is None else keep
        self.thumbs.blockSignals(True)
        self.thumbs.clear()
        n_pages = len(self.doc)
        near = set(range(max(0, keep - 6), min(n_pages, keep + 7))) if n_pages > 24 else set(range(n_pages))
        for i in range(n_pages):                              # big files: only the pages around the current one now, the rest in the background
            item = QListWidgetItem(self.thumb_icon(i) if i in near else self._thumb_placeholder(i), str(i + 1))
            item.setSizeHint(QSize(140, THUMB_STRIP))         # stacked pages: each page owns only its top strip
            item.setToolTip(f"Drag page {i + 1} to a folder or the Desktop to export it. "
                            "Select several pages to export them together.")
            self.thumbs.addItem(item)
        last = self.thumbs.item(self.thumbs.count() - 1)
        if last is not None:                                  # the last page has nothing over it: give it full height
            lh = last.icon().pixmap(THUMB_BOX).deviceIndependentSize().height()
            last.setSizeHint(QSize(140, max(THUMB_STRIP, int(lh) + THUMB_PAD + 8)))
        self.cur = max(0, min(keep, len(self.doc) - 1))
        self._thumb_gen = getattr(self, "_thumb_gen", 0) + 1
        self._thumb_todo = sorted((i for i in range(n_pages) if i not in near), key=lambda i: abs(i - self.cur), reverse=True)
        if self._thumb_todo:
            QTimer.singleShot(20, lambda g=self._thumb_gen: self._thumb_tick(g))
        self.thumbs.setCurrentRow(self.cur)
        self.thumbs.blockSignals(False)
        self.thumbs.show_selection(False)
        self.thumbs._prog.clear()
        self.thumbs.set_pinned(self.cur)
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
        """Each trash button sits in the top-left corner of its page thumbnail (only the page under the mouse shows one)."""
        vp = self.thumbs.viewport()
        hov = self.thumbs._hover
        for i, b in enumerate(self.del_btns):
            if i != hov:
                if b.isVisible():
                    b.hide()
                continue
            r = self.thumbs.thumb_rect(i) if self.thumbs.item(i) is not None else None
            if r is None or r.top() + 30 < 0 or r.top() > vp.height():
                b.hide()
                continue
            spot = QRectF(r.left() + 4, r.top() + 4, 24, 24)   # inside the thumbnail, top-left corner
            if self.thumbs.covered_by_reveal(i, spot):         # another page is sliding open over this corner
                b.hide()
                continue
            b.move(int(spot.left()), int(spot.top()))
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
        first = max(0, bisect.bisect_right(self.ys, lo) - 1)         # pages are stacked top to bottom: binary search
        last = min(len(self.ys), bisect.bisect_right(self.ys, hi))
        want = [i for i in range(first, last) if self.ys[i] + self.hs[i] >= lo and self.ys[i] <= hi]
        keep = set(want)
        for i in [k for k in self.imgs if k not in keep]:             # far away: free memory
            del self.imgs[i]
            self.labels[i].clear()
        for i in want:
            self.render_page(i)

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
            self.thumbs.show_selection(False)               # scrolling is not a click: no blue frame
            self.thumbs.set_pinned(i)                       # the page being viewed slides open in the strip
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
        self.thumbs.set_pinned(i)
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
        self.thumbs.show_selection(True)

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
            self.thumbs.show_selection(True)
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
        self.edit_tool = None
        if self._editor is not None:
            ed, self._editor = self._editor, None
            ed._done = True
            ed.hide()
            ed.deleteLater()

    @property
    def panel(self):
        if self._panel is None:
            p = self._panel = EditPanel()
            p.hide()
            p.fmt_changed.connect(self._on_fmt)
            p.shape_style_changed.connect(self._on_shape_style)
            p.tool_changed.connect(self._set_tool)
            p.border_requested.connect(self._on_border)
            p.arrange.connect(self._on_arrange)
            self._lay.insertWidget(0, p)
        return self._panel

    def set_edit(self, on):
        self.edit_mode = on
        if self._panel is not None:
            self._panel.setVisible(on)
        elif on and self.isVisible():
            self.panel.setVisible(True)
        if on:
            self._clear_armed_export()
        else:
            self._commit_editor()
            self.edit_sel, self.edit_page, self._drag, self._ghost, self._hover = [], None, None, None, None
            self._set_tool(None)
        for lab in self.labels:
            lab.setCursor(Qt.CursorShape.CrossCursor if (on and self.edit_tool) else Qt.CursorShape.ArrowCursor)
            lab.export_drag_enabled = not on
            lab.setFocusPolicy(Qt.FocusPolicy.StrongFocus if on else Qt.FocusPolicy.NoFocus)
            lab.update()

    def on_click(self, idx, pos):                           # old click-to-edit-a-line path; Edit mode now handles input itself
        return

    # -- the Edit bar --
    def _set_tool(self, kind):
        self.edit_tool = kind
        if self._panel is not None:
            self._panel.set_tool(kind)
        if kind:
            self.edit_sel = []
            self._commit_editor()
        for lab in self.labels:
            lab.setCursor(Qt.CursorShape.CrossCursor if kind else Qt.CursorShape.ArrowCursor)
            lab.update()

    @staticmethod
    def _first_span(line):
        return next((s for s in line["spans"] if s["text"].strip()), line["spans"][0])

    def _text_values(self, obj):
        """Font / size / colour / alignment / line spacing of a text object, for the Edit bar."""
        lines = obj["lines"]
        sp = self._first_span(lines[0])
        fam_key, bold, italic = _pdf_font_parts(sp["font"], sp["flags"])
        _system_fonts()
        fam = _SYS_FAMILY_NAMES.get(fam_key) or re.sub(r"^[A-Z]{6}\+", "", sp["font"] or "")
        align, spacing = None, None
        if len(lines) > 1:
            bs = [fitz.Rect(l["bbox"]) for l in lines]
            if all(abs(b.x0 - bs[0].x0) < 1.5 for b in bs):
                align = "left"
            elif all(abs(b.x1 - bs[0].x1) < 1.5 for b in bs):
                align = "right"
            elif all(abs((b.x0 + b.x1) / 2 - (bs[0].x0 + bs[0].x1) / 2) < 1.5 for b in bs):
                align = "center"
            step = self._first_span(lines[1])["origin"][1] - sp["origin"][1]
            if step > 0 and sp["size"] > 0:
                spacing = round(step / sp["size"] / 0.05) * 0.05
        return fam, sp["size"], bool(sp["flags"] & 16), bool(sp["flags"] & 2), _rgb_from_int(sp["color"]), align, spacing

    def _sel_changed(self):
        """Show the properties of what is selected in the Edit bar."""
        texts = [o for o in self.edit_sel if o["kind"] == "text"]
        draws = [o for o in self.edit_sel if o["kind"] == "draw"]
        try:
            if texts:
                self.panel.set_text_values(*self._text_values(texts[0]))
            if draws and draws[0]["paths"]:
                p = draws[0]["paths"][0]
                d = (p.get("dashes") or "").strip()
                self.panel.set_shape_values(p.get("color"), p.get("fill"), p.get("width") or 1.0,
                                            "solid" if d in ("", "[] 0") else ("dotted" if d.startswith("[0") else "dashed"))
        except Exception:
            pass

    def _on_fmt(self, key, value):
        """A text property was changed in the Edit bar."""
        ed = self._editor
        if ed is not None:
            ed.fmt[key] = value
            lk = {"family": "family", "bold": "bold", "italic": "italic", "color": "color", "align": "align",
                  "spacing": "spacing"}
            if key == "size":
                ed.set_look(px=max(8, value * self.zoom))
            elif key in lk:
                ed.set_look(**{lk[key]: value})
            return
        if value is None or self.edit_page is None:
            return
        if any(o["kind"] == "text" for o in self.edit_sel):
            self._restyle_text(self.edit_page, {key: value})

    def _on_shape_style(self, key, value):
        if self.edit_tool or self.edit_page is None:
            return
        if any(o["kind"] == "draw" for o in self.edit_sel):
            ov = {"stroke": "color", "fill": "fill", "width": "width", "dash": "dash"}[key]
            self._restyle_shapes(self.edit_page, {ov: value})

    def _on_border(self):
        if not self.doc:
            return
        dlg = BorderDialog(self.main, len(self.doc), self.cur)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        v = dlg.values()
        self.snapshot()
        try:
            for i in v["pages"]:
                _draw_border(self.doc[i], v["color"], v["width"], v["margin"], v["style"], v["radius"])
        except Exception as e:
            QMessageBox.warning(self.main, "Border failed", str(e))
        self._objs = {}
        for i in v["pages"]:
            self.render_page(i, force=True)
            self.thumbs.item(i).setIcon(self.thumb_icon(i))
        self.thumbs.viewport().update()
        for lab in self.labels:
            lab.update()

    def _on_arrange(self, mode):
        objs = list(self.edit_sel)
        if len(objs) < 2 or self.edit_page is None:
            QMessageBox.information(self.main, "PDF Studio", "Select two or more objects (Shift+click, or drag a box around them) first.")
            return
        u = self._sel_rect()
        items = []
        for o in objs:
            r, dx, dy = o["rect"], 0.0, 0.0
            if mode == "left":
                dx = u.x0 - r.x0
            elif mode == "right":
                dx = u.x1 - r.x1
            elif mode == "hcenter":
                dx = (u.x0 + u.x1) / 2 - (r.x0 + r.x1) / 2
            elif mode == "top":
                dy = u.y0 - r.y0
            elif mode == "bottom":
                dy = u.y1 - r.y1
            else:
                dy = (u.y0 + u.y1) / 2 - (r.y0 + r.y1) / 2
            items.append((o, (0.0, 0.0), 1.0, 1.0, dx, dy))
        if any(abs(it[4]) > 1e-3 or abs(it[5]) > 1e-3 for it in items):
            self._transform_each(self.edit_page, items)

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
            pr = page.rect
            parea = max(1.0, pr.width * pr.height)
            big = [p for p in paths if p["rect"].width * p["rect"].height >= 0.5 * parea]
            small = [p for p in paths if p not in big]
            for p in big:                                   # page-sized frames: kept apart so they don't swallow everything
                rr = fitz.Rect(p["rect"])
                if p.get("fill") is None:
                    objs.append({"kind": "draw", "rect": rr, "paths": [p], "hollow": True})
                elif rr.width * rr.height < 0.85 * parea:
                    objs.append({"kind": "draw", "rect": rr, "paths": [p]})   # (bigger than that = the page background)
            if small:
                for c in page.cluster_drawings(drawings=small):
                    mine = [p for p in small if (fitz.Rect(c) + (-1, -1, 1, 1)).contains(p["rect"])]
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
            r = o["rect"]
            if not (r + (-tol, -tol, tol, tol)).contains(pt):
                continue
            if o.get("hollow"):                             # an empty frame is only grabbed by its outline
                inner = r + (tol * 2, tol * 2, -tol * 2, -tol * 2)
                if inner.x1 > inner.x0 and inner.y1 > inner.y0 and inner.contains(pt):
                    continue
            area = max(1.0, r.width * r.height) * (0.5 if o["kind"] == "text" else 1.0)
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

    # -- drawing a ready-made shape --
    def _shape_box(self, kind, start, cur, shift):
        dx, dy = cur.x - start.x, cur.y - start.y
        if kind in LINE_SHAPES:
            if shift:                                       # snap to 0 / 45 / 90 degrees
                ang = round(math.atan2(dy, dx) / (math.pi / 4)) * (math.pi / 4)
                ln = math.hypot(dx, dy)
                dx, dy = ln * math.cos(ang), ln * math.sin(ang)
        elif kind in SQUARE_SHAPES or shift:
            side = max(abs(dx), abs(dy))
            dx, dy = math.copysign(side, dx or 1), math.copysign(side, dy or 1)
        return start.x, start.y, start.x + dx, start.y + dy

    def _paint_shape_preview(self, p, d):
        z = self.zoom
        x0, y0, x1, y1 = self._shape_box(d["kind"], d["start"], d["cur"], d["shift"])
        st = self.panel.shape_style()
        out, solid = _shape_qpath(d["kind"], x0 * z, y0 * z, x1 * z, y1 * z, st["width"] * z)
        stroke = QColor(*[int(c * 255) for c in st["stroke"]]) if st["stroke"] else None
        fill = QColor(*[int(c * 255) for c in st["fill"]]) if st["fill"] else None
        pen = QPen(stroke, max(1.0, st["width"] * z)) if stroke else QPen(Qt.PenStyle.NoPen)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        if st["dash"] == "dashed":
            pen.setStyle(Qt.PenStyle.DashLine)
        elif st["dash"] == "dotted":
            pen.setStyle(Qt.PenStyle.DotLine)
        p.setPen(pen)
        p.setBrush(fill if (fill and d["kind"] not in LINE_SHAPES) else Qt.BrushStyle.NoBrush)
        p.drawPath(out)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(stroke or QColor(0, 0, 0))
        p.drawPath(solid)

    # -- mouse --
    def edit_press(self, lab, pos, mods):
        idx = lab.index
        self._commit_editor()
        if not self._page_ok(idx):
            return
        pt = fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom)
        if self.edit_tool:                                  # drawing a ready-made shape
            self.edit_page, self.edit_sel = idx, []
            self._drag = {"mode": "shape", "idx": idx, "kind": self.edit_tool, "start": pt, "cur": pt,
                          "shift": bool(mods & Qt.KeyboardModifier.ShiftModifier)}
            self._ghost = None
            for l in self.labels:
                l.update()
            return
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
        self._sel_changed()
        for l in self.labels:
            l.update()

    def edit_move(self, lab, pos, buttons, mods):
        idx = lab.index
        d = self._drag
        if d is not None and d["mode"] == "shape":
            if (buttons & Qt.MouseButton.LeftButton) and d["idx"] == idx:
                d["cur"] = fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom)
                d["shift"] = bool(mods & Qt.KeyboardModifier.ShiftModifier)
                lab.update()
            return
        if d is None or not (buttons & Qt.MouseButton.LeftButton) or d["idx"] != idx:
            if d is None:                                   # hover feedback
                if self.edit_tool:
                    lab.setCursor(Qt.CursorShape.CrossCursor)
                    return
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
        if d["mode"] == "shape":
            kind = d["kind"]
            x0, y0, x1, y1 = self._shape_box(kind, d["start"], d["cur"], d["shift"])
            if math.hypot(x1 - x0, y1 - y0) < 4:            # a plain click drops a default-sized shape
                sx, sy = d["start"].x, d["start"].y
                if kind in LINE_SHAPES:
                    x0, y0, x1, y1 = sx, sy, sx + 120, sy
                elif kind in SQUARE_SHAPES:
                    x0, y0, x1, y1 = sx, sy, sx + 80, sy + 80
                else:
                    x0, y0, x1, y1 = sx, sy, sx + 110, sy + 76
            self._apply_new_shape(idx, kind, x0, y0, x1, y1)
            self._set_tool(None)
            return
        if d["mode"] == "band":
            b = self._band
            self._band = None
            if b.width() > 3 or b.height() > 3:
                rect = fitz.Rect(b.left() / self.zoom, b.top() / self.zoom, b.right() / self.zoom, b.bottom() / self.zoom)
                add = [o for o in self.page_objects(idx) if rect.contains(o["rect"]) and o not in self.edit_sel]
                self.edit_sel = self.edit_sel + add if (mods & Qt.KeyboardModifier.ShiftModifier) else add
            self._sel_changed()
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
        if self.edit_tool or not self._page_ok(idx):
            return
        pt = fitz.Point(pos.x() / self.zoom, pos.y() / self.zoom)
        o = self._hit(idx, pt)
        if o is not None and o["kind"] == "text":
            self.edit_page, self.edit_sel = idx, [o]
            self._sel_changed()
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
            if self.edit_tool:
                self._drag = None
                self._set_tool(None)
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
        d = self._drag
        if d is not None and d["mode"] == "shape" and d["idx"] == idx:
            self._paint_shape_preview(p, d)
            return
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
            if moving and img is not None and not o.get("hollow"):   # live ghost of the content being moved / scaled
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
            fam, size, bold, italic, color, align, spacing = self._text_values(obj)
            r = obj["rect"]
            x, y, w, h = r.x0 * z - 2, r.y0 * z - 2, r.width * z + 60, r.height * z + 6
        else:
            fmt = self.panel.current_text_fmt()
            text, size = "", fmt.get("size", 12.0)
            fam, bold, italic = fmt.get("family"), fmt.get("bold"), fmt.get("italic")
            color, align, spacing = fmt.get("color", (0, 0, 0)), fmt.get("align"), fmt.get("spacing")
            x, y, w, h = pt.x * z, (pt.y - size) * z, 220, size * z * 1.8
        look = {"family": fam, "px": size * z, "bold": bold, "italic": italic, "color": color, "align": align,
                "spacing": spacing}
        ed = InlineTextEditor(lab, text, w, h, look, keep=self.panel)
        ed.move(int(x), int(y))
        ed.committed.connect(lambda t, o=obj, p0=pt, i=lab.index, e=ed: self._editor_done(i, o, p0, t, dict(e.fmt)))
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

    def _editor_done(self, idx, obj, pt, text, fmt):
        self._editor_closed()
        try:
            if obj is None:
                if text.strip():
                    self._apply_new_text(idx, pt, text, self.panel.current_text_fmt())
            else:
                old = "\n".join("".join(sp["text"] for sp in ln["spans"]) for ln in obj["lines"])
                if text != old or fmt:
                    self._apply_text_edit(idx, obj, text, fmt)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))

    # -- applying changes to the PDF --
    def _finish_edit(self, idx, select_rects=None):
        """Redraw the changed page + its thumbnail (keeps the scroll position) and re-select what was edited."""
        self._objs.pop(idx, None)
        self.render_page(idx, force=True)
        self.thumbs.item(idx).setIcon(self.thumb_icon(idx))
        self.thumbs.viewport().update()
        self.edit_page = idx
        self.edit_sel = []
        if select_rects:
            for rr in select_rects:
                best, best_iou, inside = None, 0.25, []
                for o in self.page_objects(idx):
                    inter = fitz.Rect(o["rect"]) & rr
                    if inter.is_empty:
                        continue
                    ia, oa = inter.width * inter.height, max(1e-6, o["rect"].width * o["rect"].height)
                    if o["kind"] == "text" and ia / oa > 0.6:
                        inside.append(o)                        # a text block that MuPDF split into several pieces
                    iou = ia / max(1e-6, oa + rr.width * rr.height - ia)
                    if iou > best_iou:
                        best, best_iou = o, iou
                for o in inside + ([best] if best is not None else []):
                    if o not in self.edit_sel:
                        self.edit_sel.append(o)
        self._sel_changed()
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
        if draws:
            _remove_draws(page, draws)

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

    @staticmethod
    def _style_of(span, fmt=None):
        """Font / size / colour of a text span, with the changes made in the Edit bar on top."""
        flags = span["flags"]
        st = {"font": span["font"], "family": None, "size": span["size"], "color": _rgb_from_int(span["color"])}
        fmt = fmt or {}
        if "bold" in fmt:
            flags = (flags | 16) if fmt["bold"] else (flags & ~16)
        if "italic" in fmt:
            flags = (flags | 2) if fmt["italic"] else (flags & ~2)
        st["flags"] = flags
        if fmt.get("family"):
            st["family"] = fmt["family"]
        if fmt.get("size"):
            st["size"] = float(fmt["size"])
        if fmt.get("color"):
            st["color"] = fmt["color"]
        return st

    def _layout_text(self, page, obj, texts, fmt, cache, same_text=False):
        """Write a text object's lines again with the given text and format. Returns the area it now covers."""
        self._fonts_added = True
        lines, R = obj["lines"], fitz.Rect(obj["rect"])
        firsts = [self._first_span(l) for l in lines]
        n = len(texts)
        spacing, align = fmt.get("spacing"), fmt.get("align")
        ratio = 1.2
        if len(lines) > 1 and firsts[0]["size"] > 0:
            ratio = (firsts[1]["origin"][1] - firsts[0]["origin"][1]) / firsts[0]["size"]
            if ratio < 0.3:
                ratio = 1.2
        ratio = max(0.5, min(4.0, spacing if spacing else ratio))
        keep_y = not spacing and n == len(lines)
        styles = [self._style_of(firsts[min(i, len(firsts) - 1)], fmt) for i in range(n)]
        ys, y = [], firsts[0]["origin"][1]
        for i in range(n):
            if keep_y:
                y = firsts[i]["origin"][1]
            elif i > 0:
                y += ratio * styles[i]["size"]
            ys.append(y)
        used = None

        def grow(x0, x1, yb, size):
            nonlocal used
            rc = fitz.Rect(x0, yb - size, x1, yb + size * 0.3)
            used = rc if used is None else (used | rc)

        if same_text and set(fmt) <= {"align", "spacing"}:  # only alignment / spacing: move the original spans, keep their looks
            for i, ln in enumerate(lines):
                lb = fitz.Rect(ln["bbox"])
                dx = 0.0
                if align == "left":
                    dx = R.x0 - lb.x0
                elif align == "right":
                    dx = R.x1 - lb.x1
                elif align == "center":
                    dx = (R.x0 + R.x1) / 2 - (lb.x0 + lb.x1) / 2
                dy = ys[i] - firsts[i]["origin"][1]
                for sp in ln["spans"]:
                    if not sp["text"].strip():
                        continue
                    bb = fitz.Rect(sp["bbox"])
                    st = self._style_of(sp)
                    if _RTL_RE.search(sp["text"]):
                        _emit_text(page, sp["text"], sp["origin"][1] + dy, st, "right", bb.x0 + dx, bb.x1 + dx, cache)
                    else:
                        _emit_text(page, sp["text"], sp["origin"][1] + dy, st, "left", sp["origin"][0] + dx, bb.x1 + dx, cache)
                grow(lb.x0 + dx, lb.x1 + dx, ys[i], firsts[i]["size"])
            return used
        for i, t in enumerate(texts):
            if not t.strip():
                continue
            st, rtl = styles[i], bool(_RTL_RE.search(t))
            w = _text_width(page, t, st, cache) * (0.8 if rtl else 1.0)
            if align in ("left", "center", "right"):
                a, x0, x1 = align, R.x0, R.x1
            else:                                           # keep each line where it was
                ob = fitz.Rect(lines[min(i, len(lines) - 1)]["bbox"])
                a, x0, x1 = ("right" if rtl else "left"), ob.x0, ob.x1
            _emit_text(page, t, ys[i], st, a, x0, x1, cache)
            if a == "right":
                grow(x1 - w, x1, ys[i], st["size"])
            elif a == "center":
                grow((x0 + x1) / 2 - w / 2, (x0 + x1) / 2 + w / 2, ys[i], st["size"])
            else:
                grow(x0, x0 + w, ys[i], st["size"])
        return used

    def _transform_each(self, idx, items, ov=None):
        """Move / scale objects. items = [(object, anchor, sx, sy, dx, dy)], ov = new look for the vector shapes."""
        if not items:
            return
        page = self.doc[idx]
        self.snapshot()
        objs = [it[0] for it in items]
        cache = {}
        new_rects = []
        try:
            payloads = {o["xref"]: self._image_payload(o) for o in objs if o["kind"] == "image"}
            self._remove_objects(page, objs)
            for o, anchor, sx, sy, dx, dy in items:
                ax, ay = anchor
                T = lambda p, ax=ax, ay=ay, sx=sx, sy=sy, dx=dx, dy=dy: fitz.Point(ax + (p.x - ax) * sx + dx, ay + (p.y - ay) * sy + dy)
                if abs(sx - 1) < 1e-6:
                    s_text = sy
                elif abs(sy - 1) < 1e-6:
                    s_text = sx
                else:
                    s_text = (sx + sy) / 2
                avg = (abs(sx) + abs(sy)) / 2
                nr = self._tx_rect(o["rect"], (sx, sy, dx, dy), anchor)
                new_rects.append(nr)
                if o["kind"] == "text":
                    for ln in o["lines"]:
                        for sp in ln["spans"]:
                            if not sp["text"].strip():
                                continue
                            self._fonts_added = True
                            bb = fitz.Rect(sp["bbox"])
                            nb = fitz.Rect(T(bb.tl), T(bb.br))
                            st = self._style_of(sp)
                            st["size"] = sp["size"] * s_text
                            org = T(fitz.Point(sp["origin"]))
                            if _RTL_RE.search(sp["text"]):
                                _emit_text(page, sp["text"], org.y, st, "right", nb.x0, nb.x1, cache)
                            else:
                                _emit_text(page, sp["text"], org.y, st, "left", org.x, nb.x1, cache)
                elif o["kind"] == "image":
                    pl = payloads.get(o["xref"])
                    if pl:
                        page.insert_image(nr, stream=pl[0], mask=pl[1], keep_proportion=False)
                else:
                    for path in o["paths"]:
                        _redraw_path(page, path, T, avg, ov)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))
        self._finish_edit(idx, new_rects)

    def _apply_transform(self, idx, anchor, sx, sy, dx, dy):
        self._transform_each(idx, [(o, anchor, sx, sy, dx, dy) for o in self.edit_sel])

    def _restyle_shapes(self, idx, ov):
        draws = [o for o in self.edit_sel if o["kind"] == "draw"]
        self._transform_each(idx, [(o, (0.0, 0.0), 1.0, 1.0, 0.0, 0.0) for o in draws], ov)

    def _restyle_text(self, idx, fmt):
        objs = [o for o in self.edit_sel if o["kind"] == "text"]
        if not objs:
            return
        page = self.doc[idx]
        self.snapshot()
        rects, cache = [], {}
        try:
            self._remove_objects(page, objs)
            groups = [o for o in objs]
            if len(objs) > 1 and set(fmt) <= {"align", "spacing"}:     # several pieces: line them up as one block
                lines = sorted((ln for o in objs for ln in o["lines"]), key=lambda l: self._first_span(l)["origin"][1])
                r = fitz.Rect(objs[0]["rect"])
                for o in objs[1:]:
                    r |= o["rect"]
                groups = [{"kind": "text", "rect": r, "lines": lines}]
            for o in groups:
                texts = ["".join(sp["text"] for sp in ln["spans"]) for ln in o["lines"]]
                r = self._layout_text(page, o, texts, fmt, cache, same_text=True)
                if r is not None:
                    rects.append(r)
        except Exception as e:
            QMessageBox.warning(self.main, "Edit failed", str(e))
        self._finish_edit(idx, rects)

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

    def _apply_text_edit(self, idx, obj, text, fmt=None):
        page = self.doc[idx]
        self.snapshot()
        self._remove_objects(page, [obj])
        old = [("".join(sp["text"] for sp in ln["spans"])) for ln in obj["lines"]]
        new = text.split("\n")
        r = self._layout_text(page, obj, new, fmt or {}, {}, same_text=(new == old))
        self._finish_edit(idx, [r] if r is not None else None)

    def _apply_new_text(self, idx, pt, text, fmt):
        page = self.doc[idx]
        self.snapshot()
        self._fonts_added = True
        cache = {}
        flags = (16 if fmt.get("bold") else 0) | (2 if fmt.get("italic") else 0)
        st = {"font": "Helvetica", "family": fmt.get("family") or "Arial", "flags": flags,
              "size": float(fmt.get("size") or 12.0), "color": fmt.get("color") or (0, 0, 0)}
        if not fmt.get("family"):                           # no font picked: the standard sans font keeps the file small
            st["family"] = None
        lines = text.split("\n")
        ratio = float(fmt.get("spacing") or 1.2)
        widths = [_text_width(page, t, st, cache) * (0.8 if _RTL_RE.search(t) else 1.0) for t in lines]
        wmax = max(widths + [1.0])
        align = fmt.get("align") or "left"
        used = None
        for i, t in enumerate(lines):
            if t.strip():
                y = pt.y + i * st["size"] * ratio
                _emit_text(page, t, y, st, align, pt.x, pt.x + wmax, cache)
                rc = fitz.Rect(pt.x, y - st["size"], pt.x + wmax, y + st["size"] * 0.3)
                used = rc if used is None else (used | rc)
        self._finish_edit(idx, [used] if used is not None else None)

    def _apply_new_shape(self, idx, kind, x0, y0, x1, y1):
        st = self.panel.shape_style()
        page = self.doc[idx]
        self.snapshot()
        try:
            _draw_shape_pdf(page, kind, x0, y0, x1, y1, st["stroke"], st["fill"], st["width"], st["dash"])
        except Exception as e:
            QMessageBox.warning(self.main, "Shape failed", str(e))
        h = st["width"] / 2
        self._finish_edit(idx, [fitz.Rect(min(x0, x1) - h, min(y0, y1) - h, max(x0, x1) + h, max(y0, y1) + h)])


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
        self.page_total = QLabel("/ 0")
        self.page_total.setObjectName("PageText")
        self.page_box = QLineEdit()
        self.page_box.setObjectName("PageBox")
        self.page_box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_box.setFixedWidth(68)
        self.page_box.setValidator(QIntValidator(1, 999999, self))
        self.page_box.setToolTip("Type a page number and press Enter to go to that page")
        self.page_box.returnPressed.connect(self.page_box_go)
        self.page_box.installEventFilter(self)
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
        if paths and paths[0] == MERGE_TAG:
            self.open_combined(paths[1:])
            return
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

    def open_combined(self, items):
        """Explorer > Combine: build the merged PDF in memory and show it as an unsaved tab."""
        items = [p for p in items if os.path.isfile(p)]
        if not items:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            d, skipped = combine_to_doc(items)
        finally:
            QApplication.restoreOverrideCursor()
        if len(d) == 0:
            QMessageBox.warning(self, "PDF Studio", "Nothing could be combined.")
            return
        tab = DocTab(self)
        tab.load_doc(d, "Combined", _unique_target(os.path.dirname(items[0])))
        tab.set_edit(self.a_edit.isChecked())
        tab.changed.connect(lambda t=tab: self.on_tab_changed(t))
        i = self.tabs.addTab(tab, "")
        self.stack.setCurrentIndex(1)
        self.tabs.setCurrentIndex(i)
        self.on_tab_changed(tab)
        if skipped:
            QMessageBox.warning(self, "PDF Studio", "Skipped:\n" + "\n".join(skipped))

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
        name = os.path.basename(tab.path) if tab.path else (tab.title or "Untitled")
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
            if paths == ["::noop::"]:
                sock.deleteLater()
                return
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
        if self._tray is not None and not self._quitting and self._settings.value("background/keep", "1") != "0":
            for i in reversed(range(self.tabs.count())):    # free the documents, keep the process (and Qt) loaded
                t = self.tabs.widget(i)
                self.tabs.removeTab(i)
                t.deleteLater()
            self.stack.setCurrentIndex(0)
            self.update_status()
            self.hide()
            e.ignore()                                      # the next PDF you open appears instantly
            return
        e.accept()

    # ---------- stays loaded in the background so the next PDF opens instantly ----------
    _tray, _quitting = None, False

    def setup_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return
        tray = QSystemTrayIcon(app_icon(), self)
        tray.setToolTip("PDF Studio")
        m = QMenu()
        m.addAction("Open PDF Studio", self.bring_to_front)
        keep = m.addAction("Keep running in background (instant open)")
        keep.setCheckable(True)
        keep.setChecked(self._settings.value("background/keep", "1") != "0")
        keep.toggled.connect(lambda on: self._settings.setValue("background/keep", "1" if on else "0"))
        if sys.platform == "win32" and getattr(sys, "frozen", False):
            auto = m.addAction("Start with Windows")
            auto.setCheckable(True)
            auto.setChecked(_autostart_enabled())
            auto.toggled.connect(_set_autostart)
        m.addSeparator()
        m.addAction("Exit", self.quit_app)
        tray.setContextMenu(m)
        tray.activated.connect(lambda r: self.bring_to_front() if r in (QSystemTrayIcon.ActivationReason.Trigger,
                                                                        QSystemTrayIcon.ActivationReason.DoubleClick) else None)
        tray.show()
        self._tray, self._tray_menu = tray, m

    def quit_app(self):
        self._quitting = True
        if self.close():
            QApplication.quit()
        else:
            self._quitting = False

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
    paths = _open_after or [os.path.abspath(a) for a in args]
    mode, server = acquire_instance(paths)
    if mode == "forwarded":
        sys.exit(0)
    app.setQuitOnLastWindowClosed(False)                    # closing the window keeps the app loaded in the tray
    background = "--background" in sys.argv
    w = Main()
    if server:
        w.attach_server(server)
    w.setup_tray()
    app.commitDataRequest.connect(lambda *_: setattr(w, "_quitting", True))   # Windows logoff / shutdown must not be blocked
    if w._tray is None:
        app.setQuitOnLastWindowClosed(True)
    if not background or paths:
        w.show()
    w.open_paths(paths)
    import threading
    QTimer.singleShot(4000, lambda: threading.Thread(target=_cleanup_old_page_exports, daemon=True).start())
    QTimer.singleShot(2500, refresh_registration_if_stale)  # fix a dead exe path in the Explorer registration
    if not paths and not background:                       # started on its own (not to open a file): offer to install
        QTimer.singleShot(700, w.offer_install_on_first_run)
    sys.exit(app.exec())

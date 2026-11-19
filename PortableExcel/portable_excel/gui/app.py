"""Application start-up: logging, crash handling, password lock, main window."""
from __future__ import annotations

import faulthandler
import logging
import os
import sys
import traceback
from typing import List, Optional

log = logging.getLogger("PortableExcel")
_fault_file = None


def setup_logging() -> None:
    global _fault_file
    from . import config
    if log.handlers:
        return
    log.setLevel(logging.INFO)
    try:
        path = config.log_path()
        h = logging.FileHandler(path, encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        log.addHandler(h)
        _fault_file = open(path, "a", encoding="utf-8")
        faulthandler.enable(_fault_file)
    except OSError:
        try:
            faulthandler.enable()
        except (RuntimeError, ValueError, AttributeError):  # no stderr when frozen
            pass
    if sys.stderr is not None:
        sh = logging.StreamHandler(sys.stderr)
        sh.setLevel(logging.WARNING)
        log.addHandler(sh)


_in_hook = [False]
_smoke_errors: List[str] = []


def excepthook(etype, value, tb):
    text = "".join(traceback.format_exception(etype, value, tb))
    log.error("Unhandled exception:\n%s", text)
    _smoke_errors.append(text)
    if os.environ.get("PORTABLEEXCEL_SMOKE_TEST") == "1":
        return
    if _in_hook[0]:
        return
    _in_hook[0] = True
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        if QApplication.instance() is not None:
            box = QMessageBox()
            box.setIcon(QMessageBox.Critical)
            box.setWindowTitle("PortableExcel - unexpected error")
            box.setText("An unexpected error occurred. Your work is still open; please save "
                        "it under a new name.")
            box.setDetailedText(text)
            box.exec()
    except Exception:
        pass
    finally:
        _in_hook[0] = False


def create_app(argv=None):
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtWidgets import QApplication
    from . import icons, theme
    app = QApplication.instance()
    if app is None:
        QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
        app = QApplication(list(argv if argv is not None else sys.argv))
        app.setApplicationName("PortableExcel")
        app.setOrganizationName("PortableExcel")
        app.setApplicationDisplayName("PortableExcel")
        theme.apply(app)
        app.setWindowIcon(icons.app_icon())
    return app


def demo_book():
    from ..model import Workbook
    book = Workbook()
    sh = book.active
    rows = [["Product", "Units", "Price", "Total"], ["Apples", "120", "0.5", "=B2*C2"],
            ["Pears", "80", "0.75", "=B3*C3"], ["Total", "=SUM(B2:B3)", "", "=SUM(D2:D3)"]]
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            sh.set_raw(r, c, v)
    return book


def _smoke(app) -> int:
    from PySide6.QtCore import QTimer
    from .main_window import MainWindow
    code = [0]
    try:
        book = demo_book()
        win = MainWindow(book)
        win.show()
        win.engine.invalidate()
        win.engine.recalc()
        total = win.engine.display_text(book.active, 3, 3)
        log.info("smoke test: D4=%s", total)
        if total != "120":
            raise RuntimeError(f"unexpected recalc result {total!r}")
    except Exception:
        log.exception("smoke test failed")
        return 1

    def finish():
        if _smoke_errors:
            code[0] = 1
        win.modified = False
        win.close()
        app.exit(code[0])
    QTimer.singleShot(1500, finish)
    rc = app.exec()
    return 1 if (rc or _smoke_errors) else 0


def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv if argv is None else argv)
    setup_logging()
    sys.excepthook = excepthook
    log.info("PortableExcel starting (%s)", " ".join(argv))
    app = create_app(argv)
    if os.environ.get("PORTABLEEXCEL_SMOKE_TEST") == "1":
        return _smoke(app)
    from . import security
    if not security.unlock():
        return 0
    from .main_window import MainWindow
    win = MainWindow()
    win.show()
    for path in argv[1:]:
        if path and not path.startswith("-") and os.path.exists(path):
            win.open_file(os.path.abspath(path))
    return app.exec()

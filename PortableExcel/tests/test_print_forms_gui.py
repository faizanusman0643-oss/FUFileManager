"""Printing/PDF, barcode-font emulation, form window API, clipboard (offscreen)."""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["PORTABLEEXCEL_NO_LOCK"] = "1"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest  # noqa: E402
from PySide6.QtCore import QPoint, QPointF, Qt  # noqa: E402
from PySide6.QtGui import QColor, QImage, QPainter, QWheelEvent  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from portable_excel import forms  # noqa: E402
from portable_excel.gui import barcode_font as bf  # noqa: E402
from portable_excel.gui import config  # noqa: E402
from portable_excel.gui.app import create_app  # noqa: E402
from portable_excel.gui.printing import excel_filter_to_qt  # noqa: E402
from portable_excel.gui.userform_view import QtFormView  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return create_app(["test"])


@pytest.fixture
def win(app, tmp_path):
    config.set_ini_path(str(tmp_path / "PortableExcel.ini"))
    from portable_excel.gui import main_window as mw
    w = mw.MainWindow()
    w.show()
    yield w
    w.modified = False
    w.close()
    config.set_ini_path(None)


# ------------------------------------------------------------------ barcode fonts
def test_decoder_roundtrip():
    enc = bf.encode_b("ABC-123")
    vals = bf.decode(enc)
    assert vals[0] == bf.START_B and vals[-1] == bf.STOP
    assert bf.is_valid_symbol(vals)
    assert bf.human_text(vals) == "ABC-123"
    # value 0 as space or as char 194, high values 195..206
    assert bf.char_value(" ") == 0 and bf.char_value(chr(194)) == 0
    assert bf.char_value(chr(195)) == 95 and bf.char_value(chr(206)) == 106
    # IDAutomation offset variants (ReturnType 6-9)
    assert bf.char_value(chr(204 + 11500)) == 104
    # set C pairs: start C (205), 12, 34, check, stop
    vals = [105, 12, 34]
    check = (105 + 12 * 1 + 34 * 2) % 103
    s = "".join(chr(v + 100) if v > 94 else chr(v + 32) for v in vals + [check, 106])
    assert bf.human_text(bf.decode(s)) == "1234"
    # start + 7 data + check = 9 symbols of 11 modules, plus the 13-module stop
    assert sum(bf.module_runs(bf.decode(enc))) == 9 * 11 + 13


def test_font_detection_and_paint(app):
    assert bf.is_barcode_font("Code 128")
    assert bf.is_barcode_font("IDAutomationSC128XS DEMO")
    assert bf.is_barcode_font("Libre Barcode 128 Text")
    assert not bf.is_barcode_font("Calibri")
    assert bf.font_installed("Libre Barcode 128")  # bundled OFL font
    assert not bf.needs_emulation("Libre Barcode 128 Text")
    assert bf.needs_emulation("IDAutomationSC128XXS DEMO")
    img = QImage(300, 60, QImage.Format_RGB32)
    img.fill(Qt.white)
    p = QPainter(img)
    assert bf.paint(p, img.rect(), bf.encode_b("XYZ"), "Code 128 Text")
    p.end()
    dark = sum(1 for x in range(300) if QColor(img.pixel(x, 20)).value() < 100)
    assert dark > 30  # bars were drawn


def test_cell_with_barcode_font_renders_bars(win):
    sh = win.book.active
    sh.set_raw(0, 0, "'" + bf.encode_b("AB12"))
    sh.update_style(0, 0, font_name="IDAutomationSC128XS DEMO")
    sh.col_widths[0] = 220
    sh.row_heights[0] = 40
    win.view.apply_sizes()
    win.view.goto(5, 5)
    QApplication.processEvents()
    img = win.view.viewport().grab().toImage()
    r = win.view.visualRect(win.view.model().index(0, 0))
    y = r.top() + r.height() // 3
    dark = sum(1 for x in range(r.left(), r.right()) if QColor(img.pixel(x, y)).value() < 80)
    assert dark > 20


# ------------------------------------------------------------------ printing
def test_export_pdf_pages(win, tmp_path):
    sh = win.book.active
    for r in range(120):
        sh.set_raw(r, 0, f"Row {r + 1}")
        sh.set_raw(r, 1, str(r * 3))
    sh.set_raw(0, 2, "'" + bf.encode_b("PDF1"))
    sh.update_style(0, 2, font_name="Code 128")
    path = str(tmp_path / "out.pdf")
    assert win.export_pdf(sh, 0, 0, 1_048_575, 16_383, path) == path
    data = open(path, "rb").read()
    assert data.startswith(b"%PDF")
    assert data.count(b"/Type /Page") - data.count(b"/Type /Pages") >= 2
    rr = win.range_renderer(sh, 0, 0, 119, 1)
    assert rr.page_count(700, 1000) >= 2
    # VBA host entry point
    from portable_excel.gui.vba_host import GuiVBAHost
    host = GuiVBAHost(win)
    p2 = str(tmp_path / "sel.pdf")
    assert host.export_pdf(sh, 0, 0, 9, 1, p2) == p2 and os.path.getsize(p2) > 500


def test_excel_filter_conversion():
    assert excel_filter_to_qt("PDF Files (*.pdf), *.pdf") == "PDF Files (*.pdf)"
    assert excel_filter_to_qt("Excel Files (*.xls*), *.xls*, Text, *.txt;*.csv") == \
        "Excel Files (*.xls*);;Text (*.txt *.csv)"
    assert excel_filter_to_qt(None) == "All files (*)"


def test_clipboard_host(win):
    from portable_excel.gui.vba_host import GuiVBAHost
    host = GuiVBAHost(win)
    host.set_clipboard_text("hello DataObject")
    assert host.get_clipboard_text() == "hello DataObject"


# ------------------------------------------------------------------ forms
class Inst:
    def __init__(self, form):
        self.name = "UserForm1"
        self.form = forms.normalize_form(form)
        self.calls = []

    def fire(self, ctrl, event, args=()):
        self.calls.append(("fire", ctrl, event))
        return list(args)

    def user_set(self, ctrl, prop, value):
        self.calls.append(("set", ctrl, prop, value))
        if ctrl is None:
            self.form[prop] = value
        else:
            for c in self.form["controls"]:
                if c["name"] == ctrl:
                    c[prop] = value

    def request_close(self, mode=0):
        self.view.close()


def _wheel(widget, delta):
    pos = QPointF(widget.width() / 2, widget.height() / 2)
    ev = QWheelEvent(pos, QPointF(widget.mapToGlobal(pos.toPoint())), QPoint(0, 0),
                     QPoint(0, delta), Qt.NoButton, Qt.NoModifier, Qt.NoScrollPhase, False)
    QApplication.sendEvent(widget, ev)


def test_wheel_scrolls_scrollbar(app):
    f = forms.new_form()
    f["controls"] = [forms.new_control("Label", "Lbl", left=6, top=6),
                     forms.new_control("ScrollBar", "vScroll", left=200, top=6, height=120,
                                       max=100, small_change=2)]
    inst = Inst(f)
    v = QtFormView(inst)
    inst.view = v
    v.show(False)
    _wheel(v.renderer.widgets["Lbl"], -120)   # one notch down over another control
    sb = inst.form["controls"][1]
    assert sb["value"] == 6
    assert ("fire", "vScroll", "Change") in inst.calls
    _wheel(v.renderer.widgets["vScroll"], 240)  # two notches up, clamped at min
    assert sb["value"] == 0
    # controls added at run time also get the wheel
    v.add_control(forms.new_control("TextBox", "Added", left=6, top=40))
    _wheel(v.renderer.widgets["Added"], -120)
    assert sb["value"] == 6
    v.close()


def test_window_state_and_resize_event(app):
    inst = Inst(forms.new_form())
    v = QtFormView(inst)
    inst.view = v
    v.set_window_flags(resizable=True, minimize=True, maximize=True, caption=True)
    v.show(False)
    assert v.windowFlags() & Qt.WindowMaximizeButtonHint
    w0 = inst.form["width"]
    v.resize(v.width() + 100, v.height() + 40)
    QApplication.processEvents()
    assert inst.form["width"] > w0
    assert ("fire", None, "Resize") in inst.calls
    v.set_window_state("maximized")
    assert v.windowState() & Qt.WindowMaximized
    v.set_window_state("normal")
    assert not (v.windowState() & Qt.WindowMaximized)
    v.close()

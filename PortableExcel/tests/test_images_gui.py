"""Pictures / barcodes over the grid (offscreen Qt)."""
import importlib
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["PORTABLEEXCEL_NO_LOCK"] = "1"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

try:  # real module when present, else the throwaway stub
    _bc = importlib.import_module("portable_excel.barcodes")
    if not hasattr(_bc, "make_barcode_image"):
        raise ImportError
except ImportError:
    _bc = importlib.import_module("tests.stubs.barcodes")
    sys.modules["portable_excel.barcodes"] = _bc
    setattr(importlib.import_module("portable_excel"), "barcodes", _bc)

import pytest  # noqa: E402
from PySide6.QtCore import QEvent, QPoint, QPointF, Qt  # noqa: E402
from PySide6.QtGui import QImage, QMouseEvent  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from portable_excel.gui import config  # noqa: E402
from portable_excel.gui.app import create_app  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return create_app(["test"])


@pytest.fixture
def win(app, tmp_path, monkeypatch):
    config.set_ini_path(str(tmp_path / "PortableExcel.ini"))
    from portable_excel.gui import main_window as mw
    for name in ("information", "warning"):
        monkeypatch.setattr(QMessageBox, name, staticmethod(lambda *a, **k: QMessageBox.Ok))
    w = mw.MainWindow()
    w.resize(1200, 800)
    w.show()
    QApplication.processEvents()
    yield w
    w.modified = False
    w.close()
    config.set_ini_path(None)


def drag(widget, start, end, mods=Qt.NoModifier):
    QTest.mousePress(widget, Qt.LeftButton, mods, start)
    ev = QMouseEvent(QEvent.MouseMove, QPointF(end), QPointF(widget.mapToGlobal(end)),
                     Qt.NoButton, Qt.LeftButton, mods)
    QApplication.sendEvent(widget, ev)
    QTest.mouseRelease(widget, Qt.LeftButton, mods, end)


def insert_linked(win):
    win.user_set_cell(1, 0, "ABC-123")
    win.view.goto(1, 0)
    d = win.barcode_dialog(modal=False)
    assert d.source.text() == "A2" and d.target.text() == "B2"
    assert d.symbology() == "code128"
    assert d.update_preview() is None and d.preview.pixmap() is not None
    return win.insert_barcode(d.values())


def test_barcode_dialog_validation(win):
    win.user_set_cell(0, 0, "12345")
    win.view.goto(0, 0)
    d = win.barcode_dialog(modal=False)
    d.sym.setCurrentIndex(d.sym.findData("ean13"))
    assert d.error.text()  # 5 digits are not valid EAN-13
    from PySide6.QtWidgets import QDialogButtonBox
    assert not d.buttons.button(QDialogButtonBox.Ok).isEnabled()
    d.fixed.setChecked(True)
    d.text.setText("400638133393")
    assert not d.error.text()
    assert "4006381333931" in d.value_label.text()  # normalised with check digit
    # ISBT-style text value ('=...' stored with an apostrophe) defaults to Code 128
    win.user_set_cell(3, 1, "'=PK0262601234400")
    win.view.goto(3, 1)
    d2 = win.barcode_dialog(modal=False)
    assert d2.symbology() == "code128" and d2.data_text() == "=PK0262601234400"


def test_insert_move_resize_undo_refresh_delete(win):
    img = insert_linked(win)
    sheet = win.book.active
    assert len(sheet.images) == 1 and img["barcode"]["source"] == "A2"
    assert (img["row"], img["col"]) == (1, 1)
    assert win.name_box.currentText() == img["name"]
    view = win.view
    vp = view.viewport()
    r = win.images.rect(view, sheet.images[0])
    # move by dragging the middle: re-anchors to the cell under the top-left
    target_x = view.columnViewportPosition(3) + 5
    target_y = view.rowViewportPosition(4) + 3
    start = r.center()
    end = start + QPoint(target_x - r.left(), target_y - r.top())
    drag(vp, start, end)
    im = sheet.images[0]
    assert (im["row"], im["col"]) == (4, 3) and (im["dx"], im["dy"]) == (5, 3)
    # resize with the bottom-right handle, Shift keeps the aspect ratio
    r = win.images.rect(view, im)
    w0, h0 = im["width"], im["height"]
    br = win.images.handles(r)["br"].center()
    drag(vp, br, br + QPoint(40, 0), Qt.ShiftModifier)
    im = sheet.images[0]
    assert im["width"] == w0 + 40
    assert abs(im["height"] / im["width"] - h0 / w0) < 0.05
    # undo resize, undo move
    win.undo_stack.undo()
    assert sheet.images[0]["width"] == w0
    win.undo_stack.undo()
    assert (sheet.images[0]["row"], sheet.images[0]["col"]) == (1, 1)
    # auto-refresh when the source cell changes
    data0 = sheet.images[0]["data"]
    win.user_set_cell(1, 0, "XYZ-999")
    assert sheet.images[0]["data"] != data0
    win.undo_stack.undo()
    assert sheet.images[0]["data"] == data0
    # copy / paste / delete
    win.images.select(sheet.images[0]["id"])
    win.copy()
    win.view.goto(10, 0)
    win.paste()
    assert len(sheet.images) == 2 and sheet.images[1]["row"] == 10
    assert sheet.images[1]["name"] != sheet.images[0]["name"]
    win.view.setFocus()
    QTest.keyClick(win.view, Qt.Key_Delete)
    assert len(sheet.images) == 1
    win.undo_stack.undo()
    assert len(sheet.images) == 2


def test_context_menu_and_save_png(win, tmp_path):
    insert_linked(win)
    im = win.book.active.images[0]
    win.images.select(im["id"])
    m = win.images.context_menu(win.mapToGlobal(QPoint(50, 50)))
    texts = [a.text() for a in m.actions() if a.text()]
    assert "Edit Barcode..." in texts and "Save Picture As PNG..." in texts
    m.close()
    p = win.images.save_picture(str(tmp_path / "bc.png"))
    assert QImage(p).width() > 0
    win.images.bring(False)
    win.images.reset_size()


def test_batch_generation_single_undo(win):
    sheet = win.book.active
    win.user_set_cell(0, 1, "DIN NO")
    for i in range(5):
        win.user_set_cell(1 + i, 1, f"'=PK02626012344{i:02d}")
    win.view.select_range(1, 1, 5, 1)
    d = win.batch_barcodes_dialog(modal=False)
    v = d.values()
    assert v["out_col"] == 2 and v["symbology"] == "code128"
    n_undo = win.undo_stack.count()
    assert win.generate_barcodes(v) == 5
    assert len(sheet.images) == 5
    assert {im["row"] for im in sheet.images} == {1, 2, 3, 4, 5}
    assert all(im["col"] == 2 for im in sheet.images)
    assert sheet.col_widths.get(2, 64) >= sheet.images[0]["width"]
    assert win.undo_stack.count() == n_undo + 1
    win.undo_stack.undo()
    assert sheet.images == [] and 2 not in sheet.col_widths


def test_insert_picture_file(win, tmp_path):
    qi = QImage(40, 20, QImage.Format_RGB32)
    qi.fill(Qt.red)
    path = str(tmp_path / "pic.bmp")
    qi.save(path)
    img = win.images.insert_picture_file(path)
    assert img["data"].startswith(b"\x89PNG") and img["width"] == 40
    assert img["barcode"] is None and img["name"].startswith("Picture")

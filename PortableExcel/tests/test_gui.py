"""GUI tests (pytest, offscreen Qt)."""
import os
import subprocess
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["PORTABLEEXCEL_NO_LOCK"] = "1"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import importlib  # noqa: E402

import pytest  # noqa: E402

# Use throwaway stubs only for modules that do not exist (yet).
for _name in ("engine", "vba"):
    try:
        _mod = importlib.import_module(f"portable_excel.{_name}")
        if not hasattr(_mod, "FormulaEngine" if _name == "engine" else "VBARuntime"):
            raise ImportError
    except ImportError:
        _stub = importlib.import_module(f"tests.stubs.{_name}")
        sys.modules[f"portable_excel.{_name}"] = _stub
        setattr(importlib.import_module("portable_excel"), _name, _stub)

from PySide6.QtCore import Qt  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtTest import QTest  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402

from portable_excel.gui import config, ops, security  # noqa: E402
from portable_excel.gui.app import create_app  # noqa: E402


@pytest.fixture(scope="session")
def app():
    return create_app(["test"])


@pytest.fixture(autouse=True)
def temp_ini(tmp_path):
    config.set_ini_path(str(tmp_path / "PortableExcel.ini"))
    yield tmp_path
    config.set_ini_path(None)


@pytest.fixture
def win(app, monkeypatch):
    from portable_excel.gui import main_window as mw
    # never block on modal message boxes during tests
    monkeypatch.setattr(QMessageBox, "question", staticmethod(lambda *a, **k: QMessageBox.Yes))
    monkeypatch.setattr(QMessageBox, "information", staticmethod(lambda *a, **k: QMessageBox.Ok))
    monkeypatch.setattr(QMessageBox, "warning", staticmethod(lambda *a, **k: QMessageBox.Ok))
    w = mw.MainWindow()
    w.show()
    QApplication.processEvents()
    yield w
    w.modified = False
    w.close()
    QApplication.processEvents()


def type_into(w, r, c, text):
    w.view.goto(r, c)
    w.view.setFocus()
    QTest.keyClicks(w.view, text)
    QTest.keyClick(w.view.delegate.editor or w.view, Qt.Key_Return)
    QApplication.processEvents()


# ---------------------------------------------------------------- security
def test_password_default_hash():
    assert security.check_password("12345643")
    assert not security.check_password("12345644")
    assert not security.check_password("")
    src = open(security.__file__, encoding="utf-8").read()
    assert "12345643" not in src  # never stored in plain text
    assert security.ITERATIONS == 200_000


def test_change_password_persists(temp_ini):
    assert security.change_password("wrong", "abcd1234", "abcd1234") is not None
    assert security.change_password("12345643", "abcd1234", "abcd") is not None
    assert security.change_password("12345643", "abcd1234", "abcd1234") is None
    ini = temp_ini / "PortableExcel.ini"
    assert ini.exists()
    assert "abcd1234" not in ini.read_text()
    config.set_ini_path(str(ini))  # drop cache, re-read from disk
    assert security.check_password("abcd1234")
    assert not security.check_password("12345643")


def test_lock_dialog(app):
    d = security.LockDialog()
    d.edit.setText("nope")
    d.try_unlock()
    assert d.failed == 1 and d.result() == 0
    for _ in range(4):
        d.edit.setText("nope")
        d.try_unlock()
    assert not d.ok_btn.isEnabled()  # 30 s lockout after 5 failures
    d._remaining = 1
    d._tick()
    d.edit.setText("12345643")
    d.try_unlock()
    assert d.result() == 1


# ---------------------------------------------------------------- grid
def test_typing_and_formula(win):
    type_into(win, 0, 0, "5")
    sh = win.book.active
    assert sh.get_raw(0, 0) == "5"
    assert win.view.current_cell() == (1, 0)  # Enter moves down
    type_into(win, 0, 1, "=A1*2")
    assert win.view.sheet_model.display(0, 1) == "10"
    type_into(win, 0, 0, "7")
    assert win.view.sheet_model.display(0, 1) == "14"
    assert win.modified and "*" in win.windowTitle()


def test_tab_moves_right(win):
    win.view.goto(2, 2)
    QTest.keyClicks(win.view, "x")
    QTest.keyClick(win.view.delegate.editor, Qt.Key_Tab)
    assert win.view.current_cell() == (2, 3)
    assert win.book.active.get_raw(2, 2) == "x"


def test_undo_redo(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "1")
    win.user_set_cell(0, 0, "2")
    win.view.goto(0, 0)
    win.apply_style(bold=True)
    assert sh.get_style(0, 0).get("bold")
    win.undo_stack.undo()
    assert not sh.get_style(0, 0).get("bold")
    win.undo_stack.undo()
    assert sh.get_raw(0, 0) == "1"
    win.undo_stack.redo()
    assert sh.get_raw(0, 0) == "2"
    assert win.undo_stack.undoLimit() == 100


def test_copy_paste_shift(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "1")
    win.user_set_cell(1, 0, "2")
    win.user_set_cell(0, 1, "=A1*10")
    win.view.goto(0, 1)
    win.copy()
    assert "10" in QGuiApplication.clipboard().text()
    win.view.goto(1, 1)
    win.paste()
    assert sh.get_raw(1, 1) == "=A2*10"
    assert win.view.sheet_model.display(1, 1) == "20"
    # single cell tiles over a larger selection
    win.view.select_range(2, 1, 3, 1)
    win.paste()
    assert sh.get_raw(3, 1) == "=A4*10"


def test_external_tsv_paste(win):
    QGuiApplication.clipboard().setText("a\tb\r\n1\t2\r\n")
    win.clip = None
    win.view.goto(4, 0)
    win.paste()
    sh = win.book.active
    assert (sh.get_raw(4, 0), sh.get_raw(4, 1), sh.get_raw(5, 1)) == ("a", "b", "2")


def test_insert_delete_row(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "1")
    win.user_set_cell(1, 0, "2")
    win.user_set_cell(2, 0, "=A1+A2")
    win.view.goto(1, 0)
    win.insert_cells("row")
    assert sh.get_raw(2, 0) == "2"
    assert sh.get_raw(3, 0) == "=A1+A3"
    win.undo_stack.undo()
    assert sh.get_raw(2, 0) == "=A1+A2"
    win.view.goto(0, 0)
    win.delete_cells("row")
    assert sh.get_raw(0, 0) == "2"


def test_sheet_add_rename(win):
    win.user_set_cell(0, 0, "3")
    win.add_sheet()
    assert len(win.book.sheets) == 2
    assert win.book.active is win.book.sheets[1]
    new = win.book.active
    win.user_set_cell(0, 0, "=Sheet1!A1*2")
    assert win.rename_sheet(win.book.sheets[0], "Data")
    assert win.book.sheets[0].name == "Data"
    assert "Data" in new.get_raw(0, 0)
    assert win.view.sheet_model.display(0, 0) == "6"
    assert not win.rename_sheet(win.book.sheets[0], "bad/name")
    assert [win.sheet_tabs.tabText(i) for i in range(win.sheet_tabs.count())] == ["Data",
                                                                                 "Sheet2"]


def test_number_format_combo(win):
    win.user_set_cell(0, 0, "0.256")
    win.view.goto(0, 0)
    i = win.num_combo.findText("Percent")
    win.num_combo.setCurrentIndex(i)
    win.num_combo.activated.emit(i)
    assert win.book.active.get_style(0, 0)["number_format"] == "0%"
    assert win.view.sheet_model.display(0, 0) == "26%"
    i = win.num_combo.findText("General")
    win.num_combo.activated.emit(i)
    assert "number_format" not in win.book.active.get_style(0, 0)


def test_fill_down_and_series(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "1")
    win.user_set_cell(1, 0, "2")
    win.fill_from_handle((0, 0, 1, 0), (2, 0, 4, 0))
    assert [sh.get_raw(r, 0) for r in range(5)] == ["1", "2", "3", "4", "5"]
    win.user_set_cell(0, 1, "=A1*2")
    win.view.select_range(0, 1, 3, 1)
    win.fill_down()
    assert sh.get_raw(3, 1) == "=A4*2"
    assert ops.fill_series(["Mon"], 2, lambda r, k: r) == ["Tue", "Wed"]
    assert ops.fill_series(["Item9"], 2, lambda r, k: r) == ["Item10", "Item11"]


def test_sort_and_status_sum(win):
    sh = win.book.active
    for r, v in enumerate(["Value", "3", "1", "2"]):
        win.user_set_cell(r, 0, v)
    win.view.select_range(0, 0, 3, 0)
    win.update_stats()
    assert "Sum: 6" in win.stats_label.text()
    assert "Count: 4" in win.stats_label.text()
    win.view.goto(1, 0)
    win.sort_selection(False)
    assert [sh.get_raw(r, 0) for r in range(4)] == ["Value", "1", "2", "3"]
    win.sort_selection(True)
    assert [sh.get_raw(r, 0) for r in range(4)] == ["Value", "3", "2", "1"]


def test_find_replace(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "apple pie")
    win.user_set_cell(5, 2, "Apple")
    opts = {"what": "apple", "with": "pear", "case": False, "whole": False, "values": False,
            "workbook": False}
    win.show_find(True)
    win.find_next(opts)
    assert win.view.current_cell() in ((0, 0), (5, 2))
    assert win.replace_all(opts) == 2
    assert sh.get_raw(0, 0) == "pear pie" and sh.get_raw(5, 2) == "pear"


def test_delete_key_and_ctrl_arrow(win):
    sh = win.book.active
    for r in range(5):
        win.user_set_cell(r, 0, str(r))
    win.view.goto(0, 0)
    win.view.setFocus()
    QTest.keyClick(win.view, Qt.Key_Down, Qt.ControlModifier)
    assert win.view.current_cell() == (4, 0)
    QTest.keyClick(win.view, Qt.Key_Delete)
    assert sh.get_raw(4, 0) == ""


def test_freeze_and_zoom(win):
    win.freeze("row")
    assert win.view.frozen == (1, 0) and "row" in win.view.overlays
    win.freeze("col")
    assert win.view.frozen == (0, 1)
    win.freeze(None)
    assert not win.view.overlays
    win.set_zoom(1.5)
    assert win.view.columnWidth(0) == 96
    win.set_zoom(1.0)


def test_grid_grows(win):
    m = win.view.sheet_model
    assert m.rowCount() >= 1000 and m.columnCount() >= 26
    win.view.goto(5000, 40)
    assert m.rowCount() > 5000 and m.columnCount() > 40


def test_formula_bar_commit(win):
    win.view.goto(0, 0)
    win.formula_edit.setText("=1+2")
    win.formula_edit.textEdited.emit("=1+2")
    win.formula_bar_commit(move=False)
    assert win.book.active.get_raw(0, 0) == "=1+2"
    assert win.view.sheet_model.display(0, 0) == "3"
    win.name_box.setEditText("C7")
    win._name_box_entered()
    assert win.view.current_cell() == (6, 2)


def test_save_and_open_roundtrip(win, tmp_path):
    win.user_set_cell(0, 0, "42")
    path = str(tmp_path / "book.xlsx")
    assert win._save_to(path)
    assert not win.modified
    from portable_excel import fileio
    book = fileio.load_workbook(path)
    assert book.active.get_raw(0, 0) in ("42", "42.0")
    assert config.recent_files()[0] == os.path.abspath(path)


def test_msgbox_mapping(app):
    from portable_excel.gui.vba_host import VB_NO, VB_YES, build_msgbox
    box, mapping = build_msgbox("Continue?", 4 + 32, "T")
    assert sorted(mapping.values()) == [VB_YES, VB_NO]
    assert box.icon() == QMessageBox.Question


def test_vba_editor_and_macro_list(win):
    ed = win.show_vba_editor()
    m = ed.insert_module("standard")
    code_ed = ed.editors[m["name"].lower()]
    code_ed.setPlainText("Sub Hello()\n    Range(\"A1\").Value = 5\nEnd Sub\n")
    assert win.book.vba_modified
    assert "Hello" in win.list_macros()
    ed.close()


def test_smoke_subprocess():
    env = dict(os.environ, PORTABLEEXCEL_SMOKE_TEST="1", PORTABLEEXCEL_NO_LOCK="1",
               QT_QPA_PLATFORM="offscreen")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "run.py")], env=env, timeout=60,
                       cwd=ROOT, capture_output=True)
    assert r.returncode == 0, r.stderr.decode(errors="replace")


def test_unlock_skipped_with_env(app):
    assert security.lock_disabled()
    assert security.unlock() is True


def test_run_macro_and_security_bar(win, monkeypatch):
    from portable_excel.gui import vba_host
    shown = []
    monkeypatch.setattr(vba_host.GuiVBAHost, "msgbox",
                        lambda self, prompt, buttons=0, title="": shown.append(prompt) or 1)
    win.book.vba_modules = [{"name": "Module1", "type": "standard", "code":
                             'Sub Fill()\n    Range("A1").Value = 7\n    MsgBox "done"\nEnd Sub\n'}]
    win.macros_enabled = False
    win.after_load()
    assert win.message_bar.isVisible()
    win.enable_macros()
    assert not win.message_bar.isVisible()
    assert "Fill" in " ".join(win.list_macros())
    assert win.run_macro("Fill")
    assert win.view.sheet_model.display(0, 0) == "7"
    assert shown == ["done"]


def test_point_mode_inserts_reference(win):
    win.view.goto(5, 5)
    win.view.start_edit("=")
    ed = win.view.delegate.editor
    assert ed is not None and ed.can_point()
    rect = win.view.visualRect(win.view.sheet_model.index(1, 2))
    QTest.mouseClick(win.view.viewport(), Qt.LeftButton, Qt.NoModifier, rect.center())
    assert ed.text() == "=C2"
    QTest.keyClick(ed, Qt.Key_Return)
    assert win.book.active.get_raw(5, 5) == "=C2"


def test_fill_handle_drag(win):
    sh = win.book.active
    win.user_set_cell(0, 0, "Mon")
    win.view.goto(0, 0)
    vp = win.view.viewport()
    rr = win.view.range_rect(0, 0, 0, 0)
    QTest.mousePress(vp, Qt.LeftButton, Qt.NoModifier, rr.bottomRight())
    target = win.view.visualRect(win.view.sheet_model.index(2, 0)).center()
    QTest.mouseMove(vp, target)
    from PySide6.QtGui import QMouseEvent
    from PySide6.QtCore import QEvent, QPointF
    ev = QMouseEvent(QEvent.MouseMove, QPointF(target), QPointF(vp.mapToGlobal(target)),
                     Qt.NoButton, Qt.LeftButton, Qt.NoModifier)
    QApplication.sendEvent(vp, ev)
    QTest.mouseRelease(vp, Qt.LeftButton, Qt.NoModifier, target)
    assert [sh.get_raw(r, 0) for r in range(3)] == ["Mon", "Tue", "Wed"]

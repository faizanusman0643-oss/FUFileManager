"""Render screenshots of every PortableExcel feature (offscreen, no user input).

    cd /home/user/FUFileManager/PortableExcel && python3 tools/screenshots.py

PNG files are written to ./screenshots/.  Uses the real engine/vba/fileio
modules when available; otherwise falls back to the throwaway stubs in
tests/stubs (the script prints which ones were used).
"""
import importlib
import os
import sys
import tempfile

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ["PORTABLEEXCEL_NO_LOCK"] = "1"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
OUT = os.path.join(ROOT, "screenshots")

USED_STUBS = []
for _name in ("engine", "vba"):
    try:
        _m = importlib.import_module(f"portable_excel.{_name}")
        needed = "FormulaEngine" if _name == "engine" else "VBARuntime"
        if not hasattr(_m, needed):
            raise ImportError(f"{_name} lacks {needed}")
    except Exception:
        _stub = importlib.import_module(f"tests.stubs.{_name}")
        sys.modules[f"portable_excel.{_name}"] = _stub
        setattr(importlib.import_module("portable_excel"), _name, _stub)
        USED_STUBS.append(_name)

from PySide6.QtCore import QPoint, QRect, Qt, QTimer  # noqa: E402
from PySide6.QtGui import QColor, QPainter, QPixmap  # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox, QWidget  # noqa: E402

from portable_excel.gui import config  # noqa: E402

# keep the user's real ini untouched
_tmp = tempfile.mkdtemp(prefix="pe_shots_")
config.set_ini_path(os.path.join(_tmp, "PortableExcel.ini"))

from portable_excel.gui import dialogs, security  # noqa: E402
from portable_excel.gui import main_window as mw  # noqa: E402
from portable_excel.gui.app import create_app  # noqa: E402
from portable_excel.gui.vba_host import build_msgbox  # noqa: E402
from portable_excel.model import Workbook  # noqa: E402

W, H = 1366, 768

DEMO_MACRO = '''Option Explicit

' Fills a multiplication table and reports when done
Sub FillTable()
    Dim i As Integer, j As Integer
    Range("A1").Value = "x"
    For i = 1 To 10
        Cells(1, i + 1).Value = i
        Cells(i + 1, 1).Value = i
        For j = 1 To 10
            Cells(i + 1, j + 1).Value = i * j
        Next j
    Next i
    Debug.Print "Table filled"
    MsgBox "The multiplication table is ready.", vbInformation, "PortableExcel"
End Sub

Sub Greet()
    MsgBox "Hello from VBA!"
End Sub
'''

SHEET_MODULE = '''Private Sub Worksheet_Change(ByVal Target As Range)
    ' runs after every edit on this sheet
    Application.StatusBar = "Changed " & Target.Address
End Sub
'''


def pump(n=3):
    for _ in range(n):
        QApplication.processEvents()


def save(pm: QPixmap, name: str):
    path = os.path.join(OUT, name + ".png")
    pm.save(path)
    print("  wrote", os.path.relpath(path, ROOT))


def compose(base: QWidget, overlay: QWidget, pos=None) -> QPixmap:
    """Window grab with a dialog/menu grab painted on top (with a shadow)."""
    pump()
    pm = base.grab()
    op = overlay.grab()
    p = QPainter(pm)
    if pos is None:
        pos = QPoint((pm.width() - op.width()) // 2, (pm.height() - op.height()) // 2)
    for i in range(6, 0, -1):
        p.fillRect(QRect(pos.x() + i, pos.y() + i, op.width(), op.height()),
                   QColor(0, 0, 0, 12))
    p.drawPixmap(pos, op)
    p.setPen(QColor("#8A8A8A"))
    p.drawRect(QRect(pos.x() - 1, pos.y() - 1, op.width() + 1, op.height() + 1))
    p.end()
    return pm


def dialog_shot(win, dlg, name, title_bar=True):
    dlg.show()
    pump(5)
    if title_bar:
        frame = framed(dlg)
        pm = compose_pixmap(win, frame)
    else:
        pm = compose(win, dlg)
    save(pm, name)
    dlg.close()
    pump()


def framed(dlg: QWidget) -> QPixmap:
    """Dialog grab with a simple drawn window title bar (offscreen has none)."""
    inner = dlg.grab()
    bar = 30
    pm = QPixmap(inner.width() + 2, inner.height() + bar + 1)
    pm.fill(QColor("#FFFFFF"))
    p = QPainter(pm)
    p.fillRect(0, 0, pm.width(), bar, QColor("#FFFFFF"))
    p.setPen(QColor("#262626"))
    f = dlg.font()
    p.setFont(f)
    p.drawText(QRect(12, 0, pm.width() - 60, bar), Qt.AlignVCenter | Qt.AlignLeft,
               dlg.windowTitle())
    p.drawText(QRect(pm.width() - 34, 0, 30, bar), Qt.AlignCenter, "✕")
    p.drawPixmap(1, bar, inner)
    p.setPen(QColor("#217346"))
    p.drawRect(0, 0, pm.width() - 1, pm.height() - 1)
    p.end()
    return pm


def compose_pixmap(base: QWidget, op: QPixmap, pos=None) -> QPixmap:
    pump()
    pm = base.grab()
    p = QPainter(pm)
    if pos is None:
        pos = QPoint((pm.width() - op.width()) // 2, (pm.height() - op.height()) // 2)
    for i in range(8, 0, -1):
        p.fillRect(QRect(pos.x() + i, pos.y() + i, op.width(), op.height()), QColor(0, 0, 0, 10))
    p.drawPixmap(pos, op)
    p.end()
    return pm


def menu_shot(win, menu, name):
    pump(5)
    pos = win.mapFromGlobal(menu.pos())
    save(compose(win, menu, pos), name)
    menu.close()
    pump()


def sales_book() -> Workbook:
    book = Workbook()
    sh = book.active
    sh.name = "Sales"
    head = ["Region", "Product", "Units", "Unit Price", "Revenue", "Share"]
    data = [("North", "Laptop", 120, 899.0), ("North", "Monitor", 340, 189.5),
            ("South", "Laptop", 95, 905.0), ("South", "Keyboard", 610, 24.99),
            ("East", "Monitor", 275, 192.0), ("East", "Laptop", 130, 879.0),
            ("West", "Keyboard", 450, 25.49), ("West", "Monitor", 210, 186.0)]
    for c, t in enumerate(head):
        sh.set_raw(0, c, t)
        sh.update_style(0, c, bold=True, fill="#217346", color="#FFFFFF", align="center")
    for i, (reg, prod, units, price) in enumerate(data, start=1):
        r = i
        sh.set_raw(r, 0, reg)
        sh.set_raw(r, 1, prod)
        sh.set_raw(r, 2, str(units))
        sh.set_raw(r, 3, str(price))
        sh.set_raw(r, 4, f"=C{r + 1}*D{r + 1}")
        sh.set_raw(r, 5, f"=E{r + 1}/$E${len(data) + 2}")
        sh.update_style(r, 3, number_format="$#,##0.00")
        sh.update_style(r, 4, number_format="$#,##0.00")
        sh.update_style(r, 5, number_format="0.0%")
        if i % 2 == 0:
            for c in range(6):
                sh.update_style(r, c, fill="#E2EFDA")
    t = len(data) + 2
    sh.set_raw(t - 1, 0, "Total")
    sh.update_style(t - 1, 0, bold=True)
    sh.set_raw(t - 1, 2, f"=SUM(C2:C{t - 1})")
    sh.set_raw(t - 1, 4, f"=SUM(E2:E{t - 1})")
    sh.set_raw(t - 1, 5, f"=SUM(F2:F{t - 1})")
    for c in (2, 4, 5):
        sh.update_style(t - 1, c, bold=True)
    sh.update_style(t - 1, 4, bold=True, number_format="$#,##0.00")
    sh.update_style(t - 1, 5, bold=True, number_format="0.0%")
    sh.set_raw(12, 0, "Average revenue per order")
    sh.set_raw(12, 4, "=AVERAGE(E2:E9)")
    sh.update_style(12, 4, number_format="$#,##0.00", italic=True)
    sh.set_raw(13, 0, "Best product")
    sh.set_raw(13, 4, '=INDEX(B2:B9,MATCH(MAX(E2:E9),E2:E9,0))')
    sh.col_widths.update({0: 92, 1: 90, 3: 86, 4: 110, 5: 70})
    s2 = book.add_sheet("Summary")
    s2.set_raw(0, 0, "Quarter")
    s2.set_raw(0, 1, "Target")
    for i, q in enumerate(["Q1", "Q2", "Q3", "Q4"], 1):
        s2.set_raw(i, 0, q)
        s2.set_raw(i, 1, str(50000 + i * 7500))
    book.add_sheet("Notes")
    book._pe_title = "Sales2026.xlsx"
    return book


def new_window(book=None):
    w = mw.MainWindow(book)
    w.resize(W, H)
    w.show()
    pump(5)
    return w


def main():
    os.makedirs(OUT, exist_ok=True)
    app = create_app(["screenshots"])
    # never block on modal message boxes
    QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)
    QMessageBox.information = staticmethod(lambda *a, **k: QMessageBox.Ok)
    QMessageBox.warning = staticmethod(lambda *a, **k: QMessageBox.Ok)
    print("Stub modules used:", ", ".join(USED_STUBS) or "none (real engine/vba)")

    # 01 lock screen --------------------------------------------------------
    lock = security.LockDialog()
    lock.edit.setText("1234")
    lock.show()
    pump(5)
    bg = QPixmap(W, H)
    bg.fill(QColor("#E9EEF2"))
    p = QPainter(bg)
    fr = framed(lock)
    p.drawPixmap((W - fr.width()) // 2, (H - fr.height()) // 2, fr)
    p.end()
    save(bg, "01_lock_screen")
    lock.close()

    # 02 main window ----------------------------------------------------------
    win = new_window(sales_book())
    win.view.select_range(1, 4, 8, 4)
    pump(5)
    save(win.grab(), "02_main_window_home_ribbon")

    # 03 formula bar editing ------------------------------------------------------
    win.view.goto(10, 4)
    win.view.start_edit(None)
    pump()
    ed = win.view.delegate.editor
    if ed is not None:
        ed.setText("=SUM(E2:E9)*1.05")
        ed.textEdited.emit(ed.text())
    pump(5)
    save(win.grab(), "03_formula_bar_editing")
    if ed is not None:
        win.view.finish_edit(ed, commit=False)
    pump()

    # 04 insert tab ---------------------------------------------------------------------
    win.ribbon.show_page("Insert")
    win.view.goto(3, 1)
    save(win.grab(), "04_insert_tab")

    # 05 formulas tab + insert function ---------------------------------------------
    win.ribbon.show_page("Formulas")
    d = dialogs.InsertFunctionDialog(mw.engine_mod.FUNCTIONS, win)
    d.search.setText("IF")
    dialog_shot(win, d, "05_formulas_tab_insert_function_dialog")

    # 06 data tab sorted -------------------------------------------------------------------
    win.ribbon.show_page("Data")
    win.view.goto(1, 4)
    win.sort_selection(True)
    win.view.select_range(0, 0, 8, 5, cur=(1, 4))
    save(win.grab(), "06_data_tab_sorted")
    win.undo_stack.undo()

    # 07 view tab freeze -----------------------------------------------------------------------
    win.ribbon.show_page("View")
    sh = win.book.active
    for r in range(15, 60):
        sh.set_raw(r, 0, f"Row {r + 1}")
        sh.set_raw(r, 2, str(r * 3))
    win.view.refresh()
    win.freeze("row")
    win.view.verticalScrollBar().setValue(20)
    win.view.goto(25, 2)
    save(win.grab(), "07_view_tab_freeze")
    win.freeze(None)
    for r in range(15, 60):
        sh.set_raw(r, 0, "")
        sh.set_raw(r, 2, "")
    win.view.verticalScrollBar().setValue(0)
    win.view.goto(0, 0)

    # 08 developer tab -------------------------------------------------------------------------
    win.ribbon.show_page("Developer")
    save(win.grab(), "08_developer_tab")

    # 09 format cells -----------------------------------------------------------------------
    win.ribbon.show_page("Home")
    win.view.goto(1, 4)
    r, c = win.view.current_cell()
    val = win.view.sheet_model.value(r, c)
    d = dialogs.FormatCellsDialog(sh.get_style(r, c), lambda f: win.preview_format(val, f), win)
    d.cat.setCurrentRow(2)
    dialog_shot(win, d, "09_format_cells_dialog")

    # 10 find & replace -------------------------------------------------------------------------
    fd = win.show_find(True)
    fd.what.setEditText("Laptop")
    fd.with_.setEditText("Notebook")
    win.find_all(fd.options())
    pump()
    save(compose_pixmap(win, framed(fd)), "10_find_replace")
    fd.close()

    # 11 sheet tab context menu ---------------------------------------------------------------
    tabs = win.sheet_tabs
    pos = tabs.mapToGlobal(tabs.tabRect(1).center())
    m = win._sheet_menu(1, pos - QPoint(0, 230))
    menu_shot(win, m, "11_sheet_tabs_context_menu")

    # 12 macro security bar ---------------------------------------------------------------------
    mbook = sales_book()
    mbook._pe_title = "Macros2026.xlsm"
    mbook.vba_modules = [{"name": "ThisWorkbook", "type": "document",
                          "code": "Private Sub Workbook_Open()\n    MsgBox \"Welcome\"\nEnd Sub\n"},
                         {"name": "Sheet1", "type": "document", "code": SHEET_MODULE},
                         {"name": "Module1", "type": "standard", "code": DEMO_MACRO}]
    mwin = new_window(mbook)
    mwin.after_load()
    pump(5)
    save(mwin.grab(), "12_macro_security_bar")

    # enable macros without running Workbook_Open's MsgBox modally
    mwin.message_bar.hide()
    mwin.macros_enabled = True
    mwin.ensure_runtime()

    # 13 macros dialog ----------------------------------------------------------------------------
    d = dialogs.MacrosDialog(mwin.list_macros(), mwin, mwin.book_title())
    dialog_shot(mwin, d, "13_macros_dialog")

    # 14 VBA editor ---------------------------------------------------------------------------------
    ed = mwin.show_vba_editor()
    ed.resize(W, H)
    ce = ed.open_module("Module1")
    ed.immediate.print_line("?Range(\"E11\").Value")
    ed.immediate.print_line(mwin.immediate_execute('?Range("E11").Value') or "")
    if ce is not None:
        from PySide6.QtGui import QTextCursor
        blk = ce.document().findBlockByNumber(10)
        ce.setTextCursor(QTextCursor(blk))
    pump(5)
    save(ed.grab(), "14_vba_editor_with_code")
    ed.hide()

    # 15 MsgBox from a macro ------------------------------------------------------------------------
    box, _ = build_msgbox("The multiplication table is ready.\n\nDo you want to format it?",
                          4 + 32, "PortableExcel", mwin)
    box.show()
    pump(5)
    save(compose_pixmap(mwin, framed(box)), "15_msgbox_from_macro")
    box.close()

    # 16 macro result ------------------------------------------------------------------------------
    rbook = Workbook()
    rbook.vba_modules = [{"name": "Module1", "type": "standard", "code": DEMO_MACRO}]
    rwin = new_window(rbook)
    rwin.macros_enabled = True
    from portable_excel.gui import vba_host
    vba_host.GuiVBAHost.msgbox = lambda self, prompt, buttons=0, title="": 1
    rwin.run_macro("FillTable")
    rwin.set_status_message("Macro FillTable finished")
    rwin.view.select_range(1, 1, 10, 10)
    pump(5)
    save(rwin.grab(), "16_macro_result_grid")

    # 17 change password -----------------------------------------------------------------------------
    d = security.ChangePasswordDialog(win)
    d.old.setText("12345643")
    d.new.setText("newpass1")
    d.confirm.setText("newpass1")
    dialog_shot(win, d, "17_change_password_dialog")

    # 18 open password prompt ---------------------------------------------------------------------------
    d = dialogs.PasswordDialog("Password", "'Budget2026.xlsx' is protected.",
                               "The password you supplied is not correct. Verify that the CAPS "
                               "LOCK key is off and be sure to use the correct capitalization.",
                               False, win)
    d.pw.setText("secret")
    dialog_shot(win, d, "18_open_password_prompt")

    # 19 cell context menu ------------------------------------------------------------------------------
    win.view.select_range(2, 1, 4, 3)
    vp = win.view.viewport()
    rect = win.view.visualRect(win.view.sheet_model.index(3, 2))
    m = win.show_cell_menu(vp.mapToGlobal(rect.center()))
    menu_shot(win, m, "19_context_menu")

    # 20 save warnings ---------------------------------------------------------------------------------------
    warnings = []
    try:
        from portable_excel import fileio
        warnings = fileio.save_workbook(win.book, os.path.join(_tmp, "out.csv"), win.engine)
    except Exception as e:
        print("  (fileio warnings unavailable:", e, ")")
    if not warnings:
        warnings = ["CSV keeps only the active sheet; 2 other sheets were not saved.",
                    "Cell formatting (fonts, colours, number formats) is not stored in CSV.",
                    "Formulas were saved as their calculated values."]
    d = win.show_save_warnings(warnings, modal=False)
    dialog_shot(win, d, "20_save_warnings")

    for w in (win, mwin, rwin):
        w.modified = False
        w.close()
    pump()
    print("Done:", len([f for f in os.listdir(OUT) if f.endswith(".png")]), "screenshots in", OUT)
    if USED_STUBS:
        print("NOTE: rendered with stub modules:", ", ".join(USED_STUBS),
              "- re-run after integration.")


if __name__ == "__main__":
    main()

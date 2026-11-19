"""PortableExcel main window."""
from __future__ import annotations

import importlib
import logging
import os
import re
import traceback
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import QEvent, QObject, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QAction, QColor, QFont, QGuiApplication, QKeySequence, QUndoStack
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog, QFontComboBox, QFrame,
                               QHBoxLayout, QInputDialog, QLabel, QLineEdit, QMainWindow,
                               QMenu, QMessageBox, QPushButton, QScrollBar, QSizePolicy,
                               QSlider, QStackedWidget, QTabBar, QToolButton, QVBoxLayout,
                               QWidget)

from .. import engine as engine_mod
from .. import fileio
from ..model import (MAX_COLS, MAX_ROWS, Workbook, a1, col_to_name, literal_value, parse_a1,
                     parse_range_a1, value_to_text)
from . import config, dialogs, icons, ops, security
from .backstage import Backstage
from .commands import UNDO_LIMIT, CellsCommand, SnapshotCommand
from .grid import FormulaLineEdit, SpreadsheetView
from .ribbon import Ribbon, TitleStrip
from .vba_host import GuiVBAHost


def _vba():
    """The VBA interpreter module, or None (resolved lazily so a module that
    is still missing or broken never prevents the spreadsheet from working)."""
    try:
        return importlib.import_module("portable_excel.vba")
    except Exception:
        log.exception("VBA interpreter could not be imported")
        return None


log = logging.getLogger("PortableExcel")

_WINDOWS: List["MainWindow"] = []
_book_counter = [0]

LARGE_RANGE_ROWS = 5000
LARGE_RANGE_COLS = 200


def _vba_error_class():
    mod = _vba()
    return getattr(mod, "VBAError", None) if mod is not None else None


# --------------------------------------------------------------------------
# Small widgets
# --------------------------------------------------------------------------
class MessageBar(QWidget):
    enable_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("messageBar")
        self.setAttribute(Qt.WA_StyledBackground, True)
        h = QHBoxLayout(self)
        h.setContentsMargins(10, 4, 6, 4)
        ic = QLabel()
        ic.setPixmap(icons.pixmap("security", 18))
        h.addWidget(ic)
        b = QLabel("SECURITY WARNING")
        b.setObjectName("messageBold")
        h.addWidget(b)
        self.text = QLabel("Macros have been disabled.")
        h.addWidget(self.text)
        self.button = QPushButton("Enable Content")
        self.button.setObjectName("enableContent")
        self.button.clicked.connect(self.enable_clicked.emit)
        h.addWidget(self.button)
        h.addStretch(1)
        close = QToolButton()
        close.setText("✕")
        close.setAutoRaise(True)
        close.clicked.connect(self.hide)
        h.addWidget(close)


class SheetTabs(QTabBar):
    rename_requested = Signal(int, str)
    menu_requested = Signal(int, object)
    add_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("sheetTabs")
        self.setDrawBase(False)
        self.setExpanding(False)
        self.setMovable(False)
        self.setUsesScrollButtons(True)
        self.setElideMode(Qt.ElideNone)
        self._editor = None

    def mouseDoubleClickEvent(self, e):
        i = self.tabAt(e.position().toPoint())
        if i >= 0:
            self.start_rename(i)
        else:
            self.add_requested.emit()

    def start_rename(self, i):
        rect = self.tabRect(i)
        ed = QLineEdit(self.tabText(i), self)
        ed.setGeometry(rect.adjusted(2, 2, -2, -2))
        ed.selectAll()
        ed.setStyleSheet("border:1px solid #217346;")
        ed.show()
        ed.setFocus()
        self._editor = ed
        done = [False]

        def finish(commit=True):
            if done[0]:
                return
            done[0] = True
            text = ed.text().strip()
            ed.deleteLater()
            self._editor = None
            if commit and text and text != self.tabText(i):
                self.rename_requested.emit(i, text)

        ed.editingFinished.connect(lambda: finish(True))
        ed.installEventFilter(self)
        ed._finish = finish

    def eventFilter(self, obj, ev):
        if isinstance(obj, QLineEdit) and ev.type() == QEvent.KeyPress and \
                ev.key() == Qt.Key_Escape:
            obj._finish(False)
            return True
        return super().eventFilter(obj, ev)

    def contextMenuEvent(self, e):
        i = self.tabAt(e.pos())
        self.menu_requested.emit(i, e.globalPos())


class _MacroStopFilter(QObject):
    """While a macro runs: Esc / Ctrl+Break request a stop, user input to the
    workbook window is swallowed."""

    def __init__(self, win):
        super().__init__()
        self.win = win

    def eventFilter(self, obj, ev):
        t = ev.type()
        if t == QEvent.KeyPress:
            if ev.key() in (Qt.Key_Escape, Qt.Key_Pause, Qt.Key_Cancel) or (
                    ev.key() == Qt.Key_Break):
                self.win.stop_macro()
                return True
        if t in (QEvent.KeyPress, QEvent.MouseButtonPress, QEvent.MouseButtonDblClick,
                 QEvent.Wheel):
            w = obj if isinstance(obj, QWidget) else None
            if w is not None and w.window() is self.win:
                return True
        return False


# --------------------------------------------------------------------------
# Main window
# --------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, book: Optional[Workbook] = None):
        super().__init__()
        _WINDOWS.append(self)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowIcon(icons.app_icon())
        self.book = book or Workbook()
        if not getattr(self.book, "_pe_title", None) and not self.book.path:
            _book_counter[0] += 1
            self.book._pe_title = f"Book{_book_counter[0]}"
        self.engine = engine_mod.FormulaEngine(self.book)
        self.zoom = 1.0
        self.show_formulas = False
        self.modified = False
        self.save_password: Optional[str] = None
        self.clip = None
        self._marquee = None
        self.runtime = None
        self.vba_stale = True
        self.macros_enabled = not self.book.vba_modules
        self.macro_running = False
        self.immediate_log: List[str] = []
        self.vba_editor = None
        self.find_dialog = None
        self.fb_editing = False
        self._refresh_pending = False
        self._frozen: Dict[int, Tuple[int, int]] = {}
        self._sheet_sel: Dict[int, Tuple] = {}
        self._suspend_events = False
        self.undo_stack = QUndoStack(self)
        self.undo_stack.setUndoLimit(UNDO_LIMIT)
        self.book.change_listeners.append(self._on_book_change)
        self._build_ui()
        self.resize(1366, 768)
        self.update_title()
        self.rebuild_sheet_tabs()
        self.on_current_changed()

    # ======================================================================
    # UI construction
    # ======================================================================
    def _act(self, name, text, slot=None, shortcut=None, icon=None, checkable=False, tip=None):
        a = QAction(text, self)
        if icon:
            a.setIcon(icons.icon(icon))
        if shortcut:
            seqs = shortcut if isinstance(shortcut, (list, tuple)) else [shortcut]
            a.setShortcuts([QKeySequence(s) for s in seqs])
            a.setShortcutContext(Qt.WindowShortcut)
        a.setCheckable(checkable)
        if slot is not None:
            a.triggered.connect(slot)
        sc = a.shortcut().toString(QKeySequence.NativeText)
        a.setToolTip(tip or (f"{text.replace('&', '')} ({sc})" if sc else text.replace("&", "")))
        self.addAction(a)
        self.actions[name] = a
        return a

    def _build_actions(self):
        self.actions: Dict[str, QAction] = {}
        A = self._act
        A("new", "New", self.new_workbook, "Ctrl+N", "new")
        A("open", "Open", lambda: self.open_file(), "Ctrl+O", "open")
        A("save", "Save", self.save, "Ctrl+S", "save")
        A("save_as", "Save As", self.save_as, "F12")
        A("save_pw", "Save with password...", self.save_with_password, None, "lock")
        A("change_pw", "Change app password...", self.change_app_password, None, "key")
        A("exit", "Exit", self.exit_app)
        A("undo", "Undo", self.undo_stack.undo, "Ctrl+Z", "undo")
        A("redo", "Redo", self.undo_stack.redo, ["Ctrl+Y", "Ctrl+Shift+Z"], "redo")
        self.undo_stack.canUndoChanged.connect(self.actions["undo"].setEnabled)
        self.undo_stack.canRedoChanged.connect(self.actions["redo"].setEnabled)
        self.actions["undo"].setEnabled(False)
        self.actions["redo"].setEnabled(False)
        A("cut", "Cut", lambda: self.copy(cut=True), "Ctrl+X", "cut")
        A("copy", "Copy", self.copy, "Ctrl+C", "copy")
        A("paste", "Paste", self.paste, "Ctrl+V", "paste")
        A("paste_values", "Paste Values", lambda: self.paste(values_only=True), "Ctrl+Shift+V",
          "paste_values")
        A("bold", "Bold", lambda: self.toggle_style("bold"), ["Ctrl+B", "Ctrl+2"], "bold",
          checkable=True)
        A("italic", "Italic", lambda: self.toggle_style("italic"), ["Ctrl+I", "Ctrl+3"],
          "italic", checkable=True)
        A("underline", "Underline", lambda: self.toggle_style("underline"), ["Ctrl+U", "Ctrl+4"],
          "underline", checkable=True)
        A("align_left", "Align Left", lambda: self.set_align("left"), None, "align_left",
          checkable=True)
        A("align_center", "Center", lambda: self.set_align("center"), None, "align_center",
          checkable=True)
        A("align_right", "Align Right", lambda: self.set_align("right"), None, "align_right",
          checkable=True)
        A("wrap", "Wrap Text", lambda: self.toggle_style("wrap"), None, "wrap", checkable=True)
        A("font_color", "Font Color", lambda: self.apply_style(color=self._font_color),
          None, "font_color")
        A("fill_color", "Fill Color", lambda: self.apply_style(fill=self._fill_color),
          None, "fill_color")
        A("percent", "Percent Style", lambda: self.set_number_format("0%"), "Ctrl+Shift+%",
          "percent")
        A("comma", "Comma Style", lambda: self.set_number_format("#,##0.00"), None, "comma")
        A("dec_more", "Increase Decimal", lambda: self.change_decimals(1), None, "dec_more")
        A("dec_less", "Decrease Decimal", lambda: self.change_decimals(-1), None, "dec_less")
        A("insert_rows", "Insert Sheet Rows", lambda: self.insert_cells("row"), "Ctrl++",
          "insert_row")
        A("insert_cols", "Insert Sheet Columns", lambda: self.insert_cells("col"), None,
          "insert_col")
        A("delete_rows", "Delete Sheet Rows", lambda: self.delete_cells("row"), "Ctrl+-",
          "delete_row")
        A("delete_cols", "Delete Sheet Columns", lambda: self.delete_cells("col"), None,
          "delete_col")
        A("row_height", "Row Height...", self.ask_row_height)
        A("col_width", "Column Width...", self.ask_col_width)
        A("autofit_col", "AutoFit Column Width", self.autofit_selection)
        A("format_cells", "Format Cells...", self.format_cells, "Ctrl+1")
        A("autosum", "AutoSum", lambda: self.autosum("SUM"), "Alt+=", "autosum")
        A("fill_down", "Fill Down", self.fill_down, "Ctrl+D", "fill_down")
        A("fill_right", "Fill Right", self.fill_right, "Ctrl+R")
        A("clear_contents", "Clear Contents", self.clear_contents)
        A("clear_formats", "Clear Formats", self.clear_formats)
        A("clear_all", "Clear All", self.clear_all, None, "clear")
        A("sort_asc", "Sort A to Z", lambda: self.sort_selection(False), None, "sort_asc")
        A("sort_desc", "Sort Z to A", lambda: self.sort_selection(True), None, "sort_desc")
        A("remove_dups", "Remove Duplicates", self.remove_duplicates, None, "duplicates")
        A("find", "Find...", lambda: self.show_find(False), "Ctrl+F", "find")
        A("replace", "Replace...", lambda: self.show_find(True), "Ctrl+H")
        A("goto", "Go To...", self.goto_dialog, ["Ctrl+G", "F5"])
        A("insert_function", "Insert Function", self.insert_function, "Shift+F3", "function")
        A("calc_now", "Calculate Now", self.calculate_now, "F9", "calc")
        A("show_formulas", "Show Formulas", self.toggle_show_formulas, "Ctrl+`", "function",
          checkable=True)
        A("define_name", "Define Name...", self.define_name, None, "names")
        A("name_manager", "Name Manager", self.name_manager, "Ctrl+F3", "names")
        A("insert_sheet", "Insert Sheet", self.add_sheet, "Shift+F11", "sheet")
        A("format_table", "Format as Table", self.format_as_table, None, "table")
        A("insert_date", "Insert Date", self.insert_today, "Ctrl+;")
        A("zoom_in", "Zoom In", lambda: self.set_zoom(self.zoom + 0.1), "Ctrl+=", "zoom_in")
        A("zoom_out", "Zoom Out", lambda: self.set_zoom(self.zoom - 0.1), None, "zoom_out")
        A("zoom_100", "100%", lambda: self.set_zoom(1.0), None, "zoom_100")
        A("zoom_dialog", "Zoom", self.zoom_dialog, None, "find")
        A("zoom_sel", "Zoom to Selection", self.zoom_to_selection, None, "zoom_in")
        A("gridlines", "Gridlines", self.toggle_gridlines, None, "gridlines", checkable=True)
        self.actions["gridlines"].setChecked(True)
        A("headings", "Headings", self.toggle_headings, None, "names", checkable=True)
        self.actions["headings"].setChecked(True)
        A("formula_bar", "Formula Bar", self.toggle_formula_bar, None, "function",
          checkable=True)
        self.actions["formula_bar"].setChecked(True)
        A("freeze_panes", "Freeze Panes", lambda: self.freeze("panes"), None, "freeze_panes")
        A("freeze_row", "Freeze Top Row", lambda: self.freeze("row"), None, "freeze_row")
        A("freeze_col", "Freeze First Column", lambda: self.freeze("col"), None, "freeze_col")
        A("unfreeze", "Unfreeze Panes", lambda: self.freeze(None), None, "freeze_panes")
        A("vba_editor", "Visual Basic", self.show_vba_editor, "Alt+F11", "vba")
        A("macros", "Macros", self.show_macros_dialog, "Alt+F8", "macros")
        A("run_macro", "Run Macro", self.run_macro_prompt, None, "run")
        A("macro_security", "Macro Security", self.macro_security, None, "security")
        A("enable_macros", "Enable Content", self.enable_macros, None, "security")
        A("stop_macro", "Stop Macro", self.stop_macro, None, "stop")
        A("lock", "Lock PortableExcel", self.lock_app, None, "lock")
        A("next_sheet", "Next Sheet", lambda: self.step_sheet(1), "Ctrl+PgDown")
        A("prev_sheet", "Previous Sheet", lambda: self.step_sheet(-1), "Ctrl+PgUp")
        A("about", "About PortableExcel", self.about, None, "info")
        A("file", "File", self.show_backstage, "Alt+F")

    def _build_ui(self):
        self._font_color = "#FF0000"
        self._fill_color = "#FFFF00"
        self._build_actions()
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        main = QWidget()
        v = QVBoxLayout(main)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        # title strip
        self.title_strip = TitleStrip()
        self.title_strip.add_action(self.actions["save"], "save_white")
        self.title_strip.add_action(self.actions["undo"], "undo_white")
        self.title_strip.add_action(self.actions["redo"], "redo_white")
        v.addWidget(self.title_strip)
        # ribbon
        self.ribbon = Ribbon()
        self.ribbon.file_clicked.connect(self.show_backstage)
        self._build_ribbon()
        v.addWidget(self.ribbon)
        # message bar
        self.message_bar = MessageBar()
        self.message_bar.enable_clicked.connect(self.enable_macros)
        self.message_bar.hide()
        v.addWidget(self.message_bar)
        # formula bar
        self.formula_bar_widget = self._build_formula_bar()
        v.addWidget(self.formula_bar_widget)
        # grid
        self.view = SpreadsheetView(self)
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        v.addWidget(self.view, 1)
        # sheet bar
        v.addWidget(self._build_sheet_bar())
        self.stack.addWidget(main)
        self.backstage = Backstage(self)
        self.backstage.back.connect(lambda: self.stack.setCurrentIndex(0))
        self.stack.addWidget(self.backstage)
        self._build_status_bar()
        self.view.setFocus()

    # ---------------------------------------------------------------- ribbon
    def _build_ribbon(self):
        acts = self.actions
        rb = self.ribbon
        # Home ----------------------------------------------------------------
        home = rb.add_page("Home")
        g = home.add_group("Clipboard")
        pm = QMenu(self)
        pm.addAction(acts["paste"])
        pm.addAction(acts["paste_values"])
        g.add_large(acts["paste"], "paste", menu=pm)
        g.add_small_column([(acts["cut"], "cut"), (acts["copy"], "copy")])
        g = home.add_group("Font")
        self.font_box = QFontComboBox()
        self.font_box.setFixedWidth(150)
        self.font_box.setCurrentFont(QFont("Calibri"))
        self.font_box.setEditable(True)
        self.font_box.currentFontChanged.connect(
            lambda f: self.apply_style(font_name=f.family()) if not self._syncing else None)
        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26,
                                                  28, 36, 48, 72)])
        self.size_box.setCurrentText("11")
        self.size_box.setFixedWidth(52)
        self.size_box.lineEdit().returnPressed.connect(self._size_entered)
        self.size_box.activated.connect(lambda _i: self._size_entered())
        self.font_color_menu = dialogs.ColorMenu("Automatic", self)
        self.font_color_menu.picked.connect(self._pick_font_color)
        self.fill_color_menu = dialogs.ColorMenu("No Fill", self)
        self.fill_color_menu.picked.connect(self._pick_fill_color)
        self._syncing = False
        g.add_rows([[self.font_box, self.size_box],
                    [(acts["bold"], "bold"), (acts["italic"], "italic"),
                     (acts["underline"], "underline"), "|",
                     (acts["fill_color"], "fill_color", self.fill_color_menu),
                     (acts["font_color"], "font_color", self.font_color_menu)]])
        g = home.add_group("Alignment")
        g.add_rows([[(acts["align_left"], "align_left"), (acts["align_center"], "align_center"),
                     (acts["align_right"], "align_right")],
                    [(acts["wrap"], "wrap")]])
        self.align_btns = g
        g = home.add_group("Number")
        self.num_combo = QComboBox()
        self.num_combo.setFixedWidth(130)
        for label, code in dialogs.NUMBER_FORMATS:
            self.num_combo.addItem(label, code)
        self.num_combo.activated.connect(self._num_combo_chosen)
        g.add_rows([[self.num_combo],
                    [(acts["percent"], "percent"), (acts["comma"], "comma"),
                     (acts["dec_more"], "dec_more"), (acts["dec_less"], "dec_less")]])
        g = home.add_group("Cells")
        im = QMenu(self)
        im.addAction(acts["insert_rows"])
        im.addAction(acts["insert_cols"])
        im.addAction(acts["insert_sheet"])
        dm = QMenu(self)
        dm.addAction(acts["delete_rows"])
        dm.addAction(acts["delete_cols"])
        dm.addAction("Delete Sheet", self.delete_sheet)
        fm = QMenu(self)
        fm.addAction(acts["row_height"])
        fm.addAction(acts["col_width"])
        fm.addAction(acts["autofit_col"])
        fm.addSeparator()
        fm.addAction("Rename Sheet", lambda: self.rename_sheet_prompt())
        fm.addAction("Hide Sheet", self.hide_sheet)
        fm.addSeparator()
        fm.addAction(acts["format_cells"])
        ins = QAction(icons.icon("insert_row"), "Insert", self)
        ins.triggered.connect(lambda: self.insert_cells("row"))
        dele = QAction(icons.icon("delete_row"), "Delete", self)
        dele.triggered.connect(lambda: self.delete_cells("row"))
        fmt = QAction(icons.icon("table"), "Format", self)
        fmt.triggered.connect(self.format_cells)
        g.add_large(ins, "insert_row", menu=im)
        g.add_large(dele, "delete_row", menu=dm)
        g.add_large(fmt, "table", menu=fm)
        g = home.add_group("Editing")
        sm = QMenu(self)
        for label, fn in (("Sum", "SUM"), ("Average", "AVERAGE"), ("Count Numbers", "COUNT"),
                          ("Max", "MAX"), ("Min", "MIN")):
            sm.addAction(label, lambda f=fn: self.autosum(f))
        sm.addSeparator()
        sm.addAction(acts["insert_function"])
        cm = QMenu(self)
        cm.addAction(acts["clear_all"])
        cm.addAction(acts["clear_formats"])
        cm.addAction(acts["clear_contents"])
        sortm = QMenu(self)
        sortm.addAction(acts["sort_asc"])
        sortm.addAction(acts["sort_desc"])
        findm = QMenu(self)
        findm.addAction(acts["find"])
        findm.addAction(acts["replace"])
        findm.addAction(acts["goto"])
        g.add_small_column([(acts["autosum"], "autosum", sm), (acts["fill_down"], "fill_down"),
                            (acts["clear_all"], "clear", cm)])
        sort_act = QAction(icons.icon("sort_asc"), "Sort && Filter", self)
        sort_act.triggered.connect(lambda: self.sort_selection(False))
        find_act = QAction(icons.icon("find"), "Find && Select", self)
        find_act.triggered.connect(lambda: self.show_find(False))
        g.add_large(sort_act, "sort_asc", menu=sortm)
        g.add_large(find_act, "find", menu=findm)
        # Insert ----------------------------------------------------------------
        ins = rb.add_page("Insert")
        g = ins.add_group("Tables")
        g.add_large(acts["format_table"], "table")
        g = ins.add_group("Functions")
        fn_act = QAction(icons.icon("function"), "Function...", self)
        fn_act.triggered.connect(self.insert_function)
        g.add_large(fn_act, "function")
        g.add_large(acts["autosum"], "autosum")
        g = ins.add_group("Cells")
        g.add_small_column([(acts["insert_rows"], "insert_row"), (acts["insert_cols"], "insert_col"),
                            (acts["insert_sheet"], "sheet")])
        g = ins.add_group("Text")
        g.add_small_column([(acts["insert_date"], "calc")])
        # Formulas ---------------------------------------------------------------
        fo = rb.add_page("Formulas")
        g = fo.add_group("Function Library")
        g.add_large(fn_act, "function")
        g.add_large(acts["autosum"], "autosum", menu=sm)
        self._function_menus(g)
        g = fo.add_group("Defined Names")
        g.add_large(acts["name_manager"], "names")
        g.add_small_column([(acts["define_name"], "names")])
        g = fo.add_group("Formula Auditing")
        g.add_small_column([(acts["show_formulas"], "function")])
        g = fo.add_group("Calculation")
        g.add_large(acts["calc_now"], "calc")
        # Data -----------------------------------------------------------------------
        da = rb.add_page("Data")
        g = da.add_group("Sort & Filter")
        g.add_large(acts["sort_asc"], "sort_asc")
        g.add_large(acts["sort_desc"], "sort_desc")
        g = da.add_group("Data Tools")
        g.add_large(acts["remove_dups"], "duplicates")
        g = da.add_group("Find")
        g.add_large(QAction(icons.icon("find"), "Find && Replace", self,
                            triggered=lambda: self.show_find(True)), "find")
        g.add_small_column([(acts["goto"], "find")])
        g = da.add_group("Calculation")
        g.add_large(QAction(icons.icon("calc"), "Refresh All", self,
                            triggered=self.calculate_now), "calc")
        # View -------------------------------------------------------------------------
        vw = rb.add_page("View")
        g = vw.add_group("Show")
        g.add_small_column([(acts["gridlines"], "gridlines"), (acts["headings"], "names"),
                            (acts["formula_bar"], "function")])
        g = vw.add_group("Zoom")
        g.add_large(acts["zoom_dialog"], "find")
        g.add_large(acts["zoom_100"], "zoom_100")
        g.add_large(acts["zoom_sel"], "zoom_in")
        g.add_small_column([(acts["zoom_in"], "zoom_in"), (acts["zoom_out"], "zoom_out")])
        g = vw.add_group("Window")
        frm = QMenu(self)
        for k in ("freeze_panes", "freeze_row", "freeze_col", "unfreeze"):
            frm.addAction(acts[k])
        fz = QAction(icons.icon("freeze_panes"), "Freeze Panes", self)
        fz.triggered.connect(lambda: self.freeze("panes"))
        fz.setData("menu-only")
        g.add_large(fz, "freeze_panes", menu=frm)
        g.add_small_column([(acts["freeze_row"], "freeze_row"), (acts["freeze_col"], "freeze_col"),
                            (acts["unfreeze"], "freeze_panes")])
        # Developer -----------------------------------------------------------------------
        dev = rb.add_page("Developer")
        g = dev.add_group("Code")
        g.add_large(acts["vba_editor"], "vba")
        g.add_large(acts["macros"], "macros")
        g.add_large(acts["run_macro"], "run")
        g.add_small_column([(acts["macro_security"], "security"),
                            (acts["enable_macros"], "security"), (acts["stop_macro"], "stop")])
        g = dev.add_group("Security")
        g.add_large(acts["lock"], "lock")
        g.add_large(acts["change_pw"], "key")
        rb.show_page("Home")

    def _function_menus(self, g):
        cats: Dict[str, List[str]] = {}
        for name in sorted(self.function_names()):
            cats.setdefault(dialogs.function_category(name), []).append(name)
        for cat, icon in (("Logical", "function"), ("Text", "function"),
                          ("Date & Time", "calc"), ("Lookup & Reference", "find"),
                          ("Math & Trig", "autosum")):
            names = cats.get(cat, [])
            m = QMenu(self)
            for n in names:
                m.addAction(n, lambda n=n: self.insert_function_name(n))
            a = QAction(icons.icon(icon), cat.replace("&", "&&"), self)
            a.setData("menu-only")
            g.add_large(a, icon, menu=m)

    # ---------------------------------------------------------- formula bar
    def _build_formula_bar(self):
        w = QWidget()
        w.setObjectName("formulaBar")
        w.setAttribute(Qt.WA_StyledBackground, True)
        h = QHBoxLayout(w)
        h.setContentsMargins(4, 4, 4, 4)
        h.setSpacing(4)
        self.name_box = QComboBox()
        self.name_box.setObjectName("nameBox")
        self.name_box.setEditable(True)
        self.name_box.setFixedWidth(110)
        self.name_box.setInsertPolicy(QComboBox.NoInsert)
        self.name_box.lineEdit().returnPressed.connect(self._name_box_entered)
        self.name_box.activated.connect(lambda _i: self._name_box_entered())
        h.addWidget(self.name_box)
        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet("color:#D4D4D4;")
        h.addWidget(sep)
        self.fb_cancel = QToolButton()
        self.fb_cancel.setObjectName("fbButton")
        self.fb_cancel.setText("✕")
        self.fb_cancel.setToolTip("Cancel")
        self.fb_cancel.clicked.connect(self.formula_bar_cancel)
        self.fb_enter = QToolButton()
        self.fb_enter.setObjectName("fbButton")
        self.fb_enter.setText("✓")
        self.fb_enter.setToolTip("Enter")
        self.fb_enter.clicked.connect(lambda: self.formula_bar_commit(move=False))
        self.fb_fx = QToolButton()
        self.fb_fx.setObjectName("fbButton")
        self.fb_fx.setText("fx")
        f = self.fb_fx.font()
        f.setItalic(True)
        f.setFamily("Times New Roman")
        f.setPointSizeF(11)
        self.fb_fx.setFont(f)
        self.fb_fx.setToolTip("Insert Function")
        self.fb_fx.clicked.connect(self.insert_function)
        for b in (self.fb_cancel, self.fb_enter, self.fb_fx):
            h.addWidget(b)
        self.formula_edit = FormulaLineEdit(self.function_names())
        self.formula_edit.setObjectName("formulaEdit")
        self.formula_edit.textEdited.connect(self._fb_edited)
        self.formula_edit.returnPressed.connect(lambda: self.formula_bar_commit(move=True))
        self.formula_edit.installEventFilter(self)
        h.addWidget(self.formula_edit, 1)
        return w

    def eventFilter(self, obj, ev):
        if obj is getattr(self, "formula_edit", None) and ev.type() == QEvent.KeyPress:
            if ev.key() == Qt.Key_Escape:
                self.formula_bar_cancel()
                return True
            if ev.key() in (Qt.Key_Tab, Qt.Key_Backtab) and not self.formula_edit.popup_visible():
                back = ev.key() == Qt.Key_Backtab or bool(ev.modifiers() & Qt.ShiftModifier)
                self.formula_bar_commit(move=False)
                self.view.move_after_edit(0, -1 if back else 1)
                return True
        return super().eventFilter(obj, ev)

    # ---------------------------------------------------------- sheet bar
    def _build_sheet_bar(self):
        w = QWidget()
        w.setObjectName("sheetBar")
        w.setAttribute(Qt.WA_StyledBackground, True)
        h = QHBoxLayout(w)
        h.setContentsMargins(4, 0, 0, 0)
        h.setSpacing(0)
        for txt, step in (("◀", -1), ("▶", 1)):
            b = QToolButton()
            b.setObjectName("sheetNav")
            b.setText(txt)
            b.clicked.connect(lambda _=False, s=step: self.step_sheet(s))
            h.addWidget(b)
        h.addSpacing(6)
        self.sheet_tabs = SheetTabs()
        self.sheet_tabs.currentChanged.connect(self._sheet_tab_changed)
        self.sheet_tabs.rename_requested.connect(self._rename_from_tab)
        self.sheet_tabs.menu_requested.connect(self._sheet_menu)
        self.sheet_tabs.add_requested.connect(self.add_sheet)
        h.addWidget(self.sheet_tabs)
        add = QToolButton()
        add.setObjectName("addSheet")
        add.setText("+")
        add.setToolTip("New sheet")
        add.clicked.connect(self.add_sheet)
        h.addWidget(add)
        h.addStretch(1)
        self.hscroll = QScrollBar(Qt.Horizontal)
        self.hscroll.setMinimumWidth(300)
        self.hscroll.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        inner = self.view.horizontalScrollBar()
        inner.rangeChanged.connect(self.hscroll.setRange)
        inner.valueChanged.connect(self.hscroll.setValue)
        self.hscroll.valueChanged.connect(inner.setValue)
        self.hscroll.setRange(inner.minimum(), inner.maximum())
        h.addWidget(self.hscroll, 1)
        return w

    # ---------------------------------------------------------- status bar
    def _build_status_bar(self):
        sb = self.statusBar()
        sb.setSizeGripEnabled(False)
        self.mode_label = QLabel("Ready")
        self.mode_label.setMinimumWidth(60)
        sb.addWidget(self.mode_label)
        self.msg_label = QLabel("")
        sb.addWidget(self.msg_label, 1)
        self.stats_label = QLabel("")
        sb.addPermanentWidget(self.stats_label)
        zo = QToolButton()
        zo.setText("−")
        zo.clicked.connect(lambda: self.set_zoom(self.zoom - 0.1))
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(50, 200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setFixedWidth(110)
        self.zoom_slider.valueChanged.connect(lambda v: self.set_zoom(v / 100.0))
        zi = QToolButton()
        zi.setText("+")
        zi.clicked.connect(lambda: self.set_zoom(self.zoom + 0.1))
        self.zoom_label = QToolButton()
        self.zoom_label.setText("100%")
        self.zoom_label.clicked.connect(self.zoom_dialog)
        for wdg in (zo, self.zoom_slider, zi, self.zoom_label):
            sb.addPermanentWidget(wdg)

    # ======================================================================
    # Grid context (called by grid.py)
    # ======================================================================
    def function_names(self):
        try:
            return list(engine_mod.FUNCTIONS.keys())
        except Exception:
            return []

    def set_mode(self, mode: str):
        if hasattr(self, "mode_label"):
            self.mode_label.setText(mode)

    def on_editor_text(self, text: str):
        if "view" not in self.__dict__:
            return
        self.formula_edit.setText(text)
        self.view.refresh()

    def active_formula_editor(self):
        ed = self.view.delegate.editor if "view" in self.__dict__ else None
        if ed is not None:
            return ed
        if self.fb_editing:
            return self.formula_edit
        return None

    def clip_marquee(self):
        if self._marquee and self._marquee[0] is self.book.active:
            return self._marquee[1:]
        return None

    def cancel_marquee(self):
        if self._marquee:
            self._marquee = None
            self.view.refresh()

    def on_selection_changed(self):
        if "view" not in self.__dict__:
            return
        self.update_stats()
        self._update_name_box()
        self._fire_selection_change()

    def on_current_changed(self):
        if "view" not in self.__dict__ or not hasattr(self, "formula_edit"):
            return
        r, c = self.view.current_cell()
        sheet = self.book.active
        if not self.view.is_editing():
            self.formula_edit.setText(sheet.get_raw(r, c))
            self.fb_editing = False
        self._update_name_box()
        self._sync_format_controls()
        self.update_stats()

    def _update_name_box(self):
        r1, c1, r2, c2 = self.view.active_range()
        r, c = self.view.current_cell()
        if self.name_box.lineEdit().hasFocus():
            return
        for name, ref in self.book.defined_names.items():
            sheet_part, _, rng = ref.rpartition("!")
            if sheet_part.strip("'").lower() == self.book.active.name.lower():
                if parse_range_a1(rng) == (r1, c1, r2, c2) and (r1, c1, r2, c2) != (r, c, r, c) \
                        or parse_range_a1(rng) == (r1, c1, r2, c2):
                    self.name_box.setEditText(name)
                    return
        self.name_box.setEditText(a1(r, c))

    def _sync_format_controls(self):
        r, c = self.view.current_cell()
        st = self.book.active.get_style(r, c)
        self._syncing = True
        try:
            self.actions["bold"].setChecked(bool(st.get("bold")))
            self.actions["italic"].setChecked(bool(st.get("italic")))
            self.actions["underline"].setChecked(bool(st.get("underline")))
            self.actions["wrap"].setChecked(bool(st.get("wrap")))
            for k in ("left", "center", "right"):
                self.actions["align_" + k].setChecked(st.get("align") == k)
            self.font_box.blockSignals(True)
            self.font_box.setCurrentFont(QFont(st.get("font_name") or "Calibri"))
            self.font_box.blockSignals(False)
            fs = st.get("font_size")
            self.size_box.setCurrentText(value_to_text(float(fs)) if fs else "11")
            fmt = st.get("number_format")
            idx = 0
            for i in range(self.num_combo.count()):
                if self.num_combo.itemData(i) == fmt:
                    idx = i
                    break
            else:
                if fmt:
                    if self.num_combo.findText("Custom") < 0:
                        self.num_combo.addItem("Custom", "__custom__")
                    idx = self.num_combo.findText("Custom")
            self.num_combo.setCurrentIndex(idx)
        finally:
            self._syncing = False

    def update_stats(self):
        if not hasattr(self, "stats_label") or "view" not in self.__dict__:
            return
        sheet = self.book.active
        total = 0.0
        n_num = 0
        n_all = 0
        for (r1, c1, r2, c2) in self.view.sel_ranges:
            if r1 == r2 and c1 == c2 and len(self.view.sel_ranges) == 1:
                break
            for r, c, raw in sheet.iter_cells_in(r1, c1, r2, c2):
                v = self.view.sheet_model.value(r, c)
                if v is None or v == "":
                    continue
                n_all += 1
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    total += v
                    n_num += 1
        parts = []
        r, c = self.view.current_cell()
        fmt = sheet.get_style(r, c).get("number_format")

        def show(v):
            v = round(v, 10)
            if fmt and fmt not in ("General", "@"):
                txt = self.preview_format(v, fmt)
                if txt and not txt.startswith("#"):
                    return txt
            return value_to_text(v)
        if n_num:
            parts.append(f"Average: {show(total / n_num)}")
        if n_all > 1 or (n_all and n_num):
            parts.append(f"Count: {n_all}")
        if n_num:
            parts.append(f"Sum: {show(total)}")
        self.stats_label.setText("     ".join(parts))

    def show_cell_menu(self, pos):
        m = QMenu(self)
        for k in ("cut", "copy", "paste", "paste_values"):
            m.addAction(self.actions[k])
        m.addSeparator()
        m.addAction("Insert...", self.insert_dialog)
        m.addAction("Delete...", self.delete_dialog)
        m.addAction(self.actions["clear_contents"])
        m.addSeparator()
        sm = m.addMenu("Sort")
        sm.addAction(self.actions["sort_asc"])
        sm.addAction(self.actions["sort_desc"])
        m.addSeparator()
        m.addAction(self.actions["format_cells"])
        m.addAction(self.actions["define_name"])
        self._last_menu = m
        m.popup(pos)
        return m

    def show_header_menu(self, axis, pos):
        m = QMenu(self)
        for k in ("cut", "copy", "paste"):
            m.addAction(self.actions[k])
        m.addSeparator()
        m.addAction("Insert", lambda: self.insert_cells(axis))
        m.addAction("Delete", lambda: self.delete_cells(axis))
        m.addAction(self.actions["clear_contents"])
        m.addSeparator()
        m.addAction(self.actions["format_cells"])
        if axis == "col":
            m.addAction(self.actions["col_width"])
        else:
            m.addAction(self.actions["row_height"])
        m.addAction("Hide", lambda: self.hide_rows_cols(axis, True))
        m.addAction("Unhide", lambda: self.hide_rows_cols(axis, False))
        self._last_menu = m
        m.popup(pos)
        return m

    # ======================================================================
    # Edits
    # ======================================================================
    def mark_dirty(self):
        if not self.modified:
            self.modified = True
            self.update_title()

    def _clamp(self, r1, c1, r2, c2):
        rows, cols = self.book.active.used_bounds()
        if r2 - r1 + 1 > LARGE_RANGE_ROWS:
            r2 = max(r1, min(r2, max(rows, r1 + 1) - 1))
        if c2 - c1 + 1 > LARGE_RANGE_COLS:
            c2 = max(c1, min(c2, max(cols, c1 + 1) - 1))
        return r1, c1, r2, c2

    def selection_keys(self, clamp=True) -> List[Tuple[int, int]]:
        keys = []
        seen = set()
        for rg in self.view.sel_ranges:
            r1, c1, r2, c2 = self._clamp(*rg) if clamp else rg
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) not in seen:
                        seen.add((r, c))
                        keys.append((r, c))
        return keys

    def apply_changes(self, sheet, changes: Dict, text: str, select=None) -> bool:
        """changes: {(r,c): {"raw": str|None, "style": dict|None}} (missing key = keep)."""
        before, after = {}, {}
        for k, ch in changes.items():
            raw0 = sheet.cells.get(k)
            st0 = sheet.styles.get(k)
            b = (raw0, dict(st0) if st0 else None)
            raw1 = ch["raw"] if "raw" in ch else raw0
            st1 = ch["style"] if "style" in ch else st0
            raw1 = raw1 if raw1 not in ("",) else None
            a = (raw1, dict(st1) if st1 else None)
            if a != b:
                before[k] = b
                after[k] = a
        if not after:
            return False
        self.undo_stack.push(CellsCommand(self, sheet, before, after, text, select))
        return True

    def cells_changed(self, sheet, keys, select=None):
        try:
            self.engine.invalidate()
        except Exception:
            log.exception("engine.invalidate failed")
        if sheet is not self.book.active:
            self.activate_sheet(sheet)
        if select:
            self.view.select_range(*select)
        self.view.refresh()
        self.mark_dirty()
        self.on_current_changed()
        if keys and self.runtime is not None and self.macros_enabled and not self._suspend_events:
            rs = [k[0] for k in keys]
            cs = [k[1] for k in keys]
            try:
                self.runtime.notify_change(sheet, min(rs), min(cs), max(rs), max(cs))
            except Exception as e:
                self._handle_vba_exception(e)
            self.view.refresh()

    def structure_changed(self):
        try:
            self.engine.invalidate()
        except Exception:
            log.exception("engine.invalidate failed")
        self.book.active_index = max(0, min(self.book.active_index, len(self.book.sheets) - 1))
        self.rebuild_sheet_tabs()
        cur = self.view.current_cell()
        self.view.set_sheet(self.book.active)
        self.view.set_frozen(*self._frozen.get(id(self.book.active), (0, 0)))
        self.view.goto(*cur)
        self.view.refresh()
        self.mark_dirty()
        self.on_current_changed()
        if self.vba_editor is not None:
            self.vba_editor.rebuild_tree()

    def user_set_cell(self, r, c, text, fill_selection=False):
        sheet = self.book.active
        keys = self.selection_keys() if fill_selection else [(r, c)]
        changes = {}
        for k in keys:
            raw = text
            if fill_selection and text.startswith("=") and k != (r, c):
                raw = engine_mod.shift_formula(text, k[0] - r, k[1] - c)
            ch = {"raw": raw}
            # typing "12%" gives a percent format, like Excel
            if raw and not raw.startswith("=") and raw.strip().endswith("%") and \
                    isinstance(literal_value(raw), float) and \
                    not sheet.get_style(*k).get("number_format"):
                st = dict(sheet.get_style(*k))
                st["number_format"] = "0%" if "." not in raw else "0.00%"
                ch["style"] = st
            changes[k] = ch
        self.apply_changes(sheet, changes, "Typing")
        self._marquee = None if not self.clip or not self.clip.get("cut") else self._marquee
        self.on_current_changed()

    def _style_changes(self, fn) -> Dict:
        sheet = self.book.active
        changes = {}
        for k in self.selection_keys():
            st = dict(sheet.styles.get(k, {}))
            fn(st)
            for key in [x for x, v in st.items() if v is None or v is False]:
                st.pop(key)
            changes[k] = {"style": st}
        return changes

    def apply_style(self, **kw):
        if getattr(self, "_syncing", False):
            return

        def fn(st):
            for k, v in kw.items():
                st[k] = v
        self.apply_changes(self.book.active, self._style_changes(fn), "Format")
        self._sync_format_controls()

    def toggle_style(self, key):
        r, c = self.view.current_cell()
        on = not bool(self.book.active.get_style(r, c).get(key))
        self.apply_style(**{key: True if on else None})

    def set_align(self, align):
        r, c = self.view.current_cell()
        cur = self.book.active.get_style(r, c).get("align")
        self.apply_style(align=None if cur == align else align)

    def _size_entered(self):
        if self._syncing:
            return
        try:
            size = float(self.size_box.currentText())
        except ValueError:
            return
        if 1 <= size <= 409:
            self.apply_style(font_size=size)

    def _pick_font_color(self, col):
        self._font_color = col
        self.actions["font_color"].setIcon(icons.color_icon("font_color", col or "#000000"))
        self.apply_style(color=col)

    def _pick_fill_color(self, col):
        self._fill_color = col
        self.actions["fill_color"].setIcon(icons.color_icon("fill_color", col or "#FFFFFF"))
        self.apply_style(fill=col)

    def _num_combo_chosen(self, i):
        code = self.num_combo.itemData(i)
        if code == "__custom__":
            self.format_cells()
            return
        self.set_number_format(code)

    def set_number_format(self, fmt):
        self.apply_style(number_format=fmt)

    def change_decimals(self, delta):
        r, c = self.view.current_cell()
        fmt = self.book.active.get_style(r, c).get("number_format") or ""
        if not fmt or fmt == "General":
            v = self.view.sheet_model.value(r, c)
            dec = 0
            if isinstance(v, float) and not v.is_integer():
                dec = min(10, len(value_to_text(v).split(".")[-1]))
            fmt = "0" + ("." + "0" * dec if dec else "")
        m = re.search(r"(0+)?(\.(0*))?(%|E\+00)?", fmt[::-1])
        base, _, dec_part = fmt.partition(".")
        suffix = ""
        if dec_part:
            digits = len(dec_part) - len(dec_part.lstrip("0"))
            suffix = dec_part[digits:]
        else:
            digits = 0
            mm = re.match(r"^([^%E]*)(.*)$", base)
            base, suffix = mm.group(1), mm.group(2)
        digits = max(0, min(30, digits + delta))
        new = base + ("." + "0" * digits if digits else "") + suffix
        self.set_number_format(new)

    def clear_contents(self, selection=True):
        sheet = self.book.active
        if selection:
            changes = {}
            for (r1, c1, r2, c2) in self.view.sel_ranges:
                for r, c, _raw in sheet.iter_cells_in(r1, c1, r2, c2):
                    changes[(r, c)] = {"raw": None}
        else:
            r, c = self.view.current_cell()
            changes = {(r, c): {"raw": None}}
        self.apply_changes(sheet, changes, "Clear")

    def clear_formats(self):
        sheet = self.book.active
        changes = {}
        for (r1, c1, r2, c2) in self.view.sel_ranges:
            for k in list(sheet.styles):
                if r1 <= k[0] <= r2 and c1 <= k[1] <= c2:
                    changes[k] = {"style": None}
        self.apply_changes(sheet, changes, "Clear Formats")

    def clear_all(self):
        sheet = self.book.active
        changes = {}
        for (r1, c1, r2, c2) in self.view.sel_ranges:
            for k in list(sheet.styles) + list(sheet.cells):
                if r1 <= k[0] <= r2 and c1 <= k[1] <= c2:
                    changes[k] = {"style": None, "raw": None}
        self.apply_changes(sheet, changes, "Clear All")

    def insert_today(self):
        import datetime
        d = datetime.date.today()
        r, c = self.view.current_cell()
        text = f"{d.month}/{d.day}/{d.year}"
        if self.view.is_editing():
            self.view.delegate.editor.insert(text)
            return
        try:
            serial = engine_mod.date_to_serial(d)
            st = dict(self.book.active.get_style(r, c))
            st["number_format"] = "m/d/yyyy"
            self.apply_changes(self.book.active, {(r, c): {"raw": str(serial), "style": st}},
                               "Insert Date")
        except Exception:
            self.user_set_cell(r, c, text)

    # ======================================================================
    # Formula bar
    # ======================================================================
    def _fb_edited(self, text):
        self.fb_editing = True
        ed = self.view.delegate.editor
        if ed is not None and ed.text() != text:
            ed.setText(text)
        self.set_mode("Edit")
        self.view.refresh()

    def formula_bar_commit(self, move=True):
        text = self.formula_edit.text()
        if self.view.is_editing():
            self.view.finish_edit(self.view.delegate.editor, commit=True,
                                  move=(1, 0) if move else (0, 0))
            self.fb_editing = False
            return
        if not self.fb_editing:
            if move:
                self.view.setFocus()
                self.view.move_after_edit(1, 0)
            return
        self.fb_editing = False
        r, c = self.view.current_cell()
        self.user_set_cell(r, c, text)
        self.set_mode("Ready")
        self.view.setFocus()
        if move:
            self.view.move_after_edit(1, 0)

    def formula_bar_cancel(self):
        if self.view.is_editing():
            self.view.finish_edit(self.view.delegate.editor, commit=False)
        self.fb_editing = False
        r, c = self.view.current_cell()
        self.formula_edit.setText(self.book.active.get_raw(r, c))
        self.set_mode("Ready")
        self.view.setFocus()
        self.view.refresh()

    def formula_bar_focus_lost_commit(self):
        if self.fb_editing:
            self.formula_bar_commit(move=False)

    # ---------------------------------------------------------- name box
    def _name_box_entered(self):
        text = self.name_box.currentText().strip()
        if not text:
            return
        target = text
        sheet = self.book.active
        dn = {k.lower(): v for k, v in self.book.defined_names.items()}
        if text.lower() in dn:
            target = dn[text.lower()]
        if "!" in target:
            sname, _, target = target.rpartition("!")
            sh = self.book.sheet_by_name(sname.strip("'"))
            if sh is not None:
                sheet = sh
        rng = parse_range_a1(target)
        if rng is None:
            if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", text):
                r1, c1, r2, c2 = self.view.active_range()
                self.book.defined_names[text] = (f"{self._sheet_ref(self.book.active)}!"
                                                 f"{a1(r1, c1, True)}" +
                                                 (f":{a1(r2, c2, True)}" if (r1, c1) != (r2, c2)
                                                  else ""))
                self._refresh_name_list()
                self.mark_dirty()
                self.view.setFocus()
                return
            QMessageBox.warning(self, "PortableExcel", "Reference isn't valid.")
            return
        if sheet is not self.book.active:
            self.activate_sheet(sheet)
        r1, c1, r2, c2 = rng
        self.view.select_range(r1, c1, r2, c2)
        self.view.setFocus()

    @staticmethod
    def _sheet_ref(sheet):
        n = sheet.name
        return f"'{n}'" if re.search(r"[^A-Za-z0-9_]", n) else n

    def _refresh_name_list(self):
        cur = self.name_box.currentText()
        self.name_box.blockSignals(True)
        self.name_box.clear()
        self.name_box.addItems(sorted(self.book.defined_names))
        self.name_box.setEditText(cur)
        self.name_box.blockSignals(False)

    # ======================================================================
    # Clipboard
    # ======================================================================
    def copy(self, cut=False):
        if self.view.is_editing():
            self.view.delegate.editor.cut() if cut else self.view.delegate.editor.copy()
            return
        if len(self.view.sel_ranges) > 1:
            QMessageBox.information(self, "PortableExcel",
                                    "This action won't work on multiple selections.")
            return
        sheet = self.book.active
        r1, c1, r2, c2 = self._clamp(*self.view.active_range())
        cells = {}
        rows = []
        model = self.view.sheet_model
        for r in range(r1, r2 + 1):
            row = []
            for c in range(c1, c2 + 1):
                raw = sheet.cells.get((r, c))
                st = sheet.styles.get((r, c))
                if raw is not None or st:
                    cells[(r - r1, c - c1)] = (raw, dict(st) if st else None,
                                               model.value(r, c))
                row.append(model.display(r, c) if raw is not None else "")
            rows.append(row)
        text = ops.to_tsv(rows)
        self.clip = {"sheet": sheet, "r1": r1, "c1": c1, "rows": r2 - r1 + 1,
                     "cols": c2 - c1 + 1, "cells": cells, "cut": cut, "text": text}
        self._marquee = (sheet, r1, c1, r2, c2)
        QGuiApplication.clipboard().setText(text)
        self.view.refresh()
        self.msg_label.setText("Select destination and press ENTER or choose Paste")

    def paste(self, values_only=False):
        if self.view.is_editing():
            self.view.delegate.editor.paste()
            return
        sheet = self.book.active
        r1, c1, r2, c2 = self.view.active_range()
        sys_text = QGuiApplication.clipboard().text()
        clip = self.clip
        internal = clip is not None and (sys_text == clip["text"] or not sys_text)
        changes: Dict = {}
        if internal:
            rows, cols = clip["rows"], clip["cols"]
            sel_h, sel_w = r2 - r1 + 1, c2 - c1 + 1
            rep_r = sel_h // rows if sel_h % rows == 0 and sel_h > rows else 1
            rep_c = sel_w // cols if sel_w % cols == 0 and sel_w > cols else 1
            if clip["cut"]:
                rep_r = rep_c = 1
                for dr in range(rows):
                    for dc in range(cols):
                        changes[(clip["r1"] + dr, clip["c1"] + dc)] = {"raw": None,
                                                                       "style": None}
                if clip["sheet"] is not sheet:
                    self.apply_changes(clip["sheet"], changes, "Cut")
                    changes = {}
            for i in range(rep_r):
                for j in range(rep_c):
                    br, bc = r1 + i * rows, c1 + j * cols
                    for dr in range(rows):
                        for dc in range(cols):
                            raw, st, val = clip["cells"].get((dr, dc), (None, None, None))
                            tr, tc = br + dr, bc + dc
                            if tr >= MAX_ROWS or tc >= MAX_COLS:
                                continue
                            if values_only:
                                if raw is not None and raw.startswith("="):
                                    raw = self._value_raw(val)
                                changes[(tr, tc)] = {"raw": raw}
                            else:
                                if raw and raw.startswith("=") and not clip["cut"]:
                                    raw = engine_mod.shift_formula(
                                        raw, tr - (clip["r1"] + dr), tc - (clip["c1"] + dc))
                                changes[(tr, tc)] = {"raw": raw, "style": st}
            sel = (r1, c1, r1 + rows * rep_r - 1, c1 + cols * rep_c - 1)
            self.apply_changes(sheet, changes, "Paste", select=sel)
            if clip["cut"]:
                self.clip = None
                self._marquee = None
        else:
            data = ops.parse_tsv(sys_text)
            if not data:
                return
            rows = len(data)
            cols = max(len(r) for r in data)
            sel_h, sel_w = r2 - r1 + 1, c2 - c1 + 1
            rep_r = sel_h // rows if sel_h % rows == 0 and sel_h > rows else 1
            rep_c = sel_w // cols if sel_w % cols == 0 and sel_w > cols else 1
            for i in range(rep_r):
                for j in range(rep_c):
                    for dr, line in enumerate(data):
                        for dc in range(cols):
                            t = line[dc] if dc < len(line) else ""
                            changes[(r1 + i * rows + dr, c1 + j * cols + dc)] = {"raw": t or None}
            self.apply_changes(sheet, changes, "Paste",
                               select=(r1, c1, r1 + rows * rep_r - 1, c1 + cols * rep_c - 1))
        self.msg_label.setText("")
        self.view.refresh()

    @staticmethod
    def _value_raw(val) -> Optional[str]:
        if val is None:
            return None
        if isinstance(val, bool):
            return "TRUE" if val else "FALSE"
        if isinstance(val, (int, float)):
            return repr(float(val)) if not float(val).is_integer() else str(int(val))
        s = str(val)
        lit = literal_value(s)
        if not isinstance(lit, str) or lit != s:
            return "'" + s
        return s

    # ======================================================================
    # Fill
    # ======================================================================
    def _fill(self, src, dst, text="Fill"):
        sheet = self.book.active
        sr1, sc1, sr2, sc2 = src
        dr1, dc1, dr2, dc2 = dst
        changes = {}
        vertical = dc1 == sc1 and dc2 == sc2
        if vertical:
            down = dr1 > sr2
            n = dr2 - dr1 + 1
            for c in range(sc1, sc2 + 1):
                line = [sheet.get_raw(r, c) for r in range(sr1, sr2 + 1)]
                styles = [sheet.styles.get((r, c)) for r in range(sr1, sr2 + 1)]
                if not down:
                    line = line[::-1]
                    styles = styles[::-1]
                sign = 1 if down else -1
                vals = ops.fill_series(
                    line, n, lambda raw, k, s=sign: engine_mod.shift_formula(raw, s * k, 0))
                for i, v in enumerate(vals):
                    r = dr1 + i if down else dr2 - i
                    changes[(r, c)] = {"raw": v or None, "style": styles[i % len(styles)]}
        else:
            right = dc1 > sc2
            n = dc2 - dc1 + 1
            for r in range(sr1, sr2 + 1):
                line = [sheet.get_raw(r, c) for c in range(sc1, sc2 + 1)]
                styles = [sheet.styles.get((r, c)) for c in range(sc1, sc2 + 1)]
                if not right:
                    line = line[::-1]
                    styles = styles[::-1]
                sign = 1 if right else -1
                vals = ops.fill_series(
                    line, n, lambda raw, k, s=sign: engine_mod.shift_formula(raw, 0, s * k))
                for i, v in enumerate(vals):
                    c = dc1 + i if right else dc2 - i
                    changes[(r, c)] = {"raw": v or None, "style": styles[i % len(styles)]}
        sel = (min(sr1, dr1), min(sc1, dc1), max(sr2, dr2), max(sc2, dc2))
        self.apply_changes(sheet, changes, text, select=sel)

    def fill_from_handle(self, src, dst):
        self._fill(src, dst, "Auto Fill")

    def fill_handle_double_click(self):
        """Fill down as far as the neighbouring column has data."""
        r1, c1, r2, c2 = self.view.active_range()
        cells = self.book.active.cells
        nb = c1 - 1 if c1 > 0 and (r2 + 1, c1 - 1) in cells else c2 + 1
        end = r2
        while (end + 1, nb) in cells:
            end += 1
        if end > r2:
            self._fill((r1, c1, r2, c2), (r2 + 1, c1, end, c2), "Auto Fill")

    def fill_down(self):
        r1, c1, r2, c2 = self._clamp(*self.view.active_range())
        if r1 == r2:
            if r1 == 0:
                return
            src = (r1 - 1, c1, r1 - 1, c2)
            dst = (r1, c1, r2, c2)
        else:
            src = (r1, c1, r1, c2)
            dst = (r1 + 1, c1, r2, c2)
        self._copy_fill(src, dst, "Fill Down")

    def fill_right(self):
        r1, c1, r2, c2 = self._clamp(*self.view.active_range())
        if c1 == c2:
            if c1 == 0:
                return
            src = (r1, c1 - 1, r2, c1 - 1)
            dst = (r1, c1, r2, c2)
        else:
            src = (r1, c1, r2, c1)
            dst = (r1, c1 + 1, r2, c2)
        self._copy_fill(src, dst, "Fill Right")

    def _copy_fill(self, src, dst, text):
        sheet = self.book.active
        sr1, sc1, sr2, sc2 = src
        changes = {}
        for r in range(dst[0], dst[2] + 1):
            for c in range(dst[1], dst[3] + 1):
                rr = sr1 if sr1 == sr2 else r
                cc = sc1 if sc1 == sc2 else c
                raw = sheet.get_raw(rr, cc)
                if raw.startswith("="):
                    raw = engine_mod.shift_formula(raw, r - rr, c - cc)
                changes[(r, c)] = {"raw": raw or None, "style": sheet.styles.get((rr, cc))}
        self.apply_changes(sheet, changes, text, select=self.view.active_range())

    # ======================================================================
    # Rows / columns / structure
    # ======================================================================
    def _structural(self, text, fn):
        before = ops.snapshot_book(self.book)
        try:
            fn()
        except Exception as e:
            ops.restore_book(self.book, before)
            log.exception("structural change failed")
            QMessageBox.warning(self, "PortableExcel", f"{text} failed:\n{e}")
            self.structure_changed()
            return False
        after = ops.snapshot_book(self.book)
        self.undo_stack.push(SnapshotCommand(self, before, after, text))
        self.structure_changed()
        return True

    def insert_cells(self, axis):
        r1, c1, r2, c2 = self.view.active_range()
        if axis == "row":
            idx, n = r1, min(r2 - r1 + 1, 10000)
        else:
            idx, n = c1, min(c2 - c1 + 1, 1000)
        sheet = self.book.active
        self._structural("Insert " + ("Rows" if axis == "row" else "Columns"),
                         lambda: engine_mod.insert_delete(self.book, sheet, axis, idx, n))

    def delete_cells(self, axis):
        r1, c1, r2, c2 = self.view.active_range()
        if axis == "row":
            idx, n = r1, r2 - r1 + 1
        else:
            idx, n = c1, c2 - c1 + 1
        sheet = self.book.active
        self._structural("Delete " + ("Rows" if axis == "row" else "Columns"),
                         lambda: engine_mod.insert_delete(self.book, sheet, axis, idx, -n))

    def insert_dialog(self):
        item, ok = QInputDialog.getItem(self, "Insert", "Insert:", ["Entire row",
                                                                    "Entire column"], 0, False)
        if ok:
            self.insert_cells("row" if item == "Entire row" else "col")

    def delete_dialog(self):
        item, ok = QInputDialog.getItem(self, "Delete", "Delete:", ["Entire row",
                                                                    "Entire column"], 0, False)
        if ok:
            self.delete_cells("row" if item == "Entire row" else "col")

    def hide_rows_cols(self, axis, hide):
        r1, c1, r2, c2 = self.view.active_range()
        sheet = self.book.active

        def fn():
            if axis == "col":
                for c in range(c1, min(c2, c1 + 500) + 1):
                    if hide:
                        sheet.col_widths[c] = 0
                    elif sheet.col_widths.get(c) == 0:
                        sheet.col_widths.pop(c)
                if not hide and c1 > 0 and sheet.col_widths.get(c1 - 1) == 0:
                    sheet.col_widths.pop(c1 - 1)
            else:
                for r in range(r1, min(r2, r1 + 5000) + 1):
                    if hide:
                        sheet.row_heights[r] = 0
                    elif sheet.row_heights.get(r) == 0:
                        sheet.row_heights.pop(r)
                if not hide and r1 > 0 and sheet.row_heights.get(r1 - 1) == 0:
                    sheet.row_heights.pop(r1 - 1)
        self._structural("Hide" if hide else "Unhide", fn)

    def ask_row_height(self):
        r, _ = self.view.current_cell()
        cur = self.book.active.row_heights.get(r, 20)
        h, ok = QInputDialog.getInt(self, "Row Height", "Row height (pixels):", cur, 0, 800)
        if ok:
            r1, _, r2, _ = self._clamp(*self.view.active_range())
            self._structural("Row Height", lambda: [self.book.active.row_heights.__setitem__(
                rr, h) for rr in range(r1, r2 + 1)])

    def ask_col_width(self):
        _, c = self.view.current_cell()
        cur = self.book.active.col_widths.get(c, 64)
        w, ok = QInputDialog.getInt(self, "Column Width", "Column width (pixels):", cur, 0, 2000)
        if ok:
            _, c1, _, c2 = self._clamp(*self.view.active_range())
            self._structural("Column Width", lambda: [self.book.active.col_widths.__setitem__(
                cc, w) for cc in range(c1, c2 + 1)])

    def autofit_selection(self):
        _, c1, _, c2 = self._clamp(*self.view.active_range())
        for c in range(c1, c2 + 1):
            self.view.autofit_column(c)

    # ======================================================================
    # Sort / duplicates
    # ======================================================================
    def _current_region(self):
        r1, c1, r2, c2 = self.view.active_range()
        if (r1, c1) != (r2, c2):
            return self._clamp(r1, c1, r2, c2)
        cells = self.book.active.cells
        if (r1, c1) not in cells:
            return None
        top, left, bottom, right = r1, c1, r1, c1
        changed = True
        while changed:
            changed = False
            for (r, c) in ((top - 1, None), (bottom + 1, None), (None, left - 1),
                           (None, right + 1)):
                if r is not None and r >= 0 and any((r, cc) in cells
                                                    for cc in range(left - 1, right + 2)):
                    top, bottom = min(top, r), max(bottom, r)
                    changed = True
                if c is not None and c >= 0 and any((rr, c) in cells
                                                    for rr in range(top - 1, bottom + 2)):
                    left, right = min(left, c), max(right, c)
                    changed = True
        return top, left, bottom, right

    def _has_header(self, r1, c1, r2, c2):
        if r2 <= r1:
            return False
        sheet = self.book.active
        model = self.view.sheet_model
        first = [model.value(r1, c) for c in range(c1, c2 + 1)]
        second = [model.value(r1 + 1, c) for c in range(c1, c2 + 1)]
        if all(isinstance(v, str) for v in first if v is not None) and \
                any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in second):
            return True
        return any(sheet.get_style(r1, c).get("bold") for c in range(c1, c2 + 1)) and \
            not any(sheet.get_style(r1 + 1, c).get("bold") for c in range(c1, c2 + 1))

    def sort_selection(self, descending=False):
        region = self._current_region()
        if region is None:
            QMessageBox.information(self, "PortableExcel", "Select a cell in a data range.")
            return
        r1, c1, r2, c2 = region
        _, key_col = self.view.current_cell()
        key_col = min(max(key_col, c1), c2)
        start = r1 + 1 if self._has_header(r1, c1, r2, c2) else r1
        if start >= r2:
            return
        sheet = self.book.active
        model = self.view.sheet_model
        rows = list(range(start, r2 + 1))
        keyed = sorted(rows, key=lambda r: ops.sort_key(model.value(r, key_col), descending))
        changes = {}
        for new_r, old_r in zip(rows, keyed):
            for c in range(c1, c2 + 1):
                raw = sheet.cells.get((old_r, c))
                if raw and raw.startswith("="):
                    raw = engine_mod.shift_formula(raw, new_r - old_r, 0)
                changes[(new_r, c)] = {"raw": raw, "style": sheet.styles.get((old_r, c))}
        self.apply_changes(sheet, changes, "Sort", select=(r1, c1, r2, c2))

    def remove_duplicates(self):
        region = self._current_region()
        if region is None:
            return
        r1, c1, r2, c2 = region
        start = r1 + 1 if self._has_header(r1, c1, r2, c2) else r1
        sheet = self.book.active
        model = self.view.sheet_model
        seen = set()
        keep = []
        for r in range(start, r2 + 1):
            key = tuple(value_to_text(model.value(r, c)).lower() for c in range(c1, c2 + 1))
            if key in seen:
                continue
            seen.add(key)
            keep.append(r)
        removed = (r2 - start + 1) - len(keep)
        changes = {}
        for i, r in enumerate(range(start, r2 + 1)):
            for c in range(c1, c2 + 1):
                if i < len(keep):
                    src = keep[i]
                    raw = sheet.cells.get((src, c))
                    if raw and raw.startswith("="):
                        raw = engine_mod.shift_formula(raw, r - src, 0)
                    changes[(r, c)] = {"raw": raw, "style": sheet.styles.get((src, c))}
                else:
                    changes[(r, c)] = {"raw": None, "style": None}
        if removed:
            self.apply_changes(sheet, changes, "Remove Duplicates")
        QMessageBox.information(self, "PortableExcel",
                                f"{removed} duplicate values found and removed; "
                                f"{len(keep)} unique values remain.")

    # ======================================================================
    # Find & replace / go to
    # ======================================================================
    def show_find(self, replace=False):
        if self.find_dialog is None:
            d = dialogs.FindReplaceDialog(self, replace)
            d.find_next.connect(self.find_next)
            d.find_all.connect(self.find_all)
            d.replace_one.connect(self.replace_one)
            d.replace_all.connect(self.replace_all)
            self.find_dialog = d
        self.find_dialog.set_replace_mode(replace)
        r1, c1, r2, c2 = self.view.active_range()
        self.find_dialog.show()
        self.find_dialog.raise_()
        self.find_dialog.what.setFocus()
        return self.find_dialog

    def _matches(self, sheet, r, c, raw, o):
        what = o["what"]
        if o["values"]:
            text = self.engine.display_text(sheet, r, c) if raw else ""
        else:
            text = raw
        a, b = (text, what) if o["case"] else (text.lower(), what.lower())
        return a == b if o["whole"] else (b in a)

    def _search_cells(self, o):
        sheets = self.book.sheets if o.get("workbook") else [self.book.active]
        out = []
        for sh in sheets:
            for (r, c) in sorted(sh.cells):
                if self._matches(sh, r, c, sh.cells[(r, c)], o):
                    out.append((sh, r, c))
        return out

    def find_next(self, o):
        if not o["what"]:
            return
        hits = self._search_cells(o)
        d = self.find_dialog
        if not hits:
            if d:
                d.status.setText("We couldn't find what you were looking for.")
            return
        cur_sheet = self.book.active
        r, c = self.view.current_cell()
        order = [s for s in self.book.sheets]
        cur_key = (order.index(cur_sheet), r, c)
        nxt = None
        for sh, hr, hc in hits:
            if (order.index(sh), hr, hc) > cur_key:
                nxt = (sh, hr, hc)
                break
        if nxt is None:
            nxt = hits[0]
        sh, hr, hc = nxt
        if sh is not self.book.active:
            self.activate_sheet(sh)
        self.view.goto(hr, hc)
        if d:
            d.status.setText(f"{sh.name}!{a1(hr, hc)}  ({len(hits)} match"
                             f"{'es' if len(hits) != 1 else ''})")

    def find_all(self, o):
        d = self.find_dialog
        hits = self._search_cells(o) if o["what"] else []
        d.results.clear()
        for sh, r, c in hits:
            d.results.addItem(f"{sh.name}  {a1(r, c, True)}   {sh.cells[(r, c)]}")
        d.results.setVisible(True)
        d.status.setText(f"{len(hits)} cell(s) found")
        try:
            d.results.itemActivated.disconnect()
        except (RuntimeError, TypeError):
            pass
        d.results.itemActivated.connect(
            lambda it, h=hits, w=d.results: self._goto_hit(h[w.row(it)]))

    def _goto_hit(self, hit):
        sh, r, c = hit
        if sh is not self.book.active:
            self.activate_sheet(sh)
        self.view.goto(r, c)

    def _replace_text(self, raw, o):
        what, rep = o["what"], o["with"]
        if o["whole"]:
            return rep
        if o["case"]:
            return raw.replace(what, rep)
        return re.sub(re.escape(what), lambda m: rep, raw, flags=re.I)

    def replace_one(self, o):
        sheet = self.book.active
        r, c = self.view.current_cell()
        raw = sheet.cells.get((r, c))
        if raw is not None and o["what"] and self._matches(sheet, r, c, raw, dict(o, values=False)):
            self.apply_changes(sheet, {(r, c): {"raw": self._replace_text(raw, o)}}, "Replace")
        self.find_next(dict(o, values=False))

    def replace_all(self, o):
        if not o["what"]:
            return
        hits = self._search_cells(dict(o, values=False))
        by_sheet: Dict = {}
        for sh, r, c in hits:
            by_sheet.setdefault(sh, {})[(r, c)] = {"raw": self._replace_text(sh.cells[(r, c)], o)}
        self.undo_stack.beginMacro("Replace All")
        for sh, ch in by_sheet.items():
            self.apply_changes(sh, ch, "Replace")
        self.undo_stack.endMacro()
        if self.find_dialog:
            self.find_dialog.status.setText(f"All done. We made {len(hits)} replacements.")
        return len(hits)

    def goto_dialog(self):
        text, ok = QInputDialog.getText(self, "Go To", "Reference:",
                                        text=self.name_box.currentText())
        if ok and text:
            self.name_box.setEditText(text)
            self._name_box_entered()

    # ======================================================================
    # Formulas
    # ======================================================================
    def insert_function(self):
        d = dialogs.InsertFunctionDialog(engine_mod.FUNCTIONS, self)
        if d.exec() and d.selected():
            self.insert_function_name(d.selected())

    def insert_function_name(self, name):
        ed = self.view.delegate.editor
        if ed is None and not self.fb_editing:
            self.view.start_edit("=" + name + "(")
            ed = self.view.delegate.editor
            if ed is not None:
                ed.end(False)
            return
        target = ed or self.formula_edit
        if not target.text():
            target.setText("=")
        target.insert(name + "(")
        target.textEdited.emit(target.text())

    def autosum(self, fn="SUM"):
        sheet = self.book.active
        model = self.view.sheet_model
        r1, c1, r2, c2 = self._clamp(*self.view.active_range())

        def isnum(r, c):
            v = model.value(r, c)
            return isinstance(v, (int, float)) and not isinstance(v, bool)

        if (r1, c1) != (r2, c2):
            changes = {}
            for c in range(c1, c2 + 1):
                last = r2
                if any(isnum(r, c) for r in range(r1, r2 + 1)):
                    target = r2 + 1 if sheet.cells.get((r2, c)) else r2
                    end = r2 if target == r2 + 1 else r2 - 1
                    if end >= r1:
                        changes[(target, c)] = {"raw": f"={fn}({a1(r1, c)}:{a1(end, c)})"}
            self.apply_changes(sheet, changes, "AutoSum")
            return
        r, c = r1, c1
        top = r - 1
        while top >= 0 and isnum(top, c):
            top -= 1
        if top < r - 1:
            formula = f"={fn}({a1(top + 1, c)}:{a1(r - 1, c)})"
        else:
            left = c - 1
            while left >= 0 and isnum(r, left):
                left -= 1
            formula = f"={fn}({a1(r, left + 1)}:{a1(r, c - 1)})" if left < c - 1 else f"={fn}()"
        self.view.start_edit(formula)
        ed = self.view.delegate.editor
        if ed is not None and formula.endswith("()"):
            ed.setCursorPosition(len(formula) - 1)

    def calculate_now(self):
        try:
            self.engine.invalidate()
            self.engine.recalc()
        except Exception as e:
            log.exception("recalc failed")
            QMessageBox.warning(self, "PortableExcel", f"Calculation failed:\n{e}")
        self.view.refresh()
        self.update_stats()

    def toggle_show_formulas(self):
        self.show_formulas = self.actions["show_formulas"].isChecked()
        self.view.refresh()

    def define_name(self):
        r1, c1, r2, c2 = self.view.active_range()
        ref = f"{self._sheet_ref(self.book.active)}!{a1(r1, c1, True)}" + \
            (f":{a1(r2, c2, True)}" if (r1, c1) != (r2, c2) else "")
        name, ok = QInputDialog.getText(self, "New Name", f"Name for {ref}:")
        if ok and name:
            if not re.fullmatch(r"[A-Za-z_\\][A-Za-z0-9_.]*", name) or parse_a1(name):
                QMessageBox.warning(self, "PortableExcel", "The name that you entered is not valid.")
                return
            self.book.defined_names[name] = ref
            self._refresh_name_list()
            self.engine.invalidate()
            self.view.refresh()
            self.mark_dirty()

    def name_manager(self):
        names = sorted(self.book.defined_names)
        if not names:
            QMessageBox.information(self, "Name Manager",
                                    "This workbook has no defined names.\nUse Define Name or type "
                                    "a name in the Name Box.")
            return
        items = [f"{n}  =  {self.book.defined_names[n]}" for n in names]
        item, ok = QInputDialog.getItem(self, "Name Manager",
                                        "Defined names (choose one to delete):", items, 0, False)
        if ok and item:
            n = names[items.index(item)]
            if QMessageBox.question(self, "Name Manager", f"Delete the name '{n}'?") == \
                    QMessageBox.Yes:
                self.book.defined_names.pop(n, None)
                self._refresh_name_list()
                self.engine.invalidate()
                self.view.refresh()
                self.mark_dirty()

    def format_cells(self, tab=0):
        r, c = self.view.current_cell()
        st = self.book.active.get_style(r, c)
        val = self.view.sheet_model.value(r, c)
        d = dialogs.FormatCellsDialog(st, lambda fmt: self.preview_format(val, fmt), self,
                                      tab if isinstance(tab, int) else 0)
        self._format_dialog = d
        if d.exec():
            ch = d.changes()
            if ch:
                self.apply_style(**ch)

    def preview_format(self, value, fmt):
        """Display text of value with number format fmt, via the engine."""
        if value is None:
            return ""
        from ..model import Workbook as _WB
        wb = _WB()
        sh = wb.sheets[0]
        raw = self._value_raw(value) or ""
        sh.set_raw(0, 0, raw)
        if fmt:
            sh.update_style(0, 0, number_format=fmt)
        try:
            return engine_mod.FormulaEngine(wb).display_text(sh, 0, 0)
        except Exception:
            return raw

    def format_as_table(self):
        region = self._current_region()
        if region is None:
            return
        r1, c1, r2, c2 = region
        sheet = self.book.active
        changes = {}
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                st = dict(sheet.styles.get((r, c), {}))
                if r == r1:
                    st.update(bold=True, fill="#4472C4", color="#FFFFFF")
                else:
                    st["fill"] = "#D9E1F2" if (r - r1) % 2 == 1 else "#FFFFFF"
                changes[(r, c)] = {"style": st}
        self.apply_changes(sheet, changes, "Format as Table", select=region)

    # ======================================================================
    # View
    # ======================================================================
    def set_zoom(self, z):
        z = round(max(0.5, min(2.0, z)), 2)
        if abs(z - self.zoom) < 0.001:
            return
        self.zoom = z
        self.zoom_slider.blockSignals(True)
        self.zoom_slider.setValue(round(z * 100))
        self.zoom_slider.blockSignals(False)
        self.zoom_label.setText(f"{round(z * 100)}%")
        cur = self.view.current_cell()
        self.view.apply_sizes()
        if self.view.overlays:
            self.view.set_frozen(*self.view.frozen)
        self.view.scrollTo(self.view.sheet_model.index(*cur))

    def zoom_dialog(self):
        items = ["200%", "150%", "125%", "100%", "75%", "50%"]
        cur = f"{round(self.zoom * 100)}%"
        item, ok = QInputDialog.getItem(self, "Zoom", "Magnification:", items,
                                        items.index(cur) if cur in items else 3, True)
        if ok:
            try:
                self.set_zoom(float(item.strip().rstrip("%")) / 100)
            except ValueError:
                pass

    def zoom_to_selection(self):
        r1, c1, r2, c2 = self._clamp(*self.view.active_range())
        w = sum(self.view.columnWidth(c) for c in range(c1, c2 + 1)) / self.zoom
        h = sum(self.view.rowHeight(r) for r in range(r1, r2 + 1)) / self.zoom
        vw, vh = self.view.viewport().width(), self.view.viewport().height()
        if w and h:
            self.set_zoom(min(vw / w, vh / h))
        self.view.scrollTo(self.view.sheet_model.index(r1, c1), SpreadsheetView.PositionAtTop)

    def toggle_gridlines(self):
        self.view.setShowGrid(self.actions["gridlines"].isChecked())

    def toggle_headings(self):
        on = self.actions["headings"].isChecked()
        self.view.horizontalHeader().setVisible(on)
        self.view.verticalHeader().setVisible(on)

    def toggle_formula_bar(self):
        self.formula_bar_widget.setVisible(self.actions["formula_bar"].isChecked())

    def freeze(self, kind):
        r, c = self.view.current_cell()
        if kind is None:
            fr = fc = 0
        elif kind == "row":
            fr, fc = 1, 0
        elif kind == "col":
            fr, fc = 0, 1
        else:
            if self.view.frozen != (0, 0):
                fr = fc = 0
            else:
                fr, fc = r, c
                if (fr, fc) == (0, 0):
                    fr = 1
        self._frozen[id(self.book.active)] = (fr, fc)
        self.view.set_frozen(fr, fc)
        self.actions["unfreeze"].setEnabled(bool(fr or fc))

    # ======================================================================
    # Sheets
    # ======================================================================
    def rebuild_sheet_tabs(self):
        tabs = self.sheet_tabs
        tabs.blockSignals(True)
        while tabs.count():
            tabs.removeTab(0)
        self._tab_sheets = []
        for sh in self.book.sheets:
            if sh.visible or sh is self.book.active:
                tabs.addTab(sh.name)
                self._tab_sheets.append(sh)
        if self.book.active in self._tab_sheets:
            tabs.setCurrentIndex(self._tab_sheets.index(self.book.active))
        tabs.blockSignals(False)

    def _sheet_tab_changed(self, i):
        if 0 <= i < len(getattr(self, "_tab_sheets", [])):
            self.activate_sheet(self._tab_sheets[i])

    def activate_sheet(self, sheet):
        if sheet is None:
            return
        if sheet is self.book.active and self.view.sheet_model.sheet is sheet:
            return
        if self.view.is_editing():
            self.view.commit_open_editor()
        old = self.book.active
        self._sheet_sel[id(old)] = (self.view.current_cell(), self.view.active_range(),
                                    self.view.verticalScrollBar().value(),
                                    self.view.horizontalScrollBar().value())
        self._fire_event(old.code_name, "Worksheet_Deactivate")
        self.book.active_index = self.book.sheets.index(sheet)
        sheet.visible = True
        self.view.set_sheet(sheet)
        self.view.set_frozen(*self._frozen.get(id(sheet), (0, 0)))
        saved = self._sheet_sel.get(id(sheet))
        if saved:
            (r, c), rng, vs, hs = saved
            self.view.select_range(*rng, cur=(r, c))
            self.view.verticalScrollBar().setValue(vs)
            self.view.horizontalScrollBar().setValue(hs)
        else:
            self.view.goto(0, 0)
        self.rebuild_sheet_tabs()
        self.on_current_changed()
        self._fire_event(sheet.code_name, "Worksheet_Activate")
        self._fire_event("ThisWorkbook", "Workbook_SheetActivate")

    def step_sheet(self, step):
        vis = [s for s in self.book.sheets if s.visible]
        if self.book.active not in vis:
            return
        i = vis.index(self.book.active) + step
        if 0 <= i < len(vis):
            self.activate_sheet(vis[i])

    def add_sheet(self, *_args, index=None):
        if index is None or isinstance(index, bool):
            index = self.book.active_index + 1
        holder = {}

        def fn():
            holder["sheet"] = self.book.add_sheet(None, index)
        if self._structural("Insert Sheet", fn):
            self.activate_sheet(holder["sheet"])
        return holder.get("sheet")

    def delete_sheet(self, sheet=None):
        sheet = sheet if sheet is not None and not isinstance(sheet, bool) else self.book.active
        if len([s for s in self.book.sheets if s.visible]) <= 1:
            QMessageBox.information(self, "PortableExcel",
                                    "A workbook must contain at least one visible worksheet.")
            return
        if sheet.cells and QMessageBox.question(
                self, "PortableExcel", f"Sheet '{sheet.name}' will be permanently deleted. "
                "Continue?", QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        def fn():
            idx = self.book.sheets.index(sheet)
            engine_mod.delete_sheet_refs(self.book, sheet.name)
            self.book.sheets.remove(sheet)
            vis = [i for i, s in enumerate(self.book.sheets) if s.visible]
            self.book.active_index = min(vis, key=lambda i: abs(i - idx)) if vis else 0
        self._structural("Delete Sheet", fn)

    def valid_sheet_name(self, name, sheet=None) -> Optional[str]:
        if not name or len(name) > 31 or re.search(r"[\[\]:*?/\\]", name) or \
                name.startswith("'") or name.endswith("'"):
            return ("The sheet name must be 1-31 characters and cannot contain : \\ / ? * [ ] "
                    "or begin/end with an apostrophe.")
        other = self.book.sheet_by_name(name)
        if other is not None and other is not sheet:
            return "That sheet name is already taken. Try a different one."
        return None

    def rename_sheet(self, sheet, new):
        err = self.valid_sheet_name(new, sheet)
        if err:
            QMessageBox.warning(self, "PortableExcel", err)
            return False
        old = sheet.name

        def fn():
            engine_mod.rename_sheet_refs(self.book, old, new)
            sheet.name = new
            for k, v in list(self.book.defined_names.items()):
                self.book.defined_names[k] = re.sub(
                    rf"^'?{re.escape(old)}'?!", f"{self._sheet_ref(sheet)}!", v)
        return self._structural("Rename Sheet", fn)

    def rename_sheet_prompt(self, sheet=None):
        sheet = sheet or self.book.active
        if sheet in self._tab_sheets:
            self.sheet_tabs.start_rename(self._tab_sheets.index(sheet))

    def _rename_from_tab(self, i, text):
        if 0 <= i < len(self._tab_sheets):
            self.rename_sheet(self._tab_sheets[i], text)

    def move_sheet(self, sheet, step):
        i = self.book.sheets.index(sheet)
        j = i + step
        if not 0 <= j < len(self.book.sheets):
            return

        def fn():
            self.book.sheets.insert(j, self.book.sheets.pop(i))
            self.book.active_index = self.book.sheets.index(sheet)
        self._structural("Move Sheet", fn)

    def hide_sheet(self, sheet=None):
        sheet = sheet if sheet is not None and not isinstance(sheet, bool) else self.book.active
        if len([s for s in self.book.sheets if s.visible]) <= 1:
            QMessageBox.information(self, "PortableExcel",
                                    "A workbook must contain at least one visible worksheet.")
            return

        def fn():
            sheet.visible = False
            vis = [i for i, s in enumerate(self.book.sheets) if s.visible]
            idx = self.book.sheets.index(sheet)
            self.book.active_index = min(vis, key=lambda i: abs(i - idx))
        self._structural("Hide Sheet", fn)

    def unhide_sheet(self):
        hidden = [s for s in self.book.sheets if not s.visible]
        if not hidden:
            return
        name, ok = QInputDialog.getItem(self, "Unhide", "Unhide sheet:",
                                        [s.name for s in hidden], 0, False)
        if ok:
            sh = self.book.sheet_by_name(name)

            def fn():
                sh.visible = True
                self.book.active_index = self.book.sheets.index(sh)
            self._structural("Unhide Sheet", fn)

    def _sheet_menu(self, i, pos):
        sheet = self._tab_sheets[i] if 0 <= i < len(self._tab_sheets) else self.book.active
        m = QMenu(self)
        m.addAction("Insert...", lambda: self.add_sheet(index=self.book.sheets.index(sheet)))
        m.addAction("Delete", lambda: self.delete_sheet(sheet))
        m.addAction("Rename", lambda: self.rename_sheet_prompt(sheet))
        m.addSeparator()
        a = m.addAction("Move Left", lambda: self.move_sheet(sheet, -1))
        a.setEnabled(self.book.sheets.index(sheet) > 0)
        a = m.addAction("Move Right", lambda: self.move_sheet(sheet, 1))
        a.setEnabled(self.book.sheets.index(sheet) < len(self.book.sheets) - 1)
        m.addSeparator()
        m.addAction("View Code", lambda: self._view_sheet_code(sheet))
        m.addSeparator()
        m.addAction("Hide", lambda: self.hide_sheet(sheet))
        a = m.addAction("Unhide...", self.unhide_sheet)
        a.setEnabled(any(not s.visible for s in self.book.sheets))
        self._last_menu = m
        m.popup(pos)
        return m

    def _view_sheet_code(self, sheet):
        ed = self.show_vba_editor()
        ed.open_module(sheet.code_name)

    # ======================================================================
    # Files
    # ======================================================================
    def book_title(self) -> str:
        if self.book.path:
            return os.path.basename(self.book.path)
        return getattr(self.book, "_pe_title", "Book1")

    def update_title(self):
        t = f"{self.book_title()}{'*' if self.modified else ''} - PortableExcel"
        self.setWindowTitle(t)
        if hasattr(self, "title_strip"):
            self.title_strip.set_title(t)
        if self.vba_editor is not None:
            self.vba_editor._update_title()

    def is_pristine(self) -> bool:
        return (not self.modified and not self.book.path and
                all(not s.cells for s in self.book.sheets) and not self.book.vba_modules)

    def maybe_save(self) -> bool:
        """True when it is OK to discard the current book."""
        if not self.modified:
            return True
        box = QMessageBox(self)
        box.setWindowTitle("PortableExcel")
        box.setIcon(QMessageBox.Warning)
        box.setText(f"Want to save your changes to '{self.book_title()}'?")
        save = box.addButton("Save", QMessageBox.AcceptRole)
        dont = box.addButton("Don't Save", QMessageBox.DestructiveRole)
        box.addButton("Cancel", QMessageBox.RejectRole)
        box.setDefaultButton(save)
        box.exec()
        clicked = box.clickedButton()
        if clicked is save:
            return self.save()
        return clicked is dont

    def new_workbook(self):
        w = MainWindow()
        w.show()
        return w

    def open_file(self, path: Optional[str] = None):
        if not path or isinstance(path, bool):
            start = os.path.dirname(self.book.path) if self.book.path else (
                os.path.dirname(config.recent_files()[0]) if config.recent_files() else "")
            path, _ = QFileDialog.getOpenFileName(self, "Open", start, fileio.SUPPORTED_OPEN)
            if not path:
                return None
        book, password = self._load_with_password(path)
        if book is None:
            return None
        config.add_recent(path)
        if self.is_pristine():
            self.set_book(book, password)
            return self
        w = MainWindow(book)
        w.save_password = password
        w.show()
        w.after_load()
        return w

    def _load_with_password(self, path):
        password = None
        error = ""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for _attempt in range(10):
                try:
                    book = fileio.load_workbook(path, password)
                    return book, password
                except fileio.PasswordRequired:
                    QApplication.restoreOverrideCursor()
                    d = dialogs.PasswordDialog(
                        "Password", f"'{os.path.basename(path)}' is protected.", error, False, self)
                    self._password_dialog = d
                    if not d.exec():
                        return None, None
                    password = d.password()
                    error = ("The password you supplied is not correct. Verify that the CAPS "
                             "LOCK key is off and be sure to use the correct capitalization.")
                    QApplication.setOverrideCursor(Qt.WaitCursor)
            return None, None
        except fileio.FileFormatError as e:
            QMessageBox.warning(self, "PortableExcel", f"PortableExcel cannot open "
                                f"'{os.path.basename(path)}'.\n\n{e}")
            return None, None
        except OSError as e:
            QMessageBox.warning(self, "PortableExcel", f"Could not open the file:\n{e}")
            return None, None
        except Exception as e:
            log.exception("load failed")
            QMessageBox.warning(self, "PortableExcel", f"Could not open the file:\n{e}")
            return None, None
        finally:
            while QApplication.overrideCursor() is not None:
                QApplication.restoreOverrideCursor()

    def set_book(self, book, password=None):
        try:
            self.book.change_listeners.remove(self._on_book_change)
        except ValueError:
            pass
        self.book = book
        self.book.change_listeners.append(self._on_book_change)
        self.save_password = password
        self.engine = engine_mod.FormulaEngine(book)
        self.runtime = None
        self.vba_stale = True
        self.macros_enabled = not book.vba_modules
        self.undo_stack.clear()
        self._frozen.clear()
        self._sheet_sel.clear()
        self.clip = None
        self._marquee = None
        self.modified = False
        self.view.set_sheet(book.active)
        self.view.set_frozen(0, 0)
        self.view.goto(0, 0)
        self.rebuild_sheet_tabs()
        self._refresh_name_list()
        self.update_title()
        self.on_current_changed()
        if self.vba_editor is not None:
            self.vba_editor.reload_from_book()
        self.after_load()

    def after_load(self):
        """Macro security handling after a workbook was opened."""
        if not self.book.vba_modules:
            self.message_bar.hide()
            return
        mode = config.get("macro_security", "disable_notify")
        if mode == "enable_all":
            self.enable_macros()
        elif mode == "disable_all":
            self.message_bar.hide()
            self.set_status_message("Macros have been disabled (Trust Center setting).")
        else:
            self.message_bar.show()

    def save(self) -> bool:
        if self.view.is_editing():
            self.view.commit_open_editor()
        if not self.book.path:
            return self.save_as()
        return self._save_to(self.book.path)

    def save_as(self) -> bool:
        if self.view.is_editing():
            self.view.commit_open_editor()
        base = self.book.path or os.path.join(
            os.path.dirname(config.recent_files()[0]) if config.recent_files() else "",
            self.book_title() + (".xlsm" if self.book.vba_modules else ".xlsx"))
        filters = fileio.SUPPORTED_SAVE
        default_filter = next((f for f in filters.split(";;") if
                               ("xlsm" if self.book.vba_modules else "(*.xlsx)") in f), "")
        path, chosen = QFileDialog.getSaveFileName(self, "Save As", base, filters, default_filter)
        if not path:
            return False
        if not os.path.splitext(path)[1]:
            m = re.search(r"\*(\.\w+)", chosen or "")
            path += m.group(1) if m else ".xlsx"
        return self._save_to(path)

    def save_with_password(self) -> bool:
        d = dialogs.PasswordDialog("Encrypt Document",
                                   "Encrypt the contents of this file. Leave empty to remove "
                                   "the password.", "", True, self)
        self._password_dialog = d
        if not d.exec():
            return False
        self.save_password = d.password() or None
        if self.book.path and os.path.splitext(self.book.path)[1].lower() == ".csv" and \
                self.save_password:
            QMessageBox.information(self, "PortableExcel",
                                    "CSV files cannot be encrypted; choose an Excel format.")
            return self.save_as()
        self.mark_dirty()
        return self.save()

    def _save_to(self, path) -> bool:
        ext = os.path.splitext(path)[1].lower()
        if self.book.vba_modules and ext in (".xlsx", ".xltx", ".csv"):
            r = QMessageBox.question(
                self, "PortableExcel",
                "The following features cannot be saved in macro-free workbooks:\n\n"
                "• VB project\n\nTo save a file with these features, click No, and then "
                "choose a macro-enabled file type.\nTo continue saving as a macro-free "
                "workbook, click Yes.", QMessageBox.Yes | QMessageBox.No)
            if r != QMessageBox.Yes:
                return self.save_as()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self._fire_event("ThisWorkbook", "Workbook_BeforeSave")
            warnings = fileio.save_workbook(self.book, path, engine=self.engine,
                                            password=self.save_password)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            log.exception("save failed")
            QMessageBox.critical(self, "PortableExcel", f"The file could not be saved:\n{e}")
            return False
        QApplication.restoreOverrideCursor()
        self.book.path = path
        self.book.vba_modified = False
        self.modified = False
        config.add_recent(path)
        self.update_title()
        self.set_status_message(f"Saved {os.path.basename(path)}")
        if warnings:
            self.show_save_warnings(warnings)
        return True

    def show_save_warnings(self, warnings, modal=True):
        d = dialogs.WarningsDialog("PortableExcel - Compatibility",
                                   "The workbook was saved, but some features may be lost or "
                                   "changed:", list(warnings), self)
        self._warn_dialog = d
        if modal:
            d.exec()
        return d

    def change_app_password(self):
        d = security.ChangePasswordDialog(self)
        self._pw_change_dialog = d
        d.exec()

    def lock_app(self):
        wins = [w for w in _WINDOWS if w.isVisible()]
        for w in wins:
            w.hide()
        dlg = security.LockDialog()
        if dlg.exec():
            for w in wins:
                w.show()
        else:
            for w in list(wins):
                w.modified = False
                w.close()

    def about(self):
        QMessageBox.about(self, "About PortableExcel",
                          "<b>PortableExcel</b><br>A portable spreadsheet with formulas, "
                          "xlsx/xlsm/csv files and a VBA interpreter.<br><br>"
                          f"Settings: {config.ini_path()}")

    def exit_app(self):
        for w in list(_WINDOWS):
            if not w.close():
                return
        QApplication.quit()

    def show_backstage(self):
        self.backstage.show_page("open" if not self.book.path and not self.modified else "info")
        self.stack.setCurrentIndex(1)

    def closeEvent(self, e):
        if self.macro_running:
            self.stop_macro()
            e.ignore()
            return
        self._fire_event("ThisWorkbook", "Workbook_BeforeClose")
        if not self.maybe_save():
            e.ignore()
            return
        if self.vba_editor is not None:
            self.vba_editor.close()
        if self.find_dialog is not None:
            self.find_dialog.close()
        try:
            self.book.change_listeners.remove(self._on_book_change)
        except ValueError:
            pass
        if self in _WINDOWS:
            _WINDOWS.remove(self)
        e.accept()

    # ======================================================================
    # Macros
    # ======================================================================
    def set_status_message(self, text):
        if text is None or text is False or text == "":
            self.msg_label.setText("")
        else:
            self.msg_label.setText(str(text))

    def immediate_print(self, text):
        self.immediate_log.append(text)
        del self.immediate_log[:-1000]
        if self.vba_editor is not None:
            self.vba_editor.print_immediate(text)

    def vba_code_changed(self):
        self.vba_stale = True
        self.mark_dirty()

    def macro_security(self):
        d = dialogs.MacroSecurityDialog(config.get("macro_security", "disable_notify"), self)
        self._macro_sec_dialog = d
        if d.exec():
            config.put("macro_security", d.value())
            if d.value() == "enable_all" and not self.macros_enabled:
                self.enable_macros()

    def _vba_available(self) -> bool:
        mod = _vba()
        return mod is not None and hasattr(mod, "VBARuntime")

    def enable_macros(self) -> bool:
        self.message_bar.hide()
        self.macros_enabled = True
        if not self.book.vba_modules:
            return True
        if not self.ensure_runtime():
            return False
        self._fire_event("ThisWorkbook", "Workbook_Open", force=True)
        self.view.refresh()
        return True

    def ensure_runtime(self) -> bool:
        if not self._vba_available():
            QMessageBox.warning(self, "PortableExcel", "The VBA interpreter is not available.")
            return False
        if self.runtime is None:
            try:
                self.runtime = _vba().VBARuntime(self.book, self.engine, GuiVBAHost(self))
            except Exception as e:
                log.exception("VBARuntime failed")
                QMessageBox.critical(self, "PortableExcel", f"Could not start VBA:\n{e}")
                return False
            self.vba_stale = True
        if self.vba_stale:
            try:
                errors = self.runtime.load_modules(self.book.vba_modules) or []
            except Exception as e:
                errors = [e]
            self.vba_stale = False
            if errors:
                self.show_compile_errors(errors)
        return True

    def show_compile_errors(self, errors):
        lines = []
        for e in errors:
            mod = getattr(e, "module", "") or ""
            line = getattr(e, "line", None)
            msg = getattr(e, "message", None) or str(e)
            lines.append(f"{mod}{(' line ' + str(line)) if line else ''}: {msg}")
        box = QMessageBox(self)
        box.setWindowTitle("PortableExcel VBA")
        box.setIcon(QMessageBox.Warning)
        box.setText("Compile error")
        box.setInformativeText("\n".join(lines[:20]))
        box.addButton(QMessageBox.Ok)
        dbg = box.addButton("Debug", QMessageBox.ActionRole)
        self._compile_box = box
        box.exec()
        if box.clickedButton() is dbg and errors:
            e = errors[0]
            ed = self.show_vba_editor()
            ed.show_error(getattr(e, "module", None), getattr(e, "line", None))

    def _check_macros_allowed(self) -> bool:
        if self.macros_enabled:
            return True
        r = QMessageBox.question(self, "PortableExcel",
                                 "Macros in this workbook are disabled.\n\nEnable the content "
                                 "of this workbook and run the macro?",
                                 QMessageBox.Yes | QMessageBox.No)
        if r == QMessageBox.Yes:
            return self.enable_macros()
        return False

    def list_macros(self) -> List[str]:
        if self.runtime is not None and self.macros_enabled and self._vba_available():
            try:
                if self.vba_stale:
                    self.ensure_runtime()
                return list(self.runtime.list_macros())
            except Exception:
                log.exception("list_macros failed")
        from .vba_editor import procedures
        out = []
        for m in self.book.vba_modules:
            if m.get("type") not in ("standard", "document"):
                continue
            for kind, name, params, _line in procedures(m.get("code", "")):
                if kind == "Sub" and not params:
                    out.append(name if m.get("type") == "standard" else f"{m['name']}.{name}")
        return out

    def run_macro(self, name: str, module: Optional[str] = None):
        if self.macro_running:
            return False
        if self.view.is_editing():
            self.view.commit_open_editor()
        if not self._check_macros_allowed() or not self.ensure_runtime():
            return False
        target = name
        try:
            names = self.runtime.list_macros()
            if module and f"{module}.{name}" in names:
                target = f"{module}.{name}"
        except Exception:
            pass
        self.macro_running = True
        self.runtime_stop_requested = False
        filt = _MacroStopFilter(self)
        QApplication.instance().installEventFilter(filt)
        self.set_status_message(f"Running macro {name}...  (Esc to stop)")
        ok = True
        try:
            self.runtime.run(target)
        except Exception as e:
            ok = False
            self._handle_vba_exception(e)
        finally:
            QApplication.instance().removeEventFilter(filt)
            self.macro_running = False
            self.set_status_message("")
            self.refresh_from_macro()
            self.undo_stack.clear()
        return ok

    def _handle_vba_exception(self, e):
        VBAError = _vba_error_class()
        if VBAError is not None and isinstance(e, VBAError):
            self.show_runtime_error(e)
        else:
            log.error("Macro failed: %s", "".join(traceback.format_exception(e)))
            box = QMessageBox(self)
            box.setWindowTitle("PortableExcel VBA")
            box.setIcon(QMessageBox.Critical)
            box.setText(f"Run-time error\n\nInternal error while running the macro:\n{e}")
            self._runtime_box = box
            box.exec()

    def show_runtime_error(self, e, modal=True):
        msg = getattr(e, "message", None) or str(e)
        module = getattr(e, "module", None)
        line = getattr(e, "line", None)
        num = getattr(e, "number", None)
        box = QMessageBox(self)
        box.setWindowTitle("PortableExcel VBA")
        box.setIcon(QMessageBox.Critical)
        head = f"Run-time error{(' ' + repr(num).strip(chr(39))) if num else ''}:"
        where = f"\n\n{module}" + (f", line {line}" if line else "") if module else ""
        box.setText(f"{head}\n\n{msg}{where}")
        end = box.addButton("End", QMessageBox.RejectRole)
        dbg = box.addButton("Debug", QMessageBox.AcceptRole)
        box.setDefaultButton(dbg)
        self._runtime_box = box
        if not modal:
            return box
        box.exec()
        if box.clickedButton() is dbg:
            ed = self.show_vba_editor()
            ed.show_error(module, line)
        return box

    def stop_macro(self):
        if self.runtime is not None:
            try:
                self.runtime.request_stop()
            except Exception:
                pass

    def refresh_from_macro(self):
        """Repaint after VBA changed cells/sheets."""
        try:
            self.engine.invalidate()
        except Exception:
            pass
        names = [s.name for s in self._tab_sheets] if hasattr(self, "_tab_sheets") else []
        if names != [s.name for s in self.book.sheets if s.visible] or \
                self.view.sheet_model.sheet is not self.book.active:
            self.rebuild_sheet_tabs()
            if self.view.sheet_model.sheet is not self.book.active:
                self.view.set_sheet(self.book.active)
        self.view.sheet_model.ensure_cell(*[max(0, x - 1) for x in
                                            self.book.active.used_bounds()])
        self.view.apply_sizes()
        self.view.refresh()
        self.on_current_changed()
        if self.vba_editor is not None:
            self.vba_editor.rebuild_tree()

    def _on_book_change(self, sheet, r, c):
        self.modified = True
        if not self._refresh_pending:
            self._refresh_pending = True
            QTimer.singleShot(0, self._deferred_refresh)

    def _deferred_refresh(self):
        self._refresh_pending = False
        try:
            self.engine.invalidate()
        except Exception:
            pass
        self.update_title()
        self.view.refresh()
        self.on_current_changed()

    def _fire_event(self, module, proc, args=(), force=False):
        if self.runtime is None or not self.macros_enabled or self._suspend_events:
            return
        if self.macro_running and not force:
            return
        code = "".join(m.get("code", "") for m in self.book.vba_modules
                       if m.get("name", "").lower() == module.lower())
        if proc.lower() not in code.lower():
            return
        if self.vba_stale and not self.ensure_runtime():
            return
        try:
            self.runtime.fire_event(module, proc, tuple(args))
        except Exception as e:
            self._handle_vba_exception(e)
        self.view.refresh()

    def _fire_selection_change(self):
        if self.runtime is None or not self.macros_enabled:
            return
        sh = self.book.active
        code = "".join(m.get("code", "") for m in self.book.vba_modules
                       if m.get("name", "").lower() == sh.code_name.lower())
        if "worksheet_selectionchange" not in code.lower():
            return
        r1, c1, r2, c2 = self.view.active_range()
        try:
            rng = self.runtime.make_range(sh, r1, c1, r2, c2)
        except Exception:
            return
        self._fire_event(sh.code_name, "Worksheet_SelectionChange", (rng,))

    def show_vba_editor(self):
        from .vba_editor import VBAEditor
        if self.vba_editor is None:
            self.vba_editor = VBAEditor(self)
        self.vba_editor.show()
        self.vba_editor.raise_()
        self.vba_editor.activateWindow()
        return self.vba_editor

    def show_macros_dialog(self):
        d = dialogs.MacrosDialog(self.list_macros(), self, self.book_title())
        self._macros_dialog = d
        if not d.exec():
            return
        name = d.macro_name()
        if d.action == "run":
            mod = None
            if "." in name:
                mod, _, name = name.partition(".")
            self.run_macro(name, module=mod)
        elif d.action == "edit":
            self.edit_macro(name)
        elif d.action == "create":
            self.create_macro(name)
        elif d.action == "delete":
            self.delete_macro(name)

    def run_macro_prompt(self):
        macros = self.list_macros()
        if not macros:
            QMessageBox.information(self, "PortableExcel", "This workbook has no macros.\n"
                                    "Use Visual Basic (Alt+F11) to write one.")
            return
        name, ok = QInputDialog.getItem(self, "Run Macro", "Macro:", macros, 0, False)
        if ok and name:
            mod = None
            if "." in name:
                mod, _, name = name.partition(".")
            self.run_macro(name, module=mod)

    def _find_proc(self, name):
        from .vba_editor import procedures
        mod_name = None
        if "." in name:
            mod_name, _, name = name.partition(".")
        for m in self.book.vba_modules:
            if mod_name and m["name"].lower() != mod_name.lower():
                continue
            for kind, n, params, line in procedures(m.get("code", "")):
                if n.lower() == name.lower():
                    return m, line
        return None, None

    def edit_macro(self, name):
        m, line = self._find_proc(name)
        ed = self.show_vba_editor()
        if m is not None:
            code_ed = ed.open_module(m["name"])
            if code_ed is not None and line:
                from PySide6.QtGui import QTextCursor
                blk = code_ed.document().findBlockByNumber(line - 1)
                code_ed.setTextCursor(QTextCursor(blk))
                code_ed.centerCursor()

    def create_macro(self, name):
        ed = self.show_vba_editor()
        target = next((m for m in self.book.vba_modules if m.get("type") == "standard"), None)
        if target is None:
            target = ed.insert_module("standard")
        code = target.get("code", "")
        if code and not code.endswith("\n"):
            code += "\n"
        code += f"\nSub {name}()\n    \nEnd Sub\n"
        target["code"] = code
        self.book.vba_modified = True
        self.vba_code_changed()
        ed.reload_from_book()
        self.edit_macro(name)

    def delete_macro(self, name):
        m, line = self._find_proc(name)
        if m is None:
            return
        if QMessageBox.question(self, "PortableExcel", f"Do you want to delete the macro "
                                f"{name}?") != QMessageBox.Yes:
            return
        lines = m["code"].split("\n")
        start = line - 1
        end = start
        while end < len(lines) and not re.match(r"^\s*End\s+(Sub|Function)\b", lines[end], re.I):
            end += 1
        del lines[start:end + 1]
        m["code"] = "\n".join(lines)
        self.book.vba_modified = True
        self.vba_code_changed()
        if self.vba_editor is not None:
            self.vba_editor.reload_from_book()

    def immediate_execute(self, text) -> str:
        if not self._vba_available():
            return "VBA interpreter not available"
        if not self.macros_enabled and not self._check_macros_allowed():
            return "Macros are disabled"
        if not self.ensure_runtime():
            return ""
        try:
            out = self.runtime.evaluate_immediate(text)
        except Exception as e:
            out = f"Error: {getattr(e, 'message', None) or e}"
        self.refresh_from_macro()
        return "" if out is None else str(out)

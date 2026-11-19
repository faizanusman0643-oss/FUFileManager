"""PortableExcel VBA editor window (project tree, code editor, Immediate window)."""
from __future__ import annotations

import os
import re
from typing import Dict, List, Optional

from PySide6.QtCore import QRect, QRegularExpression, QSize, Qt, Signal
from PySide6.QtGui import (QAction, QColor, QFont, QFontDatabase, QKeySequence, QPainter,
                           QSyntaxHighlighter, QTextCharFormat, QTextCursor, QTextDocument,
                           QTextFormat)
from PySide6.QtWidgets import (QComboBox, QDockWidget, QFileDialog, QFormLayout, QHBoxLayout,
                               QInputDialog, QLabel, QLineEdit, QMainWindow, QMessageBox,
                               QPlainTextEdit, QTabWidget, QTextEdit, QToolBar, QTreeWidget,
                               QTreeWidgetItem, QVBoxLayout, QWidget)

from . import icons

VBA_KEYWORDS = """
AddressOf Alias And As Boolean ByRef Byte ByVal Call Case Const Currency Date Debug Declare
Dim Do Double Each Else ElseIf Empty End Enum Erase Error Event Exit Explicit False For
Friend Function Get GoSub GoTo If Implements In Integer Is Let Lib Like Long LongLong Loop
Me Mod New Next Not Nothing Null Object On Option Optional Or ParamArray Preserve Print
Private Property Public RaiseEvent ReDim Resume Return Select Set Single Static Step Stop
String Sub Then To True Type TypeOf Until Variant Wend While With WithEvents Xor Base
Compare Binary Text Module Explicit Attribute
""".split()

_PROC_RE = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                      r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*\(([^)]*)\)",
                      re.I)
_END_PROC_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
_BLOCK_OPEN = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
                         r"(Sub|Function|Property|If\b.*\bThen\s*$|For\b|Do\b|While\b|"
                         r"Select\s+Case|With\b|Else\b|ElseIf\b|Case\b|Type\b|Enum\b)", re.I)

TYPE_GROUPS = [("document", "Microsoft Excel Objects"), ("form", "Forms"),
               ("standard", "Modules"), ("class", "Class Modules")]


def procedures(code: str) -> List[tuple]:
    """[(kind, name, params, line_no 1-based)]"""
    out = []
    for i, line in enumerate(code.splitlines(), 1):
        m = _PROC_RE.match(line)
        if m:
            out.append((m.group(1).split()[0].title(), m.group(2), m.group(3).strip(), i))
    return out


def procedure_at(code: str, line_no: int) -> Optional[tuple]:
    """Procedure containing 1-based line number."""
    lines = code.splitlines()
    cur = None
    for i, line in enumerate(lines, 1):
        m = _PROC_RE.match(line)
        if m:
            cur = (m.group(1).split()[0].title(), m.group(2), m.group(3).strip(), i)
        if i >= line_no:
            return cur
        if _END_PROC_RE.match(line) and i < line_no:
            cur = None
    return None


class VBAHighlighter(QSyntaxHighlighter):
    def __init__(self, doc):
        super().__init__(doc)
        self.kw = QTextCharFormat()
        self.kw.setForeground(QColor("#0000FF"))
        self.comment = QTextCharFormat()
        self.comment.setForeground(QColor("#008000"))
        self.string = QTextCharFormat()
        self.string.setForeground(QColor("#A31515"))
        self.kw_re = QRegularExpression(r"\b(" + "|".join(VBA_KEYWORDS) + r")\b",
                                        QRegularExpression.CaseInsensitiveOption)

    def highlightBlock(self, text):
        it = self.kw_re.globalMatch(text)
        while it.hasNext():
            m = it.next()
            self.setFormat(m.capturedStart(), m.capturedLength(), self.kw)
        # strings and comments (a ' outside a string starts a comment)
        in_str = False
        start = 0
        i = 0
        while i < len(text):
            ch = text[i]
            if ch == '"':
                if in_str:
                    if i + 1 < len(text) and text[i + 1] == '"':
                        i += 2
                        continue
                    self.setFormat(start, i - start + 1, self.string)
                    in_str = False
                else:
                    in_str = True
                    start = i
            elif ch == "'" and not in_str:
                self.setFormat(i, len(text) - i, self.comment)
                return
            elif not in_str and text[i:i + 4].lower() == "rem " and text[:i].strip() == "":
                self.setFormat(i, len(text) - i, self.comment)
                return
            i += 1
        if in_str:
            self.setFormat(start, len(text) - start, self.string)


class _LineArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor

    def sizeHint(self):
        return QSize(self.editor.line_area_width(), 0)

    def paintEvent(self, e):
        self.editor.paint_line_area(e)


class CodeEditor(QPlainTextEdit):
    def __init__(self, parent=None, read_only=False):
        super().__init__(parent)
        f = QFontDatabase.systemFont(QFontDatabase.FixedFont)
        for fam in ("Consolas", "Courier New", "DejaVu Sans Mono", "Liberation Mono"):
            if fam in QFontDatabase.families():
                f = QFont(fam)
                break
        f.setPointSizeF(10)
        self.setFont(f)
        self.setLineWrapMode(QPlainTextEdit.NoWrap)
        self.setTabStopDistance(self.fontMetrics().horizontalAdvance(" ") * 4)
        self.setReadOnly(read_only)
        self.highlighter = VBAHighlighter(self.document())
        self.line_area = _LineArea(self)
        self.blockCountChanged.connect(self._update_width)
        self.updateRequest.connect(self._update_area)
        self._update_width()
        self.marked_line = None
        self.marked_color = QColor("#FFFF00")
        self.cursorPositionChanged.connect(self._refresh_extra)
        self._refresh_extra()

    def line_area_width(self):
        digits = max(3, len(str(self.blockCount())))
        return 10 + self.fontMetrics().horizontalAdvance("9") * digits

    def _update_width(self, *_):
        self.setViewportMargins(self.line_area_width(), 0, 0, 0)

    def _update_area(self, rect, dy):
        if dy:
            self.line_area.scroll(0, dy)
        else:
            self.line_area.update(0, rect.y(), self.line_area.width(), rect.height())

    def resizeEvent(self, e):
        super().resizeEvent(e)
        cr = self.contentsRect()
        self.line_area.setGeometry(QRect(cr.left(), cr.top(), self.line_area_width(), cr.height()))

    def paint_line_area(self, e):
        p = QPainter(self.line_area)
        p.fillRect(e.rect(), QColor("#F3F3F3"))
        block = self.firstVisibleBlock()
        num = block.blockNumber()
        top = round(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + round(self.blockBoundingRect(block).height())
        p.setPen(QColor("#8A8A8A"))
        while block.isValid() and top <= e.rect().bottom():
            if block.isVisible() and bottom >= e.rect().top():
                if self.marked_line == num + 1:
                    p.fillRect(0, top, self.line_area.width(), bottom - top, self.marked_color)
                    p.setPen(QColor("#000000"))
                p.drawText(0, top, self.line_area.width() - 5, self.fontMetrics().height(),
                           Qt.AlignRight, str(num + 1))
                p.setPen(QColor("#8A8A8A"))
            block = block.next()
            top = bottom
            bottom = top + round(self.blockBoundingRect(block).height())
            num += 1
        p.end()

    def _refresh_extra(self):
        sels = []
        if not self.isReadOnly():
            cur = QTextEdit.ExtraSelection()
            cur.format.setBackground(QColor("#F5FAF7"))
            cur.format.setProperty(QTextFormat.FullWidthSelection, True)
            cur.cursor = self.textCursor()
            cur.cursor.clearSelection()
            sels.append(cur)
        if self.marked_line:
            blk = self.document().findBlockByNumber(self.marked_line - 1)
            if blk.isValid():
                m = QTextEdit.ExtraSelection()
                m.format.setBackground(self.marked_color)
                m.format.setProperty(QTextFormat.FullWidthSelection, True)
                m.cursor = QTextCursor(blk)
                sels.append(m)
        self.setExtraSelections(sels)

    def mark_line(self, line: Optional[int], color="#FFFF00"):
        self.marked_line = line
        self.marked_color = QColor(color)
        if line:
            blk = self.document().findBlockByNumber(line - 1)
            if blk.isValid():
                c = QTextCursor(blk)
                self.setTextCursor(c)
                self.centerCursor()
        self._refresh_extra()
        self.line_area.update()

    def current_line(self) -> int:
        return self.textCursor().blockNumber() + 1

    def keyPressEvent(self, e):
        if self.isReadOnly():
            return super().keyPressEvent(e)
        if self.marked_line and e.text():
            self.mark_line(None)
        if e.key() == Qt.Key_Tab and not e.modifiers():
            c = self.textCursor()
            if c.hasSelection():
                self._indent_selection(+1)
            else:
                col = c.positionInBlock()
                c.insertText(" " * (4 - col % 4))
            return
        if e.key() == Qt.Key_Backtab:
            self._indent_selection(-1)
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            c = self.textCursor()
            line = c.block().text()
            indent = line[: len(line) - len(line.lstrip())]
            stripped = line.strip()
            if _BLOCK_OPEN.match(stripped) and not re.match(r"^\s*(Exit|End)\b", stripped, re.I) \
                    and not re.search(r"\bThen\s+\S", stripped, re.I) \
                    and not stripped.lower().startswith(("declare", "private declare",
                                                         "public declare")):
                indent += "    "
            super().keyPressEvent(e)
            self.textCursor().insertText(indent)
            return
        super().keyPressEvent(e)
        # dedent closing keywords
        if e.text() and e.text().isalpha():
            c = self.textCursor()
            line = c.block().text()
            if re.fullmatch(r"\s+(End|Next|Loop|Wend|Else|ElseIf|Case)", line, re.I):
                pass

    def _indent_selection(self, direction):
        c = self.textCursor()
        doc = self.document()
        start = doc.findBlock(c.selectionStart()).blockNumber()
        end = doc.findBlock(c.selectionEnd()).blockNumber()
        c.beginEditBlock()
        for n in range(start, end + 1):
            blk = doc.findBlockByNumber(n)
            cc = QTextCursor(blk)
            if direction > 0:
                cc.insertText("    ")
            else:
                t = blk.text()
                k = min(4, len(t) - len(t.lstrip(" ")))
                for _ in range(k):
                    cc.deleteChar()
        c.endEditBlock()


class ImmediateWindow(QPlainTextEdit):
    """Type a statement (or ?expr) and press Enter."""
    execute = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        f = QFontDatabase.systemFont(QFontDatabase.FixedFont)
        f.setPointSizeF(9.5)
        self.setFont(f)
        self.setPlaceholderText("Immediate: type ?expression or a statement and press Enter")

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and not e.modifiers() & Qt.ShiftModifier:
            line = self.textCursor().block().text()
            c = self.textCursor()
            c.movePosition(QTextCursor.EndOfBlock)
            self.setTextCursor(c)
            super().keyPressEvent(e)
            if line.strip():
                self.execute.emit(line.strip())
            return
        super().keyPressEvent(e)

    def print_line(self, text: str):
        c = self.textCursor()
        c.movePosition(QTextCursor.End)
        if self.toPlainText() and not self.toPlainText().endswith("\n"):
            c.insertText("\n")
        c.insertText(text + "\n")
        self.setTextCursor(c)
        self.ensureCursorVisible()


class VBAEditor(QMainWindow):
    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setWindowFlag(Qt.Window, True)
        self.setWindowIcon(icons.icon("vba"))
        self.resize(1100, 720)
        self.editors: Dict[str, CodeEditor] = {}
        self._loading = False
        self._last_find = ""
        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.setDocumentMode(True)
        self.tabs.tabCloseRequested.connect(self._close_tab)
        self.tabs.currentChanged.connect(self._tab_changed)
        center = QWidget()
        cl = QVBoxLayout(center)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        top = QHBoxLayout()
        top.setContentsMargins(4, 3, 4, 3)
        self.obj_combo = QComboBox()
        self.obj_combo.setMinimumWidth(200)
        self.proc_combo = QComboBox()
        self.proc_combo.setMinimumWidth(260)
        self.proc_combo.activated.connect(self._goto_proc)
        top.addWidget(self.obj_combo)
        top.addWidget(self.proc_combo)
        top.addStretch(1)
        cl.addLayout(top)
        cl.addWidget(self.tabs, 1)
        self.setCentralWidget(center)
        # project tree
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.itemDoubleClicked.connect(self._tree_open)
        self.tree.itemClicked.connect(self._tree_select)
        dock = QDockWidget("Project - VBAProject", self)
        dock.setObjectName("projectDock")
        dock.setWidget(self.tree)
        dock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        self.addDockWidget(Qt.LeftDockWidgetArea, dock)
        # properties
        props = QWidget()
        pf = QFormLayout(props)
        self.prop_name = QLineEdit()
        self.prop_name.editingFinished.connect(self._rename_module)
        self.prop_type = QLabel("")
        pf.addRow("(Name)", self.prop_name)
        pf.addRow("Type", self.prop_type)
        pdock = QDockWidget("Properties", self)
        pdock.setObjectName("propsDock")
        pdock.setWidget(props)
        pdock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        self.addDockWidget(Qt.LeftDockWidgetArea, pdock)
        # immediate window
        self.immediate = ImmediateWindow()
        self.immediate.execute.connect(self._immediate_exec)
        idock = QDockWidget("Immediate", self)
        idock.setObjectName("immediateDock")
        idock.setWidget(self.immediate)
        idock.setFeatures(QDockWidget.DockWidgetMovable | QDockWidget.DockWidgetFloatable)
        self.addDockWidget(Qt.BottomDockWidgetArea, idock)
        self.resizeDocks([idock], [150], Qt.Vertical)
        self.resizeDocks([dock], [240], Qt.Horizontal)
        self._build_toolbar()
        self.statusBar().showMessage("Ready")
        for line in win.immediate_log:
            self.immediate.print_line(line)
        self.rebuild_tree()
        self._update_title()

    # toolbar -------------------------------------------------------------------
    def _act(self, text, icon, slot, shortcut=None, tip=None):
        a = QAction(icons.icon(icon), text, self)
        a.triggered.connect(slot)
        if shortcut:
            a.setShortcut(QKeySequence(shortcut))
            a.setShortcutContext(Qt.WindowShortcut)
        a.setToolTip(tip or (f"{text} ({shortcut})" if shortcut else text))
        self.addAction(a)
        return a

    def _build_toolbar(self):
        tb = QToolBar("Standard")
        tb.setObjectName("vbeToolbar")
        tb.setIconSize(QSize(18, 18))
        tb.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.addToolBar(tb)
        self.a_save = self._act("Save", "save", self.win.save, "Ctrl+S")
        self.a_mod = self._act("Insert Module", "module", lambda: self.insert_module("standard"))
        self.a_cls = self._act("Insert Class Module", "class_module",
                               lambda: self.insert_module("class"))
        self.a_rem = self._act("Remove Module", "remove", self.remove_module)
        self.a_imp = self._act("Import File...", "import", self.import_file, "Ctrl+M")
        self.a_exp = self._act("Export File...", "export", self.export_file, "Ctrl+E")
        self.a_run = self._act("Run Sub", "run", self.run_current, "F5",
                               "Run the Sub where the cursor is (F5)")
        self.a_stop = self._act("Stop", "stop", self.win.stop_macro, "Ctrl+Break")
        self.a_find = self._act("Find", "find", self.find, "Ctrl+F")
        self.a_find_next = self._act("Find Next", "find", self.find_next, "F3")
        self.a_view_excel = self._act("View PortableExcel", "app", self._show_main, "Alt+F11")
        for a in (self.a_save, None, self.a_mod, self.a_cls, self.a_rem, None, self.a_imp,
                  self.a_exp, None, self.a_run, self.a_stop, None, self.a_find,
                  self.a_view_excel):
            if a is None:
                tb.addSeparator()
            else:
                tb.addAction(a)

    def _show_main(self):
        self.win.raise_()
        self.win.activateWindow()

    def _update_title(self):
        self.setWindowTitle(f"PortableExcel VBA Editor - {self.win.book_title()}")

    # modules -------------------------------------------------------------------
    @property
    def book(self):
        return self.win.book

    def module(self, name: str) -> Optional[dict]:
        for m in self.book.vba_modules:
            if m.get("name", "").lower() == name.lower():
                return m
        return None

    def _document_names(self) -> List[str]:
        names = ["ThisWorkbook"] + [s.code_name for s in self.book.sheets]
        for m in self.book.vba_modules:
            if m.get("type") == "document" and m["name"] not in names:
                names.append(m["name"])
        return names

    def rebuild_tree(self):
        self.tree.clear()
        root = QTreeWidgetItem([f"VBAProject ({self.win.book_title()})"])
        root.setIcon(0, icons.icon("folder"))
        self.tree.addTopLevelItem(root)
        groups = {}
        for t, label in TYPE_GROUPS:
            g = QTreeWidgetItem([label])
            g.setIcon(0, icons.icon("folder"))
            groups[t] = g
        sheet_by_code = {s.code_name.lower(): s for s in self.book.sheets}
        for name in self._document_names():
            sh = sheet_by_code.get(name.lower())
            label = f"{name} ({sh.name})" if sh else name
            it = QTreeWidgetItem([label])
            it.setData(0, Qt.UserRole, name)
            it.setIcon(0, icons.icon("workbook_obj" if name == "ThisWorkbook" else "sheet_obj"))
            groups["document"].addChild(it)
        for m in self.book.vba_modules:
            t = m.get("type", "standard")
            if t == "document":
                continue
            it = QTreeWidgetItem([m["name"] + (" (read-only)" if t == "form" else "")])
            it.setData(0, Qt.UserRole, m["name"])
            it.setIcon(0, icons.icon({"class": "class_module", "form": "form"}.get(t, "module")))
            groups.get(t, groups["standard"]).addChild(it)
        for t, _ in TYPE_GROUPS:
            if groups[t].childCount() or t == "document":
                root.addChild(groups[t])
        self.tree.expandAll()
        self.obj_combo.clear()
        self.obj_combo.addItems(self._document_names() +
                                [m["name"] for m in self.book.vba_modules
                                 if m.get("type") != "document"])

    def _module_type(self, name):
        m = self.module(name)
        if m:
            return m.get("type", "standard")
        return "document"

    def _tree_select(self, item, _col=0):
        name = item.data(0, Qt.UserRole)
        if not name:
            return
        self.prop_name.setText(name)
        t = self._module_type(name)
        self.prop_type.setText(t)
        self.prop_name.setReadOnly(t in ("document", "form"))

    def _tree_open(self, item, _col=0):
        name = item.data(0, Qt.UserRole)
        if name:
            self.open_module(name)

    def open_module(self, name: str, line: Optional[int] = None) -> Optional[CodeEditor]:
        m = self.module(name)
        if m is None and name not in self._document_names():
            return None
        key = (m["name"] if m else name)
        ed = self.editors.get(key.lower())
        if ed is None:
            ro = bool(m and m.get("type") == "form")
            ed = CodeEditor(read_only=ro)
            self._loading = True
            ed.setPlainText(m.get("code", "") if m else "")
            self._loading = False
            ed.textChanged.connect(lambda k=key, e=ed: self._code_changed(k, e))
            ed.cursorPositionChanged.connect(self._sync_proc_combo)
            self.editors[key.lower()] = ed
            sh = {s.code_name.lower(): s for s in self.book.sheets}.get(key.lower())
            title = f"{key} ({sh.name})" if sh else key
            self.tabs.addTab(ed, icons.icon("module"), title)
        self.tabs.setCurrentWidget(ed)
        if line:
            ed.mark_line(line)
        ed.setFocus()
        return ed

    def current_module_name(self) -> Optional[str]:
        ed = self.tabs.currentWidget()
        for k, e in self.editors.items():
            if e is ed:
                m = self.module(k)
                return m["name"] if m else next((n for n in self._document_names()
                                                 if n.lower() == k), k)
        return None

    def _close_tab(self, i):
        ed = self.tabs.widget(i)
        for k, e in list(self.editors.items()):
            if e is ed:
                del self.editors[k]
        self.tabs.removeTab(i)
        ed.deleteLater()

    def _tab_changed(self, _i):
        name = self.current_module_name()
        if name:
            self.obj_combo.setCurrentText(name)
        self._fill_proc_combo()

    def _fill_proc_combo(self):
        self.proc_combo.clear()
        ed = self.tabs.currentWidget()
        if not isinstance(ed, CodeEditor):
            return
        self.proc_combo.addItem("(General)", 0)
        for kind, name, params, line in procedures(ed.toPlainText()):
            self.proc_combo.addItem(f"{name}", line)
        self._sync_proc_combo()

    def _sync_proc_combo(self):
        ed = self.tabs.currentWidget()
        if not isinstance(ed, CodeEditor):
            return
        p = procedure_at(ed.toPlainText(), ed.current_line())
        idx = self.proc_combo.findText(p[1]) if p else 0
        if idx >= 0 and idx != self.proc_combo.currentIndex():
            self.proc_combo.blockSignals(True)
            self.proc_combo.setCurrentIndex(idx)
            self.proc_combo.blockSignals(False)

    def _goto_proc(self, i):
        ed = self.tabs.currentWidget()
        line = self.proc_combo.itemData(i)
        if isinstance(ed, CodeEditor) and line:
            blk = ed.document().findBlockByNumber(line - 1)
            ed.setTextCursor(QTextCursor(blk))
            ed.centerCursor()
            ed.setFocus()

    def _code_changed(self, key, ed):
        if self._loading:
            return
        m = self.module(key)
        if m is None:
            m = {"name": key, "type": "document", "code": ""}
            self.book.vba_modules.append(m)
        if m.get("type") == "form":
            return
        m["code"] = ed.toPlainText()
        self.book.vba_modified = True
        self.win.vba_code_changed()
        n_before = self.proc_combo.count()
        if len(procedures(m["code"])) + 1 != n_before:
            self._fill_proc_combo()

    def _unique_name(self, base):
        i = 1
        names = {m["name"].lower() for m in self.book.vba_modules} | \
            {n.lower() for n in self._document_names()}
        while f"{base}{i}".lower() in names:
            i += 1
        return f"{base}{i}"

    def insert_module(self, kind="standard", name=None, code=None):
        name = name or self._unique_name("Module" if kind == "standard" else "Class")
        m = {"name": name, "type": kind, "code": code if code is not None else
             "Option Explicit\n\n"}
        self.book.vba_modules.append(m)
        self.book.vba_modified = True
        self.win.vba_code_changed()
        self.rebuild_tree()
        ed = self.open_module(name)
        if ed:
            c = ed.textCursor()
            c.movePosition(QTextCursor.End)
            ed.setTextCursor(c)
        return m

    def remove_module(self):
        name = self.current_module_name()
        it = self.tree.currentItem()
        if it is not None and it.data(0, Qt.UserRole):
            name = it.data(0, Qt.UserRole)
        m = self.module(name) if name else None
        if not m or m.get("type") == "document":
            QMessageBox.information(self, "PortableExcel VBA Editor",
                                    "Select a standard module, class module or form to remove.")
            return
        r = QMessageBox.question(self, "PortableExcel VBA Editor",
                                 f"Do you want to export {name} before removing it?",
                                 QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        if r == QMessageBox.Cancel:
            return
        if r == QMessageBox.Yes:
            self.export_file(name)
        self.book.vba_modules.remove(m)
        ed = self.editors.pop(name.lower(), None)
        if ed is not None:
            self.tabs.removeTab(self.tabs.indexOf(ed))
            ed.deleteLater()
        self.book.vba_modified = True
        self.win.vba_code_changed()
        self.rebuild_tree()

    def export_file(self, name=None):
        name = name if isinstance(name, str) else self.current_module_name()
        m = self.module(name) if name else None
        if m is None:
            QMessageBox.information(self, "PortableExcel VBA Editor", "Open a module first.")
            return
        ext = {"class": ".cls", "form": ".frm"}.get(m.get("type"), ".bas")
        path, _ = QFileDialog.getSaveFileName(self, "Export File", name + ext,
                                              "Basic Files (*.bas);;Class Files (*.cls);;"
                                              "All files (*)")
        if not path:
            return
        code = m.get("code", "")
        if m.get("type") == "class":
            header = ("VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
                      f'Attribute VB_Name = "{name}"\r\n')
        else:
            header = f'Attribute VB_Name = "{name}"\r\n'
        with open(path, "w", encoding="utf-8", newline="") as fh:
            fh.write(header + code.replace("\r\n", "\n").replace("\n", "\r\n"))
        self.statusBar().showMessage(f"Exported {name} to {path}", 5000)

    def import_file(self, path=None):
        if not isinstance(path, str):
            path, _ = QFileDialog.getOpenFileName(self, "Import File", "",
                                                  "VB Files (*.bas *.cls *.frm);;All files (*)")
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as fh:
                text = fh.read()
        except OSError as e:
            QMessageBox.warning(self, "PortableExcel VBA Editor", str(e))
            return
        name, code, kind = parse_vb_file(text, path)
        if self.module(name):
            name = self._unique_name(name)
        self.insert_module(kind, name, code)

    # running ---------------------------------------------------------------------
    def run_current(self):
        ed = self.tabs.currentWidget()
        name = self.current_module_name()
        if isinstance(ed, CodeEditor) and name:
            p = procedure_at(ed.toPlainText(), ed.current_line())
            if p and p[0] == "Sub" and not p[2]:
                self.win.run_macro(p[1], module=name)
                return
        self.win.show_macros_dialog()

    def show_error(self, module: Optional[str], line: Optional[int]):
        if module:
            ed = self.open_module(module)
            if ed and line:
                ed.mark_line(line)
        self.show()
        self.raise_()

    def _immediate_exec(self, text):
        out = self.win.immediate_execute(text)
        if out:
            self.immediate.print_line(out)

    def print_immediate(self, text):
        self.immediate.print_line(text)

    # find --------------------------------------------------------------------------
    def find(self):
        text, ok = QInputDialog.getText(self, "Find", "Find what:", text=self._last_find)
        if ok and text:
            self._last_find = text
            self.find_next()

    def find_next(self):
        ed = self.tabs.currentWidget()
        if not isinstance(ed, CodeEditor) or not self._last_find:
            return
        if not ed.find(self._last_find):
            ed.moveCursor(QTextCursor.Start)
            if not ed.find(self._last_find):
                self.statusBar().showMessage("The specified text was not found.", 4000)

    def _rename_module(self):
        new = self.prop_name.text().strip()
        it = self.tree.currentItem()
        old = it.data(0, Qt.UserRole) if it else None
        m = self.module(old) if old else None
        if not m or m.get("type") in ("document", "form") or not new or new == old:
            return
        if not re.fullmatch(r"[A-Za-z]\w{0,30}", new) or self.module(new):
            QMessageBox.warning(self, "PortableExcel VBA Editor", "Invalid or duplicate name.")
            self.prop_name.setText(old)
            return
        m["name"] = new
        ed = self.editors.pop(old.lower(), None)
        if ed is not None:
            self.editors[new.lower()] = ed
            self.tabs.setTabText(self.tabs.indexOf(ed), new)
            try:
                ed.textChanged.disconnect()
            except (RuntimeError, TypeError):
                pass
            ed.textChanged.connect(lambda k=new, e=ed: self._code_changed(k, e))
        self.book.vba_modified = True
        self.win.vba_code_changed()
        self.rebuild_tree()

    def reload_from_book(self):
        """Book replaced (e.g. new file opened)."""
        for k, ed in list(self.editors.items()):
            self.tabs.removeTab(self.tabs.indexOf(ed))
            ed.deleteLater()
        self.editors.clear()
        self.rebuild_tree()
        self._update_title()


def parse_vb_file(text: str, path: str = ""):
    """Return (name, code, kind) from an exported .bas/.cls/.frm file."""
    ext = os.path.splitext(path)[1].lower()
    kind = {".cls": "class", ".frm": "form"}.get(ext, "standard")
    lines = text.replace("\r\n", "\n").split("\n")
    name = os.path.splitext(os.path.basename(path))[0] or "Module1"
    out = []
    in_header = False
    for i, line in enumerate(lines):
        s = line.strip()
        if i == 0 and s.upper().startswith("VERSION "):
            in_header = True
            continue
        if in_header:
            if s.upper() == "END":
                in_header = False
            continue
        m = re.match(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', s, re.I)
        if m:
            name = m.group(1)
            continue
        if re.match(r"Attribute\s+VB_", s, re.I):
            continue
        out.append(line)
    return name, "\n".join(out).lstrip("\n"), kind

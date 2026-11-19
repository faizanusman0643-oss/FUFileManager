"""Spreadsheet grid: a lazy table model over model.Sheet, an Excel-like
delegate/editor and the QTableView with Excel keyboard/mouse behaviour.

The model never creates per-cell items; every paint asks the formula engine
for the display text of the visible cells only.
"""
from __future__ import annotations

import re
from typing import List, Optional, Tuple

from PySide6.QtCore import (QAbstractTableModel, QEvent, QItemSelection,
                            QItemSelectionModel, QModelIndex, QPoint, QRect, QSize,
                            QStringListModel, Qt, QTimer, Signal)
from PySide6.QtGui import (QBrush, QColor, QFont, QFontMetrics, QKeySequence,
                           QPainter, QPalette, QPen, QTextOption)
from PySide6.QtWidgets import (QAbstractItemDelegate, QAbstractItemView, QApplication,
                               QCompleter, QHeaderView, QLineEdit, QStyle,
                               QStyledItemDelegate, QStyleOptionHeader, QTableView)

from ..model import (MAX_COLS, MAX_ROWS, XLError, a1, col_to_name, parse_a1,
                     parse_range_a1, value_to_text)

GREEN = QColor("#217346")
DARK_GREEN = QColor("#185C37")
GRID_COLOR = QColor("#E1E1E1")
HEADER_BG = QColor("#F3F3F3")
HEADER_SEL_BG = QColor("#D3D3D3")
HEADER_FULL_BG = QColor("#C6E7D3")
HEADER_LINE = QColor("#D4D4D4")
SEL_TINT = QColor(33, 115, 70, 34)

DEFAULT_COL_WIDTH = 64
DEFAULT_ROW_HEIGHT = 20
MIN_ROWS = 1000
MIN_COLS = 26
ROW_GROW = 1000
COL_GROW = 26

REF_COLORS = ["#4472C4", "#C0504D", "#7030A0", "#00A651", "#ED7D31", "#0E8A9A"]
_REF_RE = re.compile(r"(?<![A-Za-z0-9_.!])(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)"
                     r"(?![A-Za-z0-9_(])")
_POINT_CHARS = set("=(,+-*/^&<>:;% ")


def formula_refs(text: str) -> List[Tuple[int, int, int, int]]:
    """A1-style references (same sheet only) inside formula text."""
    if not text.startswith("="):
        return []
    out = []
    # ignore string literals
    stripped = re.sub(r'"[^"]*"', lambda m: " " * len(m.group(0)), text)
    for m in _REF_RE.finditer(stripped):
        start = m.start()
        if start > 0 and stripped[start - 1] == "!":
            continue
        rng = parse_range_a1(m.group(1))
        if rng:
            out.append(rng)
    return out


# --------------------------------------------------------------------------
# Model
# --------------------------------------------------------------------------
class SheetModel(QAbstractTableModel):
    def __init__(self, ctx, parent=None):
        super().__init__(parent)
        self.ctx = ctx
        self.sheet = ctx.book.active
        self.n_rows = MIN_ROWS
        self.n_cols = MIN_COLS
        self._fit_used()

    # sizing ---------------------------------------------------------------
    def _fit_used(self):
        rows, cols = self.sheet.used_bounds()
        self.n_rows = min(MAX_ROWS, max(MIN_ROWS, rows + ROW_GROW))
        self.n_cols = min(MAX_COLS, max(MIN_COLS, cols + COL_GROW))

    def set_sheet(self, sheet):
        self.beginResetModel()
        self.sheet = sheet
        self._fit_used()
        self.endResetModel()

    def reset(self):
        self.set_sheet(self.sheet)

    def ensure_rows(self, n: int):
        n = min(MAX_ROWS, n)
        if n > self.n_rows:
            self.beginInsertRows(QModelIndex(), self.n_rows, n - 1)
            self.n_rows = n
            self.endInsertRows()

    def ensure_cols(self, n: int):
        n = min(MAX_COLS, n)
        if n > self.n_cols:
            self.beginInsertColumns(QModelIndex(), self.n_cols, n - 1)
            self.n_cols = n
            self.endInsertColumns()

    def ensure_cell(self, r: int, c: int):
        if r >= self.n_rows - 1:
            self.ensure_rows(r + ROW_GROW)
        if c >= self.n_cols - 1:
            self.ensure_cols(c + COL_GROW)

    # Qt API ----------------------------------------------------------------
    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.n_rows

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else self.n_cols

    def display(self, r: int, c: int) -> str:
        raw = self.sheet.cells.get((r, c))
        if raw is None:
            return ""
        if self.ctx.show_formulas and raw.startswith("="):
            return raw
        try:
            return self.ctx.engine.display_text(self.sheet, r, c)
        except Exception:  # engine bug must never break painting
            return "#ERROR!"

    def value(self, r: int, c: int):
        raw = self.sheet.cells.get((r, c))
        if raw is None:
            return None
        try:
            return self.ctx.engine.value(self.sheet, r, c)
        except Exception:
            return None

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        if role == Qt.DisplayRole:
            return self.display(r, c)
        if role == Qt.EditRole:
            return self.sheet.get_raw(r, c)
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if role != Qt.EditRole or not index.isValid():
            return False
        self.ctx.user_set_cell(index.row(), index.column(), "" if value is None else str(value))
        return True

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return col_to_name(section) if orientation == Qt.Horizontal else str(section + 1)
        return None


# --------------------------------------------------------------------------
# Editor with function-name completion
# --------------------------------------------------------------------------
_TOKEN_RE = re.compile(r"([A-Za-z][A-Za-z0-9_.]*)$")


class FormulaLineEdit(QLineEdit):
    """QLineEdit that pops up function names while typing a formula."""

    def __init__(self, names=(), parent=None):
        super().__init__(parent)
        self._names_model = QStringListModel(sorted(names), self)
        self.completer_ = QCompleter(self._names_model, self)
        self.completer_.setCaseSensitivity(Qt.CaseInsensitive)
        self.completer_.setCompletionMode(QCompleter.PopupCompletion)
        self.completer_.setWidget(self)
        self.completer_.activated.connect(self._insert_completion)
        self.textEdited.connect(self._maybe_complete)

    def set_function_names(self, names):
        self._names_model.setStringList(sorted(names))

    def _token(self):
        text = self.text()[: self.cursorPosition()]
        if not text.startswith("="):
            return None
        m = _TOKEN_RE.search(text)
        if not m:
            return None
        before = text[: m.start()]
        if before and (before[-1].isalnum() or before[-1] in "$\"!"):
            return None
        return m.group(1)

    def _maybe_complete(self, _text):
        tok = self._token()
        if not tok or (len(tok) >= 2 and re.fullmatch(r"[A-Za-z]{1,3}\d+", tok)):
            self.completer_.popup().hide()
            return
        self.completer_.setCompletionPrefix(tok)
        if self.completer_.completionCount() == 0:
            self.completer_.popup().hide()
            return
        rect = self.cursorRect()
        rect.setWidth(220)
        self.completer_.complete(rect)

    def _insert_completion(self, name):
        tok = self._token() or ""
        pos = self.cursorPosition()
        text = self.text()
        new = text[: pos - len(tok)] + name + "(" + text[pos:]
        self.setText(new)
        self.setCursorPosition(pos - len(tok) + len(name) + 1)
        self.textEdited.emit(self.text())
        self.completer_.popup().hide()

    def popup_visible(self) -> bool:
        return self.completer_.popup().isVisible()

    def can_point(self) -> bool:
        """True when a click on the grid should insert a reference."""
        t = self.text()
        if not t.startswith("="):
            return False
        if self.hasSelectedText():
            return False
        pos = self.cursorPosition()
        return pos > 0 and t[pos - 1] in _POINT_CHARS

    def insert_ref(self, ref: str, replace_len: int = 0):
        pos = self.cursorPosition()
        t = self.text()
        start = pos - replace_len
        self.setText(t[:start] + ref + t[pos:])
        self.setCursorPosition(start + len(ref))
        self.textEdited.emit(self.text())


class CellEditor(FormulaLineEdit):
    def __init__(self, names, parent=None):
        super().__init__(names, parent)
        self.mode = "edit"  # "enter": arrows commit and move; "edit": arrows move the caret
        self.setFrame(False)
        self.setStyleSheet("QLineEdit{border:2px solid #217346; padding:0 1px;"
                           "background:white; selection-background-color:#9fd5b7;"
                           "selection-color:black;}")


# --------------------------------------------------------------------------
# Delegate
# --------------------------------------------------------------------------
class CellDelegate(QStyledItemDelegate):
    def __init__(self, view):
        super().__init__(view)
        self.view = view  # main SpreadsheetView
        self.editor: Optional[CellEditor] = None
        self._font_cache = {}

    # painting ----------------------------------------------------------------
    def _font(self, st: Optional[dict]) -> QFont:
        z = self.view.ctx.zoom
        if not st:
            key = (None, z)
        else:
            key = (st.get("font_name"), st.get("font_size"), bool(st.get("bold")),
                   bool(st.get("italic")), bool(st.get("underline")), z)
        f = self._font_cache.get(key)
        if f is None:
            f = QFont(self.view.base_font)
            if st:
                if st.get("font_name"):
                    f.setFamily(st["font_name"])
                if st.get("font_size"):
                    try:
                        f.setPointSizeF(max(1.0, float(st["font_size"])) * 0.75 * 96 / 72)
                    except (TypeError, ValueError):
                        pass
                f.setBold(bool(st.get("bold")))
                f.setItalic(bool(st.get("italic")))
                f.setUnderline(bool(st.get("underline")))
            f.setPointSizeF(max(1.0, f.pointSizeF() * z))
            self._font_cache[key] = f
        return f

    def paint(self, p: QPainter, opt, index):
        model = index.model()
        r, c = index.row(), index.column()
        sheet = model.sheet
        st = sheet.styles.get((r, c))
        rect = opt.rect
        if st and st.get("fill"):
            p.fillRect(rect, QColor(st["fill"]))
        raw = sheet.cells.get((r, c))
        if raw is not None:
            text = model.display(r, c)
            if text:
                self._paint_text(p, rect, r, c, raw, text, st, model)
        if opt.state & QStyle.State_Selected:
            cur = self.view.currentIndex()
            if not (cur.row() == r and cur.column() == c):
                p.fillRect(rect, SEL_TINT)

    def _paint_text(self, p, rect, r, c, raw, text, st, model):
        st = st or {}
        font = self._font(st if st else None)
        fm = QFontMetrics(font)
        align = st.get("align")
        val = None
        if not align or True:
            val = model.value(r, c)
        is_num = isinstance(val, (int, float)) and not isinstance(val, bool)
        if self.view.ctx.show_formulas and raw.startswith("="):
            is_num = False
        if align == "left":
            h = Qt.AlignLeft
        elif align == "center":
            h = Qt.AlignHCenter
        elif align == "right":
            h = Qt.AlignRight
        elif is_num:
            h = Qt.AlignRight
        elif isinstance(val, (bool, XLError)):
            h = Qt.AlignHCenter
        else:
            h = Qt.AlignLeft
        wrap = bool(st.get("wrap"))
        color = QColor(st["color"]) if st.get("color") else QColor("#000000")
        pad = max(2, int(3 * self.view.ctx.zoom))
        tr = rect.adjusted(pad, 0, -pad, 0)
        p.save()
        p.setFont(font)
        p.setPen(color)
        if wrap:
            p.setClipRect(rect)
            p.drawText(tr, h | Qt.AlignVCenter | Qt.TextWordWrap, text)
        else:
            width = fm.horizontalAdvance(text)
            if width > tr.width():
                if is_num and h == Qt.AlignRight and not self.view.ctx.show_formulas:
                    n = max(1, tr.width() // max(1, fm.horizontalAdvance("#")))
                    text = "#" * n
                    p.setClipRect(rect)
                elif h == Qt.AlignLeft:
                    # overflow into empty neighbours on the right
                    extra = 0
                    cells = model.sheet.cells
                    cc = c + 1
                    need = width - tr.width()
                    view = self.view
                    limit = view.viewport().width() + 400
                    while extra < need and cc < model.n_cols and (r, cc) not in cells \
                            and rect.right() + extra < limit:
                        extra += view.columnWidth(cc)
                        cc += 1
                    clip = QRect(rect.left(), rect.top(), rect.width() + extra, rect.height())
                    p.setClipRect(clip)
                    tr = QRect(tr.left(), tr.top(), tr.width() + extra + pad, tr.height())
                else:
                    p.setClipRect(rect)
            else:
                p.setClipRect(rect)
            p.drawText(tr, h | Qt.AlignVCenter | Qt.TextSingleLine, text)
        p.restore()

    def sizeHint(self, opt, index):
        return QSize(DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT)

    # editing -----------------------------------------------------------------
    def createEditor(self, parent, opt, index):
        ed = CellEditor(self.view.ctx.function_names(), parent)
        st = index.model().sheet.styles.get((index.row(), index.column()))
        ed.setFont(self._font(st))
        self.editor = ed
        ed.textEdited.connect(self.view.ctx.on_editor_text)
        ed.destroyed.connect(self._editor_gone)
        return ed

    def _editor_gone(self, obj=None):
        if self.editor is obj or obj is None:
            self.editor = None
        try:
            self.view.ctx.set_mode("Ready")
            self.view.viewport().update()
        except RuntimeError:  # view already deleted
            pass

    def setEditorData(self, editor, index):
        start = self.view.pending_text
        if start is not None:
            editor.setText(start)
            editor.mode = "enter"
            self.view.ctx.set_mode("Enter")
            editor._maybe_complete(start)
        else:
            editor.setText(index.model().sheet.get_raw(index.row(), index.column()))
            editor.mode = "edit"
            self.view.ctx.set_mode("Edit")
        self.view.ctx.on_editor_text(editor.text())

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text(), Qt.EditRole)

    def updateEditorGeometry(self, editor, opt, index):
        r = opt.rect
        fm = QFontMetrics(editor.font())
        w = max(r.width() + 2, fm.horizontalAdvance(editor.text()) + 16)
        editor.setGeometry(r.left() - 1, r.top() - 1, w, r.height() + 2)
        editor.textChanged.connect(lambda t, e=editor, rr=QRect(r): self._grow(e, rr))

    def _grow(self, editor, r):
        fm = QFontMetrics(editor.font())
        w = max(r.width() + 2, fm.horizontalAdvance(editor.text()) + 16)
        if w != editor.width():
            editor.resize(min(w, max(r.width() + 2, editor.parent().width() - r.left())),
                          editor.height())

    def eventFilter(self, editor, event):
        et = event.type()
        if et == QEvent.KeyPress and isinstance(editor, CellEditor):
            key = event.key()
            mods = event.modifiers()
            if editor.popup_visible() and key in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Tab):
                return False
            if key in (Qt.Key_Return, Qt.Key_Enter):
                if mods & Qt.AltModifier:
                    return False
                fill = bool(mods & Qt.ControlModifier)
                self.view.finish_edit(editor, commit=True,
                                      move=(0, 0) if fill else ((-1, 0) if mods & Qt.ShiftModifier
                                                                else (1, 0)),
                                      fill_selection=fill)
                return True
            if key in (Qt.Key_Tab, Qt.Key_Backtab):
                back = key == Qt.Key_Backtab or bool(mods & Qt.ShiftModifier)
                self.view.finish_edit(editor, commit=True, move=(0, -1) if back else (0, 1))
                return True
            if key == Qt.Key_Escape:
                self.view.finish_edit(editor, commit=False)
                return True
            if key == Qt.Key_F2:
                editor.mode = "edit" if editor.mode == "enter" else "enter"
                self.view.ctx.set_mode("Edit" if editor.mode == "edit" else "Enter")
                return True
            if editor.mode == "enter" and key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left,
                                                  Qt.Key_Right):
                ps = self.view.point_state
                cont = bool(ps and ps.get("editor") is editor and ps.get("text") == editor.text())
                if (editor.can_point() or cont) and not editor.popup_visible():
                    self.view.keyboard_point(editor, key)
                    return True
                move = {Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0), Qt.Key_Left: (0, -1),
                        Qt.Key_Right: (0, 1)}[key]
                self.view.finish_edit(editor, commit=True, move=move)
                return True
            return False
        if et == QEvent.FocusOut and isinstance(editor, CellEditor):
            if self.view.point_click_pending(editor):
                return True
            fw = QApplication.focusWidget()
            if fw is not None and (fw is editor.completer_.popup() or
                                   editor.isAncestorOf(fw)):
                return True
        return super().eventFilter(editor, event)


# --------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------
class ExcelHeader(QHeaderView):
    def __init__(self, orientation, view):
        super().__init__(orientation, view)
        self.grid = view
        self.setSectionsClickable(True)
        self.setHighlightSections(False)
        self.setDefaultAlignment(Qt.AlignCenter)

    def paintSection(self, p: QPainter, rect: QRect, idx: int):
        if not rect.isValid():
            return
        sel, full = self.grid.header_state(self.orientation(), idx)
        p.save()
        p.fillRect(rect, HEADER_FULL_BG if full else (HEADER_SEL_BG if sel else HEADER_BG))
        p.setPen(HEADER_LINE)
        p.drawLine(rect.topRight(), rect.bottomRight())
        p.drawLine(rect.bottomLeft(), rect.bottomRight())
        if sel or full:
            p.setPen(QPen(GREEN, 2))
            if self.orientation() == Qt.Horizontal:
                p.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
            else:
                p.drawLine(rect.right(), rect.top(), rect.right(), rect.bottom())
        f = QFont(self.grid.base_font)
        f.setPointSizeF(max(1.0, f.pointSizeF() * self.grid.ctx.zoom))
        p.setFont(f)
        p.setPen(DARK_GREEN if (sel or full) else QColor("#444444"))
        label = col_to_name(idx) if self.orientation() == Qt.Horizontal else str(idx + 1)
        p.drawText(rect, Qt.AlignCenter, label)
        p.restore()


# --------------------------------------------------------------------------
# Views
# --------------------------------------------------------------------------
class GridBase(QTableView):
    """Common painting (selection border, reference highlights) for the main
    view and the frozen-pane overlays."""

    def __init__(self, ctx, parent=None):
        super().__init__(parent)
        self.ctx = ctx
        self.main = self
        self.setHorizontalHeader(ExcelHeader(Qt.Horizontal, self))
        self.setVerticalHeader(ExcelHeader(Qt.Vertical, self))
        self.setFrameShape(QTableView.NoFrame)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.setWordWrap(False)
        self.setTabKeyNavigation(False)
        self.setAutoScroll(True)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerItem)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollPerItem)
        self.setMouseTracking(True)
        pal = self.palette()
        pal.setColor(QPalette.Highlight, SEL_TINT)
        pal.setColor(QPalette.HighlightedText, QColor("#000000"))
        pal.setColor(QPalette.Base, QColor("#FFFFFF"))
        self.setPalette(pal)
        self.setStyleSheet("QTableView{gridline-color:#E1E1E1; background:white;"
                           "selection-background-color:transparent;}"
                           "QTableView QTableCornerButton::section{background:#F3F3F3;"
                           "border:none; border-right:1px solid #D4D4D4;"
                           "border-bottom:1px solid #D4D4D4;}")

    @property
    def base_font(self) -> QFont:
        return self.main._base_font

    # header helpers --------------------------------------------------------
    def header_state(self, orientation, idx):
        main = self.main
        sel = full = False
        for (r1, c1, r2, c2) in main.sel_ranges:
            if orientation == Qt.Horizontal:
                if c1 <= idx <= c2:
                    sel = True
                    if r1 == 0 and r2 >= main.model().n_rows - 1:
                        full = True
            else:
                if r1 <= idx <= r2:
                    sel = True
                    if c1 == 0 and c2 >= main.model().n_cols - 1:
                        full = True
        return sel, full

    def range_rect(self, r1, c1, r2, c2) -> QRect:
        x1 = self.columnViewportPosition(c1)
        x2 = self.columnViewportPosition(c2) + self.columnWidth(c2)
        y1 = self.rowViewportPosition(r1)
        y2 = self.rowViewportPosition(r2) + self.rowHeight(r2)
        return QRect(x1, y1, x2 - x1, y2 - y1)

    def paintEvent(self, e):
        super().paintEvent(e)
        main = self.main
        p = QPainter(self.viewport())
        vw = self.viewport().rect().adjusted(-4, -4, 4, 4)
        # references of the formula being edited
        target = main.ctx.active_formula_editor()
        if target is not None:
            for i, (r1, c1, r2, c2) in enumerate(formula_refs(target.text())):
                rr = self.range_rect(r1, c1, min(r2, main.model().n_rows - 1),
                                     min(c2, main.model().n_cols - 1)).intersected(vw)
                if rr.isEmpty():
                    continue
                col = QColor(REF_COLORS[i % len(REF_COLORS)])
                fill = QColor(col)
                fill.setAlpha(28)
                p.fillRect(rr, fill)
                p.setPen(QPen(col, 2))
                p.drawRect(rr.adjusted(1, 1, -1, -1))
        # copy marquee
        clip = main.ctx.clip_marquee()
        if clip is not None:
            r1, c1, r2, c2 = clip
            rr = self.range_rect(r1, c1, r2, c2).intersected(vw)
            if not rr.isEmpty():
                pen = QPen(GREEN, 2, Qt.DashLine)
                p.setPen(pen)
                p.drawRect(rr.adjusted(0, 0, -1, -1))
        # selection border + fill handle
        cur = main.currentIndex()
        if main.sel_ranges and cur.isValid():
            r1, c1, r2, c2 = main.active_range()
            rr = self.range_rect(r1, c1, r2, c2)
            if rr.intersects(vw):
                p.setPen(QPen(GREEN, 2))
                p.setBrush(Qt.NoBrush)
                p.drawRect(rr.adjusted(0, 0, -1, -1))
                if len(main.sel_ranges) == 1:
                    h = main.handle_rect(rr)
                    p.fillRect(h.adjusted(-1, -1, 1, 1), QColor("#ffffff"))
                    p.fillRect(h, GREEN)
            # active cell outline inside a multi-cell selection
        imgs = getattr(main.ctx, "images", None)
        if imgs is not None:
            imgs.paint(self, p)
        if main.fill_drag is not None and main.fill_drag.get("dst"):
            r1, c1, r2, c2 = main.fill_drag["dst"]
            rr = self.range_rect(r1, c1, r2, c2).intersected(vw)
            p.setPen(QPen(QColor("#666666"), 1, Qt.DashLine))
            p.drawRect(rr.adjusted(0, 0, -1, -1))
        p.end()

    def wheelEvent(self, e):
        if self.main is not self:
            self.main.wheelEvent(e)
            return
        if e.modifiers() & Qt.ControlModifier:
            step = 0.1 if e.angleDelta().y() > 0 else -0.1
            self.ctx.set_zoom(self.ctx.zoom + step)
            return
        super().wheelEvent(e)


class FrozenView(GridBase):
    """Overlay that shows the frozen top rows and/or first columns."""

    def __init__(self, main, kind):
        super().__init__(main.ctx, main)
        self.main = main
        self.kind = kind  # "row" | "col" | "corner"
        self.setModel(main.model())
        self.setSelectionModel(main.selectionModel())
        self.setItemDelegate(main.delegate)
        self.setFocusPolicy(Qt.NoFocus)
        self.setFocusProxy(main)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setShowGrid(main.showGrid())
        self.horizontalHeader().sectionResized.connect(self._resized_h)
        self.verticalHeader().sectionResized.connect(self._resized_v)
        self.horizontalHeader().sectionPressed.connect(main.horizontalHeader().sectionPressed)
        self.verticalHeader().sectionPressed.connect(main.verticalHeader().sectionPressed)
        self.horizontalHeader().sectionEntered.connect(main.horizontalHeader().sectionEntered)
        self.verticalHeader().sectionEntered.connect(main.verticalHeader().sectionEntered)
        self.viewport().setCursor(main.viewport().cursor())

    def scrollTo(self, index, hint=QAbstractItemView.EnsureVisible):
        # the overlay never scrolls by itself; the main view keeps it in sync
        pass

    def _resized_h(self, i, old, new):
        if self.main.columnWidth(i) != new:
            self.main.setColumnWidth(i, new)

    def _resized_v(self, i, old, new):
        if self.main.rowHeight(i) != new:
            self.main.setRowHeight(i, new)

    def mouseDoubleClickEvent(self, e):
        idx = self.indexAt(e.position().toPoint())
        if idx.isValid():
            self.main.setCurrentIndex(idx)
            self.main.edit(idx)

    def mousePressEvent(self, e):
        if self.main.handle_point_click(self, e):
            return
        super().mousePressEvent(e)
        self.main.setFocus()

    def mouseMoveEvent(self, e):
        if self.main.handle_point_move(self, e):
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        ps = self.main.point_state
        if ps and ps.get("drag"):
            ps["drag"] = False
            ps["editor"].setFocus()
            return
        super().mouseReleaseEvent(e)

    def contextMenuEvent(self, e):
        self.main.ctx.show_cell_menu(e.globalPos())

    def paintEvent(self, e):
        super().paintEvent(e)
        p = QPainter(self.viewport())
        p.setPen(QPen(QColor("#9e9e9e"), 1))
        fr, fc = self.main.frozen
        if self.kind in ("row", "corner") and fr:
            y = self.rowViewportPosition(fr - 1) + self.rowHeight(fr - 1) - 1
            p.drawLine(0, y, self.viewport().width(), y)
        if self.kind in ("col", "corner") and fc:
            x = self.columnViewportPosition(fc - 1) + self.columnWidth(fc - 1) - 1
            p.drawLine(x, 0, x, self.viewport().height())
        p.end()


class SpreadsheetView(GridBase):
    """The main grid view."""

    def __init__(self, ctx, parent=None):
        super().__init__(ctx, parent)
        self._base_font = QFont(self.font())
        self.pending_text: Optional[str] = None
        self.sel_ranges: List[Tuple[int, int, int, int]] = [(0, 0, 0, 0)]
        self.fill_drag = None
        self.point_state = None
        self._point_pending = False
        self.frozen = (0, 0)
        self.overlays = {}
        self._applying_sizes = False
        self.sheet_model = SheetModel(ctx, self)
        self.setModel(self.sheet_model)
        self.delegate = CellDelegate(self)
        self.setItemDelegate(self.delegate)
        self.setCornerButtonEnabled(True)
        self.horizontalHeader().sectionResized.connect(self._col_resized)
        self.verticalHeader().sectionResized.connect(self._row_resized)
        self.horizontalHeader().sectionHandleDoubleClicked.connect(self.autofit_column)
        self.verticalHeader().sectionHandleDoubleClicked.connect(self.autofit_row)
        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(
            lambda pos: ctx.show_header_menu("col", self.horizontalHeader().mapToGlobal(pos)))
        self.verticalHeader().customContextMenuRequested.connect(
            lambda pos: ctx.show_header_menu("row", self.verticalHeader().mapToGlobal(pos)))
        self.verticalScrollBar().valueChanged.connect(self._v_scrolled)
        self.horizontalScrollBar().valueChanged.connect(self._h_scrolled)
        self._connect_selection()
        self.apply_sizes()
        self.setCurrentIndex(self.sheet_model.index(0, 0))

    # sheet / sizes ---------------------------------------------------------
    def _connect_selection(self):
        sm = self.selectionModel()
        sm.selectionChanged.connect(self._sel_changed)
        sm.currentChanged.connect(self._cur_changed)

    def set_sheet(self, sheet):
        self.sheet_model.set_sheet(sheet)
        self.apply_sizes()

    def apply_sizes(self):
        """Apply zoom, default sizes and the sheet's custom widths/heights."""
        z = self.ctx.zoom
        sheet = self.sheet_model.sheet
        self._applying_sizes = True
        try:
            f = QFont(self._base_font)
            f.setPointSizeF(max(1.0, f.pointSizeF() * z))
            hh, vh = self.horizontalHeader(), self.verticalHeader()
            hh.setMinimumSectionSize(0)
            vh.setMinimumSectionSize(0)
            hh.setDefaultSectionSize(max(8, round(DEFAULT_COL_WIDTH * z)))
            vh.setDefaultSectionSize(max(6, round(DEFAULT_ROW_HEIGHT * z)))
            hh.setFixedHeight(max(12, round(DEFAULT_ROW_HEIGHT * z)))
            # reset previously customised sections, then apply the sheet's
            for c in list(getattr(self, "_custom_cols", ())):
                if c < self.sheet_model.n_cols:
                    self.setColumnHidden(c, False)
                    hh.resizeSection(c, hh.defaultSectionSize())
            for r in list(getattr(self, "_custom_rows", ())):
                if r < self.sheet_model.n_rows:
                    self.setRowHidden(r, False)
                    vh.resizeSection(r, vh.defaultSectionSize())
            self._custom_cols = set()
            self._custom_rows = set()
            for c, w in sheet.col_widths.items():
                if c >= self.sheet_model.n_cols:
                    self.sheet_model.ensure_cols(c + COL_GROW)
                self._custom_cols.add(c)
                if w <= 0:
                    self.setColumnHidden(c, True)
                else:
                    self.setColumnHidden(c, False)
                    hh.resizeSection(c, max(2, round(w * z)))
            for r, h in sheet.row_heights.items():
                if r >= self.sheet_model.n_rows:
                    self.sheet_model.ensure_rows(r + ROW_GROW)
                self._custom_rows.add(r)
                if h <= 0:
                    self.setRowHidden(r, True)
                else:
                    self.setRowHidden(r, False)
                    vh.resizeSection(r, max(2, round(h * z)))
            self.delegate._font_cache.clear()
            self._update_row_header_width()
        finally:
            self._applying_sizes = False
        self.update_frozen()
        self.viewport().update()

    def _update_row_header_width(self):
        f = QFont(self._base_font)
        f.setPointSizeF(max(1.0, f.pointSizeF() * self.ctx.zoom))
        fm = QFontMetrics(f)
        last = self.rowAt(self.viewport().height() - 1)
        if last < 0:
            last = self.sheet_model.n_rows - 1
        digits = max(3, len(str(last + 1)))
        w = fm.horizontalAdvance("9" * digits) + max(8, round(10 * self.ctx.zoom))
        if self.verticalHeader().width() != w:
            self.verticalHeader().setFixedWidth(w)
            for ov in self.overlays.values():
                ov.verticalHeader().setFixedWidth(w)
            self.update_frozen()

    def _col_resized(self, c, old, new):
        if self._applying_sizes:
            return
        z = self.ctx.zoom
        sheet = self.sheet_model.sheet
        w = round(new / z)
        if w == DEFAULT_COL_WIDTH and c in sheet.col_widths:
            sheet.col_widths.pop(c, None)
        elif w != DEFAULT_COL_WIDTH:
            sheet.col_widths[c] = w
        self._custom_cols.add(c)
        for ov in self.overlays.values():
            if ov.columnWidth(c) != new:
                ov.setColumnWidth(c, new)
        self.ctx.mark_dirty()
        self.update_frozen()
        self.viewport().update()

    def _row_resized(self, r, old, new):
        if self._applying_sizes:
            return
        z = self.ctx.zoom
        sheet = self.sheet_model.sheet
        h = round(new / z)
        if h == DEFAULT_ROW_HEIGHT:
            sheet.row_heights.pop(r, None)
        else:
            sheet.row_heights[r] = h
        self._custom_rows.add(r)
        for ov in self.overlays.values():
            if ov.rowHeight(r) != new:
                ov.setRowHeight(r, new)
        self.ctx.mark_dirty()
        self.update_frozen()
        self.viewport().update()

    def autofit_column(self, c):
        sheet = self.sheet_model.sheet
        best = 0
        for (r, cc), raw in sheet.cells.items():
            if cc != c:
                continue
            fm = QFontMetrics(self.delegate._font(sheet.styles.get((r, cc))))
            best = max(best, fm.horizontalAdvance(self.sheet_model.display(r, cc)))
        if best == 0:
            best = round(DEFAULT_COL_WIDTH * self.ctx.zoom) - 10
        self.setColumnWidth(c, best + max(8, round(10 * self.ctx.zoom)))

    def autofit_row(self, r):
        self.setRowHeight(r, round(DEFAULT_ROW_HEIGHT * self.ctx.zoom))

    def setShowGrid(self, on):
        super().setShowGrid(on)
        for ov in getattr(self, "overlays", {}).values():
            ov.setShowGrid(on)

    # freeze panes ------------------------------------------------------------
    def set_frozen(self, rows: int, cols: int):
        self.frozen = (max(0, rows), max(0, cols))
        for ov in self.overlays.values():
            ov.hide()
            ov.deleteLater()
        self.overlays = {}
        fr, fc = self.frozen
        if fr:
            self.overlays["row"] = FrozenView(self, "row")
        if fc:
            self.overlays["col"] = FrozenView(self, "col")
        if fr and fc:
            self.overlays["corner"] = FrozenView(self, "corner")
        for ov in self.overlays.values():
            ov.horizontalHeader().setDefaultSectionSize(self.horizontalHeader().defaultSectionSize())
            ov.verticalHeader().setDefaultSectionSize(self.verticalHeader().defaultSectionSize())
            ov.horizontalHeader().setFixedHeight(self.horizontalHeader().height())
            ov.verticalHeader().setFixedWidth(self.verticalHeader().width())
            for c in range(min(self.sheet_model.n_cols, max(self._custom_cols | {0}) + 1)):
                if self.columnWidth(c) != ov.columnWidth(c):
                    ov.setColumnWidth(c, self.columnWidth(c))
                ov.setColumnHidden(c, self.isColumnHidden(c))
            for r in range(min(self.sheet_model.n_rows, max(self._custom_rows | {0}) + 1)):
                if self.rowHeight(r) != ov.rowHeight(r):
                    ov.setRowHeight(r, self.rowHeight(r))
                ov.setRowHidden(r, self.isRowHidden(r))
        if "row" in self.overlays:
            self.overlays["row"].horizontalScrollBar().setValue(self.horizontalScrollBar().value())
        if "col" in self.overlays:
            self.overlays["col"].verticalScrollBar().setValue(self.verticalScrollBar().value())
        self.update_frozen()
        for ov in self.overlays.values():
            ov.show()
        if "corner" in self.overlays:
            self.overlays["corner"].raise_()
        if fr or fc:
            self.verticalScrollBar().setValue(0)
            self.horizontalScrollBar().setValue(0)

    def update_frozen(self):
        if not getattr(self, "overlays", None):
            return
        fr, fc = self.frozen
        hh = self.horizontalHeader().height()
        vw = self.verticalHeader().width()
        rows_h = sum(self.rowHeight(r) for r in range(fr))
        cols_w = sum(self.columnWidth(c) for c in range(fc))
        full_w = self.viewport().width() + vw
        full_h = self.viewport().height() + hh
        geo = {"row": (0, 0, full_w, hh + rows_h),
               "col": (0, 0, vw + cols_w, full_h),
               "corner": (0, 0, vw + cols_w, hh + rows_h)}
        for k, ov in self.overlays.items():
            ov.setGeometry(*geo[k])
            ov.verticalHeader().setFixedWidth(vw)
            ov.horizontalHeader().setFixedHeight(hh)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.update_frozen()
        self._update_row_header_width()

    def scrollTo(self, index, hint=QAbstractItemView.EnsureVisible):
        super().scrollTo(index, hint)
        fr, fc = self.frozen
        if not index.isValid():
            return
        r, c = index.row(), index.column()
        if fr and r >= fr:
            top = self.verticalScrollBar().value()
            if r < top + fr:
                self.verticalScrollBar().setValue(max(0, r - fr))
        if fc and c >= fc:
            left = self.horizontalScrollBar().value()
            if c < left + fc:
                self.horizontalScrollBar().setValue(max(0, c - fc))

    def _v_scrolled(self, v):
        sb = self.verticalScrollBar()
        if v >= sb.maximum() - 5 and self.sheet_model.n_rows < MAX_ROWS:
            self.sheet_model.ensure_rows(self.sheet_model.n_rows + ROW_GROW)
        if "col" in self.overlays:
            self.overlays["col"].verticalScrollBar().setValue(v)
        self._update_row_header_width()
        self.viewport().update()

    def _h_scrolled(self, v):
        sb = self.horizontalScrollBar()
        if v >= sb.maximum() - 2 and self.sheet_model.n_cols < MAX_COLS:
            self.sheet_model.ensure_cols(self.sheet_model.n_cols + COL_GROW)
        if "row" in self.overlays:
            self.overlays["row"].horizontalScrollBar().setValue(v)
        self.viewport().update()

    def scrollContentsBy(self, dx, dy):
        super().scrollContentsBy(dx, dy)
        self.viewport().update()

    def refresh(self):
        self.viewport().update()
        self.horizontalHeader().viewport().update()
        self.verticalHeader().viewport().update()
        for ov in self.overlays.values():
            ov.viewport().update()
            ov.horizontalHeader().viewport().update()
            ov.verticalHeader().viewport().update()

    # selection -------------------------------------------------------------
    def _sel_changed(self, *_):
        ranges = []
        for rg in self.selectionModel().selection():
            ranges.append((rg.top(), rg.left(), rg.bottom(), rg.right()))
        cur = self.currentIndex()
        if not ranges and cur.isValid():
            ranges = [(cur.row(), cur.column(), cur.row(), cur.column())]
        self.sel_ranges = self._merge(ranges)
        self.refresh()
        self.ctx.on_selection_changed()

    @staticmethod
    def _merge(ranges):
        """QItemSelection may split a rectangle into several ranges; merge
        ranges that together form one rectangle."""
        if len(ranges) <= 1:
            return ranges
        r1 = min(r[0] for r in ranges)
        c1 = min(r[1] for r in ranges)
        r2 = max(r[2] for r in ranges)
        c2 = max(r[3] for r in ranges)
        area = sum((r[2] - r[0] + 1) * (r[3] - r[1] + 1) for r in ranges)
        if area == (r2 - r1 + 1) * (c2 - c1 + 1):
            # could still overlap; accept when no two ranges overlap
            ok = True
            for i, a in enumerate(ranges):
                for b in ranges[i + 1:]:
                    if a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]:
                        ok = False
            if ok:
                return [(r1, c1, r2, c2)]
        return ranges

    def _cur_changed(self, cur, prev):
        if not self.sel_ranges or not self.selectionModel().hasSelection():
            if cur.isValid():
                self.sel_ranges = [(cur.row(), cur.column(), cur.row(), cur.column())]
        if cur.isValid():
            self.sheet_model.ensure_cell(cur.row(), cur.column())
        self.refresh()
        self.ctx.on_current_changed()

    def active_range(self) -> Tuple[int, int, int, int]:
        cur = self.currentIndex()
        r, c = (cur.row(), cur.column()) if cur.isValid() else (0, 0)
        for rg in reversed(self.sel_ranges):
            if rg[0] <= r <= rg[2] and rg[1] <= c <= rg[3]:
                return rg
        if self.sel_ranges:
            return self.sel_ranges[-1]
        return (r, c, r, c)

    def current_cell(self) -> Tuple[int, int]:
        cur = self.currentIndex()
        return (cur.row(), cur.column()) if cur.isValid() else (0, 0)

    def select_range(self, r1, c1, r2, c2, cur=None):
        r1, r2 = sorted((max(0, r1), max(0, r2)))
        c1, c2 = sorted((max(0, c1), max(0, c2)))
        self.sheet_model.ensure_cell(r2, c2)
        r2 = min(r2, self.sheet_model.n_rows - 1)
        c2 = min(c2, self.sheet_model.n_cols - 1)
        m = self.sheet_model
        cr, cc = cur if cur else (r1, c1)
        sm = self.selectionModel()
        sm.setCurrentIndex(m.index(cr, cc), QItemSelectionModel.NoUpdate)
        sel = QItemSelection(m.index(r1, c1), m.index(r2, c2))
        sm.select(sel, QItemSelectionModel.ClearAndSelect)
        self.scrollTo(m.index(cr, cc))

    def goto(self, r, c, extend=False):
        r = max(0, min(MAX_ROWS - 1, r))
        c = max(0, min(MAX_COLS - 1, c))
        self.sheet_model.ensure_cell(r, c)
        m = self.sheet_model
        idx = m.index(r, c)
        if extend:
            anchor = self._anchor()
            sm = self.selectionModel()
            sel = QItemSelection(m.index(min(anchor[0], r), min(anchor[1], c)),
                                 m.index(max(anchor[0], r), max(anchor[1], c)))
            sm.select(sel, QItemSelectionModel.ClearAndSelect)
            self._extend_end = (r, c)
            self.scrollTo(idx)
            self.sel_ranges = [(min(anchor[0], r), min(anchor[1], c),
                                max(anchor[0], r), max(anchor[1], c))]
            self.refresh()
            self.ctx.on_selection_changed()
        else:
            self._extend_end = None
            self.selectionModel().setCurrentIndex(idx, QItemSelectionModel.ClearAndSelect)
            self.scrollTo(idx)

    def _anchor(self):
        return self.current_cell()

    def _extend_pos(self):
        end = getattr(self, "_extend_end", None)
        if end and self.selectionModel().hasSelection():
            r1, c1, r2, c2 = self.active_range()
            if r1 <= end[0] <= r2 and c1 <= end[1] <= c2:
                return end
        return self.current_cell()

    # data edge (Ctrl+arrow) ------------------------------------------------
    def data_edge(self, r, c, dr, dc):
        cells = self.sheet_model.sheet.cells
        maxr, maxc = MAX_ROWS - 1, MAX_COLS - 1

        def filled(rr, cc):
            return (rr, cc) in cells

        def inside(rr, cc):
            return 0 <= rr <= maxr and 0 <= cc <= maxc

        nr, nc = r + dr, c + dc
        if not inside(nr, nc):
            return r, c
        if filled(r, c) and filled(nr, nc):
            while inside(nr + dr, nc + dc) and filled(nr + dr, nc + dc):
                nr, nc = nr + dr, nc + dc
            return nr, nc
        # jump to next filled cell along the line
        if dr:
            col = sorted(k[0] for k in cells if k[1] == c and (k[0] - r) * dr > 0)
            if col:
                return (col[0] if dr > 0 else col[-1]), c
            return (maxr if dr > 0 else 0), c
        row = sorted(k[1] for k in cells if k[0] == r and (k[1] - c) * dc > 0)
        if row:
            return r, (row[0] if dc > 0 else row[-1])
        return r, (maxc if dc > 0 else 0)

    # keyboard ------------------------------------------------------------------
    def keyPressEvent(self, e):
        ed = self.delegate.editor
        if ed is not None:  # keys that reached the view while an editor is open
            if not getattr(self, "_forwarding", False):
                self._forwarding = True
                try:
                    QApplication.sendEvent(ed, e)
                finally:
                    self._forwarding = False
            return
        key = e.key()
        mods = e.modifiers()
        ctrl = bool(mods & Qt.ControlModifier)
        shift = bool(mods & Qt.ShiftModifier)
        alt = bool(mods & Qt.AltModifier)
        r, c = self.current_cell()
        imgs = self.ctx.images
        if imgs.selected:
            if key in (Qt.Key_Delete, Qt.Key_Backspace):
                imgs.delete_selected()
                return
            if key == Qt.Key_Escape:
                imgs.select(None)
                return
            if not ctrl:
                imgs.select(None)
        arrows = {Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0), Qt.Key_Left: (0, -1),
                  Qt.Key_Right: (0, 1)}
        if key in arrows and not alt:
            dr, dc = arrows[key]
            if shift:
                er, ec = self._extend_pos()
                if ctrl:
                    nr, nc = self.data_edge(er, ec, dr, dc)
                else:
                    nr, nc = er + dr, ec + dc
                self.goto(nr, nc, extend=True)
            else:
                if ctrl:
                    nr, nc = self.data_edge(r, c, dr, dc)
                else:
                    nr, nc = r + dr, c + dc
                self.goto(nr, nc)
            return
        if key in (Qt.Key_Return, Qt.Key_Enter):
            self._move_in_selection(-1 if shift else 1, 0)
            return
        if key in (Qt.Key_Tab, Qt.Key_Backtab):
            self._move_in_selection(0, -1 if (shift or key == Qt.Key_Backtab) else 1)
            return
        if key == Qt.Key_Home:
            if ctrl:
                self.goto(0, 0, extend=shift)
            else:
                self.goto(r, 0, extend=shift)
            return
        if key == Qt.Key_End and ctrl:
            rows, cols = self.sheet_model.sheet.used_bounds()
            self.goto(max(0, rows - 1), max(0, cols - 1), extend=shift)
            return
        if key in (Qt.Key_PageDown, Qt.Key_PageUp) and not ctrl:
            rows_vis = max(1, self.viewport().height() // max(1, self.rowHeight(r) or 1))
            if alt:
                cols_vis = max(1, self.viewport().width() // max(1, self.columnWidth(c) or 1))
                self.goto(r, c + (cols_vis if key == Qt.Key_PageDown else -cols_vis), extend=shift)
            else:
                self.goto(r + (rows_vis if key == Qt.Key_PageDown else -rows_vis), c,
                          extend=shift)
            return
        if key in (Qt.Key_Delete,) and not ctrl:
            self.ctx.clear_contents()
            return
        if key == Qt.Key_Backspace and not ctrl:
            self.ctx.clear_contents(selection=False)
            self.start_edit("")
            return
        if key == Qt.Key_F2:
            self.start_edit(None)
            return
        if key == Qt.Key_Escape:
            self.ctx.cancel_marquee()
            return
        if key == Qt.Key_Space and ctrl and not shift:
            r1, c1, r2, c2 = self.active_range()
            self.select_range(0, c1, self.sheet_model.n_rows - 1, c2, cur=(r, c))
            return
        if key == Qt.Key_Space and shift and not ctrl:
            r1, c1, r2, c2 = self.active_range()
            self.select_range(r1, 0, r2, self.sheet_model.n_cols - 1, cur=(r, c))
            return
        if ctrl and key == Qt.Key_A:
            self.select_range(0, 0, self.sheet_model.n_rows - 1, self.sheet_model.n_cols - 1,
                              cur=(r, c))
            return
        text = e.text()
        if text and text.isprintable() and not ctrl and not alt:
            self.start_edit(text)
            return
        super().keyPressEvent(e)

    def _move_in_selection(self, dr, dc):
        r, c = self.current_cell()
        r1, c1, r2, c2 = self.active_range()
        if (r1, c1) != (r2, c2) and len(self.sel_ranges) == 1:
            nr, nc = r + dr, c + dc
            if nc > c2:
                nc, nr = c1, r + 1
            if nc < c1:
                nc, nr = c2, r - 1
            if nr > r2:
                nr = r1
                if dc == 0:
                    nc = c + 1 if c < c2 else c1
            if nr < r1:
                nr = r2
                if dc == 0:
                    nc = c - 1 if c > c1 else c2
            self.selectionModel().setCurrentIndex(self.sheet_model.index(nr, nc),
                                                  QItemSelectionModel.NoUpdate)
            self.scrollTo(self.sheet_model.index(nr, nc))
            self.refresh()
            return
        self.goto(r + dr, c + dc)

    def move_after_edit(self, dr, dc):
        if dr or dc:
            self._move_in_selection(dr, dc)

    def start_edit(self, text: Optional[str]):
        """text=None: edit existing content (F2); otherwise replace with text."""
        idx = self.currentIndex()
        if not idx.isValid():
            return
        self.pending_text = text
        try:
            self.edit(idx)
        finally:
            self.pending_text = None
        ed = self.delegate.editor
        if ed is not None:
            ed.setFocus()
            ed.deselect()
            ed.end(False)

    def finish_edit(self, editor, commit=True, move=(0, 0), fill_selection=False):
        text = editor.text()
        self.delegate.closeEditor.emit(editor, QAbstractItemDelegate.NoHint)
        if commit:
            r, c = self.current_cell()
            self.ctx.user_set_cell(r, c, text, fill_selection=fill_selection)
        self.setFocus()
        self.move_after_edit(*move)
        self.ctx.on_current_changed()

    def is_editing(self) -> bool:
        return self.delegate.editor is not None

    def commit_open_editor(self):
        ed = self.delegate.editor
        if ed is not None:
            self.finish_edit(ed, commit=True)

    def closeEditor(self, editor, hint):
        if self.delegate.editor is editor:
            self.delegate.editor = None
        super().closeEditor(editor, hint)
        self.ctx.set_mode("Ready")
        self.refresh()

    # mouse: fill handle & point mode ---------------------------------------
    def handle_rect(self, rr: QRect) -> QRect:
        s = max(5, round(6 * self.ctx.zoom))
        return QRect(rr.right() - s // 2, rr.bottom() - s // 2, s, s)

    def _on_handle(self, pos) -> bool:
        if len(self.sel_ranges) != 1 or self.is_editing():
            return False
        r1, c1, r2, c2 = self.active_range()
        h = self.handle_rect(self.range_rect(r1, c1, r2, c2)).adjusted(-3, -3, 3, 3)
        return h.contains(pos)

    def point_click_pending(self, editor) -> bool:
        if not editor.can_point():
            return False
        if not (QApplication.mouseButtons() & Qt.LeftButton):
            return False
        fw = QApplication.focusWidget()
        if fw is self or fw in self.overlays.values():
            return True
        w = QApplication.widgetAt(self.cursor().pos())
        while w is not None:
            if w is self or w in self.overlays.values():
                return True
            w = w.parentWidget()
        return False

    def handle_point_click(self, view, e) -> bool:
        target = self.ctx.active_formula_editor()
        if target is None or e.button() != Qt.LeftButton:
            return False
        ps = self.point_state
        cont = bool(ps and ps.get("editor") is target and ps.get("text") == target.text())
        if not target.can_point() and not cont:
            return False
        idx = view.indexAt(e.position().toPoint())
        if not idx.isValid():
            return False
        ref = a1(idx.row(), idx.column())
        target.insert_ref(ref, ps["len"] if cont else 0)
        self.point_state = {"editor": target, "anchor": (idx.row(), idx.column()), "len": len(ref),
                            "text": target.text(), "drag": True}
        self.ctx.set_mode("Point")
        QTimer.singleShot(0, target.setFocus)
        target.setFocus()
        self.refresh()
        return True

    def handle_point_move(self, view, e) -> bool:
        ps = self.point_state
        if not ps or not ps.get("drag") or not (e.buttons() & Qt.LeftButton):
            return False
        idx = view.indexAt(e.position().toPoint())
        if not idx.isValid():
            return True
        ar, ac = ps["anchor"]
        r, c = idx.row(), idx.column()
        if (r, c) == (ar, ac):
            ref = a1(ar, ac)
        else:
            ref = f"{a1(min(ar, r), min(ac, c))}:{a1(max(ar, r), max(ac, c))}"
        ps["editor"].insert_ref(ref, ps["len"])
        ps["len"] = len(ref)
        ps["text"] = ps["editor"].text()
        self.refresh()
        return True

    def keyboard_point(self, editor, key):
        """Arrow keys in Enter mode after an operator insert a moving reference."""
        ps = self.point_state
        dr, dc = {Qt.Key_Up: (-1, 0), Qt.Key_Down: (1, 0), Qt.Key_Left: (0, -1),
                  Qt.Key_Right: (0, 1)}[key]
        if ps and ps.get("editor") is editor and ps.get("text") == editor.text():
            r, c = ps["pos"]
            r, c = max(0, r + dr), max(0, c + dc)
            ref = a1(r, c)
            editor.insert_ref(ref, ps["len"])
        else:
            r, c = self.current_cell()
            r, c = max(0, r + dr), max(0, c + dc)
            ref = a1(r, c)
            editor.insert_ref(ref)
        self.point_state = {"editor": editor, "pos": (r, c), "len": len(ref),
                            "text": editor.text()}
        self.ctx.set_mode("Point")
        self.refresh()

    def mousePressEvent(self, e):
        if self.handle_point_click(self, e):
            return
        if self.is_editing() and e.button() == Qt.LeftButton:
            self.commit_open_editor()
        self.ctx.formula_bar_focus_lost_commit()
        if self.ctx.images.mouse_press(self, e.position().toPoint(), e):
            return
        if e.button() == Qt.LeftButton and self._on_handle(e.position().toPoint()):
            self.fill_drag = {"src": self.active_range(), "dst": None}
            return
        if e.button() == Qt.RightButton:
            idx = self.indexAt(e.position().toPoint())
            if idx.isValid():
                r1, c1, r2, c2 = self.active_range()
                inside = any(a[0] <= idx.row() <= a[2] and a[1] <= idx.column() <= a[3]
                             for a in self.sel_ranges)
                if not inside:
                    self.goto(idx.row(), idx.column())
            return
        self._extend_end = None
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self.handle_point_move(self, e):
            return
        if self.ctx.images.mouse_move(self, e.position().toPoint(), e):
            return
        if self.fill_drag is not None:
            idx = self.indexAt(e.position().toPoint())
            if idx.isValid():
                r1, c1, r2, c2 = self.fill_drag["src"]
                r, c = idx.row(), idx.column()
                dst = None
                down = r - r2 if r > r2 else (r - r1 if r < r1 else 0)
                right = c - c2 if c > c2 else (c - c1 if c < c1 else 0)
                if abs(down) >= abs(right) and down:
                    dst = (r2 + 1, c1, r, c2) if down > 0 else (r, c1, r1 - 1, c2)
                elif right:
                    dst = (r1, c2 + 1, r2, c) if right > 0 else (r1, c, r2, c1 - 1)
                self.fill_drag["dst"] = dst
                self.refresh()
            return
        if not e.buttons():
            self.viewport().setCursor(Qt.CrossCursor if self._on_handle(e.position().toPoint())
                                      else Qt.ArrowCursor)
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self.point_state and self.point_state.get("drag"):
            ed = self.point_state["editor"]
            self.point_state["drag"] = False
            ed.setFocus()
            return
        if self.ctx.images.mouse_release(self, e.position().toPoint(), e):
            return
        if self.fill_drag is not None:
            fd, self.fill_drag = self.fill_drag, None
            if fd.get("dst"):
                self.ctx.fill_from_handle(fd["src"], fd["dst"])
            self.refresh()
            return
        super().mouseReleaseEvent(e)

    def mouseDoubleClickEvent(self, e):
        if self.ctx.images.mouse_double(self, e.position().toPoint()):
            return
        if self._on_handle(e.position().toPoint()):
            self.ctx.fill_handle_double_click()
            return
        idx = self.indexAt(e.position().toPoint())
        if idx.isValid():
            self.setCurrentIndex(idx)
            self.start_edit(None)
            return
        super().mouseDoubleClickEvent(e)

    def contextMenuEvent(self, e):
        img, _h = self.ctx.images.hit(self, e.pos())
        if img is not None:
            return
        self.ctx.show_cell_menu(e.globalPos())

    def focusInEvent(self, e):
        super().focusInEvent(e)

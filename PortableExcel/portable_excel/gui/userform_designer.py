"""Visual UserForm designer (design surface, Toolbox, Properties window)."""
from __future__ import annotations

import copy
import re
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import QPoint, QPointF, QRect, QRectF, QSize, Qt, Signal
from PySide6.QtGui import (QBrush, QColor, QFont, QIcon, QKeySequence, QLinearGradient,
                           QPainter, QPen, QPixmap, QPolygonF)
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QColorDialog,
                               QComboBox, QFontDialog, QFrame, QGridLayout, QHBoxLayout,
                               QHeaderView, QInputDialog, QLabel, QLineEdit, QMessageBox,
                               QPushButton, QScrollArea, QSizePolicy, QSplitter, QTableWidget,
                               QToolButton, QVBoxLayout, QWidget)

from .. import forms
from ..forms import points_to_px, px_to_points
from .userform_view import DEFAULT_EVENT, FORM_BG, FORM_BORDER_H, FORM_BORDER_W, \
    FormRenderer, make_font, px, to_color

GRID_PT = 6.0
TITLE_H = 24          # drawn caption bar of the design-time window (px)
FRAME_PAD = 4         # window border (px)
MARGIN = 18           # space around the form on the surface (px)
HANDLE = 6

_IDENT = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")


def snap(v: float) -> float:
    return round(round(v / GRID_PT) * GRID_PT, 2)


# --------------------------------------------------------------------------
# Toolbox icons (original artwork)
# --------------------------------------------------------------------------
def tool_pixmap(kind: str, size: int = 24) -> QPixmap:
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.scale(size / 24.0, size / 24.0)
    dark = QColor("#333333")
    blue = QColor("#2B579A")
    pen = QPen(dark, 1.2)
    p.setPen(pen)
    f = QFont("Segoe UI")
    if kind == "Pointer":
        p.setBrush(QColor("#FFFFFF"))
        p.drawPolygon(QPolygonF([QPointF(6, 3), QPointF(6, 19), QPointF(10, 15),
                                 QPointF(13, 21), QPointF(15, 20), QPointF(12, 14),
                                 QPointF(17, 14)]))
    elif kind == "Label":
        f.setPixelSize(15)
        f.setBold(True)
        p.setFont(f)
        p.setPen(blue)
        p.drawText(QRectF(0, 0, 24, 24), Qt.AlignCenter, "A")
    elif kind == "TextBox":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(2, 6, 20, 12))
        f.setPixelSize(9)
        p.setFont(f)
        p.drawText(QRectF(3, 6, 12, 12), Qt.AlignCenter, "ab")
        p.drawLine(QPointF(16, 8), QPointF(16, 16))
    elif kind == "ComboBox":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(2, 7, 20, 10))
        p.setBrush(QColor("#DDDDDD"))
        p.drawRect(QRectF(15, 7, 7, 10))
        p.setBrush(dark)
        p.drawPolygon(QPolygonF([QPointF(16.5, 10.5), QPointF(20.5, 10.5), QPointF(18.5, 13.5)]))
    elif kind == "ListBox":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(3, 3, 18, 18))
        p.setBrush(QColor("#9FC5E8"))
        p.setPen(Qt.NoPen)
        p.drawRect(QRectF(4, 9, 16, 4))
        p.setPen(pen)
        for y in (6, 11, 16):
            p.drawLine(QPointF(6, y), QPointF(17, y))
    elif kind == "CheckBox":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(4, 6, 12, 12))
        p.setPen(QPen(QColor("#217346"), 2))
        p.drawPolyline(QPolygonF([QPointF(6, 12), QPointF(9, 15), QPointF(14, 8)]))
    elif kind == "OptionButton":
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(QRectF(5, 5, 14, 14))
        p.setBrush(dark)
        p.drawEllipse(QRectF(9, 9, 6, 6))
    elif kind == "ToggleButton":
        p.setBrush(QColor("#C8C8C8"))
        p.drawRoundedRect(QRectF(2, 7, 20, 11), 2, 2)
        p.setBrush(QColor("#FFFFFF"))
        p.drawRoundedRect(QRectF(3, 5, 11, 11), 2, 2)
    elif kind == "Frame":
        p.setBrush(Qt.NoBrush)
        p.drawRect(QRectF(3, 7, 18, 14))
        p.fillRect(QRectF(5, 4, 9, 6), QColor("#F0F0F0"))
        f.setPixelSize(8)
        p.setFont(f)
        p.drawText(QRectF(5, 2, 10, 9), Qt.AlignCenter, "xy")
    elif kind == "CommandButton":
        p.setBrush(QColor("#E8E8E8"))
        p.drawRoundedRect(QRectF(2, 6, 20, 12), 2, 2)
        p.setPen(QPen(QColor("#FFFFFF"), 1))
        p.drawLine(QPointF(4, 8), QPointF(20, 8))
    elif kind == "TabStrip":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(2, 9, 20, 12))
        p.drawRect(QRectF(2, 4, 8, 5))
        p.drawRect(QRectF(10, 5, 8, 4))
    elif kind == "MultiPage":
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(4, 10, 18, 11))
        p.drawRect(QRectF(2, 8, 18, 11))
        p.drawRect(QRectF(2, 4, 7, 4))
        p.setPen(QPen(blue, 1))
        p.drawLine(QPointF(5, 13), QPointF(16, 13))
    elif kind == "ScrollBar":
        p.setBrush(QColor("#EEEEEE"))
        p.drawRect(QRectF(8, 2, 8, 20))
        p.setBrush(QColor("#AAAAAA"))
        p.drawRect(QRectF(9, 8, 6, 6))
        p.setBrush(dark)
        p.drawPolygon(QPolygonF([QPointF(12, 3.5), QPointF(9.5, 6), QPointF(14.5, 6)]))
        p.drawPolygon(QPolygonF([QPointF(12, 20.5), QPointF(9.5, 18), QPointF(14.5, 18)]))
    elif kind == "SpinButton":
        p.setBrush(QColor("#EEEEEE"))
        p.drawRect(QRectF(7, 3, 10, 9))
        p.drawRect(QRectF(7, 12, 10, 9))
        p.setBrush(dark)
        p.drawPolygon(QPolygonF([QPointF(12, 5), QPointF(9, 9.5), QPointF(15, 9.5)]))
        p.drawPolygon(QPolygonF([QPointF(12, 19), QPointF(9, 14.5), QPointF(15, 14.5)]))
    elif kind == "Image":
        p.setBrush(QColor("#DDEBF7"))
        p.drawRect(QRectF(2, 4, 20, 16))
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#70AD47"))
        p.drawPolygon(QPolygonF([QPointF(3, 19), QPointF(9, 11), QPointF(14, 19)]))
        p.setBrush(QColor("#548235"))
        p.drawPolygon(QPolygonF([QPointF(10, 19), QPointF(15, 13), QPointF(21, 19)]))
        p.setBrush(QColor("#FFC000"))
        p.drawEllipse(QRectF(15, 6, 4, 4))
    p.end()
    return pm


def tool_icon(kind: str) -> QIcon:
    ic = QIcon()
    for s in (16, 24, 32):
        ic.addPixmap(tool_pixmap(kind, s))
    return ic


# --------------------------------------------------------------------------
# Design surface
# --------------------------------------------------------------------------
class DesignClient(QWidget):
    """Client area of the design-time form: background with the dotted grid;
    all mouse input goes to the designer (control widgets are transparent)."""

    def __init__(self, designer):
        super().__init__()
        self.designer = designer
        self.setMouseTracking(True)
        self.setAutoFillBackground(True)
        self.setFocusPolicy(Qt.NoFocus)

    def paintEvent(self, e):
        super().paintEvent(e)
        p = QPainter(self)
        p.setPen(QColor(0, 0, 0, 150))
        step = px(GRID_PT)
        if step >= 3:
            for y in range(0, self.height(), step):
                for x in range(0, self.width(), step):
                    p.drawPoint(x, y)
        p.end()

    def mousePressEvent(self, e):
        self.designer.mouse_press(e.position().toPoint(), e)

    def mouseMoveEvent(self, e):
        self.designer.mouse_move(e.position().toPoint(), e)

    def mouseReleaseEvent(self, e):
        self.designer.mouse_release(e.position().toPoint(), e)

    def mouseDoubleClickEvent(self, e):
        self.designer.mouse_double(e.position().toPoint(), e)


class SelectionOverlay(QWidget):
    def __init__(self, client, designer):
        super().__init__(client)
        self.designer = designer
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)

    def paintEvent(self, e):
        d = self.designer
        p = QPainter(self)
        # control selection
        for name in d.selected:
            r = d.ctrl_rect(name)
            if r is None:
                continue
            p.setPen(QPen(QColor("#333333"), 1, Qt.DotLine))
            p.setBrush(Qt.NoBrush)
            p.drawRect(r.adjusted(-2, -2, 1, 1))
            primary = name == d.selected[-1]
            for h in d.handle_rects(r).values():
                p.fillRect(h, QColor("#FFFFFF") if not primary else QColor("#000000"))
                p.setPen(QColor("#000000"))
                p.drawRect(h)
        if d.band is not None:
            p.setPen(QPen(QColor("#000000"), 1, Qt.DashLine))
            p.setBrush(QColor(0, 120, 215, 30))
            p.drawRect(d.band.normalized())
        p.end()


class DesignSurface(QWidget):
    """Grey workspace with the form drawn as a window (caption + border)."""

    def __init__(self, designer):
        super().__init__()
        self.designer = designer
        self.setMouseTracking(True)
        self.client = DesignClient(designer)
        self.client.setParent(self)
        self.overlay = SelectionOverlay(self.client, designer)

    def layout_form(self):
        f = self.designer.form
        cw = max(40, px(float(f.get("width") or 240) - FORM_BORDER_W))
        ch = max(30, px(float(f.get("height") or 180) - FORM_BORDER_H))
        self.client.setGeometry(MARGIN + FRAME_PAD, MARGIN + TITLE_H, cw, ch)
        self.overlay.setGeometry(0, 0, cw, ch)
        self.overlay.raise_()
        self.setMinimumSize(MARGIN * 2 + cw + FRAME_PAD * 2 + 20,
                            MARGIN * 2 + TITLE_H + ch + FRAME_PAD + 20)
        self.update()

    def window_rect(self) -> QRect:
        c = self.client.geometry()
        return QRect(c.left() - FRAME_PAD, c.top() - TITLE_H, c.width() + 2 * FRAME_PAD,
                     c.height() + TITLE_H + FRAME_PAD)

    def form_handles(self) -> Dict[str, QRect]:
        w = self.window_rect()
        s = HANDLE
        return {"r": QRect(w.right() + 1, w.center().y() - s // 2, s, s),
                "b": QRect(w.center().x() - s // 2, w.bottom() + 1, s, s),
                "rb": QRect(w.right() + 1, w.bottom() + 1, s, s)}

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor("#ABABAB"))
        w = self.window_rect()
        # window frame
        p.fillRect(w, QColor("#FFFFFF"))
        p.setPen(QColor("#6E6E6E"))
        p.drawRect(w.adjusted(0, 0, -1, -1))
        title = QRect(w.left() + 1, w.top() + 1, w.width() - 2, TITLE_H - 2)
        g = QLinearGradient(title.topLeft(), title.bottomLeft())
        g.setColorAt(0, QColor("#FFFFFF"))
        g.setColorAt(1, QColor("#E9E9E9"))
        p.fillRect(title, g)
        f = make_font(self.designer.form)
        f.setPointSizeF(9)
        p.setFont(f)
        p.setPen(QColor("#1E1E1E"))
        p.drawText(title.adjusted(8, 0, -30, 0), Qt.AlignVCenter | Qt.AlignLeft,
                   str(self.designer.form.get("caption") or ""))
        p.drawText(QRect(title.right() - 26, title.top(), 24, title.height()), Qt.AlignCenter,
                   "✕")
        # form handles when the form itself is selected
        if not self.designer.selected:
            p.setPen(QColor("#000000"))
            for h in self.form_handles().values():
                p.fillRect(h, QColor("#FFFFFF"))
                p.drawRect(h)
        p.end()

    def mousePressEvent(self, e):
        pos = e.position().toPoint()
        for k, h in self.form_handles().items():
            if h.adjusted(-2, -2, 2, 2).contains(pos):
                self.designer.begin_form_resize(k, pos)
                return
        self.designer.select([])

    def mouseMoveEvent(self, e):
        pos = e.position().toPoint()
        if self.designer.form_resize is not None:
            self.designer.update_form_resize(pos)
            return
        cur = Qt.ArrowCursor
        for k, h in self.form_handles().items():
            if h.adjusted(-2, -2, 2, 2).contains(pos):
                cur = {"r": Qt.SizeHorCursor, "b": Qt.SizeVerCursor,
                       "rb": Qt.SizeFDiagCursor}[k]
        self.setCursor(cur)

    def mouseReleaseEvent(self, e):
        if self.designer.form_resize is not None:
            self.designer.end_form_resize()


# --------------------------------------------------------------------------
# Toolbox
# --------------------------------------------------------------------------
class Toolbox(QFrame):
    def __init__(self, designer):
        super().__init__()
        self.designer = designer
        self.setObjectName("ufToolbox")
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("QFrame#ufToolbox{background:#F0F0F0;}"
                           "QToolButton{border:1px solid transparent; padding:2px;}"
                           "QToolButton:hover{border:1px solid #7FA7D3; background:#E5F1FB;}"
                           "QToolButton:checked{border:1px solid #3C7FB1; background:#CCE4F7;}")
        v = QVBoxLayout(self)
        v.setContentsMargins(4, 4, 4, 4)
        v.setSpacing(2)
        t = QLabel("Toolbox")
        t.setStyleSheet("font-weight:bold; padding:2px;")
        v.addWidget(t)
        tab = QLabel(" Controls ")
        tab.setStyleSheet("background:white; border:1px solid #ABABAB; border-bottom:none;"
                          "padding:2px 6px;")
        tab.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        v.addWidget(tab)
        box = QFrame()
        box.setStyleSheet("QFrame{background:white; border:1px solid #ABABAB;}"
                          "QToolButton{background:transparent;}")
        g = QGridLayout(box)
        g.setContentsMargins(3, 3, 3, 3)
        g.setSpacing(1)
        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        self.buttons: Dict[str, QToolButton] = {}
        for i, kind in enumerate(("Pointer",) + forms.CONTROL_TYPES):
            b = QToolButton()
            b.setIcon(tool_icon(kind))
            b.setIconSize(QSize(24, 24))
            b.setCheckable(True)
            b.setToolTip(kind)
            b.setAutoRaise(True)
            self.group.addButton(b)
            self.buttons[kind] = b
            b.clicked.connect(lambda _=False, k=kind: designer.set_tool(k))
            g.addWidget(b, i // 3, i % 3)
        self.buttons["Pointer"].setChecked(True)
        v.addWidget(box)
        v.addStretch(1)
        self.setFixedWidth(118)

    def set_tool(self, kind):
        b = self.buttons.get(kind)
        if b is not None:
            b.setChecked(True)


# --------------------------------------------------------------------------
# Properties window
# --------------------------------------------------------------------------
_FORM_PROPS = [("(Name)", "name", "name"), ("BackColor", "back_color", "color"),
               ("Caption", "caption", "text"), ("Font", "font", "font"),
               ("ForeColor", "fore_color", "color"), ("Height", "height", "num"),
               ("StartUpPosition", "start_up_position",
                "enum:0 - Manual|1 - CenterOwner|2 - CenterScreen"),
               ("Width", "width", "num")]

_COMMON = [("(Name)", "name", "name"), ("BackColor", "back_color", "color"),
           ("ControlSource", "control_source", "text"),
           ("ControlTipText", "control_tip_text", "text"), ("Enabled", "enabled", "bool"),
           ("Font", "font", "font"), ("ForeColor", "fore_color", "color"),
           ("Height", "height", "num"), ("Left", "left", "num"), ("Locked", "locked", "bool"),
           ("TabIndex", "tab_index", "int"), ("TabStop", "tab_stop", "bool"),
           ("Tag", "tag", "text"), ("Top", "top", "num"), ("Visible", "visible", "bool"),
           ("Width", "width", "num")]

_EXTRA = {
    "Label": [("Caption", "caption", "text"), ("TextAlign", "text_align",
                                               "enum:1 - Left|2 - Center|3 - Right"),
              ("WordWrap", "word_wrap", "bool")],
    "TextBox": [("MaxLength", "max_length", "int"), ("MultiLine", "multi_line", "bool"),
                ("PasswordChar", "password_char", "char"),
                ("TextAlign", "text_align", "enum:1 - Left|2 - Center|3 - Right"),
                ("Value", "value", "text")],
    "CommandButton": [("Cancel", "cancel", "bool"), ("Caption", "caption", "text"),
                      ("Default", "default", "bool"), ("Accelerator", "accelerator", "char")],
    "CheckBox": [("Caption", "caption", "text"), ("Value", "value", "bool")],
    "OptionButton": [("Caption", "caption", "text"), ("GroupName", "group_name", "text"),
                     ("Value", "value", "bool")],
    "ToggleButton": [("Caption", "caption", "text"), ("Value", "value", "bool")],
    "ComboBox": [("BoundColumn", "bound_column", "int"), ("ColumnCount", "column_count", "int"),
                 ("ColumnWidths", "column_widths", "text"), ("List", "list", "list"),
                 ("ListStyle", "list_style", "enum:0 - Plain|1 - Option"),
                 ("RowSource", "row_source", "text"),
                 ("Style", "style", "enum:0 - DropDownCombo|2 - DropDownList"),
                 ("Value", "value", "text")],
    "ListBox": [("BoundColumn", "bound_column", "int"), ("ColumnCount", "column_count", "int"),
                ("ColumnWidths", "column_widths", "text"), ("List", "list", "list"),
                ("ListStyle", "list_style", "enum:0 - Plain|1 - Option"),
                ("MultiSelect", "multi_select", "enum:0 - Single|1 - Multi|2 - Extended"),
                ("RowSource", "row_source", "text")],
    "Frame": [("Caption", "caption", "text")],
    "MultiPage": [("Pages", "pages", "pages"), ("Value", "value", "int")],
    "TabStrip": [("Tabs", "tabs", "pages"), ("Value", "value", "int")],
    "SpinButton": [("Max", "max", "int"), ("Min", "min", "int"),
                   ("Orientation", "orientation", "enum:-1 - Auto|0 - Vertical|1 - Horizontal"),
                   ("SmallChange", "small_change", "int"), ("Value", "value", "int")],
    "ScrollBar": [("LargeChange", "large_change", "int"), ("Max", "max", "int"),
                  ("Min", "min", "int"),
                  ("Orientation", "orientation", "enum:-1 - Auto|0 - Vertical|1 - Horizontal"),
                  ("SmallChange", "small_change", "int"), ("Value", "value", "int")],
    "Image": [("PictureSizeMode", "picture_size_mode", "enum:0 - Clip|1 - Stretch|3 - Zoom")],
}


def props_for(ctrl: Optional[dict]):
    if ctrl is None:
        return _FORM_PROPS
    rows = {r[0]: r for r in _COMMON}
    for r in _EXTRA.get(ctrl.get("type"), []):
        rows[r[0]] = r
    return [rows[k] for k in sorted(rows, key=lambda s: (s != "(Name)", s.lower()))]


class PropertyGrid(QWidget):
    def __init__(self, designer):
        super().__init__()
        self.designer = designer
        self._building = False
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(2)
        self.title = QLabel("Properties")
        self.title.setStyleSheet("font-weight:bold; padding:3px; background:#F0F0F0;")
        v.addWidget(self.title)
        self.combo = QComboBox()
        self.combo.activated.connect(self._combo_chosen)
        v.addWidget(self.combo)
        self.table = QTableWidget(0, 2)
        self.table.horizontalHeader().setVisible(False)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 110)
        self.table.setSelectionMode(QAbstractItemView.NoSelection)
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.table.setStyleSheet("QTableWidget{gridline-color:#D0D0D0; font-size:8pt;}"
                                 "QLineEdit, QComboBox{border:none;}")
        v.addWidget(self.table, 1)
        self.setMinimumWidth(250)

    def _combo_chosen(self, i):
        name = self.combo.itemData(i)
        self.designer.select([name] if name else [])

    def refresh(self):
        d = self.designer
        self._building = True
        try:
            self.combo.clear()
            self.combo.addItem(f"{d.module['name']} UserForm", None)
            for c in d.form.get("controls") or []:
                self.combo.addItem(f"{c['name']} {c['type']}", c["name"])
            target = d.primary()
            if target is not None:
                idx = self.combo.findData(target["name"])
                self.combo.setCurrentIndex(max(0, idx))
            else:
                self.combo.setCurrentIndex(0)
            name = target["name"] if target else d.module["name"]
            self.title.setText(f"Properties - {name}")
            rows = props_for(target)
            obj = target if target is not None else d.form
            self.table.clearContents()
            self.table.setRowCount(0)
            self.table.setRowCount(len(rows))
            for i, (label, key, kind) in enumerate(rows):
                it = self.table.item(i, 0)
                from PySide6.QtWidgets import QTableWidgetItem
                it = QTableWidgetItem(label)
                it.setFlags(Qt.ItemIsEnabled)
                self.table.setItem(i, 0, it)
                value = d.module["name"] if (key == "name" and target is None) else obj.get(key)
                self.table.setCellWidget(i, 1, self._editor(key, kind, value, obj))
        finally:
            self._building = False

    def _commit(self, key, value):
        if self._building:
            return
        self.designer.set_property(key, value)

    def _editor(self, key, kind, value, obj):
        if kind in ("text", "name", "char"):
            e = QLineEdit("" if value is None else str(value))
            if kind == "char":
                e.setMaxLength(1)
            e.editingFinished.connect(lambda e=e, k=key, kd=kind: self._text_done(k, kd, e))
            return e
        if kind in ("num", "int"):
            e = QLineEdit(_fmt_num(value))
            e.editingFinished.connect(lambda e=e, k=key, kd=kind: self._num_done(k, kd, e))
            return e
        if kind == "bool":
            cb = QComboBox()
            cb.addItems(["False", "True"])
            cb.setCurrentIndex(1 if value else 0)
            cb.activated.connect(lambda i, k=key: self._commit(k, bool(i)))
            return cb
        if kind.startswith("enum:"):
            opts = kind[5:].split("|")
            cb = QComboBox()
            cb.addItems(opts)
            vals = [int(o.split(" - ")[0]) for o in opts]
            try:
                cb.setCurrentIndex(vals.index(int(value if value is not None else vals[0])))
            except ValueError:
                pass
            cb.activated.connect(lambda i, k=key, vs=vals: self._commit(k, vs[i]))
            return cb
        if kind == "color":
            b = QPushButton(str(value) if value else "(default)")
            col = to_color(value)
            pm = QPixmap(12, 12)
            pm.fill(col or QColor(FORM_BG))
            b.setIcon(QIcon(pm))
            b.setStyleSheet("text-align:left; border:none; padding:1px 4px;")
            b.clicked.connect(lambda _=False, k=key, v=value: self._pick_color(k, v))
            return b
        if kind == "font":
            fam = obj.get("font_name") or self.designer.form.get("font_name") or "Tahoma"
            size = obj.get("font_size") or self.designer.form.get("font_size") or 8
            b = QPushButton(f"{fam}, {_fmt_num(size)}pt")
            b.setStyleSheet("text-align:left; border:none; padding:1px 4px;")
            b.clicked.connect(self._pick_font)
            return b
        if kind == "list":
            items = value or []
            b = QPushButton(f"({len(items)} items) ...")
            b.setStyleSheet("text-align:left; border:none; padding:1px 4px;")
            b.clicked.connect(lambda _=False, k=key: self._edit_list(k))
            return b
        if kind == "pages":
            items = value or []
            b = QPushButton(", ".join(str(p.get("caption", "")) for p in items) + " ...")
            b.setStyleSheet("text-align:left; border:none; padding:1px 4px;")
            b.clicked.connect(lambda _=False, k=key: self._edit_pages(k))
            return b
        return QLabel(str(value))

    def _text_done(self, key, kind, e):
        if self._building:
            return
        text = e.text()
        if key == "name":
            if not self.designer.rename(text):
                target = self.designer.primary()
                e.setText(target["name"] if target else self.designer.module["name"])
            return
        cur = self.designer.get_property(key)
        if (cur or "") != text:
            self._commit(key, text if text or key not in ("control_source", "row_source")
                         else None)

    def _num_done(self, key, kind, e):
        try:
            v = float(e.text().replace(",", "."))
        except ValueError:
            e.setText(_fmt_num(self.designer.get_property(key)))
            return
        if kind == "int":
            v = int(v)
        if self.designer.get_property(key) != v:
            self._commit(key, v)

    def _pick_color(self, key, value):
        c = QColorDialog.getColor(to_color(value) or QColor(FORM_BG), self, key)
        if c.isValid():
            self._commit(key, c.name().upper())

    def _pick_font(self):
        d = self.designer
        obj = d.primary() or d.form
        ok, f = QFontDialog.getFont(make_font(d.form, obj if d.primary() else None), self)
        if isinstance(ok, QFont):  # PySide returns (ok, font) or (font, ok) by version
            ok, f = f, ok
        if ok:
            d.set_property("font", f)

    def _edit_list(self, key):
        d = self.designer
        cur = d.get_property(key) or []
        text = "\n".join("\t".join(map(str, r)) if isinstance(r, (list, tuple)) else str(r)
                         for r in cur)
        new, ok = QInputDialog.getMultiLineText(self, "List", "One item per line "
                                                "(separate columns with Tab):", text)
        if ok:
            rows = []
            for line in new.splitlines():
                if not line.strip():
                    continue
                rows.append(line.split("\t") if "\t" in line else line)
            d.set_property(key, rows)

    def _edit_pages(self, key):
        d = self.designer
        cur = d.get_property(key) or []
        text = "\n".join(str(p.get("caption", p.get("name", ""))) for p in cur)
        new, ok = QInputDialog.getMultiLineText(self, "Pages", "One caption per line:", text)
        if ok:
            caps = [x.strip() for x in new.splitlines() if x.strip()]
            base = "Page" if key == "pages" else "Tab"
            out = []
            for i, cap in enumerate(caps):
                old = cur[i] if i < len(cur) else {}
                out.append({"name": old.get("name") or d.unique_page_name(base),
                            "caption": cap})
            d.set_property(key, out)


def _fmt_num(v) -> str:
    if v is None:
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(f)) if f.is_integer() else f"{f:.2f}".rstrip("0").rstrip(".")


# --------------------------------------------------------------------------
# Designer widget
# --------------------------------------------------------------------------
class FormDesigner(QWidget):
    changed = Signal()
    _clipboard: List[dict] = []

    def __init__(self, editor, module: dict, parent=None):
        super().__init__(parent)
        self.editor = editor            # VBAEditor (may be None in tests)
        self.module = module
        module["form"] = forms.normalize_form(module.get("form"))
        if not module["form"].get("caption"):
            module["form"]["caption"] = module["name"]
        self.selected: List[str] = []
        self.tool = "Pointer"
        self.drag = None
        self.band: Optional[QRect] = None
        self.form_resize = None
        self.undo_stack: List[dict] = []
        self.redo_stack: List[dict] = []
        self.setFocusPolicy(Qt.StrongFocus)
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        self.toolbox = Toolbox(self)
        self.surface = DesignSurface(self)
        self.scroll = QScrollArea()
        self.scroll.setWidget(self.surface)
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("QScrollArea{background:#ABABAB; border:none;}")
        self.props = PropertyGrid(self)
        split = QSplitter(Qt.Horizontal)
        split.addWidget(self.toolbox)
        split.addWidget(self.scroll)
        split.addWidget(self.props)
        split.setStretchFactor(1, 1)
        split.setSizes([118, 700, 270])
        h.addWidget(split)
        self.renderer = FormRenderer(self.surface.client, self.form, live=False)
        self.surface.layout_form()
        self.props.refresh()

    # -- model helpers -----------------------------------------------------------
    @property
    def form(self) -> dict:
        return self.module["form"]

    def controls(self) -> List[dict]:
        return self.form.setdefault("controls", [])

    def ctrl(self, name) -> Optional[dict]:
        if not name:
            return None
        for c in self.controls():
            if c["name"].lower() == str(name).lower():
                return c
        return None

    def primary(self) -> Optional[dict]:
        return self.ctrl(self.selected[-1]) if self.selected else None

    def get_property(self, key):
        obj = self.primary() or self.form
        return obj.get(key)

    def _snapshot(self):
        self.undo_stack.append(copy.deepcopy(self.form))
        del self.undo_stack[:-100]
        self.redo_stack.clear()

    def _restore(self, snap):
        self.form.clear()
        self.form.update(copy.deepcopy(snap))
        self.selected = [n for n in self.selected if self.ctrl(n)]
        self.rebuild()
        self._modified()

    def undo(self):
        if self.undo_stack:
            self.redo_stack.append(copy.deepcopy(self.form))
            self._restore(self.undo_stack.pop())

    def redo(self):
        if self.redo_stack:
            self.undo_stack.append(copy.deepcopy(self.form))
            self._restore(self.redo_stack.pop())

    def _modified(self):
        self.module["form"] = self.form
        if self.editor is not None:
            self.editor.book.vba_modified = True
            self.editor.win.vba_code_changed()
        self.changed.emit()

    def rebuild(self):
        self.renderer.form = self.form
        self.renderer.build()
        self.surface.layout_form()
        self.surface.overlay.raise_()
        self.surface.overlay.update()
        self.surface.update()
        self.props.refresh()

    def refresh_ctrl(self, c):
        self.renderer.apply(c)
        if c.get("type") == "Frame":
            for ch in self.controls():
                if (ch.get("parent") or "").lower() == c["name"].lower():
                    self.renderer.apply(ch)
        self.surface.overlay.update()

    # -- geometry ----------------------------------------------------------------
    def container_origin(self, parent: Optional[str]) -> QPoint:
        w = self.renderer.container_for(parent)
        if w is self.surface.client:
            return QPoint(0, 0)
        return w.mapTo(self.surface.client, QPoint(0, 0))

    def ctrl_rect(self, name) -> Optional[QRect]:
        w = self.renderer.widgets.get(name)
        c = self.ctrl(name)
        if w is None or c is None:
            return None
        if not w.isVisibleTo(self.surface.client):
            return None
        return QRect(w.mapTo(self.surface.client, QPoint(0, 0)), w.size())

    def handle_rects(self, r: QRect) -> Dict[str, QRect]:
        s = HANDLE
        xs = {"l": r.left() - s, "c": r.center().x() - s // 2, "r": r.right() + 1}
        ys = {"t": r.top() - s, "m": r.center().y() - s // 2, "b": r.bottom() + 1}
        out = {}
        for vk, y in ys.items():
            for hk, x in xs.items():
                if vk == "m" and hk == "c":
                    continue
                out[vk + hk] = QRect(x, y, s, s)
        return out

    def depth(self, c) -> int:
        d = 0
        p = c.get("parent")
        while p and d < 10:
            d += 1
            owner = self.ctrl(p) or self.ctrl(self.renderer.page_owner.get(p, ""))
            p = owner.get("parent") if owner else None
        return d

    def hit(self, pos: QPoint) -> Optional[dict]:
        best = None
        best_key = None
        for i, c in enumerate(self.controls()):
            r = self.ctrl_rect(c["name"])
            if r is not None and r.contains(pos):
                key = (self.depth(c), i)
                if best_key is None or key > best_key:
                    best, best_key = c, key
        return best

    def container_at(self, pos: QPoint) -> Optional[str]:
        """Name of the Frame / MultiPage page under pos, or None (form)."""
        best = None
        best_key = None
        for i, c in enumerate(self.controls()):
            if c.get("type") not in ("Frame", "MultiPage"):
                continue
            r = self.ctrl_rect(c["name"])
            if r is None or not r.contains(pos):
                continue
            if c["type"] == "Frame":
                cont = c["name"]
            else:
                pages = c.get("pages") or []
                idx = int(c.get("value") or 0)
                if not pages:
                    continue
                cont = pages[max(0, min(idx, len(pages) - 1))]["name"]
            inner = self.renderer.containers.get(cont)
            if inner is None:
                continue
            ir = QRect(inner.mapTo(self.surface.client, QPoint(0, 0)), inner.size())
            if not ir.contains(pos):
                continue
            key = (self.depth(c), i)
            if best_key is None or key > best_key:
                best, best_key = cont, key
        return best

    # -- API used by tests and the property grid --------------------------------------
    def set_tool(self, kind):
        self.tool = kind
        self.toolbox.set_tool(kind)
        self.surface.client.setCursor(Qt.ArrowCursor if kind == "Pointer" else Qt.CrossCursor)

    def select(self, names: List[str]):
        self.selected = [n for n in names if n and self.ctrl(n)]
        self.surface.overlay.update()
        self.surface.update()
        self.props.refresh()

    def place_control(self, ctype: str, left: float, top: float, width: float = None,
                      height: float = None, parent: Optional[str] = None,
                      **props) -> dict:
        """Add a control (positions in points relative to parent)."""
        self._snapshot()
        name = props.pop("name", None) or forms.unique_control_name(self.form, ctype)
        c = forms.new_control(ctype, name, **props)
        c["left"] = snap(max(0.0, float(left)))
        c["top"] = snap(max(0.0, float(top)))
        if width:
            c["width"] = max(GRID_PT, snap(float(width)))
        if height:
            c["height"] = max(GRID_PT, snap(float(height)))
        c["parent"] = parent
        c["tab_index"] = sum(1 for x in self.controls()
                             if (x.get("parent") or "") == (parent or ""))
        if ctype == "MultiPage":
            c["pages"] = [{"name": self.unique_page_name("Page"), "caption": "Page1"}]
            c["pages"].append({"name": self.unique_page_name("Page", c["pages"]),
                               "caption": "Page2"})
        self.controls().append(c)
        self.rebuild()
        self.select([name])
        self._modified()
        return c

    def unique_page_name(self, base="Page", extra=()) -> str:
        used = {c["name"].lower() for c in self.controls()}
        for c in self.controls():
            for p in (c.get("pages") or []) + (c.get("tabs") or []):
                used.add(str(p.get("name", "")).lower())
        for p in extra:
            used.add(str(p.get("name", "")).lower())
        i = 1
        while f"{base}{i}".lower() in used:
            i += 1
        return f"{base}{i}"

    def set_property(self, key: str, value):
        targets = [self.ctrl(n) for n in self.selected] or [None]
        self._snapshot()
        structural = False
        for c in targets:
            obj = c if c is not None else self.form
            if key == "font":
                f: QFont = value
                obj["font_name"] = f.family()
                obj["font_size"] = round(f.pointSizeF(), 1)
                if c is not None:
                    obj["bold"] = f.bold()
                    obj["italic"] = f.italic()
                    obj["underline"] = f.underline()
                structural = c is None
                continue
            if key in ("left", "top", "width", "height") and c is not None:
                value = max(0.0, float(value))
            if c is None and key in ("width", "height"):
                value = max(30.0, float(value))
            if key == "default" and value and c is not None:
                for o in self.controls():
                    o["default"] = False
            if key == "cancel" and value and c is not None:
                for o in self.controls():
                    o["cancel"] = False
            if key == "tab_index" and c is not None:
                self._set_tab_index(c, int(value))
                continue
            obj[key] = value
            if c is None or key in ("multi_line", "style", "column_count", "pages", "tabs",
                                    "list", "default", "cancel"):
                structural = True
            else:
                self.refresh_ctrl(c)
        if structural:
            self.rebuild()
        else:
            self.props.refresh()
        self.surface.update()
        self._modified()

    def _set_tab_index(self, c, new):
        sibs = sorted((x for x in self.controls()
                       if (x.get("parent") or "") == (c.get("parent") or "") and x is not c),
                      key=lambda x: int(x.get("tab_index") or 0))
        new = max(0, min(new, len(sibs)))
        sibs.insert(new, c)
        for i, x in enumerate(sibs):
            x["tab_index"] = i

    def rename(self, new: str) -> bool:
        """Rename the primary selection (or the form module when nothing is
        selected). Event procedures in the code are renamed too."""
        new = new.strip()
        target = self.primary()
        if target is None:
            if self.editor is not None:
                return self.editor.rename_module(self.module["name"], new)
            if not _IDENT.match(new):
                return False
            self.module["name"] = new
            return True
        old = target["name"]
        if new == old:
            return True
        if not _IDENT.match(new) or (self.ctrl(new) is not None and
                                     new.lower() != old.lower()):
            QMessageBox.warning(self, "PortableExcel VBA Editor",
                                f"Could not set the Name property. '{new}' is not a valid "
                                "or unique name.")
            return False
        self._snapshot()
        target["name"] = new
        for c in self.controls():
            if (c.get("parent") or "").lower() == old.lower():
                c["parent"] = new
        self.module["code"] = rename_event_procs(self.module.get("code", ""), old, new)
        if self.editor is not None:
            self.editor.code_replaced(self.module["name"])
        self.selected = [new if n == old else n for n in self.selected]
        self.rebuild()
        self._modified()
        return True

    def delete_selected(self):
        if not self.selected:
            return
        self._snapshot()
        gone = {n.lower() for n in self.selected}
        # also delete children of deleted containers
        changed = True
        while changed:
            changed = False
            for c in self.controls():
                par = (c.get("parent") or "").lower()
                owner = self.renderer.page_owner.get(c.get("parent") or "", "").lower()
                if c["name"].lower() not in gone and (par in gone or owner in gone):
                    gone.add(c["name"].lower())
                    changed = True
        self.form["controls"] = [c for c in self.controls() if c["name"].lower() not in gone]
        self.selected = []
        self.rebuild()
        self._modified()

    def copy_selected(self):
        FormDesigner._clipboard = [copy.deepcopy(self.ctrl(n)) for n in self.selected
                                   if self.ctrl(n)]

    def cut_selected(self):
        self.copy_selected()
        self.delete_selected()

    def paste(self):
        if not FormDesigner._clipboard:
            return []
        self._snapshot()
        names = []
        mapping = {}
        for c in FormDesigner._clipboard:
            n = copy.deepcopy(c)
            new = forms.unique_control_name(self.form, n["type"])
            mapping[c["name"]] = new
            n["name"] = new
            par = n.get("parent")
            if par in mapping:
                n["parent"] = mapping[par]
            elif par and not self.ctrl(par) and par not in self.renderer.containers:
                n["parent"] = None
            if n.get("parent") not in mapping.values():
                n["left"] = float(n.get("left") or 0) + GRID_PT
                n["top"] = float(n.get("top") or 0) + GRID_PT
            if n.get("type") == "MultiPage":
                pages = []
                for p in n.get("pages") or []:
                    pages.append({"name": self.unique_page_name("Page", pages),
                                  "caption": p.get("caption", "")})
                n["pages"] = pages
            self.controls().append(n)
            names.append(new)
        self.rebuild()
        self.select(names)
        self._modified()
        return names

    def nudge(self, dx_pt: float, dy_pt: float, resize=False):
        if not self.selected:
            return
        self._snapshot()
        for n in self.selected:
            c = self.ctrl(n)
            if resize:
                c["width"] = max(GRID_PT / 2, float(c["width"]) + dx_pt)
                c["height"] = max(GRID_PT / 2, float(c["height"]) + dy_pt)
            else:
                c["left"] = max(0.0, float(c["left"]) + dx_pt)
                c["top"] = max(0.0, float(c["top"]) + dy_pt)
            self.refresh_ctrl(c)
        self.props.refresh()
        self._modified()

    def view_code(self, name: Optional[str] = None) -> str:
        """Open the code module; create/jump to the default event procedure."""
        c = self.ctrl(name) if name else self.primary()
        if c is None:
            proc = "UserForm_Click"
        else:
            proc = f"{c['name']}_{DEFAULT_EVENT.get(c['type'], 'Click')}"
        code = self.module.get("code", "") or ""
        if not re.search(rf"^\s*(Private\s+|Public\s+)?Sub\s+{re.escape(proc)}\s*\(",
                         code, re.I | re.M):
            args = ""
            if proc.endswith(("_KeyPress",)):
                args = "ByVal KeyAscii As MSForms.ReturnInteger"
            if code and not code.endswith("\n"):
                code += "\n"
            code += f"\nPrivate Sub {proc}({args})\n    \nEnd Sub\n"
            self.module["code"] = code
            self._modified()
            if self.editor is not None:
                self.editor.code_replaced(self.module["name"])
        if self.editor is not None:
            self.editor.open_code(self.module["name"], proc)
        return proc

    # -- mouse ---------------------------------------------------------------------
    def mouse_press(self, pos: QPoint, e):
        self.setFocus()
        if e.button() != Qt.LeftButton:
            return
        if self.tool != "Pointer":
            self.drag = {"mode": "place", "start": pos, "cur": pos}
            self.band = QRect(pos, pos)
            self.surface.overlay.update()
            return
        # resize handles of the primary selection
        if len(self.selected) == 1:
            r = self.ctrl_rect(self.selected[0])
            if r is not None:
                for k, h in self.handle_rects(r).items():
                    if h.adjusted(-2, -2, 2, 2).contains(pos):
                        c = self.ctrl(self.selected[0])
                        self.drag = {"mode": "resize", "handle": k, "start": pos,
                                     "orig": {kk: float(c[kk]) for kk in
                                              ("left", "top", "width", "height")},
                                     "moved": False}
                        return
        c = self.hit(pos)
        mods = e.modifiers()
        if c is None:
            if not (mods & (Qt.ShiftModifier | Qt.ControlModifier)):
                self.select([])
            self.drag = {"mode": "band", "start": pos}
            self.band = QRect(pos, pos)
            return
        name = c["name"]
        if mods & (Qt.ShiftModifier | Qt.ControlModifier):
            sel = list(self.selected)
            if name in sel:
                sel.remove(name)
            else:
                sel.append(name)
            self.select(sel)
        elif name not in self.selected:
            self.select([name])
        else:  # make it primary
            self.selected.remove(name)
            self.selected.append(name)
            self.props.refresh()
        if c.get("type") == "MultiPage":
            self._maybe_switch_page(c, pos)
        self.drag = {"mode": "move", "start": pos, "moved": False,
                     "orig": {n: (float(self.ctrl(n)["left"]), float(self.ctrl(n)["top"]))
                              for n in self.selected}}

    def _maybe_switch_page(self, c, pos):
        w = self.renderer.widgets.get(c["name"])
        if w is None:
            return
        bar = w.tabBar()
        local = w.mapFrom(self.surface.client, pos)
        local = bar.mapFrom(w, local)
        i = bar.tabAt(local)
        if i >= 0 and i != int(c.get("value") or 0):
            c["value"] = i
            w.setCurrentIndex(i)
            self.surface.overlay.update()

    def mouse_move(self, pos: QPoint, e):
        d = self.drag
        if d is None:
            if self.tool == "Pointer" and len(self.selected) == 1:
                r = self.ctrl_rect(self.selected[0])
                cur = Qt.ArrowCursor
                if r is not None:
                    for k, h in self.handle_rects(r).items():
                        if h.adjusted(-2, -2, 2, 2).contains(pos):
                            cur = {"tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
                                   "bl": Qt.SizeBDiagCursor, "tr": Qt.SizeBDiagCursor,
                                   "tc": Qt.SizeVerCursor, "bc": Qt.SizeVerCursor,
                                   "ml": Qt.SizeHorCursor, "mr": Qt.SizeHorCursor}[k]
                    if cur == Qt.ArrowCursor and r.contains(pos):
                        cur = Qt.SizeAllCursor
                self.surface.client.setCursor(cur)
            return
        if d["mode"] in ("place", "band"):
            self.band = QRect(d["start"], pos)
            self.surface.overlay.update()
            return
        dx = px_to_points(pos.x() - d["start"].x())
        dy = px_to_points(pos.y() - d["start"].y())
        if d["mode"] == "move":
            if not d["moved"] and abs(pos.x() - d["start"].x()) + \
                    abs(pos.y() - d["start"].y()) < 3:
                return
            if not d["moved"]:
                self._snapshot()
                d["moved"] = True
            for n, (l0, t0) in d["orig"].items():
                c = self.ctrl(n)
                c["left"] = max(0.0, snap(l0 + dx))
                c["top"] = max(0.0, snap(t0 + dy))
                self.refresh_ctrl(c)
            return
        if d["mode"] == "resize":
            if not d["moved"]:
                self._snapshot()
                d["moved"] = True
            c = self.ctrl(self.selected[0])
            o = d["orig"]
            k = d["handle"]
            left, top, w, h = o["left"], o["top"], o["width"], o["height"]
            if "l" in k[1]:
                nl = min(snap(left + dx), left + w - GRID_PT)
                w = w + (left - nl)
                left = nl
            if "r" in k[1]:
                w = max(GRID_PT, snap(left + w + dx) - left)
            if k[0] == "t":
                nt = min(snap(top + dy), top + h - GRID_PT)
                h = h + (top - nt)
                top = nt
            if k[0] == "b":
                h = max(GRID_PT, snap(top + h + dy) - top)
            c.update(left=max(0.0, left), top=max(0.0, top), width=w, height=h)
            self.refresh_ctrl(c)

    def mouse_release(self, pos: QPoint, e):
        d, self.drag = self.drag, None
        if d is None:
            return
        if d["mode"] == "place":
            self.band = None
            r = QRect(d["start"], pos).normalized()
            parent = self.container_at(d["start"])
            origin = self.container_origin(parent)
            left = px_to_points(r.left() - origin.x())
            top = px_to_points(r.top() - origin.y())
            if r.width() < 6 and r.height() < 6:
                w = h = None
            else:
                w, h = px_to_points(r.width()), px_to_points(r.height())
            kind = self.tool
            self.place_control(kind, left, top, w, h, parent=parent)
            self.set_tool("Pointer")
            return
        if d["mode"] == "band":
            r = (self.band or QRect()).normalized()
            self.band = None
            if r.width() > 3 or r.height() > 3:
                names = [c["name"] for c in self.controls()
                         if (rr := self.ctrl_rect(c["name"])) is not None and r.intersects(rr)
                         and not c.get("parent")]
                self.select(names)
            self.surface.overlay.update()
            return
        if d.get("moved"):
            self.props.refresh()
            self._modified()

    def mouse_double(self, pos: QPoint, e):
        c = self.hit(pos)
        self.view_code(c["name"] if c else None)

    # form resize -----------------------------------------------------------------
    def begin_form_resize(self, handle, pos):
        self._snapshot()
        self.form_resize = {"handle": handle, "start": pos,
                            "w": float(self.form["width"]), "h": float(self.form["height"])}

    def update_form_resize(self, pos):
        fr = self.form_resize
        dx = px_to_points(pos.x() - fr["start"].x())
        dy = px_to_points(pos.y() - fr["start"].y())
        if "r" in fr["handle"]:
            self.form["width"] = max(60.0, snap(fr["w"] + dx))
        if "b" in fr["handle"]:
            self.form["height"] = max(48.0, snap(fr["h"] + dy))
        self.surface.layout_form()

    def end_form_resize(self):
        self.form_resize = None
        self.props.refresh()
        self._modified()

    # -- keyboard --------------------------------------------------------------------
    def keyPressEvent(self, e):
        k = e.key()
        ctrl = bool(e.modifiers() & Qt.ControlModifier)
        shift = bool(e.modifiers() & Qt.ShiftModifier)
        step = 0.75 if ctrl else GRID_PT
        if k == Qt.Key_Delete:
            self.delete_selected()
        elif k in (Qt.Key_Left, Qt.Key_Right, Qt.Key_Up, Qt.Key_Down):
            dx = {Qt.Key_Left: -step, Qt.Key_Right: step}.get(k, 0)
            dy = {Qt.Key_Up: -step, Qt.Key_Down: step}.get(k, 0)
            self.nudge(dx, dy, resize=shift)
        elif ctrl and k == Qt.Key_C:
            self.copy_selected()
        elif ctrl and k == Qt.Key_X:
            self.cut_selected()
        elif ctrl and k == Qt.Key_V:
            self.paste()
        elif ctrl and k == Qt.Key_Z:
            self.undo()
        elif ctrl and k == Qt.Key_Y:
            self.redo()
        elif ctrl and k == Qt.Key_A:
            self.select([c["name"] for c in self.controls() if not c.get("parent")])
        elif k == Qt.Key_Escape:
            self.set_tool("Pointer")
            self.select([])
        elif k == Qt.Key_F7:
            self.view_code()
        else:
            super().keyPressEvent(e)


def rename_event_procs(code: str, old: str, new: str) -> str:
    """Rename 'Sub Old_Event(' declarations (and Call sites 'Old_Event') to New_."""
    pat = re.compile(rf"\b{re.escape(old)}_(\w+)\b", re.I)
    events = set()
    for ev_list in forms.CONTROL_EVENTS.values():
        events.update(e.lower() for e in ev_list)
    events.update(("dropbuttonclick", "spinup", "spindown", "scroll", "change",
                   "beforedragover", "beforedroporpaste", "error", "layout", "addcontrol",
                   "removecontrol", "zoom"))

    def rep(m):
        if m.group(1).lower() in events:
            return f"{new}_{m.group(1)}"
        return m.group(0)
    return pat.sub(rep, code)

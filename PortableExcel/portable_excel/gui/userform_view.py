"""UserForm rendering with real Qt widgets.

FormRenderer builds the controls of a FORM_DICT into a client widget; it is
shared by the live form window (QtFormView, implements forms.FormView) and the
designer surface (non-interactive mode).
"""
from __future__ import annotations

import base64
import logging
import re
from typing import Callable, Dict, List, Optional

from PySide6.QtCore import QEvent, QEventLoop, QObject, QPoint, QRect, QSize, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QFontInfo, QGuiApplication, QPainter, QPalette, QPixmap
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
                               QFrame, QGroupBox, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QPlainTextEdit, QPushButton, QRadioButton,
                               QScrollBar, QTabBar, QTabWidget, QToolButton, QTreeWidget,
                               QTreeWidgetItem, QVBoxLayout, QWidget)

from .. import forms
from ..forms import points_to_px

log = logging.getLogger("PortableExcel")

FORM_BG = "#F0F0F0"
# outer form size minus the client ("Inside") size, in points
FORM_BORDER_W = 4.0
FORM_BORDER_H = 22.0
DEFAULT_EVENT = {
    "Label": "Click", "TextBox": "Change", "CommandButton": "Click", "CheckBox": "Click",
    "OptionButton": "Click", "ToggleButton": "Click", "ComboBox": "Change",
    "ListBox": "Click", "Frame": "Click", "MultiPage": "Change", "TabStrip": "Change",
    "SpinButton": "Change", "ScrollBar": "Change", "Image": "Click",
}

_subst_done = [False]


def _font_substitutions():
    if _subst_done[0]:
        return
    _subst_done[0] = True
    for fam in ("Tahoma", "MS Sans Serif"):
        QFont.insertSubstitutions(fam, ["Segoe UI", "DejaVu Sans", "Liberation Sans",
                                        "Noto Sans", "Inter"])


def norm_prop(p: str) -> str:
    """'BackColor' -> 'back_color'; snake_case passes through."""
    if not p:
        return p
    if "_" in p or p.islower():
        return p.lower()
    return re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", p).lower()


def to_color(v) -> Optional[QColor]:
    """'#RRGGBB' or an OLE colour int (BGR); system colours -> None."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        n = int(v)
        if n < 0 or n >= 0x80000000:
            return None
        return QColor(n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)
    c = QColor(str(v))
    return c if c.isValid() else None


def px(points) -> int:
    try:
        return points_to_px(float(points or 0))
    except (TypeError, ValueError):
        return 0


def _bool(v) -> bool:
    if isinstance(v, str):
        return v.strip().lower() in ("true", "-1", "1", "yes")
    return bool(v)


def make_font(form: dict, ctrl: Optional[dict] = None) -> QFont:
    _font_substitutions()
    name = (ctrl or {}).get("font_name") or form.get("font_name") or "Tahoma"
    size = (ctrl or {}).get("font_size") or form.get("font_size") or 8.0
    f = QFont(name)
    try:
        f.setPointSizeF(max(1.0, float(size)))
    except (TypeError, ValueError):
        f.setPointSizeF(8.0)
    if ctrl:
        f.setBold(_bool(ctrl.get("bold")))
        f.setItalic(_bool(ctrl.get("italic")))
        f.setUnderline(_bool(ctrl.get("underline")))
    return f


def list_rows(ctrl: dict) -> List[List[str]]:
    rows = []
    for it in ctrl.get("list") or []:
        if isinstance(it, (list, tuple)):
            rows.append(["" if x is None else str(x) for x in it])
        else:
            s = "" if it is None else str(it)
            rows.append(s.split("\t") if "\t" in s else [s])
    return rows


def column_widths_px(ctrl: dict, n: int) -> List[Optional[int]]:
    out: List[Optional[int]] = []
    spec = str(ctrl.get("column_widths") or "")
    for part in spec.replace(",", ";").split(";")[:n]:
        part = part.strip().lower()
        m = re.match(r"([\d.]+)\s*(pt|cm|in)?", part)
        if not m:
            out.append(None)
            continue
        val = float(m.group(1))
        unit = m.group(2) or "pt"
        val *= {"pt": 1, "cm": 28.35, "in": 72}[unit]
        out.append(px(val))
    while len(out) < n:
        out.append(None)
    return out


class _SpinButton(QWidget):
    """Small up/down (or left/right) button pair like the Forms SpinButton."""

    def __init__(self, parent, on_step: Callable[[int], None]):
        super().__init__(parent)
        self.on_step = on_step
        self.horizontal = False
        self.up = QToolButton(self)
        self.down = QToolButton(self)
        for b in (self.up, self.down):
            b.setAutoRepeat(True)
            b.setFocusPolicy(Qt.NoFocus)
        self.up.clicked.connect(lambda: self.on_step(+1))
        self.down.clicked.connect(lambda: self.on_step(-1))
        self._arrows()

    def _arrows(self):
        if self.horizontal:
            self.up.setArrowType(Qt.RightArrow)
            self.down.setArrowType(Qt.LeftArrow)
        else:
            self.up.setArrowType(Qt.UpArrow)
            self.down.setArrowType(Qt.DownArrow)

    def set_horizontal(self, h):
        self.horizontal = h
        self._arrows()
        self.resizeEvent(None)

    def resizeEvent(self, e):
        w, h = self.width(), self.height()
        if self.horizontal:
            self.down.setGeometry(0, 0, w // 2, h)
            self.up.setGeometry(w // 2, 0, w - w // 2, h)
        else:
            self.up.setGeometry(0, 0, w, h // 2)
            self.down.setGeometry(0, h // 2, w, h - h // 2)


class _ClickLabel(QLabel):
    def __init__(self, parent, fire):
        super().__init__(parent)
        self._fire = fire

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if e.button() == Qt.LeftButton:
            self._fire("Click")

    def mouseDoubleClickEvent(self, e):
        self._fire("DblClick")


class _ClickGroup(QGroupBox):
    def __init__(self, parent, fire):
        super().__init__(parent)
        self._fire = fire

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if e.button() == Qt.LeftButton and self.childAt(e.position().toPoint()) is None:
            self._fire("Click")

    def mouseDoubleClickEvent(self, e):
        self._fire("DblClick")


class _Client(QWidget):
    """Form client area; clicks on the background raise the form's Click."""

    def __init__(self, parent=None, fire=None):
        super().__init__(parent)
        self._fire = fire
        self.setAutoFillBackground(True)

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if self._fire and e.button() == Qt.LeftButton and \
                self.childAt(e.position().toPoint()) is None:
            self._fire(None, "Click")

    def mouseDoubleClickEvent(self, e):
        if self._fire and self.childAt(e.position().toPoint()) is None:
            self._fire(None, "DblClick")


# --------------------------------------------------------------------------
# Renderer
# --------------------------------------------------------------------------
class FormRenderer(QObject):
    """Creates Qt widgets for the controls of `form` inside `client`.

    live=False: design mode (no events, widgets transparent for the mouse).
    fire(ctrl_name, event, args) / user_set(ctrl, prop, value) are used in live mode.
    """

    def __init__(self, client: QWidget, form: dict, live: bool = False,
                 fire: Optional[Callable] = None, user_set: Optional[Callable] = None):
        super().__init__(client)
        self.client = client
        self.form = form
        self.live = live
        self._fire_cb = fire
        self._user_set_cb = user_set
        self.widgets: Dict[str, QWidget] = {}       # control name -> outer widget
        self.containers: Dict[str, QWidget] = {}    # frame / page name -> client widget
        self.page_owner: Dict[str, str] = {}         # page name -> multipage name
        self.ctrl_by_name: Dict[str, dict] = {}
        self._updating = 0
        self._entered_value: Dict[str, object] = {}
        self.build()

    # -- helpers -------------------------------------------------------------
    def controls(self) -> List[dict]:
        return self.form.get("controls") or []

    def ctrl(self, name: str) -> Optional[dict]:
        if name is None:
            return None
        c = self.ctrl_by_name.get(name.lower())
        if c is None:
            for cc in self.controls():
                if cc.get("name", "").lower() == name.lower():
                    self.ctrl_by_name[name.lower()] = cc
                    return cc
        return c

    def fire(self, name, event, args=()):
        if not self.live or self._updating or self._fire_cb is None:
            return None
        return self._fire_cb(name, event, tuple(args))

    def user_set(self, name, prop, value):
        if not self.live or self._updating or self._user_set_cb is None:
            return
        self._user_set_cb(name, prop, value)

    def clear(self):
        for w in list(self.widgets.values()):
            if w.parent() is not None:
                w.setParent(None)
            w.deleteLater()
        self.widgets.clear()
        self.containers.clear()
        self.page_owner.clear()
        self.ctrl_by_name.clear()

    def build(self):
        self.clear()
        self.apply_form()
        ctrls = self.controls()
        # containers first so children can be placed inside them
        pending = list(ctrls)
        for _round in range(8):
            rest = []
            for c in pending:
                parent = c.get("parent")
                if parent and parent.lower() not in {k.lower() for k in self.containers}:
                    if any(cc.get("name", "").lower() == parent.lower() or
                           parent.lower() in [p.get("name", "").lower()
                                              for p in cc.get("pages") or []]
                           for cc in ctrls):
                        rest.append(c)
                        continue
                self._create(c)
            if not rest or len(rest) == len(pending):
                for c in rest:
                    self._create(c)
                break
            pending = rest
        self.apply_tab_order()

    def container_for(self, parent: Optional[str]) -> QWidget:
        if parent:
            for k, w in self.containers.items():
                if k.lower() == parent.lower():
                    return w
        return self.client

    def apply_form(self):
        f = self.form
        pal = self.client.palette()
        bg = to_color(f.get("back_color")) or QColor(FORM_BG)
        pal.setColor(QPalette.Window, bg)
        fg = to_color(f.get("fore_color"))
        if fg:
            pal.setColor(QPalette.WindowText, fg)
        self.client.setPalette(pal)
        self.client.setAutoFillBackground(True)
        self.client.setFont(make_font(f))

    # -- creation -------------------------------------------------------------
    def _create(self, c: dict):
        t = c.get("type")
        name = c.get("name") or "Control"
        parent = self.container_for(c.get("parent"))
        firer = lambda ev, args=(), n=name: self.fire(n, ev, args)  # noqa: E731
        w: QWidget
        if t == "Label":
            w = _ClickLabel(parent, firer)
        elif t == "TextBox":
            if _bool(c.get("multi_line")):
                w = QPlainTextEdit(parent)
                w.textChanged.connect(lambda n=name, ww=None: self._text_changed(n))
            else:
                w = QLineEdit(parent)
                w.textEdited.connect(lambda _t, n=name: self._text_changed(n))
            self._install_keys(w, name)
        elif t == "CommandButton":
            w = QPushButton(parent)
            w.setAutoDefault(False)
            w.clicked.connect(lambda _=False, n=name: self.fire(n, "Click"))
        elif t == "ToggleButton":
            w = QPushButton(parent)
            w.setCheckable(True)
            w.setAutoDefault(False)
            w.toggled.connect(lambda on, n=name: self._bool_changed(n, on))
        elif t == "CheckBox":
            w = QCheckBox(parent)
            w.toggled.connect(lambda on, n=name: self._bool_changed(n, on))
        elif t == "OptionButton":
            w = QRadioButton(parent)
            w.setAutoExclusive(False)
            w.toggled.connect(lambda on, n=name: self._option_toggled(n, on))
        elif t == "ComboBox":
            w = QComboBox(parent)
            w.setEditable(int(c.get("style") or 0) != 2)
            w.currentIndexChanged.connect(lambda i, n=name: self._combo_index(n, i))
            w.activated.connect(lambda _i, n=name: self.fire(n, "Click"))
            if w.isEditable():
                w.lineEdit().textEdited.connect(lambda t_, n=name: self._combo_text(n, t_))
                self._install_keys(w.lineEdit(), name)
        elif t == "ListBox":
            if int(c.get("column_count") or 1) > 1:
                w = QTreeWidget(parent)
                w.setRootIsDecorated(False)
                w.setHeaderHidden(True)
                w.setIndentation(0)
                w.itemSelectionChanged.connect(lambda n=name: self._list_changed(n))
                w.itemDoubleClicked.connect(lambda *_a, n=name: self.fire(n, "DblClick"))
            else:
                w = QListWidget(parent)
                w.itemSelectionChanged.connect(lambda n=name: self._list_changed(n))
                w.itemDoubleClicked.connect(lambda *_a, n=name: self.fire(n, "DblClick"))
        elif t == "Frame":
            w = _ClickGroup(parent, firer)
            inner = QWidget(w)
            inner.setObjectName("frameClient")
            self.containers[name] = inner
        elif t == "MultiPage":
            w = QTabWidget(parent)
            for p in c.get("pages") or []:
                page = QWidget()
                page.setAutoFillBackground(True)
                w.addTab(page, str(p.get("caption", p.get("name", ""))))
                pname = p.get("name") or f"Page{w.count()}"
                self.containers[pname] = page
                self.page_owner[pname] = name
            w.currentChanged.connect(lambda i, n=name: self._index_changed(n, i))
        elif t == "TabStrip":
            w = QTabBar(parent)
            w.setDrawBase(True)
            for tb in c.get("tabs") or []:
                w.addTab(str(tb.get("caption", tb.get("name", ""))))
            w.currentChanged.connect(lambda i, n=name: self._index_changed(n, i))
        elif t == "SpinButton":
            w = _SpinButton(parent, lambda step, n=name: self._spin(n, step))
        elif t == "ScrollBar":
            w = QScrollBar(parent)
            w.valueChanged.connect(lambda v, n=name: self._scroll(n, v))
        elif t == "Image":
            w = _ClickLabel(parent, firer)
            w.setAlignment(Qt.AlignCenter)
        else:
            w = QLabel(parent)
        w.setObjectName(name)
        self.widgets[name] = w
        self.ctrl_by_name[name.lower()] = c
        if self.live:
            w.installEventFilter(self)
            if t in ("CommandButton", "CheckBox", "OptionButton", "ToggleButton", "ListBox",
                     "ComboBox", "ScrollBar", "MultiPage", "TabStrip"):
                self._watch_focus(w, name)
        else:
            self._design_mode(w)
        self.apply(c)
        w.show()

    def _design_mode(self, w: QWidget):
        for x in [w] + w.findChildren(QWidget):
            x.setAttribute(Qt.WA_TransparentForMouseEvents, True)
            x.setFocusPolicy(Qt.NoFocus)

    def _watch_focus(self, w, name):
        pass  # focus events reach eventFilter via installEventFilter

    def _install_keys(self, w: QWidget, name: str):
        w.setProperty("pe_ctrl", name)
        if self.live:
            w.installEventFilter(self)

    # -- property application -----------------------------------------------
    def geometry_px(self, c: dict) -> QRect:
        return QRect(px(c.get("left")), px(c.get("top")), max(1, px(c.get("width"))),
                     max(1, px(c.get("height"))))

    def apply(self, c: dict):
        name = c.get("name")
        w = self.widgets.get(name)
        if w is None:
            return
        self._updating += 1
        try:
            self._apply(c, w)
        finally:
            self._updating -= 1

    def _apply(self, c: dict, w: QWidget):
        t = c.get("type")
        w.setGeometry(self.geometry_px(c))
        w.setFont(make_font(self.form, c))
        w.setToolTip(str(c.get("control_tip_text") or ""))
        vis = _bool(c.get("visible", True))
        if self.live:
            w.setVisible(vis)
        else:
            w.setVisible(True)
            w.setGraphicsEffect(None)
        if self.live:
            w.setEnabled(_bool(c.get("enabled", True)))
        pal = w.palette()
        fg = to_color(c.get("fore_color"))
        bg = to_color(c.get("back_color"))
        if fg:
            for role in (QPalette.WindowText, QPalette.Text, QPalette.ButtonText):
                pal.setColor(role, fg)
        if bg and int(c.get("back_style", 1) or 0) != 0:
            for role in (QPalette.Window, QPalette.Base, QPalette.Button):
                pal.setColor(role, bg)
            w.setAutoFillBackground(t in ("Label", "Frame", "CheckBox", "OptionButton",
                                          "Image"))
        w.setPalette(pal)
        caption = str(c.get("caption") or "")
        align = {1: Qt.AlignLeft, 2: Qt.AlignHCenter, 3: Qt.AlignRight}.get(
            int(c.get("text_align") or 1), Qt.AlignLeft)
        locked = _bool(c.get("locked"))
        if t == "Label":
            w.setText(caption)
            w.setWordWrap(_bool(c.get("word_wrap")))
            w.setAlignment(align | Qt.AlignTop)
        elif t == "TextBox":
            val = c.get("value")
            text = "" if val is None else str(val)
            if isinstance(w, QPlainTextEdit):
                if w.toPlainText() != text:
                    w.setPlainText(text)
                w.setReadOnly(locked)
            else:
                if w.text() != text:
                    w.setText(text)
                w.setReadOnly(locked)
                w.setAlignment(align | Qt.AlignVCenter)
                pc = str(c.get("password_char") or "")
                w.setEchoMode(QLineEdit.Password if pc else QLineEdit.Normal)
                ml = int(c.get("max_length") or 0)
                w.setMaxLength(ml if ml > 0 else 32767)
        elif t in ("CommandButton", "ToggleButton"):
            w.setText(caption.replace("&&", "\0").replace("&", "").replace("\0", "&"))
            if t == "CommandButton":
                w.setDefault(_bool(c.get("default")))
            else:
                w.setChecked(_bool(c.get("value")))
        elif t in ("CheckBox", "OptionButton"):
            w.setText(caption)
            w.setChecked(_bool(c.get("value")))
        elif t == "ComboBox":
            items = ["\t".join(r) if len(r) > 1 else (r[0] if r else "")
                     for r in list_rows(c)]
            if [w.itemText(i) for i in range(w.count())] != items:
                w.clear()
                w.addItems(items)
            li = int(c.get("list_index", -1) if c.get("list_index") is not None else -1)
            val = c.get("value")
            if 0 <= li < w.count():
                w.setCurrentIndex(li)
            elif val not in (None, "") and str(val) in items:
                w.setCurrentIndex(items.index(str(val)))
            else:
                w.setCurrentIndex(-1)
            if w.isEditable():
                txt = "" if val is None else str(val)
                if w.currentText() != txt:
                    w.setEditText(txt)
                w.lineEdit().setReadOnly(locked)
        elif t == "ListBox":
            self._fill_list(c, w)
        elif t == "Frame":
            w.setTitle(caption)
            inner = self.containers.get(c.get("name"))
            if inner is not None:
                top = max(12, QFontInfo(w.font()).pixelSize() + 4) if caption else 4
                inner.setGeometry(2, top, max(1, w.width() - 4), max(1, w.height() - top - 2))
        elif t == "MultiPage":
            pages = c.get("pages") or []
            for i, p in enumerate(pages[:w.count()]):
                w.setTabText(i, str(p.get("caption", p.get("name", ""))))
            idx = int(c.get("value") or 0)
            if 0 <= idx < w.count():
                w.setCurrentIndex(idx)
        elif t == "TabStrip":
            idx = int(c.get("value") or 0)
            if 0 <= idx < w.count():
                w.setCurrentIndex(idx)
        elif t in ("SpinButton", "ScrollBar"):
            horiz = self._horizontal(c)
            if t == "ScrollBar":
                w.setOrientation(Qt.Horizontal if horiz else Qt.Vertical)
                lo, hi = int(c.get("min") or 0), int(c.get("max") or 0)
                w.setRange(min(lo, hi), max(lo, hi))
                w.setSingleStep(int(c.get("small_change") or 1))
                w.setPageStep(int(c.get("large_change") or 1))
                w.setValue(int(c.get("value") or 0))
            else:
                w.set_horizontal(horiz)
        elif t == "Image":
            data = c.get("picture")
            pm = QPixmap()
            if data:
                try:
                    pm.loadFromData(base64.b64decode(data))
                except (ValueError, TypeError):
                    pass
            if not pm.isNull():
                mode = int(c.get("picture_size_mode") or 0)
                if mode == 1:
                    pm = pm.scaled(w.size(), Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
                elif mode == 3:
                    pm = pm.scaled(w.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                w.setPixmap(pm)
            else:
                w.setPixmap(QPixmap())
            w.setFrameShape(QFrame.Box if not self.live or not data else QFrame.NoFrame)
        if self.live and not _bool(c.get("tab_stop", True)) and \
                t not in ("Label", "Frame", "Image"):
            w.setFocusPolicy(Qt.ClickFocus)

    @staticmethod
    def _horizontal(c):
        o = int(c.get("orientation", -1) if c.get("orientation") is not None else -1)
        if o == -1:
            return float(c.get("width") or 0) > float(c.get("height") or 0)
        return o == 1

    def _fill_list(self, c, w):
        rows = list_rows(c)
        ms = int(c.get("multi_select") or 0)
        w.setSelectionMode({0: QAbstractItemView.SingleSelection,
                            1: QAbstractItemView.MultiSelection,
                            2: QAbstractItemView.ExtendedSelection}.get(ms,
                                                                       QAbstractItemView.SingleSelection))
        selected = c.get("selected") or []
        li = c.get("list_index")
        li = -1 if li is None else int(li)
        if isinstance(w, QTreeWidget):
            n = max(int(c.get("column_count") or 1), max((len(r) for r in rows), default=1))
            w.clear()
            w.setColumnCount(n)
            for r in rows:
                QTreeWidgetItem(w, r + [""] * (n - len(r)))
            for i, cw in enumerate(column_widths_px(c, n)):
                w.setColumnWidth(i, cw if cw is not None else max(30, w.width() // n))
            items = [w.topLevelItem(i) for i in range(w.topLevelItemCount())]
        else:
            w.clear()
            for r in rows:
                QListWidgetItem("    ".join(r), w)
            items = [w.item(i) for i in range(w.count())]
        for i, it in enumerate(items):
            on = (i == li) if ms == 0 else (i < len(selected) and _bool(selected[i]))
            it.setSelected(on)
        if 0 <= li < len(items):
            if isinstance(w, QTreeWidget):
                w.setCurrentItem(items[li], 0, QItemSelectionModelNoUpdate())
            else:
                w.setCurrentRow(li, QItemSelectionModelNoUpdate())

    def apply_tab_order(self):
        if not self.live:
            return
        order = sorted((c for c in self.controls() if c.get("name") in self.widgets),
                       key=lambda c: int(c.get("tab_index") or 0))
        ws = []
        for c in order:
            w = self.widgets[c["name"]]
            if c.get("type") in ("Label", "Frame", "Image"):
                continue
            if c.get("type") == "ComboBox" and w.isEditable():
                w = w.lineEdit() or w
            ws.append(w)
        for a, b in zip(ws, ws[1:]):
            QWidget.setTabOrder(a, b)

    # -- user interaction -> runtime ----------------------------------------
    def _text_changed(self, name):
        if self._updating:
            return
        w = self.widgets.get(name)
        text = w.toPlainText() if isinstance(w, QPlainTextEdit) else w.text()
        self.user_set(name, "value", text)
        self.user_set(name, "text", text)
        self.fire(name, "Change")

    def _bool_changed(self, name, on):
        if self._updating:
            return
        self.user_set(name, "value", bool(on))
        self.fire(name, "Change")
        self.fire(name, "Click")
        self.fire(name, "AfterUpdate")

    def _option_toggled(self, name, on):
        if self._updating:
            return
        c = self.ctrl(name)
        if on:
            # exclusive within same container + group name
            grp = (c.get("parent") or "").lower(), (c.get("group_name") or "").lower()
            self.user_set(name, "value", True)
            for other in self.controls():
                if other is c or other.get("type") != "OptionButton":
                    continue
                og = (other.get("parent") or "").lower(), (other.get("group_name") or "").lower()
                ow = self.widgets.get(other.get("name"))
                if og == grp and ow is not None and ow.isChecked():
                    self._updating += 1
                    ow.setChecked(False)
                    self._updating -= 1
                    self.user_set(other["name"], "value", False)
                    self.fire(other["name"], "Change")
            self.fire(name, "Change")
            self.fire(name, "Click")
            self.fire(name, "AfterUpdate")
        else:
            # radio buttons cannot be unticked by clicking them again
            w = self.widgets.get(name)
            self._updating += 1
            w.setChecked(True)
            self._updating -= 1

    def _combo_index(self, name, i):
        if self._updating:
            return
        self.user_set(name, "list_index", i)
        self.fire(name, "Change")

    def _combo_text(self, name, text):
        if self._updating:
            return
        w = self.widgets[name]
        items = [w.itemText(i) for i in range(w.count())]
        self.user_set(name, "list_index", items.index(text) if text in items else -1)
        self.user_set(name, "value", text)
        self.user_set(name, "text", text)
        self.fire(name, "Change")

    def _list_changed(self, name):
        if self._updating:
            return
        c = self.ctrl(name)
        w = self.widgets[name]
        rows = list_rows(c)
        if isinstance(w, QTreeWidget):
            items = [w.topLevelItem(i) for i in range(w.topLevelItemCount())]
            cur = w.currentItem()
            cur_i = items.index(cur) if cur in items else -1
        else:
            items = [w.item(i) for i in range(w.count())]
            cur_i = w.currentRow()
        sel = [it.isSelected() for it in items]
        ms = int(c.get("multi_select") or 0)
        if ms == 0:
            idx = sel.index(True) if True in sel else -1
        else:
            idx = cur_i
        if ms:
            self.user_set(name, "selected", sel)
        self.user_set(name, "list_index", idx)
        self.fire(name, "Change")
        self.fire(name, "Click")
        self.fire(name, "AfterUpdate")

    def _index_changed(self, name, i):
        if self._updating:
            return
        self.user_set(name, "value", i)
        self.fire(name, "Change")
        self.fire(name, "Click")

    def _spin(self, name, step):
        c = self.ctrl(name)
        v = int(c.get("value") or 0)
        lo, hi = int(c.get("min") or 0), int(c.get("max") or 0)
        sc = int(c.get("small_change") or 1)
        nv = v + step * sc
        nv = max(min(lo, hi), min(max(lo, hi), nv))
        self.fire(name, "SpinUp" if step > 0 else "SpinDown")
        if nv != v:
            self.user_set(name, "value", nv)
            self.fire(name, "Change")

    def _scroll(self, name, v):
        if self._updating:
            return
        self.user_set(name, "value", v)
        self.fire(name, "Change")

    # -- keyboard / focus ---------------------------------------------------------
    def _owner(self, obj) -> Optional[str]:
        n = obj.property("pe_ctrl") if isinstance(obj, QObject) else None
        if n:
            return n
        if isinstance(obj, QWidget) and obj.objectName() in self.widgets:
            return obj.objectName()
        return None

    def eventFilter(self, obj, ev):
        if not self.live:
            return False
        et = ev.type()
        name = self._owner(obj)
        if name is None:
            return False
        c = self.ctrl(name)
        if c is None:
            return False
        if et == QEvent.KeyPress:
            text = ev.text()
            if text and len(text) == 1 and (text.isprintable() or text in "\b\r"):
                if ev.modifiers() & (Qt.ControlModifier | Qt.AltModifier):
                    return False
                res = self.fire(name, "KeyPress", (ord(text),))
                if res:
                    try:
                        ka = int(res[0])
                    except (TypeError, ValueError):
                        ka = ord(text)
                    if ka == 0:
                        return True
                    if ka != ord(text) and hasattr(obj, "insert"):
                        obj.insert(chr(ka))
                        return True
            return False
        if et == QEvent.FocusIn:
            self._entered_value[name] = c.get("value")
            self.fire(name, "Enter")
            return False
        if et == QEvent.FocusOut:
            reason = ev.reason()
            if reason in (Qt.ActiveWindowFocusReason, Qt.PopupFocusReason):
                return False
            w = self.widgets.get(name)
            changed = self._entered_value.get(name, c.get("value")) != c.get("value")
            if changed and c.get("type") in ("TextBox", "ComboBox", "ListBox"):
                res = self.fire(name, "BeforeUpdate", (False,))
                if res and _bool(res[0]):
                    QTimer.singleShot(0, lambda ww=obj: ww.setFocus())
                    return False
                self.fire(name, "AfterUpdate")
                self._entered_value[name] = c.get("value")
            res = self.fire(name, "Exit", (False,))
            if res and _bool(res[0]):
                QTimer.singleShot(0, lambda ww=obj: ww.setFocus())
            return False
        return False


def QItemSelectionModelNoUpdate():
    from PySide6.QtCore import QItemSelectionModel
    return QItemSelectionModel.NoUpdate


# --------------------------------------------------------------------------
# Live form window
# --------------------------------------------------------------------------
class QtFormView(QDialog):
    """forms.FormView implemented as a QDialog."""

    def __init__(self, instance, host=None, parent=None):
        if parent is None and host is not None:
            parent = getattr(host, "win", None)
        super().__init__(parent)
        self.instance = instance
        self.host = host
        self._loop: Optional[QEventLoop] = None
        self._allow_close = False
        self._closed = False
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint | Qt.WindowCloseButtonHint |
                            Qt.CustomizeWindowHint)
        self.setAttribute(Qt.WA_DeleteOnClose, False)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.client = _Client(self, self._fire)
        lay.addWidget(self.client)
        self.form = instance.form
        self.renderer = FormRenderer(self.client, self.form, live=True, fire=self._fire,
                                     user_set=self._user_set)
        self._apply_form_props()

    # -- bridge ---------------------------------------------------------------
    @property
    def name(self):
        return getattr(self.instance, "name", "UserForm")

    def _fire(self, ctrl, event, args=()):
        try:
            return self.instance.fire(ctrl, event, tuple(args))
        except Exception as e:  # VBAError or interpreter bug: never crash the GUI
            self._report(e)
            return None

    def _user_set(self, ctrl, prop, value):
        try:
            self.instance.user_set(ctrl, prop, value)
        except Exception as e:
            self._report(e)

    def _report(self, e):
        win = getattr(self.host, "win", None)
        if win is not None and hasattr(win, "_handle_vba_exception"):
            win._handle_vba_exception(e)
        else:
            log.exception("UserForm event failed: %s", e)

    def _apply_form_props(self):
        f = self.form
        self.setWindowTitle(str(f.get("caption") or self.name))
        w = max(40, px(float(f.get("width") or 240) - FORM_BORDER_W))
        h = max(30, px(float(f.get("height") or 180) - FORM_BORDER_H))
        self.setFixedSize(w, h)
        self.renderer.apply_form()

    # -- FormView API ---------------------------------------------------------
    def show(self, modal: bool = False):  # noqa: D401 - FormView.show
        self._closed = False
        self._position()
        self.setWindowModality(Qt.ApplicationModal if modal else Qt.NonModal)
        QDialog.show(self)
        self.raise_()
        self.activateWindow()
        self._focus_first()
        if modal:
            self._loop = QEventLoop(self)
            self._loop.exec()
            self._loop = None

    def _focus_first(self):
        order = sorted((c for c in self.form.get("controls") or []
                        if c.get("type") not in ("Label", "Frame", "Image")
                        and _bool(c.get("visible", True)) and _bool(c.get("enabled", True))
                        and _bool(c.get("tab_stop", True))),
                       key=lambda c: int(c.get("tab_index") or 0))
        if order:
            w = self.renderer.widgets.get(order[0]["name"])
            if w is not None:
                w.setFocus(Qt.OtherFocusReason)

    def _position(self):
        mode = int(self.form.get("start_up_position", 1) or 0)
        if mode == 0:
            left = self.form.get("left")
            top = self.form.get("top")
            if left is not None and top is not None:
                self.move(px(left), px(top))
            return
        owner = self.parentWidget()
        if mode == 1 and owner is not None and owner.isVisible():
            geo = owner.frameGeometry()
        else:
            scr = (owner.screen() if owner is not None else None) or QGuiApplication.primaryScreen()
            geo = scr.availableGeometry() if scr else QRect(0, 0, 1024, 768)
        self.move(geo.center() - QPoint(self.width() // 2, self.height() // 2))

    def _end_loop(self):
        if self._loop is not None and self._loop.isRunning():
            self._loop.quit()

    def hide(self):
        QDialog.hide(self)
        self._end_loop()

    def close(self):
        self._allow_close = True
        self._closed = True
        try:
            QDialog.close(self)
        finally:
            self._allow_close = False
        self._end_loop()
        return True

    def set_prop(self, ctrl_name, prop, value):
        prop = norm_prop(prop)
        if ctrl_name is None or ctrl_name == "" or \
                str(ctrl_name).lower() == str(self.name).lower():
            self.form[prop] = value
            if prop in ("caption", "width", "height", "back_color", "fore_color", "font_name",
                        "font_size"):
                self._apply_form_props()
            if prop in ("font_name", "font_size"):
                for c in self.form.get("controls") or []:
                    self.renderer.apply(c)
            return
        c = self.renderer.ctrl(ctrl_name)
        if c is None:
            return
        c[prop] = value
        if prop in ("list",) and c.get("type") == "ComboBox":
            pass
        if prop in ("multi_line", "style", "column_count", "pages", "tabs", "parent", "type"):
            self.renderer.build()  # structural change
            return
        self.renderer.apply(c)
        if prop == "tab_index":
            self.renderer.apply_tab_order()

    def add_control(self, ctrl: dict):
        full = forms.normalize_form({"controls": [ctrl]})["controls"][0]
        ctrls = self.form.setdefault("controls", [])
        if not any(c.get("name", "").lower() == full["name"].lower() for c in ctrls):
            ctrls.append(full)
        else:
            full = next(c for c in ctrls if c.get("name", "").lower() == full["name"].lower())
        self.renderer._create(full)
        self.renderer.apply_tab_order()

    def remove_control(self, ctrl_name: str):
        ctrls = self.form.get("controls") or []
        self.form["controls"] = [c for c in ctrls
                                 if c.get("name", "").lower() != ctrl_name.lower()]
        self.renderer.build()

    def set_focus(self, ctrl_name: str):
        w = self.renderer.widgets.get(ctrl_name)
        if w is None:
            for k, v in self.renderer.widgets.items():
                if k.lower() == str(ctrl_name).lower():
                    w = v
        if w is not None:
            if isinstance(w, QComboBox) and w.isEditable():
                w.lineEdit().setFocus(Qt.OtherFocusReason)
            else:
                w.setFocus(Qt.OtherFocusReason)

    def repaint(self):
        QWidget.repaint(self)
        QApplication.processEvents()

    # -- Qt events ------------------------------------------------------------
    def closeEvent(self, e):
        if self._allow_close:
            e.accept()
            return
        e.ignore()
        try:
            self.instance.request_close(0)
        except Exception as ex:
            self._report(ex)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            for c in self.form.get("controls") or []:
                if c.get("type") == "CommandButton" and _bool(c.get("cancel")):
                    w = self.renderer.widgets.get(c["name"])
                    if w is not None and w.isEnabled() and w.isVisible():
                        w.setFocus(Qt.OtherFocusReason)
                        w.click()
                    return
            return  # Esc does nothing without a Cancel button
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            fw = self.focusWidget()
            if isinstance(fw, QPlainTextEdit):
                return
            for c in self.form.get("controls") or []:
                if c.get("type") == "CommandButton" and _bool(c.get("default")):
                    w = self.renderer.widgets.get(c["name"])
                    if w is not None and w.isEnabled() and w.isVisible():
                        w.setFocus(Qt.OtherFocusReason)
                        w.click()
                    return
            return
        super().keyPressEvent(e)

    def reject(self):
        pass  # handled in keyPressEvent / closeEvent


class PreviewInstance:
    """Minimal stand-in used to preview a form when the interpreter has no
    UserForm support: values change, events are ignored, X closes."""

    def __init__(self, name, form):
        self.name = name
        self.form = forms.normalize_form(form)
        self.view = None
        self.events: List[tuple] = []

    def fire(self, ctrl, event, args=()):
        self.events.append((ctrl, event, tuple(args)))
        return list(args)

    def user_set(self, ctrl, prop, value):
        self.events.append(("set", ctrl, prop, value))

    def request_close(self, close_mode=0):
        if self.view is not None:
            self.view.close()

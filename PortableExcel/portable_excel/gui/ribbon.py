"""Excel-like title strip (quick access toolbar) and ribbon."""
from __future__ import annotations

from typing import Dict, List, Optional

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QAction, QColor, QPainter
from PySide6.QtWidgets import (QFrame, QHBoxLayout, QLabel, QSizePolicy, QStackedWidget,
                               QTabBar, QToolButton, QVBoxLayout, QWidget)

from . import icons


class TitleStrip(QWidget):
    """Green strip: app icon, Save/Undo/Redo, centered "Book1 - PortableExcel"."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("titleStrip")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setFixedHeight(32)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(8, 0, 8, 0)
        lay.setSpacing(2)
        logo = QLabel()
        logo.setPixmap(icons.app_pixmap(20))
        lay.addWidget(logo)
        lay.addSpacing(6)
        self.qat = QHBoxLayout()
        self.qat.setSpacing(0)
        lay.addLayout(self.qat)
        lay.addStretch(1)
        self.title = QLabel("Book1 - PortableExcel")
        self.title.setObjectName("titleLabel")
        self.title.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.title)
        lay.addStretch(1)
        self.right = QLabel("")
        self.right.setObjectName("titleRight")
        lay.addWidget(self.right)

    def add_action(self, action: QAction, icon_name: str):
        b = QToolButton(self)
        b.setObjectName("qatButton")
        b.setDefaultAction(action)
        b.setIcon(icons.icon(icon_name))
        b.setIconSize(QSize(16, 16))
        b.setAutoRaise(True)
        b.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.qat.addWidget(b)
        return b

    def set_title(self, text: str):
        self.title.setText(text)


class RibbonGroup(QFrame):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setObjectName("ribbonGroup")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(4, 2, 4, 0)
        outer.setSpacing(0)
        self.body = QHBoxLayout()
        self.body.setSpacing(2)
        self.body.setContentsMargins(0, 0, 0, 0)
        outer.addLayout(self.body, 1)
        self.caption = QLabel(title)
        self.caption.setObjectName("groupCaption")
        self.caption.setAlignment(Qt.AlignCenter)
        outer.addWidget(self.caption)
        self.buttons: Dict[str, QToolButton] = {}

    def _button(self, action: QAction, icon_name: Optional[str], large: bool,
                text_beside=True, menu=None) -> QToolButton:
        b = QToolButton(self)
        b.setDefaultAction(action)
        if icon_name:
            b.setIcon(icons.icon(icon_name))
        b.setAutoRaise(True)
        if large:
            b.setObjectName("largeButton")
            b.setIconSize(QSize(32, 32))
            b.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            b.setMinimumWidth(48)
            b.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        else:
            b.setObjectName("smallButton")
            b.setIconSize(QSize(16, 16))
            b.setToolButtonStyle(Qt.ToolButtonTextBesideIcon if text_beside
                                 else Qt.ToolButtonIconOnly)
        if menu is not None:
            b.setMenu(menu)
            b.setPopupMode(QToolButton.MenuButtonPopup if action.isEnabled() and
                           action.data() != "menu-only" else QToolButton.InstantPopup)
        self.buttons[action.text()] = b
        return b

    def add_large(self, action: QAction, icon_name: str, menu=None) -> QToolButton:
        b = self._button(action, icon_name, True, menu=menu)
        self.body.addWidget(b)
        return b

    def add_small_column(self, items) -> QWidget:
        """Column of up to three small buttons; items = [(action, icon) | QWidget]."""
        w = QWidget(self)
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        for it in items:
            if isinstance(it, QWidget):
                v.addWidget(it)
            else:
                action, icon_name = it[0], it[1]
                menu = it[2] if len(it) > 2 else None
                v.addWidget(self._button(action, icon_name, False, menu=menu))
        v.addStretch(1)
        self.body.addWidget(w)
        return w

    def add_rows(self, rows) -> QWidget:
        """Rows of compact controls; each row = list of (action, icon) | QWidget."""
        w = QWidget(self)
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 2, 0, 0)
        v.setSpacing(3)
        for row in rows:
            h = QHBoxLayout()
            h.setSpacing(1)
            for it in row:
                if isinstance(it, QWidget):
                    h.addWidget(it)
                elif it == "|":
                    sep = QFrame()
                    sep.setFrameShape(QFrame.VLine)
                    sep.setObjectName("rowSep")
                    h.addWidget(sep)
                else:
                    action, icon_name = it[0], it[1]
                    menu = it[2] if len(it) > 2 else None
                    h.addWidget(self._button(action, icon_name, False, text_beside=False,
                                             menu=menu))
            h.addStretch(1)
            v.addLayout(h)
        v.addStretch(1)
        self.body.addWidget(w)
        return w

    def add_widget(self, w: QWidget):
        self.body.addWidget(w)
        return w


class RibbonPage(QWidget):
    def __init__(self, name: str, parent=None):
        super().__init__(parent)
        self.name = name
        self.lay = QHBoxLayout(self)
        self.lay.setContentsMargins(4, 2, 4, 2)
        self.lay.setSpacing(0)
        self.lay.addStretch(1)
        self.groups: Dict[str, RibbonGroup] = {}

    def add_group(self, title: str) -> RibbonGroup:
        g = RibbonGroup(title, self)
        self.lay.insertWidget(self.lay.count() - 1, g)
        sep = QFrame(self)
        sep.setObjectName("groupSep")
        sep.setFrameShape(QFrame.VLine)
        self.lay.insertWidget(self.lay.count() - 1, sep)
        self.groups[title] = g
        return g


class Ribbon(QWidget):
    file_clicked = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("ribbon")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        self.tab_strip = QWidget(self)
        self.tab_strip.setObjectName("ribbonTabStrip")
        self.tab_strip.setAttribute(Qt.WA_StyledBackground, True)
        ts = QHBoxLayout(self.tab_strip)
        ts.setContentsMargins(0, 0, 0, 0)
        ts.setSpacing(0)
        self.tabs = QTabBar(self.tab_strip)
        self.tabs.setObjectName("ribbonTabs")
        self.tabs.setDrawBase(False)
        self.tabs.setExpanding(False)
        self.tabs.setDocumentMode(True)
        ts.addWidget(self.tabs)
        ts.addStretch(1)
        lay.addWidget(self.tab_strip)
        self.stack = QStackedWidget(self)
        self.stack.setObjectName("ribbonBody")
        self.stack.setFixedHeight(96)
        lay.addWidget(self.stack)
        self.pages: Dict[str, RibbonPage] = {}
        self._names: List[str] = []
        self._last = 1
        self.tabs.addTab("File")
        self._names.append("File")
        self.stack.addWidget(QWidget())
        self.tabs.currentChanged.connect(self._tab_changed)
        self.tabs.tabBarDoubleClicked.connect(self._toggle_collapse)
        self.collapsed = False

    def add_page(self, name: str) -> RibbonPage:
        page = RibbonPage(name)
        self.pages[name] = page
        self._names.append(name)
        self.tabs.addTab(name)
        self.stack.addWidget(page)
        if len(self._names) == 2:
            self.tabs.setCurrentIndex(1)
        return page

    def _tab_changed(self, i: int):
        if i == 0:
            self.tabs.blockSignals(True)
            self.tabs.setCurrentIndex(self._last)
            self.tabs.blockSignals(False)
            self.file_clicked.emit()
            return
        self._last = i
        self.stack.setCurrentIndex(i)
        if self.collapsed:
            self.stack.show()

    def _toggle_collapse(self, i: int):
        if i <= 0:
            return
        self.collapsed = not self.collapsed
        self.stack.setVisible(not self.collapsed)

    def show_page(self, name: str):
        i = self._names.index(name)
        self.tabs.setCurrentIndex(i)

    def current_page(self) -> str:
        return self._names[self.tabs.currentIndex()]

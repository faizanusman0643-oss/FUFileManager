"""File tab 'backstage' page (New / Open / Info / Save ...)."""
from __future__ import annotations

import os

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtWidgets import (QButtonGroup, QHBoxLayout, QLabel, QListWidget,
                               QListWidgetItem, QPushButton, QStackedWidget, QVBoxLayout,
                               QWidget)

from . import config, icons


class Backstage(QWidget):
    back = Signal()

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setObjectName("backstage")
        self.setAttribute(Qt.WA_StyledBackground, True)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        nav = QWidget()
        nav.setObjectName("backstageNav")
        nav.setAttribute(Qt.WA_StyledBackground, True)
        nav.setFixedWidth(200)
        nv = QVBoxLayout(nav)
        nv.setContentsMargins(0, 10, 0, 10)
        nv.setSpacing(0)
        self.group = QButtonGroup(self)
        self.pages = QStackedWidget()

        def item(text, icon=None, slot=None, page=None):
            b = QPushButton(("   " if icon is None else "") + text)
            b.setObjectName("bsItem")
            if icon:
                b.setIcon(icons.icon(icon))
                b.setIconSize(QSize(20, 20))
            if page is not None:
                b.setCheckable(True)
                self.group.addButton(b)
                b.clicked.connect(lambda: self.show_page(page))
            if slot is not None:
                b.clicked.connect(slot)
            nv.addWidget(b)
            return b

        item("", "back", self.back.emit).setToolTip("Back")
        self.b_new = item("New", None, page="new")
        self.b_open = item("Open", None, page="open")
        self.b_info = item("Info", None, page="info")
        item("Save", None, lambda: self._run(self.win.save))
        item("Save As", None, lambda: self._run(self.win.save_as))
        item("Save with password...", None, lambda: self._run(self.win.save_with_password))
        item("Print...", None, lambda: self._run(lambda: self.win.print_dialog(preview=True)))
        item("Export PDF...", None, lambda: self._run(self.win.export_pdf_dialog))
        nv.addSpacing(16)
        item("Change app password...", None, lambda: self._run(self.win.change_app_password))
        item("Macro security...", None, lambda: self._run(self.win.macro_security))
        nv.addStretch(1)
        item("About", None, lambda: self._run(self.win.about))
        item("Close", None, lambda: self._run(self.win.close))
        item("Exit", None, lambda: self._run(self.win.exit_app))
        lay.addWidget(nav)
        lay.addWidget(self.pages, 1)
        self._page_widgets = {"new": self._build_new(), "open": self._build_open(),
                              "info": self._build_info()}
        for w in self._page_widgets.values():
            self.pages.addWidget(w)

    def _run(self, fn):
        self.back.emit()
        fn()

    def _title(self, text):
        t = QLabel(text)
        t.setObjectName("bsTitle")
        return t

    def _build_new(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(40, 30, 40, 30)
        v.addWidget(self._title("New"))
        b = QPushButton("  Blank workbook")
        b.setObjectName("bsTile")
        b.setIcon(icons.icon("new"))
        b.setIconSize(QSize(48, 48))
        b.setFixedSize(240, 90)
        b.clicked.connect(lambda: self._run(self.win.new_workbook))
        v.addWidget(b)
        v.addStretch(1)
        return w

    def _build_open(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(40, 30, 40, 30)
        v.addWidget(self._title("Open"))
        b = QPushButton("  Browse...")
        b.setObjectName("bsTile")
        b.setIcon(icons.icon("open"))
        b.setIconSize(QSize(32, 32))
        b.setFixedSize(240, 60)
        b.clicked.connect(lambda: self._run(self.win.open_file))
        v.addWidget(b)
        h = QLabel("Recent")
        h.setObjectName("bsHeading")
        v.addSpacing(12)
        v.addWidget(h)
        self.recent = QListWidget()
        self.recent.setObjectName("recentList")
        self.recent.itemActivated.connect(self._open_recent)
        self.recent.itemClicked.connect(self._open_recent)
        v.addWidget(self.recent, 1)
        return w

    def _open_recent(self, item):
        path = item.data(Qt.UserRole)
        if path:
            self._run(lambda: self.win.open_file(path))

    def _build_info(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(40, 30, 40, 30)
        v.addWidget(self._title("Info"))
        self.info_label = QLabel("")
        self.info_label.setWordWrap(True)
        self.info_label.setTextFormat(Qt.RichText)
        v.addWidget(self.info_label)
        row = QHBoxLayout()
        for text, icon, slot in (("Protect Workbook\n(encrypt with password)", "lock",
                                  self.win.save_with_password),
                                 ("Change app\npassword", "key", self.win.change_app_password),
                                 ("Macro\nsecurity", "security", self.win.macro_security)):
            b = QPushButton(text)
            b.setObjectName("bsTile")
            b.setIcon(icons.icon(icon))
            b.setIconSize(QSize(32, 32))
            b.setFixedSize(200, 70)
            b.clicked.connect(lambda _=False, s=slot: self._run(s))
            row.addWidget(b)
        row.addStretch(1)
        v.addLayout(row)
        v.addStretch(1)
        return w

    def refresh(self):
        self.recent.clear()
        for p in config.recent_files():
            it = QListWidgetItem(icons.icon("sheet"), f"{os.path.basename(p)}\n{os.path.dirname(p)}")
            it.setData(Qt.UserRole, p)
            self.recent.addItem(it)
        if not self.recent.count():
            self.recent.addItem("No recent workbooks")
        b = self.win.book
        macros = len(b.vba_modules)
        self.info_label.setText(
            f"<p><b>{self.win.book_title()}</b><br>{b.path or 'Not saved yet'}</p>"
            f"<p>Sheets: {len(b.sheets)} &nbsp;&nbsp; VBA modules: {macros} &nbsp;&nbsp; "
            f"Macros: {'enabled' if self.win.macros_enabled else 'disabled'}<br>"
            f"Encryption on save: {'yes' if self.win.save_password else 'no'}</p>")

    def show_page(self, name):
        self.refresh()
        self.pages.setCurrentWidget(self._page_widgets[name])
        {"new": self.b_new, "open": self.b_open, "info": self.b_info}[name].setChecked(True)

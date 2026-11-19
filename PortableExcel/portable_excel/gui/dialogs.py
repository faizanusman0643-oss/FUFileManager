"""Dialogs: Format Cells, Find & Replace, Insert Function, Macros, Macro
security, passwords, save warnings, colour menu."""
from __future__ import annotations

from typing import Callable, Dict, List, Optional

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QColor, QFont, QIcon, QPixmap
from PySide6.QtWidgets import (QButtonGroup, QCheckBox, QColorDialog, QComboBox,
                               QDialog, QDialogButtonBox, QFontComboBox, QFormLayout,
                               QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QMenu, QPushButton,
                               QRadioButton, QSpinBox, QTabWidget, QToolButton,
                               QVBoxLayout, QWidget, QWidgetAction)

from . import icons

THEME_COLORS = [
    "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
    "#5B9BD5", "#70AD47",
    "#F2F2F2", "#7F7F7F", "#D0CECE", "#D6DCE4", "#D9E2F3", "#FBE5D5", "#EDEDED", "#FFF2CC",
    "#DEEBF6", "#E2EFD9",
    "#D8D8D8", "#595959", "#AEAAAA", "#ADB9CA", "#B4C6E7", "#F7CBAC", "#DBDBDB", "#FEE599",
    "#BDD7EE", "#C5E0B3",
]
STANDARD_COLORS = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050",
                   "#00B0F0", "#0070C0", "#002060", "#7030A0"]

NUMBER_FORMATS = [  # (label shown in the ribbon combo, format code or None for General)
    ("General", None),
    ("Number", "0.00"),
    ("Comma", "#,##0.00"),
    ("Currency", "$#,##0.00"),
    ("Percent", "0%"),
    ("Short Date", "m/d/yyyy"),
    ("Long Date", "dddd, mmmm d, yyyy"),
    ("Time", "h:mm:ss AM/PM"),
    ("Scientific", "0.00E+00"),
    ("Text", "@"),
]


def swatch_icon(color: Optional[str], size=16) -> QIcon:
    pm = QPixmap(size, size)
    pm.fill(QColor(color) if color else Qt.white)
    if not color:
        from PySide6.QtGui import QPainter, QPen
        p = QPainter(pm)
        p.setPen(QPen(QColor("#C00000"), 2))
        p.drawLine(0, size, size, 0)
        p.end()
    return QIcon(pm)


class ColorMenu(QMenu):
    """Drop-down palette like Excel's font/fill colour pickers."""
    picked = Signal(object)  # "#RRGGBB" or None

    def __init__(self, none_label="No Fill", parent=None):
        super().__init__(parent)
        self.none_label = none_label
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(6, 6, 6, 6)
        none_btn = QPushButton(none_label)
        none_btn.clicked.connect(lambda: self._pick(None))
        v.addWidget(none_btn)
        v.addWidget(QLabel("Theme Colors"))
        g = QGridLayout()
        g.setSpacing(2)
        for i, col in enumerate(THEME_COLORS):
            g.addWidget(self._sw(col), i // 10, i % 10)
        v.addLayout(g)
        v.addWidget(QLabel("Standard Colors"))
        g2 = QGridLayout()
        g2.setSpacing(2)
        for i, col in enumerate(STANDARD_COLORS):
            g2.addWidget(self._sw(col), 0, i)
        v.addLayout(g2)
        more = QPushButton("More Colors...")
        more.clicked.connect(self._more)
        v.addWidget(more)
        act = QWidgetAction(self)
        act.setDefaultWidget(w)
        self.addAction(act)

    def _sw(self, col):
        b = QToolButton()
        b.setFixedSize(18, 18)
        b.setStyleSheet(f"QToolButton{{background:{col}; border:1px solid #C6C6C6;}}"
                        f"QToolButton:hover{{border:1px solid #E86E00;}}")
        b.setToolTip(col)
        b.clicked.connect(lambda _=False, c=col: self._pick(c))
        return b

    def _pick(self, col):
        self.hide()
        self.picked.emit(col)

    def _more(self):
        self.hide()
        c = QColorDialog.getColor(Qt.white, self.parentWidget(), "Colors")
        if c.isValid():
            self.picked.emit(c.name().upper())


class ColorButton(QPushButton):
    def __init__(self, color: Optional[str], none_label="Automatic", parent=None):
        super().__init__(parent)
        self.color = color
        self.menu_ = ColorMenu(none_label, self)
        self.menu_.picked.connect(self.set_color)
        self.setMenu(self.menu_)
        self.set_color(color)

    def set_color(self, color):
        self.color = color
        self.setIcon(swatch_icon(color, 14))
        self.setText(color or self.menu_.none_label)


# --------------------------------------------------------------------------
# Format Cells
# --------------------------------------------------------------------------
class FormatCellsDialog(QDialog):
    CATEGORIES = ["General", "Number", "Currency", "Date", "Time", "Percentage", "Fraction",
                  "Scientific", "Text", "Custom"]

    def __init__(self, style: dict, preview: Callable[[Optional[str]], str], parent=None,
                 tab: int = 0):
        super().__init__(parent)
        self.setWindowTitle("Format Cells")
        self.style0 = dict(style or {})
        self.preview_fn = preview
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        lay.addWidget(self.tabs)
        self._build_number()
        self._build_alignment()
        self._build_font()
        self._build_fill()
        self.tabs.setCurrentIndex(tab)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.resize(560, 440)

    # Number ----------------------------------------------------------------
    def _build_number(self):
        w = QWidget()
        h = QHBoxLayout(w)
        left = QVBoxLayout()
        left.addWidget(QLabel("Category:"))
        self.cat = QListWidget()
        self.cat.addItems(self.CATEGORIES)
        self.cat.setMaximumWidth(150)
        left.addWidget(self.cat)
        h.addLayout(left)
        right = QVBoxLayout()
        box = QGroupBox("Sample")
        bl = QVBoxLayout(box)
        self.sample = QLabel("")
        self.sample.setMinimumHeight(22)
        bl.addWidget(self.sample)
        right.addWidget(box)
        form = QFormLayout()
        self.decimals = QSpinBox()
        self.decimals.setRange(0, 30)
        self.decimals.setValue(2)
        self.thousands = QCheckBox("Use 1000 Separator (,)")
        self.date_type = QComboBox()
        self.date_type.addItems(["m/d/yyyy", "d-mmm-yy", "dddd, mmmm d, yyyy", "yyyy-mm-dd",
                                 "mmm-yy"])
        self.time_type = QComboBox()
        self.time_type.addItems(["h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yyyy h:mm"])
        self.code = QLineEdit()
        form.addRow("Decimal places:", self.decimals)
        form.addRow("", self.thousands)
        form.addRow("Date type:", self.date_type)
        form.addRow("Time type:", self.time_type)
        form.addRow("Type:", self.code)
        right.addLayout(form)
        self.num_desc = QLabel("")
        self.num_desc.setWordWrap(True)
        self.num_desc.setStyleSheet("color:#555;")
        right.addWidget(self.num_desc)
        right.addStretch(1)
        h.addLayout(right, 1)
        self.tabs.addTab(w, "Number")
        fmt = self.style0.get("number_format")
        self.cat.setCurrentRow(self._category_for(fmt))
        self.code.setText(fmt or "General")
        if fmt and "." in fmt:
            self.decimals.setValue(len(fmt.split(".")[1].split("%")[0].split("E")[0]
                                       .rstrip(")_ ")) if fmt else 2)
        self.thousands.setChecked(bool(fmt and "#,##" in fmt))
        self.cat.currentRowChanged.connect(self._cat_changed)
        for sig in (self.decimals.valueChanged, self.thousands.toggled,
                    self.date_type.currentIndexChanged, self.time_type.currentIndexChanged):
            sig.connect(self._rebuild_code)
        self.code.textChanged.connect(self._update_sample)
        self._cat_changed(self.cat.currentRow(), rebuild=False)
        self._update_sample()

    @staticmethod
    def _category_for(fmt):
        if not fmt or fmt == "General":
            return 0
        if fmt == "@":
            return 8
        if "%" in fmt:
            return 5
        if "E+" in fmt.upper():
            return 7
        if "?/?" in fmt:
            return 6
        low = fmt.lower()
        if any(t in low for t in ("yy", "dd", "mmm")) or low.startswith(("m/", "d")):
            return 3
        if "h:" in low or ":ss" in low:
            return 4
        if "$" in fmt or "€" in fmt:
            return 2
        if set(fmt) <= set("0#,.") and fmt:
            return 1
        return 9

    def _cat_changed(self, row, rebuild=True):
        cat = self.CATEGORIES[row] if row >= 0 else "General"
        self.decimals.setEnabled(cat in ("Number", "Currency", "Percentage", "Scientific"))
        self.thousands.setEnabled(cat == "Number")
        self.date_type.setEnabled(cat == "Date")
        self.time_type.setEnabled(cat == "Time")
        self.code.setReadOnly(cat != "Custom")
        desc = {
            "General": "General format cells have no specific number format.",
            "Number": "Number is used for general display of numbers.",
            "Currency": "Currency formats are used for general monetary values.",
            "Date": "Date formats display date serial numbers as date values.",
            "Time": "Time formats display date serial numbers as time values.",
            "Percentage": "Percentage formats multiply the cell value by 100 and display "
                          "the result with a percent symbol.",
            "Text": "Text format cells are treated as text even when a number is in the cell.",
            "Custom": "Type the number format code, using one of the existing codes as a "
                      "starting point.",
        }.get(cat, "")
        self.num_desc.setText(desc)
        if rebuild:
            self._rebuild_code()

    def _rebuild_code(self, *_):
        cat = self.CATEGORIES[self.cat.currentRow()]
        d = self.decimals.value()
        dec = ("." + "0" * d) if d else ""
        code = {
            "General": "General",
            "Number": ("#,##0" if self.thousands.isChecked() else "0") + dec,
            "Currency": "$#,##0" + dec,
            "Date": self.date_type.currentText(),
            "Time": self.time_type.currentText(),
            "Percentage": "0" + dec + "%",
            "Fraction": "# ?/?",
            "Scientific": "0" + dec + "E+00",
            "Text": "@",
        }.get(cat)
        if code is not None:
            self.code.setText(code)

    def _update_sample(self, *_):
        try:
            self.sample.setText(self.preview_fn(self.number_format()))
        except Exception:
            self.sample.setText("")

    def number_format(self) -> Optional[str]:
        t = self.code.text().strip()
        return None if not t or t == "General" else t

    # Alignment ---------------------------------------------------------------
    def _build_alignment(self):
        w = QWidget()
        f = QFormLayout(w)
        self.halign = QComboBox()
        self.halign.addItems(["General", "Left", "Center", "Right"])
        a = self.style0.get("align")
        self.halign.setCurrentIndex({"left": 1, "center": 2, "right": 3}.get(a, 0))
        self.wrap = QCheckBox("Wrap text")
        self.wrap.setChecked(bool(self.style0.get("wrap")))
        f.addRow("Horizontal:", self.halign)
        f.addRow("Text control:", self.wrap)
        self.tabs.addTab(w, "Alignment")

    # Font --------------------------------------------------------------------
    def _build_font(self):
        w = QWidget()
        f = QFormLayout(w)
        self.font_box = QFontComboBox()
        if self.style0.get("font_name"):
            self.font_box.setCurrentFont(QFont(self.style0["font_name"]))
        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems([str(s) for s in (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26,
                                                  28, 36, 48, 72)])
        self.size_box.setCurrentText(str(self.style0.get("font_size", 11)).rstrip("0").rstrip(".")
                                     if self.style0.get("font_size") else "11")
        self.bold = QCheckBox("Bold")
        self.bold.setChecked(bool(self.style0.get("bold")))
        self.italic = QCheckBox("Italic")
        self.italic.setChecked(bool(self.style0.get("italic")))
        self.underline = QCheckBox("Underline")
        self.underline.setChecked(bool(self.style0.get("underline")))
        self.font_color = ColorButton(self.style0.get("color"), "Automatic")
        st = QHBoxLayout()
        for cb in (self.bold, self.italic, self.underline):
            st.addWidget(cb)
        st.addStretch(1)
        f.addRow("Font:", self.font_box)
        f.addRow("Size:", self.size_box)
        f.addRow("Font style:", st)
        f.addRow("Color:", self.font_color)
        self.font_preview = QLabel("AaBbCcYyZz")
        self.font_preview.setMinimumHeight(48)
        self.font_preview.setAlignment(Qt.AlignCenter)
        self.font_preview.setStyleSheet("background:white; border:1px solid #D4D4D4;")
        f.addRow("Preview:", self.font_preview)
        for sig in (self.font_box.currentFontChanged, self.size_box.currentTextChanged,
                    self.bold.toggled, self.italic.toggled, self.underline.toggled):
            sig.connect(self._font_preview)
        self.font_color.menu_.picked.connect(self._font_preview)
        self._font_changed = False
        self._size_changed = False
        self.font_box.currentFontChanged.connect(lambda *_: setattr(self, "_font_changed", True))
        self.size_box.currentTextChanged.connect(lambda *_: setattr(self, "_size_changed", True))
        self._font_preview()
        self.tabs.addTab(w, "Font")

    def _font_preview(self, *_):
        f = QFont(self.font_box.currentFont())
        try:
            f.setPointSizeF(float(self.size_box.currentText()))
        except ValueError:
            pass
        f.setBold(self.bold.isChecked())
        f.setItalic(self.italic.isChecked())
        f.setUnderline(self.underline.isChecked())
        self.font_preview.setFont(f)
        col = self.font_color.color or "#000000"
        self.font_preview.setStyleSheet(f"background:white; border:1px solid #D4D4D4; color:{col};")

    # Fill --------------------------------------------------------------------
    def _build_fill(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.addWidget(QLabel("Background Color:"))
        self.fill_color = self.style0.get("fill")
        nob = QPushButton("No Color")
        nob.clicked.connect(lambda: self._set_fill(None))
        v.addWidget(nob)
        g = QGridLayout()
        g.setSpacing(3)
        for i, col in enumerate(THEME_COLORS + STANDARD_COLORS):
            b = QToolButton()
            b.setFixedSize(22, 22)
            b.setStyleSheet(f"QToolButton{{background:{col}; border:1px solid #C6C6C6;}}")
            b.clicked.connect(lambda _=False, c=col: self._set_fill(c))
            g.addWidget(b, i // 10, i % 10)
        v.addLayout(g)
        more = QPushButton("More Colors...")
        more.clicked.connect(self._more_fill)
        v.addWidget(more)
        v.addWidget(QLabel("Sample:"))
        self.fill_sample = QLabel("")
        self.fill_sample.setMinimumHeight(40)
        v.addWidget(self.fill_sample)
        v.addStretch(1)
        self._set_fill(self.fill_color)
        self.tabs.addTab(w, "Fill")

    def _more_fill(self):
        c = QColorDialog.getColor(QColor(self.fill_color or "#FFFFFF"), self, "Colors")
        if c.isValid():
            self._set_fill(c.name().upper())

    def _set_fill(self, col):
        self.fill_color = col
        self.fill_sample.setStyleSheet(f"background:{col or 'white'}; border:1px solid #D4D4D4;")

    # result --------------------------------------------------------------------
    def changes(self) -> Dict[str, object]:
        """Style keys to set; value None removes the key."""
        ch: Dict[str, object] = {}
        ch["number_format"] = self.number_format()
        ch["align"] = [None, "left", "center", "right"][self.halign.currentIndex()]
        ch["wrap"] = True if self.wrap.isChecked() else None
        if self._font_changed:
            ch["font_name"] = self.font_box.currentFont().family()
        if self._size_changed:
            try:
                ch["font_size"] = float(self.size_box.currentText())
            except ValueError:
                pass
        ch["bold"] = True if self.bold.isChecked() else None
        ch["italic"] = True if self.italic.isChecked() else None
        ch["underline"] = True if self.underline.isChecked() else None
        ch["color"] = self.font_color.color
        ch["fill"] = self.fill_color
        # only report keys that actually changed
        return {k: v for k, v in ch.items() if self.style0.get(k) != v}


# --------------------------------------------------------------------------
# Find & Replace
# --------------------------------------------------------------------------
class FindReplaceDialog(QDialog):
    find_next = Signal(dict)
    find_all = Signal(dict)
    replace_one = Signal(dict)
    replace_all = Signal(dict)

    def __init__(self, parent=None, replace=False):
        super().__init__(parent)
        self.setWindowTitle("Find and Replace")
        self.setModal(False)
        lay = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.tabs.addTab(QWidget(), "Find")
        self.tabs.addTab(QWidget(), "Replace")
        self.tabs.setMaximumHeight(28)
        lay.addWidget(self.tabs)
        form = QFormLayout()
        self.what = QComboBox()
        self.what.setEditable(True)
        self.what.setMinimumWidth(320)
        self.with_ = QComboBox()
        self.with_.setEditable(True)
        self.with_label = QLabel("Replace with:")
        form.addRow("Find what:", self.what)
        form.addRow(self.with_label, self.with_)
        lay.addLayout(form)
        opts = QHBoxLayout()
        self.case = QCheckBox("Match case")
        self.whole = QCheckBox("Match entire cell contents")
        self.look_in = QComboBox()
        self.look_in.addItems(["Formulas", "Values"])
        self.scope = QComboBox()
        self.scope.addItems(["Sheet", "Workbook"])
        opts.addWidget(QLabel("Within:"))
        opts.addWidget(self.scope)
        opts.addWidget(QLabel("Look in:"))
        opts.addWidget(self.look_in)
        lay.addLayout(opts)
        o2 = QHBoxLayout()
        o2.addWidget(self.case)
        o2.addWidget(self.whole)
        o2.addStretch(1)
        lay.addLayout(o2)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#444;")
        lay.addWidget(self.status)
        btns = QHBoxLayout()
        self.b_replace_all = QPushButton("Replace All")
        self.b_replace = QPushButton("Replace")
        self.b_find_all = QPushButton("Find All")
        self.b_next = QPushButton("Find Next")
        self.b_next.setDefault(True)
        close = QPushButton("Close")
        for b in (self.b_replace_all, self.b_replace, self.b_find_all, self.b_next, close):
            btns.addWidget(b)
        lay.addLayout(btns)
        self.results = QListWidget()
        self.results.setMaximumHeight(140)
        self.results.hide()
        lay.addWidget(self.results)
        self.b_next.clicked.connect(lambda: self.find_next.emit(self.options()))
        self.b_find_all.clicked.connect(lambda: self.find_all.emit(self.options()))
        self.b_replace.clicked.connect(lambda: self.replace_one.emit(self.options()))
        self.b_replace_all.clicked.connect(lambda: self.replace_all.emit(self.options()))
        close.clicked.connect(self.close)
        self.tabs.currentChanged.connect(self._mode)
        self.tabs.setCurrentIndex(1 if replace else 0)
        self._mode(1 if replace else 0)

    def _mode(self, i):
        rep = i == 1
        for w in (self.with_, self.with_label, self.b_replace, self.b_replace_all):
            w.setVisible(rep)
        self.adjustSize()

    def set_replace_mode(self, rep: bool):
        self.tabs.setCurrentIndex(1 if rep else 0)

    def options(self) -> dict:
        return {"what": self.what.currentText(), "with": self.with_.currentText(),
                "case": self.case.isChecked(), "whole": self.whole.isChecked(),
                "values": self.look_in.currentIndex() == 1,
                "workbook": self.scope.currentIndex() == 1}


# --------------------------------------------------------------------------
# Insert Function
# --------------------------------------------------------------------------
FUNCTION_CATEGORIES = {
    "Math & Trig": "ABS ACOS ASIN ATAN ATAN2 CEILING COS DEGREES EVEN EXP FACT FLOOR INT LN LOG "
                   "LOG10 MOD MROUND ODD PI POWER PRODUCT RADIANS RAND RANDBETWEEN ROUND "
                   "ROUNDDOWN ROUNDUP SIGN SIN SQRT SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SUMSQ "
                   "TAN TRUNC GCD LCM QUOTIENT CEILING.MATH FLOOR.MATH SEQUENCE",
    "Statistical": "AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF "
                   "COUNTIFS LARGE MAX MAXA MAXIFS MEDIAN MIN MINA MINIFS MODE RANK RANK.EQ SMALL "
                   "STDEV STDEV.S STDEV.P STDEVP VAR VAR.S VAR.P VARP PERCENTILE QUARTILE "
                   "CORREL FORECAST SLOPE INTERCEPT",
    "Logical": "AND FALSE IF IFERROR IFNA IFS NOT OR SWITCH TRUE XOR",
    "Text": "CHAR CLEAN CODE CONCAT CONCATENATE DOLLAR EXACT FIND FIXED LEFT LEN LOWER MID "
            "PROPER REPLACE REPT RIGHT SEARCH SUBSTITUTE T TEXT TEXTJOIN TRIM UPPER VALUE "
            "NUMBERVALUE UNICHAR UNICODE",
    "Date & Time": "DATE DATEDIF DATEVALUE DAY DAYS EDATE EOMONTH HOUR MINUTE MONTH NETWORKDAYS "
                   "NOW SECOND TIME TIMEVALUE TODAY WEEKDAY WEEKNUM WORKDAY YEAR YEARFRAC",
    "Lookup & Reference": "ADDRESS CHOOSE COLUMN COLUMNS HLOOKUP INDEX INDIRECT LOOKUP MATCH "
                          "OFFSET ROW ROWS TRANSPOSE VLOOKUP XLOOKUP XMATCH UNIQUE SORT FILTER",
    "Financial": "FV IPMT NPER NPV PMT PPMT PV RATE IRR SLN DDB DB",
    "Information": "CELL ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD "
                   "ISTEXT ISREF N NA TYPE ERROR.TYPE ISFORMULA",
}


def function_category(name: str) -> str:
    for cat, names in FUNCTION_CATEGORIES.items():
        if name in names.split():
            return cat
    return "Other"


class InsertFunctionDialog(QDialog):
    def __init__(self, functions, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Function")
        self.functions = functions
        self.names = sorted(functions)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Search for a function:"))
        sh = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText("Type a brief description or name, then click Go")
        sh.addWidget(self.search)
        go = QPushButton("Go")
        sh.addWidget(go)
        lay.addLayout(sh)
        ch = QHBoxLayout()
        ch.addWidget(QLabel("Or select a category:"))
        self.cat = QComboBox()
        self.cat.addItems(["All"] + list(FUNCTION_CATEGORIES) + ["Other"])
        ch.addWidget(self.cat, 1)
        lay.addLayout(ch)
        lay.addWidget(QLabel("Select a function:"))
        self.list = QListWidget()
        lay.addWidget(self.list, 1)
        self.desc = QLabel("")
        self.desc.setWordWrap(True)
        self.desc.setMinimumHeight(40)
        f = self.desc.font()
        f.setBold(True)
        self.desc.setFont(f)
        lay.addWidget(self.desc)
        self.count = QLabel("")
        self.count.setStyleSheet("color:#666;")
        lay.addWidget(self.count)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.search.textChanged.connect(self._filter)
        go.clicked.connect(self._filter)
        self.cat.currentIndexChanged.connect(self._filter)
        self.list.currentTextChanged.connect(self._show_desc)
        self.list.itemDoubleClicked.connect(lambda _: self.accept())
        self._filter()
        self.resize(440, 480)

    def _filter(self, *_):
        q = self.search.text().strip().upper()
        cat = self.cat.currentText()
        self.list.clear()
        for n in self.names:
            if q and q not in n.upper():
                continue
            if cat != "All" and function_category(n) != cat:
                continue
            self.list.addItem(n)
        self.count.setText(f"{self.list.count()} of {len(self.names)} functions")
        if self.list.count():
            self.list.setCurrentRow(0)

    def _show_desc(self, name):
        fn = self.functions.get(name) if hasattr(self.functions, "get") else None
        doc = (getattr(fn, "__doc__", None) or "").strip().split("\n")[0] if fn else ""
        self.desc.setText(f"{name}(...)\n{doc}" if name else "")

    def selected(self) -> Optional[str]:
        it = self.list.currentItem()
        return it.text() if it else None


# --------------------------------------------------------------------------
# Macros
# --------------------------------------------------------------------------
class MacrosDialog(QDialog):
    def __init__(self, macros: List[str], parent=None, book_name="This Workbook"):
        super().__init__(parent)
        self.setWindowTitle("Macro")
        self.action = None
        lay = QHBoxLayout(self)
        left = QVBoxLayout()
        left.addWidget(QLabel("Macro name:"))
        self.name = QLineEdit()
        left.addWidget(self.name)
        self.list = QListWidget()
        self.list.addItems(macros)
        left.addWidget(self.list, 1)
        row = QHBoxLayout()
        row.addWidget(QLabel("Macros in:"))
        cb = QComboBox()
        cb.addItems([book_name])
        row.addWidget(cb, 1)
        left.addLayout(row)
        self.descr = QLabel("Description")
        self.descr.setStyleSheet("color:#666;")
        left.addWidget(self.descr)
        lay.addLayout(left, 1)
        right = QVBoxLayout()
        self.buttons = {}
        for label in ("Run", "Step Into", "Edit", "Create", "Delete", "Options...", "Cancel"):
            b = QPushButton(label)
            self.buttons[label] = b
            right.addWidget(b)
            if label in ("Run", "Edit", "Create", "Delete"):
                b.clicked.connect(lambda _=False, a=label.lower(): self._do(a))
        self.buttons["Step Into"].setEnabled(False)
        self.buttons["Options..."].setEnabled(False)
        self.buttons["Cancel"].clicked.connect(self.reject)
        self.buttons["Run"].setDefault(True)
        right.addStretch(1)
        lay.addLayout(right)
        self.list.currentTextChanged.connect(self.name.setText)
        self.list.itemDoubleClicked.connect(lambda _: self._do("run"))
        self.name.textChanged.connect(self._sync)
        if macros:
            self.list.setCurrentRow(0)
        self._sync()
        self.resize(460, 340)

    def _sync(self, *_):
        n = self.name.text().strip()
        exists = n in [self.list.item(i).text() for i in range(self.list.count())]
        self.buttons["Run"].setEnabled(exists)
        self.buttons["Edit"].setEnabled(exists)
        self.buttons["Delete"].setEnabled(exists)
        self.buttons["Create"].setEnabled(bool(n) and not exists and n.replace("_", "a")
                                          .replace(".", "a").isalnum())

    def _do(self, action):
        if not self.buttons[action.capitalize()].isEnabled():
            return
        self.action = action
        self.accept()

    def macro_name(self) -> str:
        return self.name.text().strip()


MACRO_SECURITY = [
    ("disable_notify", "Disable all macros with notification"),
    ("disable_all", "Disable all macros without notification"),
    ("enable_all", "Enable all macros (not recommended; potentially dangerous code can run)"),
]


class MacroSecurityDialog(QDialog):
    def __init__(self, current: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Trust Center - Macro Settings")
        lay = QVBoxLayout(self)
        head = QLabel("Macro Settings")
        f = head.font()
        f.setBold(True)
        f.setPointSizeF(f.pointSizeF() + 2)
        head.setFont(f)
        lay.addWidget(head)
        lay.addWidget(QLabel("For macros in documents opened in PortableExcel:"))
        self.group = QButtonGroup(self)
        self.radios = {}
        for key, text in MACRO_SECURITY:
            rb = QRadioButton(text)
            self.radios[key] = rb
            self.group.addButton(rb)
            lay.addWidget(rb)
            if key == current:
                rb.setChecked(True)
        if not self.group.checkedButton():
            self.radios["disable_notify"].setChecked(True)
        note = QLabel("With notification, a yellow message bar lets you enable the content "
                      "each time a workbook with macros is opened.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#666;")
        lay.addWidget(note)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def value(self) -> str:
        for k, rb in self.radios.items():
            if rb.isChecked():
                return k
        return "disable_notify"


# --------------------------------------------------------------------------
# Passwords & warnings
# --------------------------------------------------------------------------
class PasswordDialog(QDialog):
    def __init__(self, title: str, message: str, error: str = "", confirm: bool = False,
                 parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(icons.pixmap("lock", 32))
        top.addWidget(ic, 0, Qt.AlignTop)
        lbl = QLabel(message)
        lbl.setWordWrap(True)
        top.addWidget(lbl, 1)
        lay.addLayout(top)
        form = QFormLayout()
        self.pw = QLineEdit()
        self.pw.setEchoMode(QLineEdit.Password)
        self.pw.setMinimumWidth(260)
        form.addRow("Password:", self.pw)
        self.pw2 = None
        if confirm:
            self.pw2 = QLineEdit()
            self.pw2.setEchoMode(QLineEdit.Password)
            form.addRow("Reenter password:", self.pw2)
        lay.addLayout(form)
        self.err = QLabel(error)
        self.err.setStyleSheet("color:#C00000;")
        self.err.setVisible(bool(error))
        lay.addWidget(self.err)
        if confirm:
            caution = QLabel("Caution: if you lose or forget the password, it cannot be "
                             "recovered.")
            caution.setWordWrap(True)
            caution.setStyleSheet("color:#666;")
            lay.addWidget(caution)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _ok(self):
        if self.pw2 is not None and self.pw.text() != self.pw2.text():
            self.err.setText("The passwords do not match.")
            self.err.show()
            return
        self.accept()

    def password(self) -> str:
        return self.pw.text()


class WarningsDialog(QDialog):
    def __init__(self, title: str, intro: str, warnings: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        lay = QVBoxLayout(self)
        top = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(self.style().standardIcon(self.style().StandardPixmap.SP_MessageBoxWarning)
                     .pixmap(32, 32))
        top.addWidget(ic, 0, Qt.AlignTop)
        lbl = QLabel(intro)
        lbl.setWordWrap(True)
        top.addWidget(lbl, 1)
        lay.addLayout(top)
        self.list = QListWidget()
        for w in warnings:
            QListWidgetItem("• " + w, self.list)
        self.list.setWordWrap(True)
        lay.addWidget(self.list)
        bb = QDialogButtonBox(QDialogButtonBox.Ok)
        bb.accepted.connect(self.accept)
        lay.addWidget(bb)
        self.resize(520, 300)

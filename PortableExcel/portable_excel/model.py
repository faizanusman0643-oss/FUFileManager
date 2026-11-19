"""Shared data model for PortableExcel.

Every other module (formula engine, VBA interpreter, file I/O, GUI) builds on
these classes. Keep this module free of GUI and third-party imports.

Coordinates are 0-based everywhere: row 0 / col 0 is cell A1.
"""
from __future__ import annotations

import re
from typing import Callable, Dict, Iterator, List, Optional, Tuple

MAX_ROWS = 1_048_576
MAX_COLS = 16_384  # A .. XFD

CellKey = Tuple[int, int]


# --------------------------------------------------------------------------
# A1 helpers
# --------------------------------------------------------------------------
def col_to_name(c: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA'."""
    s = ""
    c += 1
    while c:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def name_to_col(name: str) -> int:
    """'A' -> 0, 'AA' -> 26. Case-insensitive."""
    n = 0
    for ch in name.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def parse_a1(text: str) -> Optional[CellKey]:
    """'B3' / '$B$3' -> (2, 1). None if invalid."""
    m = _CELL_RE.match(text.strip())
    if not m:
        return None
    r = int(m.group(2)) - 1
    c = name_to_col(m.group(1))
    if 0 <= r < MAX_ROWS and 0 <= c < MAX_COLS:
        return r, c
    return None


def parse_range_a1(text: str) -> Optional[Tuple[int, int, int, int]]:
    """'A1:C5' or 'B2' -> (r1, c1, r2, c2) normalised. Also 'A:A' and '3:3'."""
    t = text.strip().replace("$", "")
    if ":" in t:
        a, b = t.split(":", 1)
        if a.isalpha() and b.isalpha():  # whole columns
            c1, c2 = sorted((name_to_col(a), name_to_col(b)))
            return 0, c1, MAX_ROWS - 1, c2
        if a.isdigit() and b.isdigit():  # whole rows
            r1, r2 = sorted((int(a) - 1, int(b) - 1))
            return r1, 0, r2, MAX_COLS - 1
        p, q = parse_a1(a), parse_a1(b)
        if p is None or q is None:
            return None
        return min(p[0], q[0]), min(p[1], q[1]), max(p[0], q[0]), max(p[1], q[1])
    p = parse_a1(t)
    if p is None:
        return None
    return p[0], p[1], p[0], p[1]


def a1(r: int, c: int, absolute: bool = False) -> str:
    if absolute:
        return f"${col_to_name(c)}${r + 1}"
    return f"{col_to_name(c)}{r + 1}"


# --------------------------------------------------------------------------
# Errors
# --------------------------------------------------------------------------
class XLError:
    """An Excel error value such as #DIV/0!. Compared by code."""

    __slots__ = ("code",)

    def __init__(self, code: str):
        self.code = code

    def __repr__(self) -> str:
        return f"XLError({self.code!r})"

    def __str__(self) -> str:
        return self.code

    def __eq__(self, other) -> bool:
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self) -> int:
        return hash(self.code)


ERR_DIV0 = XLError("#DIV/0!")
ERR_VALUE = XLError("#VALUE!")
ERR_REF = XLError("#REF!")
ERR_NAME = XLError("#NAME?")
ERR_NA = XLError("#N/A")
ERR_NUM = XLError("#NUM!")
ERR_NULL = XLError("#NULL!")
ERR_CYCLE = XLError("#CYCLE!")   # circular reference (Excel shows 0 + warning)
ERR_SYNTAX = XLError("#ERROR!")  # formula could not be parsed

ERRORS_BY_CODE = {e.code: e for e in
                  (ERR_DIV0, ERR_VALUE, ERR_REF, ERR_NAME, ERR_NA, ERR_NUM, ERR_NULL)}


# --------------------------------------------------------------------------
# Literal values typed into a cell
# --------------------------------------------------------------------------
def literal_value(raw: str):
    """Convert non-formula cell text into a value.

    Returns float for numbers ("12", "-3.5", "1e3", "50%" -> 0.5, "1,234" -> 1234),
    bool for TRUE/FALSE, XLError for '#N/A' etc, otherwise the string itself.
    A leading apostrophe forces text: "'123" -> "123".
    """
    if raw is None or raw == "":
        return None
    if raw.startswith("'"):
        return raw[1:]
    s = raw.strip()
    up = s.upper()
    if up == "TRUE":
        return True
    if up == "FALSE":
        return False
    if up in ERRORS_BY_CODE:
        return ERRORS_BY_CODE[up]
    t = s.replace(",", "") if re.fullmatch(r"[-+]?\d{1,3}(,\d{3})+(\.\d*)?", s) else s
    pct = t.endswith("%")
    if pct:
        t = t[:-1]
    try:
        if t and not t.lower().lstrip("+-").startswith(("nan", "inf")):
            v = float(t)
            return v / 100.0 if pct else v
    except ValueError:
        pass
    return raw


def format_general(v: float) -> str:
    """Excel 'General' style number text."""
    if v != v:  # NaN
        return "#NUM!"
    if float(v).is_integer() and abs(v) < 1e15:
        return str(int(v))
    s = f"{v:.10g}"
    if "e" in s:
        mant, exp = s.split("e")
        s = f"{mant}E{exp[0]}{exp[1:].lstrip('0').zfill(2)}"
    return s


def value_to_text(v) -> str:
    """Text shown for a value with General format."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return format_general(float(v))
    return str(v)


# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------
# A style is a plain dict with any of these optional keys:
#   bold: bool, italic: bool, underline: bool
#   color: "#RRGGBB" (font colour), fill: "#RRGGBB" (background)
#   align: "left" | "center" | "right"
#   font_size: float (points), font_name: str
#   number_format: Excel format string, e.g. "0.00", "#,##0.00", "0%"
#   wrap: bool
STYLE_KEYS = ("bold", "italic", "underline", "color", "fill", "align",
              "font_size", "font_name", "number_format", "wrap")


# --------------------------------------------------------------------------
# Sheet / Workbook
# --------------------------------------------------------------------------
class Sheet:
    def __init__(self, name: str, code_name: Optional[str] = None):
        self.name = name
        self.code_name = code_name or name  # VBA module name, e.g. "Sheet1"
        self.cells: Dict[CellKey, str] = {}   # raw text as typed: "12", "=SUM(A1:A3)"
        self.styles: Dict[CellKey, dict] = {}
        self.col_widths: Dict[int, int] = {}  # pixels
        self.row_heights: Dict[int, int] = {}  # pixels
        self.visible = True
        # Floating pictures anchored to a cell (one-cell anchor), e.g. barcodes.
        # Each image is a dict:
        #   id: str (unique in the workbook), name: str ("Barcode 1")
        #   row, col: anchor cell (0-based); dx, dy: offset in px from that cell's top-left
        #   width, height: size in px
        #   data: PNG bytes; fmt: "png" (others converted to PNG on load)
        #   barcode: None, or {"symbology": "code128"|"code39"|"ean13"|"ean8"|"upca"
        #                      |"itf"|"qr"|"datamatrix", "source": "A2" (cell on this
        #                      sheet whose displayed value is encoded) or None,
        #                      "text": str (fixed data when source is None),
        #                      "show_text": bool, "module_px": int, "bar_height_px": int}
        #   A barcode with a source is re-rendered when that cell's value changes.
        self.images: List[dict] = []

    # raw access -----------------------------------------------------------
    def get_raw(self, r: int, c: int) -> str:
        return self.cells.get((r, c), "")

    def set_raw(self, r: int, c: int, raw) -> None:
        if raw is None or raw == "":
            self.cells.pop((r, c), None)
        else:
            self.cells[(r, c)] = str(raw)

    def get_style(self, r: int, c: int) -> dict:
        return self.styles.get((r, c), {})

    def update_style(self, r: int, c: int, **changes) -> None:
        st = dict(self.styles.get((r, c), {}))
        for k, v in changes.items():
            if v is None or v is False and k in ("bold", "italic", "underline", "wrap"):
                st.pop(k, None)
            else:
                st[k] = v
        if st:
            self.styles[(r, c)] = st
        else:
            self.styles.pop((r, c), None)

    def used_bounds(self) -> Tuple[int, int]:
        """(n_rows, n_cols) of the used area counting from A1; (0, 0) if empty."""
        rows = cols = 0
        for (r, c) in self.cells:
            rows = max(rows, r + 1)
            cols = max(cols, c + 1)
        return rows, cols

    def iter_cells_in(self, r1, c1, r2, c2) -> Iterator[Tuple[int, int, str]]:
        """Non-empty cells inside a rectangle, row-major order."""
        area = (r2 - r1 + 1) * (c2 - c1 + 1)
        if area <= len(self.cells):
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    raw = self.cells.get((r, c))
                    if raw is not None:
                        yield r, c, raw
        else:
            for (r, c) in sorted(k for k in self.cells
                                 if r1 <= k[0] <= r2 and c1 <= k[1] <= c2):
                yield r, c, self.cells[(r, c)]


class Workbook:
    """A set of sheets plus VBA modules.

    vba_modules: list of dicts {"name": str, "type": "standard"|"document"|"class"|"form",
                                "code": str}
      - "document" modules are ThisWorkbook and the sheet modules (name == sheet.code_name)
    vba_project_bin: original vbaProject.bin bytes when loaded from .xlsm (kept on save)
    """

    def __init__(self):
        self.sheets: List[Sheet] = [Sheet("Sheet1")]
        self.active_index = 0
        self.path: Optional[str] = None
        self.vba_modules: List[dict] = []
        self.vba_project_bin: Optional[bytes] = None
        self.vba_modified = False
        self.defined_names: Dict[str, str] = {}  # name -> "Sheet1!$A$1:$B$5"
        # listeners called as fn(sheet, r, c) after a cell changes through
        # set_cell(); the GUI uses this to refresh, the VBA host for
        # Worksheet_Change events.
        self.change_listeners: List[Callable] = []

    @property
    def active(self) -> Sheet:
        return self.sheets[self.active_index]

    def sheet_by_name(self, name: str) -> Optional[Sheet]:
        low = name.lower()
        for s in self.sheets:
            if s.name.lower() == low:
                return s
        return None

    def unique_sheet_name(self, base: str = "Sheet") -> str:
        i = len(self.sheets) + 1
        while self.sheet_by_name(f"{base}{i}"):
            i += 1
        return f"{base}{i}"

    def add_sheet(self, name: Optional[str] = None, index: Optional[int] = None) -> Sheet:
        name = name or self.unique_sheet_name()
        code = "Sheet1"
        n = 1
        codes = {s.code_name.lower() for s in self.sheets}
        while f"sheet{n}" in codes:
            n += 1
        code = f"Sheet{n}"
        sh = Sheet(name, code)
        if index is None:
            self.sheets.append(sh)
        else:
            self.sheets.insert(index, sh)
        return sh

    def set_cell(self, sheet: Sheet, r: int, c: int, raw) -> None:
        """Set a cell and notify listeners. Use this for user/VBA edits."""
        sheet.set_raw(r, c, raw)
        for fn in list(self.change_listeners):
            fn(sheet, r, c)

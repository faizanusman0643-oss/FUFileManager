"""Pure helpers used by the GUI (no Qt): fill series, sorting, TSV, snapshots."""
from __future__ import annotations

import copy
import csv
import io
import re
from typing import Callable, Dict, List, Optional, Tuple

from ..model import XLError, literal_value, value_to_text

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]
LISTS = [DAYS, [d[:3] for d in DAYS], MONTHS, [m[:3] for m in MONTHS]]
_TRAIL_NUM = re.compile(r"^(.*?)(\d+)$")


# --------------------------------------------------------------------------
# Fill series
# --------------------------------------------------------------------------
def _num(raw: str) -> Optional[float]:
    if not raw or raw.startswith("="):
        return None
    v = literal_value(raw)
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)


def _num_text(v: float) -> str:
    return value_to_text(round(v, 10))


def _list_match(raw: str):
    for lst in LISTS:
        for i, name in enumerate(lst):
            if raw.lower() == name.lower():
                case = "upper" if raw.isupper() and len(raw) > 1 else (
                    "lower" if raw.islower() else "title")
                return lst, i, case
    return None


def _case(s, case):
    return s.upper() if case == "upper" else (s.lower() if case == "lower" else s)


def fill_series(src: List[str], n: int, shift: Callable[[str, int], str],
                copy_only: bool = False) -> List[str]:
    """Extend a line of raw cell texts by n cells (in the direction of the line).

    shift(raw, k) must return a formula shifted by k cells along the line.
    """
    out: List[str] = []
    L = len(src)
    if L == 0 or n <= 0:
        return out
    nums = [_num(s) for s in src]
    if not copy_only and all(v is not None for v in nums) and L >= 2:
        step = (nums[-1] - nums[0]) / (L - 1)
        return [_num_text(nums[-1] + step * (k + 1)) for k in range(n)]
    if not copy_only and L == 1 or (not copy_only and all(
            _list_match(s or "") for s in src)):
        s0 = src[-1] or ""
        lm = _list_match(s0)
        if lm and s0:
            lst, i, case = lm
            step = 1
            if L >= 2:
                lm0 = _list_match(src[0] or "")
                if lm0 and lm0[0] is lst:
                    step = ((i - lm0[1]) // (L - 1)) or 1
            return [_case(lst[(i + step * (k + 1)) % len(lst)], case) for k in range(n)]
        m = _TRAIL_NUM.match(s0) if L == 1 and not s0.startswith("=") else None
        if m and m.group(1) and _num(s0) is None:
            base, digits = m.group(1), m.group(2)
            val = int(digits)
            return [f"{base}{str(val + k + 1).zfill(len(digits))}" for k in range(n)]
    for k in range(n):
        j = k % L
        raw = src[j]
        if raw and raw.startswith("="):
            out.append(shift(raw, L * (k // L + 1)))
        else:
            out.append(raw)
    return out


# --------------------------------------------------------------------------
# TSV clipboard text
# --------------------------------------------------------------------------
def to_tsv(rows: List[List[str]]) -> str:
    buf = io.StringIO()
    for row in rows:
        cells = []
        for t in row:
            if any(ch in t for ch in '\t\n\r"'):
                t = '"' + t.replace('"', '""') + '"'
            cells.append(t)
        buf.write("\t".join(cells))
        buf.write("\r\n")
    return buf.getvalue()


def parse_tsv(text: str) -> List[List[str]]:
    if not text:
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if text.endswith("\n"):
        text = text[:-1]
    if '"' in text:
        rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    else:
        rows = [line.split("\t") for line in text.split("\n")]
    return rows


# --------------------------------------------------------------------------
# Sorting
# --------------------------------------------------------------------------
def sort_key(v, descending: bool = False):
    """Excel order: numbers < text < logical < errors; blanks always last."""
    if v is None or v == "":
        return (1, 0, 0)
    if isinstance(v, bool):
        k = (2, int(v))
    elif isinstance(v, (int, float)):
        k = (0, float(v))
    elif isinstance(v, XLError):
        k = (3, 0)
    else:
        k = (1, str(v).lower())
    if descending:
        # invert inside the non-blank group
        grp, val = k
        if isinstance(val, str):
            return (0, -grp, _Rev(val))
        return (0, -grp, -val)
    return (0,) + k


class _Rev:
    __slots__ = ("s",)

    def __init__(self, s):
        self.s = s

    def __lt__(self, other):
        return self.s > other.s

    def __eq__(self, other):
        return self.s == other.s


# --------------------------------------------------------------------------
# Book snapshots (undo of structural changes)
# --------------------------------------------------------------------------
def snapshot_book(book) -> dict:
    return {
        "sheets": [(sh, sh.name, sh.code_name, dict(sh.cells),
                    {k: dict(v) for k, v in sh.styles.items()},
                    dict(sh.col_widths), dict(sh.row_heights), sh.visible,
                    copy.deepcopy(getattr(sh, "images", [])))
                   for sh in book.sheets],
        "active": book.active_index,
        "names": dict(book.defined_names),
    }


def restore_book(book, snap: dict) -> None:
    sheets = []
    for (sh, name, code, cells, styles, cw, rh, vis, images) in snap["sheets"]:
        sh.name = name
        sh.code_name = code
        sh.cells = dict(cells)
        sh.styles = {k: dict(v) for k, v in styles.items()}
        sh.col_widths = dict(cw)
        sh.row_heights = dict(rh)
        sh.visible = vis
        if hasattr(sh, "images"):
            sh.images[:] = copy.deepcopy(images)
        sheets.append(sh)
    book.sheets[:] = sheets
    book.active_index = min(snap["active"], len(sheets) - 1)
    book.defined_names.clear()
    book.defined_names.update(snap["names"])


def shift_dict_keys(d: Dict, axis: str, index: int, count: int, is_cell=True) -> Dict:
    """Shift keys of styles (cell keys) or widths/heights (int keys)."""
    out = {}
    for k, v in d.items():
        if is_cell:
            pos = k[0] if axis == "row" else k[1]
        else:
            pos = k
        if count < 0 and index <= pos < index - count:
            continue
        if pos >= index:
            pos += count
        if is_cell:
            nk = (pos, k[1]) if axis == "row" else (k[0], pos)
        else:
            nk = pos
        out[nk] = copy.copy(v)
    return out

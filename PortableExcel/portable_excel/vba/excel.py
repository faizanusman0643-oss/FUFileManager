"""Excel object model (Application, Workbook, Worksheet(s), Range, Font, Interior...)."""
from __future__ import annotations

import bisect
import re
import time

from ..model import MAX_ROWS, MAX_COLS, col_to_name, name_to_col, a1 as _a1
from .errors import rt_error, VBARuntimeErr
from .values import (VBAObject, VBAArray, EMPTY, NULL, NOTHING, MISSING, VDate, VErr,
                     to_str, to_num, to_bool, to_long, to_double, index_int, to_value,
                     str_to_num_or_none, VARIANT)
from .objects import bind, opt, NoOpObj, Collection

XL_UP, XL_DOWN, XL_LEFT, XL_RIGHT = -4162, -4121, -4159, -4161
FULL_ROW = MAX_COLS - 1
FULL_COL = MAX_ROWS - 1

PALETTE = [None, 0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF,
           0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080,
           0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF,
           0x000080, 0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF,
           0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99, 0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99,
           0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699, 0x969696,
           0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333]

ALIGN_IN = {-4108: "center", -4131: "left", -4152: "right", 7: "center", 1: None, -4130: "left",
            5: "left", -4117: "center"}
ALIGN_OUT = {"center": -4108, "left": -4131, "right": -4152, None: 1}


def vba_color_to_hex(v) -> str:
    n = to_long(v)
    r = n & 255
    g = (n >> 8) & 255
    b = (n >> 16) & 255
    return f"#{r:02X}{g:02X}{b:02X}"


def hex_to_vba_color(h, default=0):
    if not h or not isinstance(h, str) or not h.startswith("#") or len(h) < 7:
        return default
    try:
        r = int(h[1:3], 16)
        g = int(h[3:5], 16)
        b = int(h[5:7], 16)
    except ValueError:
        return default
    return r + g * 256 + b * 65536


def rgb_hex_from_index(i):
    i = index_int(i)
    if 1 <= i < len(PALETTE):
        v = PALETTE[i]
        return f"#{v:06X}"
    return None


def index_from_hex(h):
    if not h:
        return None
    try:
        v = int(h[1:7], 16)
    except (ValueError, TypeError):
        return None
    for i in range(1, len(PALETTE)):
        if PALETTE[i] == v:
            return i
    return None


def _clamp_rect(sheet, r1, c1, r2, c2):
    """Clamp whole-row/whole-column style operations to the used area."""
    if r2 - r1 + 1 > 100000 or c2 - c1 + 1 > 5000:
        nr, nc = sheet.used_bounds()
        r2 = min(r2, max(nr - 1, r1))
        c2 = min(c2, max(nc - 1, c1))
    return r1, c1, r2, c2


# --------------------------------------------------------------------------
# Sheet index for End()
# --------------------------------------------------------------------------
class SheetIndex:
    def __init__(self, sheet):
        cols = {}
        rows = {}
        for (r, c) in sheet.cells:
            cols.setdefault(c, []).append(r)
            rows.setdefault(r, []).append(c)
        for v in cols.values():
            v.sort()
        for v in rows.values():
            v.sort()
        self.cols = cols
        self.rows = rows

    def add(self, r, c):
        lst = self.cols.setdefault(c, [])
        i = bisect.bisect_left(lst, r)
        if i >= len(lst) or lst[i] != r:
            lst.insert(i, r)
        lst = self.rows.setdefault(r, [])
        i = bisect.bisect_left(lst, c)
        if i >= len(lst) or lst[i] != c:
            lst.insert(i, c)

    def remove(self, r, c):
        lst = self.cols.get(c)
        if lst:
            i = bisect.bisect_left(lst, r)
            if i < len(lst) and lst[i] == r:
                lst.pop(i)
        lst = self.rows.get(r)
        if lst:
            i = bisect.bisect_left(lst, c)
            if i < len(lst) and lst[i] == c:
                lst.pop(i)


def _end_along(lst, pos, forward, limit, has):
    """Excel Ctrl+Arrow along one line. lst: sorted positions with data."""
    if forward:
        if pos >= limit:
            return limit
        cur = has(pos)
        nxt = has(pos + 1)
        if cur and nxt:
            p = pos + 1
            i = bisect.bisect_left(lst, p)
            while i + 1 < len(lst) and lst[i + 1] == lst[i] + 1:
                i += 1
            return lst[i]
        i = bisect.bisect_right(lst, pos)
        return lst[i] if i < len(lst) else limit
    else:
        if pos <= 0:
            return 0
        cur = has(pos)
        prv = has(pos - 1)
        if cur and prv:
            i = bisect.bisect_left(lst, pos - 1)
            while i - 1 >= 0 and lst[i - 1] == lst[i] - 1:
                i -= 1
            return lst[i]
        i = bisect.bisect_left(lst, pos) - 1
        return lst[i] if i >= 0 else 0


# --------------------------------------------------------------------------
# Range
# --------------------------------------------------------------------------
class RangeObj(VBAObject):
    vba_typename = "Range"
    __slots__ = ("rt", "sheet", "r1", "c1", "r2", "c2", "kind", "areas")

    def __init__(self, rt, sheet, r1, c1, r2, c2, kind="cells", areas=None):
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        if r1 < 0 or c1 < 0 or r2 >= MAX_ROWS or c2 >= MAX_COLS:
            raise rt_error(1004)
        self.rt = rt
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.kind = kind
        self.areas = areas   # list of (r1, c1, r2, c2) for multi-area ranges

    def __repr__(self):
        return f"<Range {self.sheet.name}!{self.address()}>"

    def vba_same(self, other):
        return isinstance(other, RangeObj) and other.sheet is self.sheet and \
            (other.r1, other.c1, other.r2, other.c2) == (self.r1, self.c1, self.r2, self.c2)

    @property
    def nrows(self):
        return self.r2 - self.r1 + 1

    @property
    def ncols(self):
        return self.c2 - self.c1 + 1

    def all_areas(self):
        return self.areas or [(self.r1, self.c1, self.r2, self.c2)]

    def sub(self, r1, c1, r2, c2, kind="cells"):
        return RangeObj(self.rt, self.sheet, r1, c1, r2, c2, kind)

    def cell(self, r, c):
        return RangeObj(self.rt, self.sheet, r, c, r, c)

    def address(self, rabs=True, cabs=True, a1=True, rel_r=0, rel_c=0, external=False):
        def one(r1, c1, r2, c2):
            def cr(r):
                return f"{'$' if rabs else ''}{r + 1}"

            def cc(c):
                return f"{'$' if cabs else ''}{col_to_name(c)}"
            if not a1:
                def rr(r):
                    return f"R{r + 1}" if rabs else (f"R[{r - rel_r}]" if r != rel_r else "R")

                def rc(c):
                    return f"C{c + 1}" if cabs else (f"C[{c - rel_c}]" if c != rel_c else "C")
                if c1 == 0 and c2 == FULL_ROW:
                    return rr(r1) if r1 == r2 else f"{rr(r1)}:{rr(r2)}"
                if r1 == 0 and r2 == FULL_COL:
                    return rc(c1) if c1 == c2 else f"{rc(c1)}:{rc(c2)}"
                s = rr(r1) + rc(c1)
                if (r1, c1) != (r2, c2):
                    s += ":" + rr(r2) + rc(c2)
                return s
            if c1 == 0 and c2 == FULL_ROW:
                return f"{cr(r1)}:{cr(r2)}"
            if r1 == 0 and r2 == FULL_COL:
                return f"{cc(c1)}:{cc(c2)}"
            s = cc(c1) + cr(r1)
            if (r1, c1) != (r2, c2):
                s += ":" + cc(c2) + cr(r2)
            return s
        txt = ",".join(one(*a) for a in self.all_areas())
        if external:
            name = self.sheet.name
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
                name = "'" + name.replace("'", "''") + "'"
            txt = f"[{self.rt.workbook_name()}]{name}!{txt}"
        return txt

    # -- default member ---------------------------------------------------
    def g__default(self, args, named):
        if args:
            return self.g_item(args, named)
        return self.g_value(args, named)

    def s__default(self, args, named, value):
        if args:
            return self.g_item(args, named).s_value((), {}, value)
        return self.s_value(args, named, value)

    def vba_iter(self):
        out = []
        rt = self.rt
        sh = self.sheet
        for (r1, c1, r2, c2) in self.all_areas():
            if self.kind == "rows":
                r1, c1, r2, c2 = _clamp_iter(sh, r1, c1, r2, c2)
                out.extend(RangeObj(rt, sh, r, c1, r, c2) for r in range(r1, r2 + 1))
            elif self.kind == "cols":
                r1, c1, r2, c2 = _clamp_iter(sh, r1, c1, r2, c2)
                out.extend(RangeObj(rt, sh, r1, c, r2, c) for c in range(c1, c2 + 1))
            else:
                r1, c1, r2, c2 = _clamp_iter(sh, r1, c1, r2, c2)
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        out.append(RangeObj(rt, sh, r, c, r, c))
        return out

    # -- values -----------------------------------------------------------
    def g_value(self, args, named):
        rt = self.rt
        if self.r1 == self.r2 and self.c1 == self.c2:
            return rt.cell_value(self.sheet, self.r1, self.c1)
        r2, c2 = self.r2, self.c2
        if self.nrows * self.ncols > 5_000_000:
            raise rt_error(7)
        sh = self.sheet
        data = []
        cv = rt.cell_value
        for r in range(self.r1, r2 + 1):
            for c in range(self.c1, c2 + 1):
                data.append(cv(sh, r, c))
        return VBAArray([(1, self.nrows), (1, self.ncols)], None, data)

    def g_value2(self, args, named):
        rt = self.rt
        if self.r1 == self.r2 and self.c1 == self.c2:
            return rt.cell_value(self.sheet, self.r1, self.c1, dates=False)
        data = [rt.cell_value(self.sheet, r, c, dates=False)
                for r in range(self.r1, self.r2 + 1) for c in range(self.c1, self.c2 + 1)]
        return VBAArray([(1, self.nrows), (1, self.ncols)], None, data)

    def s_value(self, args, named, value):
        self.rt.write_values(self, value, formula=False)

    s_value2 = s_value

    def g_formula(self, args, named):
        sh = self.sheet
        rt = self.rt
        if self.r1 == self.r2 and self.c1 == self.c2:
            return rt.cell_formula(sh, self.r1, self.c1)
        data = [rt.cell_formula(sh, r, c) for r in range(self.r1, self.r2 + 1)
                for c in range(self.c1, self.c2 + 1)]
        return VBAArray([(1, self.nrows), (1, self.ncols)], None, data)

    def s_formula(self, args, named, value):
        self.rt.write_values(self, value, formula=True)

    g_formulalocal = g_formula
    s_formulalocal = s_formula
    g_formula2 = g_formula
    s_formula2 = s_formula

    def s_formulaarray(self, args, named, value):
        self.rt.write_values(self, value, formula=True)

    g_formulaarray = g_formula

    def g_formular1c1(self, args, named):
        sh = self.sheet
        rt = self.rt
        if self.r1 == self.r2 and self.c1 == self.c2:
            return a1_to_r1c1(rt.cell_formula(sh, self.r1, self.c1), self.r1, self.c1)
        data = [a1_to_r1c1(rt.cell_formula(sh, r, c), r, c) for r in range(self.r1, self.r2 + 1)
                for c in range(self.c1, self.c2 + 1)]
        return VBAArray([(1, self.nrows), (1, self.ncols)], None, data)

    def s_formular1c1(self, args, named, value):
        rt = self.rt
        if isinstance(value, VBAObject):
            value = to_value(value)
        if type(value) is VBAArray:
            rows = value.to_rows()
            for i, r in enumerate(range(self.r1, self.r2 + 1)):
                for j, c in enumerate(range(self.c1, self.c2 + 1)):
                    if i < len(rows) and j < len(rows[i]):
                        rt.write_cell(self.sheet, r, c, r1c1_to_a1(to_str(rows[i][j]), r, c))
        else:
            s = to_str(value)
            for r in range(self.r1, self.r2 + 1):
                for c in range(self.c1, self.c2 + 1):
                    rt.write_cell(self.sheet, r, c, r1c1_to_a1(s, r, c))
        rt.after_write(self.sheet, self.r1, self.c1, self.r2, self.c2)

    g_formular1c1local = g_formular1c1
    s_formular1c1local = s_formular1c1

    def g_text(self, args, named):
        rt = self.rt
        if self.r1 == self.r2 and self.c1 == self.c2:
            return rt.cell_text(self.sheet, self.r1, self.c1)
        texts = {rt.cell_text(self.sheet, r, c) for r in range(self.r1, min(self.r2, self.r1 + 1000) + 1)
                 for c in range(self.c1, min(self.c2, self.c1 + 100) + 1)}
        return texts.pop() if len(texts) == 1 else NULL

    def g_hasformula(self, args, named):
        vals = {self.sheet.get_raw(r, c).startswith("=")
                for r in range(self.r1, min(self.r2, self.r1 + 1000) + 1)
                for c in range(self.c1, min(self.c2, self.c1 + 100) + 1)}
        return vals.pop() if len(vals) == 1 else NULL

    # -- geometry ---------------------------------------------------------
    def g_row(self, a, n):
        return self.r1 + 1

    def g_column(self, a, n):
        return self.c1 + 1

    def g_count(self, a, n):
        if self.kind == "rows":
            return sum(a[2] - a[0] + 1 for a in self.all_areas())
        if self.kind == "cols":
            return sum(a[3] - a[1] + 1 for a in self.all_areas())
        return sum((a[2] - a[0] + 1) * (a[3] - a[1] + 1) for a in self.all_areas())

    g_countlarge = g_count

    def g_rows(self, args, named):
        if args:
            ix = to_value(args[0])
            if isinstance(ix, str):
                p = parse_rows_spec(ix)
                return self.sub(self.r1 + p[0], self.c1, self.r1 + p[1], self.c2, "rows")
            i = index_int(ix)
            return self.sub(self.r1 + i - 1, self.c1, self.r1 + i - 1, self.c2, "rows")
        return RangeObj(self.rt, self.sheet, self.r1, self.c1, self.r2, self.c2, "rows", self.areas)

    def g_columns(self, args, named):
        if args:
            ix = to_value(args[0])
            if isinstance(ix, str):
                p = parse_cols_spec(ix)
                return self.sub(self.r1, self.c1 + p[0], self.r2, self.c1 + p[1], "cols")
            i = index_int(ix)
            return self.sub(self.r1, self.c1 + i - 1, self.r2, self.c1 + i - 1, "cols")
        return RangeObj(self.rt, self.sheet, self.r1, self.c1, self.r2, self.c2, "cols", self.areas)

    def g_address(self, args, named):
        ra, ca, style, ext, rel = bind(args, named, ["rowabsolute", "columnabsolute",
                                                      "referencestyle", "external", "relativeto"])
        ra = True if ra is MISSING else to_bool(ra)
        ca = True if ca is MISSING else to_bool(ca)
        a1 = style is MISSING or index_int(style) != -4150
        ext = False if ext is MISSING else to_bool(ext)
        rr = rc = 0
        if isinstance(rel, RangeObj):
            rr, rc = rel.r1, rel.c1
        return self.address(ra, ca, a1, rr, rc, ext)

    g_addresslocal = g_address

    def g_item(self, args, named):
        ri, ci = bind(args, named, ["rowindex", "columnindex"], 1)
        ri = to_value(ri)
        if ci is MISSING:
            i = index_int(ri)
            if self.kind == "rows":
                return self.sub(self.r1 + i - 1, self.c1, self.r1 + i - 1, self.c2)
            if self.kind == "cols":
                return self.sub(self.r1, self.c1 + i - 1, self.r2, self.c1 + i - 1)
            w = self.ncols
            i -= 1
            return self.cell(self.r1 + i // w, self.c1 + i % w)
        ci = to_value(ci)
        c = name_to_col(ci) if isinstance(ci, str) and ci.isalpha() else index_int(ci) - 1
        return self.cell(self.r1 + index_int(ri) - 1, self.c1 + c)

    def g_cells(self, args, named):
        if not args:
            return RangeObj(self.rt, self.sheet, self.r1, self.c1, self.r2, self.c2)
        r = RangeObj(self.rt, self.sheet, self.r1, self.c1, self.r2, self.c2)
        return r.g_item(args, named)

    def g_offset(self, args, named):
        ro, co = bind(args, named, ["rowoffset", "columnoffset"])
        dr = 0 if ro is MISSING else index_int(ro)
        dc = 0 if co is MISSING else index_int(co)
        return RangeObj(self.rt, self.sheet, self.r1 + dr, self.c1 + dc, self.r2 + dr, self.c2 + dc)

    def g_resize(self, args, named):
        rs, cs = bind(args, named, ["rowsize", "columnsize"])
        nr = self.nrows if rs is MISSING else index_int(rs)
        nc = self.ncols if cs is MISSING else index_int(cs)
        if nr < 1 or nc < 1:
            raise rt_error(1004)
        return RangeObj(self.rt, self.sheet, self.r1, self.c1, self.r1 + nr - 1, self.c1 + nc - 1)

    def g_range(self, args, named):
        c1, c2 = bind(args, named, ["cell1", "cell2"], 1)
        base = self.rt.make_range_from_args(self.sheet, c1, c2)
        return RangeObj(self.rt, self.sheet, base.r1 + self.r1, base.c1 + self.c1,
                        base.r2 + self.r1, base.c2 + self.c1)

    def g_end(self, args, named):
        (d,) = bind(args, named, ["direction"], 1)
        d = index_int(d)
        rt = self.rt
        sh = self.sheet
        idx = rt.sheet_index(sh)
        r, c = self.r1, self.c1
        cells = sh.cells
        if d in (XL_UP, XL_DOWN):
            lst = idx.cols.get(c, [])
            nr = _end_along(lst, r, d == XL_DOWN, MAX_ROWS - 1, lambda p: (p, c) in cells)
            return self.cell(nr, c)
        if d in (XL_LEFT, XL_RIGHT):
            lst = idx.rows.get(r, [])
            nc = _end_along(lst, c, d == XL_RIGHT, MAX_COLS - 1, lambda p: (r, p) in cells)
            return self.cell(r, nc)
        raise rt_error(5)

    def g_currentregion(self, args, named):
        sh = self.sheet
        r1, c1, r2, c2 = self.r1, self.c1, self.r2, self.c2

        def any_in(a, b, cc, d):
            if a < 0:
                a = 0
            if b < 0:
                b = 0
            cc = min(cc, MAX_ROWS - 1)
            d = min(d, MAX_COLS - 1)
            if a > cc or b > d:
                return False
            for _ in sh.iter_cells_in(a, b, cc, d):
                return True
            return False
        changed = True
        while changed:
            changed = False
            if r1 > 0 and any_in(r1 - 1, c1 - 1, r1 - 1, c2 + 1):
                r1 -= 1
                changed = True
            if r2 < MAX_ROWS - 1 and any_in(r2 + 1, c1 - 1, r2 + 1, c2 + 1):
                r2 += 1
                changed = True
            if c1 > 0 and any_in(r1 - 1, c1 - 1, r2 + 1, c1 - 1):
                c1 -= 1
                changed = True
            if c2 < MAX_COLS - 1 and any_in(r1 - 1, c2 + 1, r2 + 1, c2 + 1):
                c2 += 1
                changed = True
        # trim empty single-cell start
        return self.sub(r1, c1, r2, c2)

    def g_entirerow(self, a, n):
        return RangeObj(self.rt, self.sheet, self.r1, 0, self.r2, FULL_ROW, "rows")

    def g_entirecolumn(self, a, n):
        return RangeObj(self.rt, self.sheet, 0, self.c1, FULL_COL, self.c2, "cols")

    def g_worksheet(self, a, n):
        return self.rt.sheet_object(self.sheet)

    g_parent = g_worksheet

    def g_areas(self, a, n):
        c = Collection()
        for ar in self.all_areas():
            c.items.append([None, self.sub(*ar)])
        if a:
            return c.g_item(a, n)
        return c

    def g_mergecells(self, a, n):
        return False

    def g_mergearea(self, a, n):
        return self

    def g_application(self, a, n):
        return self.rt.app

    def g_dependents(self, a, n):
        raise rt_error(1004, "No cells were found.")

    g_precedents = g_dependents

    # -- editing ------------------------------------------------------------
    def g_clear(self, a, n):
        self.rt.clear_range(self, contents=True, formats=True)
        return EMPTY

    def g_clearcontents(self, a, n):
        self.rt.clear_range(self, contents=True, formats=False)
        return EMPTY

    def g_clearformats(self, a, n):
        self.rt.clear_range(self, contents=False, formats=True)
        return EMPTY

    def g_clearcomments(self, a, n):
        return EMPTY

    g_clearnotes = g_clearcomments
    g_clearhyperlinks = g_clearcomments
    g_clearoutline = g_clearcomments

    def g_delete(self, args, named):
        (shift,) = bind(args, named, ["shift"])
        self.rt.delete_range(self, None if shift is MISSING else index_int(shift))
        return EMPTY

    def g_insert(self, args, named):
        shift, origin = bind(args, named, ["shift", "copyorigin"])
        self.rt.insert_range(self, None if shift is MISSING else index_int(shift))
        return EMPTY

    def g_copy(self, args, named):
        (dest,) = bind(args, named, ["destination"])
        rt = self.rt
        if dest is MISSING:
            rt.clipboard = rt.snapshot(self, cut=False)
            return True
        if not isinstance(dest, RangeObj):
            raise rt_error(424)
        rt.paste(rt.snapshot(self, cut=False), dest, "all")
        return True

    def g_cut(self, args, named):
        (dest,) = bind(args, named, ["destination"])
        rt = self.rt
        if dest is MISSING:
            rt.clipboard = rt.snapshot(self, cut=True)
            return True
        snap = rt.snapshot(self, cut=True)
        rt.paste(snap, dest, "all")
        return True

    def g_pastespecial(self, args, named):
        paste, op, skip, transpose = bind(args, named, ["paste", "operation", "skipblanks", "transpose"])
        rt = self.rt
        if rt.clipboard is None:
            raise rt_error(1004, "PasteSpecial method of Range class failed")
        p = -4104 if paste is MISSING else index_int(paste)
        mode = {-4104: "all", -4163: "values", -4122: "formats", -4123: "formulas",
                12: "values+formats", 11: "formulas", 7: "all", 8: "widths"}.get(p, "all")
        rt.paste(rt.clipboard, self, mode,
                 skip_blanks=(skip is not MISSING and to_bool(skip)),
                 transpose=(transpose is not MISSING and to_bool(transpose)),
                 operation=(0 if op is MISSING else index_int(op)))
        return True

    def g_select(self, a, n):
        self.rt.select_range(self)
        return True

    def g_activate(self, a, n):
        self.rt.select_range(self, activate=True)
        return True

    def g_show(self, a, n):
        return True

    def g_calculate(self, a, n):
        self.rt.engine_recalc()
        return True

    def g_autofit(self, a, n):
        return True

    def g_merge(self, a, n):
        return EMPTY

    g_unmerge = g_merge

    def s_mergecells(self, a, n, v):
        pass

    def g_find(self, args, named):
        return self.rt.range_find(self, args, named)

    def g_findnext(self, args, named):
        (after,) = bind(args, named, ["after"])
        return self.rt.range_find_next(self, after, forward=True)

    def g_findprevious(self, args, named):
        (after,) = bind(args, named, ["after"])
        return self.rt.range_find_next(self, after, forward=False)

    def g_replace(self, args, named):
        return self.rt.range_replace(self, args, named)

    def g_sort(self, args, named):
        return self.rt.range_sort(self, args, named)

    def g_autofill(self, args, named):
        dest, typ = bind(args, named, ["destination", "type"], 1)
        self.rt.autofill(self, dest, 0 if typ is MISSING else index_int(typ))
        return True

    def g_filldown(self, a, n):
        if self.nrows > 1:
            src = self.sub(self.r1, self.c1, self.r1, self.c2)
            self.rt.autofill(src, self, 1)
        return True

    def g_fillright(self, a, n):
        if self.ncols > 1:
            src = self.sub(self.r1, self.c1, self.r2, self.c1)
            self.rt.autofill(src, self, 1)
        return True

    def g_removeduplicates(self, args, named):
        cols, header = bind(args, named, ["columns", "header"])
        self.rt.remove_duplicates(self, cols, header)
        return EMPTY

    def g_specialcells(self, args, named):
        typ, val = bind(args, named, ["type", "value"], 1)
        return self.rt.special_cells(self, index_int(typ), val)

    def g_texttocolumns(self, args, named):
        raise VBARuntimeErr(5, "Range.TextToColumns is not supported in PortableExcel")

    def g_printout(self, a, n):
        raise VBARuntimeErr(5, "Printing is not supported in PortableExcel")

    def g_addcomment(self, a, n):
        return NoOpObj("Comment")

    def g_comment(self, a, n):
        return NOTHING

    def g_name(self, a, n):
        rt = self.rt
        addr = self.address()
        for k, v in rt.book.defined_names.items():
            if v.replace("$", "").lower().endswith(addr.replace("$", "").lower()):
                return NoOpObj("Name", {"name": k, "": k, "refersto": "=" + v})
        raise rt_error(1004)

    def s_name(self, a, n, v):
        nm = to_str(v)
        sheet = self.sheet.name
        q = sheet if re.fullmatch(r"\w+", sheet) else "'" + sheet + "'"
        self.rt.book.defined_names[nm] = f"{q}!{self.address()}"

    # -- formatting -----------------------------------------------------------
    def g_font(self, a, n):
        return FontObj(self)

    def g_interior(self, a, n):
        return InteriorObj(self)

    def g_borders(self, a, n):
        b = NoOpObj("Borders")
        if a:
            return NoOpObj("Border")
        return b

    def g_borderaround(self, a, n):
        return True

    def _style_get(self, key, default=None):
        vals = set()
        for r in range(self.r1, min(self.r2, self.r1 + 200) + 1):
            for c in range(self.c1, min(self.c2, self.c1 + 50) + 1):
                vals.add(self.sheet.get_style(r, c).get(key, default))
                if len(vals) > 1:
                    return NULL
        return vals.pop() if vals else default

    def _style_set(self, **kw):
        self.rt.style_range(self, **kw)

    def g_numberformat(self, a, n):
        v = self._style_get("number_format", "General")
        return v

    def s_numberformat(self, a, n, v):
        s = to_str(v)
        self._style_set(number_format=None if s.lower() == "general" else s)

    g_numberformatlocal = g_numberformat
    s_numberformatlocal = s_numberformat

    def g_horizontalalignment(self, a, n):
        v = self._style_get("align", None)
        if v is NULL:
            return NULL
        return ALIGN_OUT.get(v, 1)

    def s_horizontalalignment(self, a, n, v):
        self._style_set(align=ALIGN_IN.get(index_int(v)))

    def g_verticalalignment(self, a, n):
        return -4107

    def s_verticalalignment(self, a, n, v):
        pass

    def g_wraptext(self, a, n):
        v = self._style_get("wrap", False)
        return v

    def s_wraptext(self, a, n, v):
        self._style_set(wrap=bool(to_bool(v)))

    def g_columnwidth(self, a, n):
        px = self.sheet.col_widths.get(self.c1)
        if px is None:
            return 8.43
        return round(max(px - 5, 0) / 7.0, 2)

    def s_columnwidth(self, a, n, v):
        w = to_double(v)
        px = max(0, int(round(w * 7 + 5))) if w > 0 else 0
        c2 = min(self.c2, self.c1 + 16383)
        for c in range(self.c1, c2 + 1):
            self.sheet.col_widths[c] = px
        self.rt.host_refresh_soft()

    def g_rowheight(self, a, n):
        px = self.sheet.row_heights.get(self.r1)
        if px is None:
            return 15.0
        return round(px * 0.75, 2)

    def s_rowheight(self, a, n, v):
        pt = to_double(v)
        px = max(0, int(round(pt * 4 / 3)))
        r2 = min(self.r2, self.r1 + 100000)
        for r in range(self.r1, r2 + 1):
            self.sheet.row_heights[r] = px
        self.rt.host_refresh_soft()

    def g_width(self, a, n):
        return sum((self.sheet.col_widths.get(c, 64)) * 0.75 for c in range(self.c1, min(self.c2, self.c1 + 500) + 1))

    def g_height(self, a, n):
        return sum((self.sheet.row_heights.get(r, 20)) * 0.75 for r in range(self.r1, min(self.r2, self.r1 + 5000) + 1))

    def g_left(self, a, n):
        return sum(self.sheet.col_widths.get(c, 64) * 0.75 for c in range(0, min(self.c1, 2000)))

    def g_top(self, a, n):
        return sum(self.sheet.row_heights.get(r, 20) * 0.75 for r in range(0, min(self.r1, 20000)))

    def g_hidden(self, a, n):
        if self.kind == "cols" or (self.r1 == 0 and self.r2 == FULL_COL):
            return self.sheet.col_widths.get(self.c1) == 0
        return self.sheet.row_heights.get(self.r1) == 0

    def s_hidden(self, a, n, v):
        hide = to_bool(v)
        if self.kind == "cols" or (self.r1 == 0 and self.r2 == FULL_COL):
            for c in range(self.c1, min(self.c2, self.c1 + 16383) + 1):
                if hide:
                    self.sheet.col_widths[c] = 0
                elif self.sheet.col_widths.get(c) == 0:
                    del self.sheet.col_widths[c]
        else:
            for r in range(self.r1, min(self.r2, self.r1 + 100000) + 1):
                if hide:
                    self.sheet.row_heights[r] = 0
                elif self.sheet.row_heights.get(r) == 0:
                    del self.sheet.row_heights[r]
        self.rt.host_refresh_soft()

    def g_style(self, a, n):
        return "Normal"

    def s_style(self, a, n, v):
        pass

    def g_locked(self, a, n):
        return True

    def s_locked(self, a, n, v):
        pass

    def vba_let(self, name, args, named, value, is_set=False):
        m = getattr(self, "s_" + (name or "_default"), None)
        if m is not None:
            return m(args, named, value)
        if name in _RANGE_NOOP_SET:
            return None
        return VBAObject.vba_let(self, name, args, named, value, is_set)

    def vba_get(self, name, args, named):
        m = getattr(self, "g_" + (name or "_default"), None)
        if m is None:
            if name in _RANGE_NOOP_GET:
                return NoOpObj(name)
            raise rt_error(438)
        return m(args, named)


_RANGE_NOOP_SET = {"indentlevel", "orientation", "shrinktofit", "addindent", "readingorder",
                   "formulahidden", "showdetail", "useStandardHeight".lower(), "outlinelevel"}
_RANGE_NOOP_GET = {"formatconditions", "validation", "hyperlinks", "characters", "errors",
                   "displayformat", "listobject", "pivottable", "sparklinegroups"}


def _clamp_iter(sheet, r1, c1, r2, c2):
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000:
        nr, nc = sheet.used_bounds()
        r2 = min(r2, max(nr - 1, r1))
        c2 = min(c2, max(nc - 1, c1))
    return r1, c1, r2, c2


def parse_rows_spec(s):
    s = s.replace("$", "")
    if ":" in s:
        a, b = s.split(":", 1)
        return int(a) - 1, int(b) - 1
    return int(s) - 1, int(s) - 1


def parse_cols_spec(s):
    s = s.replace("$", "")
    if ":" in s:
        a, b = s.split(":", 1)
        return name_to_col(a), name_to_col(b)
    return name_to_col(s), name_to_col(s)


# --------------------------------------------------------------------------
# R1C1 conversion
# --------------------------------------------------------------------------
_A1_REF = re.compile(r"(?<![A-Za-z0-9_.\"!])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])")
_R1C1_REF = re.compile(r"(?<![A-Za-z0-9_])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(])")


def _split_strings(f):
    """Yield (is_string, text) parts of a formula."""
    parts = re.split(r'("(?:[^"]|"")*")', f)
    for i, p in enumerate(parts):
        yield i % 2 == 1, p


def a1_to_r1c1(f, r, c):
    if not isinstance(f, str) or not f.startswith("="):
        return f

    def rep(m):
        col = name_to_col(m.group(2))
        row = int(m.group(4)) - 1
        if col >= MAX_COLS or row >= MAX_ROWS:
            return m.group(0)
        rs = f"R{row + 1}" if m.group(3) else ("R" if row == r else f"R[{row - r}]")
        cs = f"C{col + 1}" if m.group(1) else ("C" if col == c else f"C[{col - c}]")
        return rs + cs
    return "".join(p if s else _A1_REF.sub(rep, p) for s, p in _split_strings(f))


def r1c1_to_a1(f, r, c):
    if not isinstance(f, str) or not f.startswith("="):
        return f

    def rep(m):
        rg, cg = m.group(1), m.group(2)
        if rg is None:
            row, rabs = r, False
        elif rg.startswith("["):
            row, rabs = r + int(rg[1:-1]), False
        else:
            row, rabs = int(rg) - 1, True
        if cg is None:
            col, cabs = c, False
        elif cg.startswith("["):
            col, cabs = c + int(cg[1:-1]), False
        else:
            col, cabs = int(cg) - 1, True
        if row < 0 or col < 0:
            return "#REF!"
        return f"{'$' if cabs else ''}{col_to_name(col)}{'$' if rabs else ''}{row + 1}"
    return "".join(p if s else _R1C1_REF.sub(rep, p) for s, p in _split_strings(f))


# --------------------------------------------------------------------------
# Font / Interior
# --------------------------------------------------------------------------
class FontObj(VBAObject):
    vba_typename = "Font"

    def __init__(self, rng):
        self.rng = rng

    def _get(self, key, default):
        return self.rng._style_get(key, default)

    def _set(self, **kw):
        self.rng._style_set(**kw)

    def g_bold(self, a, n):
        return self._get("bold", False)

    def s_bold(self, a, n, v):
        self._set(bold=bool(to_bool(v)))

    def g_italic(self, a, n):
        return self._get("italic", False)

    def s_italic(self, a, n, v):
        self._set(italic=bool(to_bool(v)))

    def g_underline(self, a, n):
        u = self._get("underline", False)
        if u is NULL:
            return NULL
        return 2 if u else -4142

    def s_underline(self, a, n, v):
        v = to_value(v)
        if type(v) is bool:
            on = v
        else:
            on = index_int(v) not in (-4142, 0)
        self._set(underline=on)

    def g_color(self, a, n):
        c = self._get("color", None)
        if c is NULL:
            return NULL
        return hex_to_vba_color(c, 0)

    def s_color(self, a, n, v):
        self._set(color=vba_color_to_hex(v))

    def g_colorindex(self, a, n):
        c = self._get("color", None)
        if c is None:
            return -4105
        return index_from_hex(c) or -4105

    def s_colorindex(self, a, n, v):
        i = index_int(v)
        self._set(color=None if i in (-4142, -4105, 0) else rgb_hex_from_index(i))

    def g_size(self, a, n):
        return self._get("font_size", 11.0)

    def s_size(self, a, n, v):
        self._set(font_size=to_double(v))

    def g_name(self, a, n):
        return self._get("font_name", "Calibri")

    def s_name(self, a, n, v):
        self._set(font_name=to_str(v))

    def g_fontstyle(self, a, n):
        b = self._get("bold", False)
        i = self._get("italic", False)
        return "Bold Italic" if b and i else "Bold" if b else "Italic" if i else "Regular"

    def s_fontstyle(self, a, n, v):
        s = to_str(v).lower()
        self._set(bold="bold" in s, italic="italic" in s)

    def g_themecolor(self, a, n):
        return 0

    def vba_let(self, name, args, named, value, is_set=False):
        m = getattr(self, "s_" + name, None)
        if m is not None:
            return m(args, named, value)
        if name in ("strikethrough", "superscript", "subscript", "themecolor", "tintandshade",
                    "shadow", "outlinefont", "themefont"):
            return None
        raise rt_error(438)

    def vba_get(self, name, args, named):
        m = getattr(self, "g_" + name, None)
        if m is not None:
            return m(args, named)
        if name in ("strikethrough", "superscript", "subscript", "shadow", "outlinefont"):
            return False
        if name in ("tintandshade", "themefont"):
            return 0
        raise rt_error(438)


class InteriorObj(VBAObject):
    vba_typename = "Interior"

    def __init__(self, rng):
        self.rng = rng

    def g_color(self, a, n):
        c = self.rng._style_get("fill", None)
        if c is NULL:
            return NULL
        return hex_to_vba_color(c, 16777215)

    def s_color(self, a, n, v):
        self.rng._style_set(fill=vba_color_to_hex(v))

    def g_colorindex(self, a, n):
        c = self.rng._style_get("fill", None)
        if c is None:
            return -4142
        return index_from_hex(c) or -4142

    def s_colorindex(self, a, n, v):
        i = index_int(v)
        self.rng._style_set(fill=None if i in (-4142, -4105, 0) else rgb_hex_from_index(i))

    def g_pattern(self, a, n):
        return -4142 if self.rng._style_get("fill", None) is None else 1

    def s_pattern(self, a, n, v):
        if index_int(v) == -4142:
            self.rng._style_set(fill=None)

    def vba_let(self, name, args, named, value, is_set=False):
        m = getattr(self, "s_" + name, None)
        if m is not None:
            return m(args, named, value)
        if name in ("themecolor", "tintandshade", "patterncolor", "patterncolorindex",
                    "patternthemecolor", "patterntintandshade"):
            return None
        raise rt_error(438)

    def vba_get(self, name, args, named):
        m = getattr(self, "g_" + name, None)
        if m is not None:
            return m(args, named)
        if name in ("themecolor", "tintandshade", "patterncolor", "patterncolorindex"):
            return 0
        raise rt_error(438)


# --------------------------------------------------------------------------
# Worksheet, Sheets, Workbook
# --------------------------------------------------------------------------
class DocMixin:
    """Members defined in the document module's VBA code take precedence."""

    def _doc_module(self):
        raise NotImplementedError

    def vba_get(self, name, args, named):
        mod = self._doc_module()
        if mod is not None:
            r = self.rt.doc_member_get(mod, self, name, args, named)
            if r is not _NOTFOUND:
                return r
        m = getattr(self, "g_" + (name or "_default"), None)
        if m is None:
            raise rt_error(438)
        return m(args, named)

    def vba_let(self, name, args, named, value, is_set=False):
        mod = self._doc_module()
        if mod is not None:
            if self.rt.doc_member_let(mod, self, name, args, named, value, is_set):
                return
        return VBAObject.vba_let(self, name, args, named, value, is_set)


_NOTFOUND = object()


class WorksheetObj(DocMixin, VBAObject):
    vba_typename = "Worksheet"

    def __init__(self, rt, sheet):
        self.rt = rt
        self.sheet = sheet

    def _doc_module(self):
        return self.rt.doc_module(self.sheet.code_name)

    def vba_same(self, other):
        return isinstance(other, WorksheetObj) and other.sheet is self.sheet

    def _all(self):
        return RangeObj(self.rt, self.sheet, 0, 0, MAX_ROWS - 1, MAX_COLS - 1)

    def g_name(self, a, n):
        return self.sheet.name

    def s_name(self, a, n, v):
        self.rt.rename_sheet(self.sheet, to_str(v))

    def g_codename(self, a, n):
        return self.sheet.code_name

    g__codename = g_codename

    def g_index(self, a, n):
        return self.rt.book.sheets.index(self.sheet) + 1

    def g_range(self, args, named):
        c1, c2 = bind(args, named, ["cell1", "cell2"], 1)
        return self.rt.make_range_from_args(self.sheet, c1, c2)

    def g_cells(self, args, named):
        if not args:
            return self._all()
        return self._all().g_item(args, named)

    def g_rows(self, args, named):
        return self._all().g_rows(args, named)

    def g_columns(self, args, named):
        return self._all().g_columns(args, named)

    def g_usedrange(self, a, n):
        cells = self.sheet.cells
        if not cells:
            return RangeObj(self.rt, self.sheet, 0, 0, 0, 0)
        rs = [k[0] for k in cells]
        cs = [k[1] for k in cells]
        return RangeObj(self.rt, self.sheet, min(rs), min(cs), max(rs), max(cs))

    def g_activate(self, a, n):
        self.rt.activate_sheet(self.sheet)
        return EMPTY

    def g_select(self, a, n):
        self.rt.activate_sheet(self.sheet)
        return EMPTY

    def g_delete(self, a, n):
        return self.rt.delete_sheet(self.sheet)

    def g_copy(self, args, named):
        before, after = bind(args, named, ["before", "after"])
        self.rt.copy_sheet(self.sheet, before, after)
        return EMPTY

    def g_move(self, args, named):
        before, after = bind(args, named, ["before", "after"])
        self.rt.move_sheet(self.sheet, before, after)
        return EMPTY

    def g_visible(self, a, n):
        if self.sheet.visible:
            return -1
        return 2 if getattr(self.sheet, "very_hidden", False) else 0

    def s_visible(self, a, n, v):
        v = to_value(v)
        if type(v) is bool:
            vis = v
            very = False
        else:
            i = index_int(v)
            vis = i == -1 or i == 1
            very = i == 2
        if not vis and sum(1 for s in self.rt.book.sheets if s.visible) <= 1 and self.sheet.visible:
            raise rt_error(1004, "Unable to set the Visible property of the Worksheet class")
        self.sheet.visible = vis
        try:
            self.sheet.very_hidden = very
        except AttributeError:
            pass
        self.rt.host_sheets_changed()

    def g_protect(self, a, n):
        self.rt.protected.add(id(self.sheet))
        return EMPTY

    def g_unprotect(self, a, n):
        self.rt.protected.discard(id(self.sheet))
        return EMPTY

    def g_protectcontents(self, a, n):
        return id(self.sheet) in self.rt.protected

    g_protectdrawingobjects = g_protectcontents
    g_protectscenarios = g_protectcontents

    def g_parent(self, a, n):
        return self.rt.workbook

    def g_application(self, a, n):
        return self.rt.app

    def g_next(self, a, n):
        sh = self.rt.book.sheets
        i = sh.index(self.sheet)
        return self.rt.sheet_object(sh[i + 1]) if i + 1 < len(sh) else NOTHING

    def g_previous(self, a, n):
        sh = self.rt.book.sheets
        i = sh.index(self.sheet)
        return self.rt.sheet_object(sh[i - 1]) if i > 0 else NOTHING

    def g_evaluate(self, args, named):
        return self.rt.evaluate_text(to_str(args[0]), self.sheet)

    def g_calculate(self, a, n):
        self.rt.engine_recalc()
        return EMPTY

    def g_paste(self, args, named):
        dest, link = bind(args, named, ["destination", "link"])
        rt = self.rt
        if rt.clipboard is None:
            raise rt_error(1004)
        if dest is MISSING:
            s = rt.selection_tuple()
            dest = RangeObj(rt, self.sheet, s[1], s[2], s[1], s[2])
        rt.paste(rt.clipboard, dest, "all")
        return EMPTY

    def g_sort(self, a, n):
        so = self.rt.sort_objects.get(id(self.sheet))
        if so is None:
            so = self.rt.sort_objects[id(self.sheet)] = SortObj(self.rt, self)
        return so

    def g_names(self, a, n):
        return NamesObj(self.rt)

    def g_type(self, a, n):
        return -4167

    def g_standardwidth(self, a, n):
        return 8.43

    def g_shapes(self, a, n):
        return NoOpObj("Shapes", {"count": 0})

    def g_listobjects(self, a, n):
        return NoOpObj("ListObjects", {"count": 0})

    def g_chartobjects(self, a, n):
        return NoOpObj("ChartObjects", {"count": 0})

    def g_hyperlinks(self, a, n):
        return NoOpObj("Hyperlinks", {"count": 0})

    def g_pagesetup(self, a, n):
        return NoOpObj("PageSetup")

    def g_tab(self, a, n):
        return NoOpObj("Tab")

    def g_outline(self, a, n):
        return NoOpObj("Outline")

    def g_autofiltermode(self, a, n):
        return False

    def s_autofiltermode(self, a, n, v):
        pass

    def g_filtermode(self, a, n):
        return False

    def g_showalldata(self, a, n):
        return EMPTY

    def g_enablecalculation(self, a, n):
        return True

    def s_enablecalculation(self, a, n, v):
        pass

    def g_printout(self, a, n):
        raise VBARuntimeErr(5, "Printing is not supported in PortableExcel")

    g_printpreview = g_printout

    def g_scrollarea(self, a, n):
        return ""

    def s_scrollarea(self, a, n, v):
        pass

    def g_displaypagebreaks(self, a, n):
        return False

    def s_displaypagebreaks(self, a, n, v):
        pass

    def g_enableselection(self, a, n):
        return 0

    def s_enableselection(self, a, n, v):
        pass


class SheetsColl(VBAObject):
    def __init__(self, rt, worksheets):
        self.rt = rt
        self.worksheets = worksheets
        self.vba_typename = "Sheets"

    def _find(self, ix):
        ix = to_value(ix)
        sheets = self.rt.book.sheets
        if isinstance(ix, str):
            s = self.rt.book.sheet_by_name(ix)
            if s is None:
                raise rt_error(9)
            return s
        i = index_int(ix)
        if 1 <= i <= len(sheets):
            return sheets[i - 1]
        raise rt_error(9)

    def g_item(self, args, named):
        (ix,) = bind(args, named, ["index"], 1)
        if type(ix) is VBAArray:
            return self.rt.sheet_object(self._find(ix.data[0]))
        return self.rt.sheet_object(self._find(ix))

    g__default = g_item

    def g_count(self, a, n):
        return len(self.rt.book.sheets)

    def vba_iter(self):
        return [self.rt.sheet_object(s) for s in list(self.rt.book.sheets)]

    def g_add(self, args, named):
        before, after, count, typ = bind(args, named, ["before", "after", "count", "type"])
        n = 1 if count is MISSING else index_int(count)
        res = None
        for _ in range(max(1, n)):
            res = self.rt.add_sheet(before, after)
        return res

    g_add2 = g_add

    def g_select(self, a, n):
        return EMPTY

    def g_application(self, a, n):
        return self.rt.app

    def g_parent(self, a, n):
        return self.rt.workbook

    def g_visible(self, a, n):
        return -1

    def g_copy(self, args, named):
        raise VBARuntimeErr(5, "Copying all sheets to a new workbook is not supported")

    def g_printout(self, a, n):
        raise VBARuntimeErr(5, "Printing is not supported in PortableExcel")


class NamesObj(VBAObject):
    vba_typename = "Names"

    def __init__(self, rt):
        self.rt = rt

    def g_add(self, args, named):
        name, refers, visible, mk, mkk, cat, cl, nl, rl, r1, rr = bind(
            args, named, ["name", "refersto", "visible", "macrotype", "shortcutkey", "category",
                          "namelocal", "referstolocal", "categorylocal", "referstor1c1",
                          "referstor1c1local"], 0)
        if name is MISSING:
            name = nl
        if refers is MISSING:
            refers = rl if rl is not MISSING else r1
        nm = to_str(name)
        if isinstance(refers, RangeObj):
            sh = refers.sheet.name
            q = sh if re.fullmatch(r"\w+", sh) else "'" + sh + "'"
            ref = f"{q}!{refers.address()}"
        else:
            ref = to_str(refers)
            if ref.startswith("="):
                ref = ref[1:]
        self.rt.book.defined_names[nm] = ref
        self.rt.engine_invalidate()
        return NameObj(self.rt, nm)

    def g_item(self, args, named):
        ix = to_value(args[0])
        names = list(self.rt.book.defined_names)
        if isinstance(ix, str):
            for k in names:
                if k.lower() == ix.lower():
                    return NameObj(self.rt, k)
            raise rt_error(1004)
        i = index_int(ix)
        if 1 <= i <= len(names):
            return NameObj(self.rt, names[i - 1])
        raise rt_error(9)

    g__default = g_item

    def g_count(self, a, n):
        return len(self.rt.book.defined_names)

    def vba_iter(self):
        return [NameObj(self.rt, k) for k in self.rt.book.defined_names]


class NameObj(VBAObject):
    vba_typename = "Name"

    def __init__(self, rt, name):
        self.rt = rt
        self.name = name

    def g_name(self, a, n):
        return self.name

    g__default = g_name

    def g_refersto(self, a, n):
        return "=" + self.rt.book.defined_names.get(self.name, "")

    def s_refersto(self, a, n, v):
        s = to_str(v)
        self.rt.book.defined_names[self.name] = s[1:] if s.startswith("=") else s
        self.rt.engine_invalidate()

    def g_referstorange(self, a, n):
        return self.rt.range_from_address(self.rt.book.defined_names.get(self.name, ""), None)

    def g_delete(self, a, n):
        self.rt.book.defined_names.pop(self.name, None)
        self.rt.engine_invalidate()
        return EMPTY

    def g_visible(self, a, n):
        return True


class WorkbookObj(DocMixin, VBAObject):
    vba_typename = "Workbook"

    def __init__(self, rt):
        self.rt = rt

    def _doc_module(self):
        return self.rt.doc_module("thisworkbook")

    def vba_same(self, other):
        return isinstance(other, WorkbookObj)

    def g_name(self, a, n):
        return self.rt.workbook_name()

    g__default = g_name

    def g_path(self, a, n):
        import os
        p = self.rt.book.path
        return os.path.dirname(os.path.abspath(p)) if p else ""

    def g_fullname(self, a, n):
        import os
        p = self.rt.book.path
        return os.path.abspath(p) if p else self.rt.workbook_name()

    def g_sheets(self, a, n):
        c = SheetsColl(self.rt, False)
        return c.g_item(a, n) if a else c

    def g_worksheets(self, a, n):
        c = SheetsColl(self.rt, True)
        return c.g_item(a, n) if a else c

    def g_activesheet(self, a, n):
        return self.rt.sheet_object(self.rt.book.active)

    def g_save(self, a, n):
        self.rt.save_workbook()
        return EMPTY

    def g_saveas(self, args, named):
        (fn,) = bind(args, named, ["filename"], 0)[:1]
        h = getattr(self.rt.host, "save_workbook_as", None)
        if h is None:
            raise VBARuntimeErr(5, "Workbook.SaveAs is not supported in PortableExcel")
        h(to_str(fn))
        return EMPTY

    g_savecopyas = g_saveas

    def g_close(self, a, n):
        return EMPTY

    def g_activate(self, a, n):
        return EMPTY

    def g_saved(self, a, n):
        return not getattr(self.rt.book, "modified", False)

    def s_saved(self, a, n, v):
        pass

    def g_readonly(self, a, n):
        return False

    def g_names(self, a, n):
        c = NamesObj(self.rt)
        return c.g_item(a, n) if a else c

    def g_codename(self, a, n):
        return "ThisWorkbook"

    def g_application(self, a, n):
        return self.rt.app

    g_parent = g_application

    def g_refreshall(self, a, n):
        self.rt.engine_recalc()
        return EMPTY

    def g_protect(self, a, n):
        return EMPTY

    g_unprotect = g_protect

    def g_windows(self, a, n):
        w = self.rt.window
        return w if not a else w

    def g_vbproject(self, a, n):
        raise VBARuntimeErr(5, "Access to the VBA project object model is not supported")

    def g_fileformat(self, a, n):
        p = (self.rt.book.path or "").lower()
        return 52 if p.endswith(".xlsm") else 51

    def g_charts(self, a, n):
        return NoOpObj("Sheets", {"count": 0})


class WorkbooksColl(VBAObject):
    vba_typename = "Workbooks"

    def __init__(self, rt):
        self.rt = rt

    def g_count(self, a, n):
        return 1

    def g_item(self, args, named):
        ix = to_value(args[0])
        if isinstance(ix, str):
            if ix.lower() in (self.rt.workbook_name().lower(),
                              self.rt.workbook_name().lower().rsplit(".", 1)[0]):
                return self.rt.workbook
            raise rt_error(9)
        if index_int(ix) == 1:
            return self.rt.workbook
        raise rt_error(9)

    g__default = g_item

    def vba_iter(self):
        return [self.rt.workbook]

    def g_open(self, a, n):
        raise VBARuntimeErr(1004, "Opening other workbooks from VBA is not supported in PortableExcel")

    def g_add(self, a, n):
        raise VBARuntimeErr(1004, "Creating new workbooks from VBA is not supported in PortableExcel")


class WindowObj(VBAObject):
    vba_typename = "Window"

    def __init__(self, rt):
        self.rt = rt
        self.props = {"freezepanes": False, "zoom": 100, "displaygridlines": True,
                      "displayheadings": True, "windowstate": -4143, "splitrow": 0,
                      "splitcolumn": 0, "scrollrow": 1, "scrollcolumn": 1, "caption": "",
                      "displayworkbooktabs": True, "visible": True}

    def vba_get(self, name, args, named):
        if name in ("", "item"):
            return self
        if name == "activesheet":
            return self.rt.sheet_object(self.rt.book.active)
        if name == "activecell":
            return self.rt.app.g_activecell((), {})
        if name == "selection":
            return self.rt.app.g_selection((), {})
        if name == "count":
            return 1
        if name in ("smallscroll", "largescroll", "scrollworkbooktabs", "activate", "close"):
            return EMPTY
        if name in self.props:
            return self.props[name]
        return NoOpObj(name)

    def vba_let(self, name, args, named, value, is_set=False):
        self.props[name] = value


class SortObj(VBAObject):
    vba_typename = "Sort"

    def __init__(self, rt, ws):
        self.rt = rt
        self.ws = ws
        self.fields = SortFields()
        self.rng = None
        self.header = 2
        self.matchcase = False
        self.orientation = 1

    def g_sortfields(self, a, n):
        return self.fields

    def g_setrange(self, args, named):
        self.rng = args[0]
        return EMPTY

    def g_header(self, a, n):
        return self.header

    def s_header(self, a, n, v):
        self.header = index_int(v)

    def g_matchcase(self, a, n):
        return self.matchcase

    def s_matchcase(self, a, n, v):
        self.matchcase = to_bool(v)

    def g_orientation(self, a, n):
        return self.orientation

    def s_orientation(self, a, n, v):
        self.orientation = index_int(v)

    def s_sortmethod(self, a, n, v):
        pass

    def g_sortmethod(self, a, n):
        return 1

    def g_rng(self, a, n):
        return self.rng if self.rng is not None else NOTHING

    def g_apply(self, a, n):
        if not isinstance(self.rng, RangeObj):
            raise rt_error(1004, "The sort reference is not valid.")
        keys = [(k, o) for k, o in self.fields.items]
        self.rt.do_sort(self.rng, keys, self.header, self.matchcase, self.orientation == 2)
        return EMPTY


class SortFields(VBAObject):
    vba_typename = "SortFields"

    def __init__(self):
        self.items = []

    def g_clear(self, a, n):
        self.items = []
        return EMPTY

    def g_add(self, args, named):
        key, sorton, order, color, dataopt = bind(args, named, ["key", "sorton", "order",
                                                                 "customorder", "dataoption"], 1)
        self.items.append((key, 1 if order is MISSING else index_int(order)))
        return NoOpObj("SortField")

    g_add2 = g_add

    def g_count(self, a, n):
        return len(self.items)


# --------------------------------------------------------------------------
# WorksheetFunction
# --------------------------------------------------------------------------
class WorksheetFunctionObj(VBAObject):
    vba_typename = "WorksheetFunction"

    def __init__(self, rt, raise_errors=True):
        self.rt = rt
        self.raise_errors = raise_errors

    def vba_get(self, name, args, named):
        return self.rt.call_worksheet_function(name, args, self.raise_errors)

    def vba_let(self, name, args, named, value, is_set=False):
        raise rt_error(438)


# --------------------------------------------------------------------------
# Application
# --------------------------------------------------------------------------
class ApplicationObj(VBAObject):
    vba_typename = "Application"

    def __init__(self, rt):
        self.rt = rt
        self.screenupdating = True
        self.displayalerts = True
        self.enableevents = True
        self.calculation = -4105
        self.statusbar = False
        self.caption = "Microsoft Excel"
        self.cutcopymode = False
        self.extra = {}

    def reset(self):
        self.screenupdating = True
        self.statusbar = False

    def vba_get(self, name, args, named):
        m = getattr(self, "g_" + (name or "_default"), None)
        if m is not None:
            return m(args, named)
        if name in self.extra:
            return self.extra[name]
        if name in _APP_NOOP:
            return _APP_NOOP[name]
        if self.rt.has_worksheet_function(name):
            return self.rt.call_worksheet_function(name, args, False)
        raise rt_error(438)

    def vba_let(self, name, args, named, value, is_set=False):
        m = getattr(self, "s_" + name, None)
        if m is not None:
            return m(args, named, value)
        if name in _APP_NOOP or name in self.extra:
            self.extra[name] = value
            return
        raise rt_error(438)

    def g__default(self, a, n):
        return "Microsoft Excel"

    g_name = g__default

    def g_application(self, a, n):
        return self

    g_parent = g_application

    def g_activeworkbook(self, a, n):
        return self.rt.workbook

    g_thisworkbook = g_activeworkbook

    def g_workbooks(self, a, n):
        c = WorkbooksColl(self.rt)
        return c.g_item(a, n) if a else c

    def g_activesheet(self, a, n):
        return self.rt.sheet_object(self.rt.book.active)

    def g_sheets(self, a, n):
        return self.rt.workbook.g_sheets(a, n)

    def g_worksheets(self, a, n):
        return self.rt.workbook.g_worksheets(a, n)

    def g_names(self, a, n):
        return self.rt.workbook.g_names(a, n)

    def g_range(self, args, named):
        c1, c2 = bind(args, named, ["cell1", "cell2"], 1)
        return self.rt.make_range_from_args(None, c1, c2)

    def _ws(self):
        return self.rt.sheet_object(self.rt.book.active)

    def g_cells(self, a, n):
        return self._ws().g_cells(a, n)

    def g_rows(self, a, n):
        return self._ws().g_rows(a, n)

    def g_columns(self, a, n):
        return self._ws().g_columns(a, n)

    def g_activecell(self, a, n):
        s = self.rt.selection_tuple()
        return RangeObj(self.rt, s[0], s[1], s[2], s[1], s[2])

    def g_selection(self, a, n):
        s = self.rt.selection_tuple()
        return RangeObj(self.rt, s[0], s[1], s[2], s[3], s[4])

    def g_activewindow(self, a, n):
        return self.rt.window

    def g_windows(self, a, n):
        return self.rt.window

    def g_screenupdating(self, a, n):
        return self.screenupdating

    def s_screenupdating(self, a, n, v):
        self.screenupdating = to_bool(v)
        if self.screenupdating:
            self.rt.host.refresh()

    def g_displayalerts(self, a, n):
        return self.displayalerts

    def s_displayalerts(self, a, n, v):
        self.displayalerts = to_bool(v)

    def g_enableevents(self, a, n):
        return self.enableevents

    def s_enableevents(self, a, n, v):
        self.enableevents = to_bool(v)

    def g_calculation(self, a, n):
        return self.calculation

    def s_calculation(self, a, n, v):
        self.calculation = index_int(v)
        if self.calculation == -4105:
            self.rt.engine_invalidate()

    def g_calculate(self, a, n):
        self.rt.engine_recalc()
        return EMPTY

    g_calculatefull = g_calculate
    g_calculatefullrebuild = g_calculate

    def g_statusbar(self, a, n):
        return self.statusbar

    def s_statusbar(self, a, n, v):
        v = to_value(v)
        self.statusbar = v
        self.rt.host.status_bar("" if v is False else to_str(v))

    def g_cutcopymode(self, a, n):
        return self.rt.clipboard is not None and (2 if self.rt.clipboard.get("cut") else 1) or False

    def s_cutcopymode(self, a, n, v):
        v = to_value(v)
        if v is False or v == 0:
            self.rt.clipboard = None

    def g_caption(self, a, n):
        return self.caption

    def s_caption(self, a, n, v):
        self.caption = to_str(v)

    def g_username(self, a, n):
        import getpass
        try:
            return getpass.getuser()
        except Exception:
            return "User"

    def s_username(self, a, n, v):
        pass

    def g_version(self, a, n):
        return "16.0"

    def g_build(self, a, n):
        return 16827

    def g_operatingsystem(self, a, n):
        import platform
        return platform.system()

    def g_pathseparator(self, a, n):
        import os
        return os.sep

    def g_path(self, a, n):
        return ""

    def g_defaultfilepath(self, a, n):
        import os
        return os.path.expanduser("~")

    def g_decimalseparator(self, a, n):
        return "."

    def g_thousandsseparator(self, a, n):
        return ","

    def g_international(self, a, n):
        return 1

    def g_worksheetfunction(self, a, n):
        return WorksheetFunctionObj(self.rt, True)

    def g_evaluate(self, args, named):
        return self.rt.evaluate_text(to_str(args[0]), None)

    def g_run(self, args, named):
        if not args:
            raise rt_error(449)
        name = to_str(args[0])
        return self.rt.run_nested(name, list(args[1:]))

    def g_wait(self, args, named):
        t = to_value(args[0]) if args else 0
        from .values import to_date
        from . import dates as D
        from datetime import datetime
        target = float(to_date(t))
        now = D.dt_to_serial(datetime.now())
        secs = (target - now) * 86400 if target > 1 else target * 86400
        end = time.time() + max(0.0, min(secs, 30.0))
        while time.time() < end:
            self.rt.host.process_events()
            self.rt.check_stop()
            time.sleep(0.02)
        return True

    def g_ontime(self, a, n):
        raise VBARuntimeErr(5, "Application.OnTime is not supported in PortableExcel")

    def g_onkey(self, a, n):
        return EMPTY

    def g_volatile(self, a, n):
        return EMPTY

    def g_inputbox(self, args, named):
        prompt, title, default, left, top, hf, hc, typ = bind(
            args, named, ["prompt", "title", "default", "left", "top", "helpfile",
                          "helpcontextid", "type"], 1)
        r = self.rt.host.inputbox(to_str(prompt), "Input" if title is MISSING else to_str(title),
                                  "" if default is MISSING else to_str(default))
        if r is None:
            return False
        t = 2 if typ is MISSING else index_int(typ)
        if t & 8:
            try:
                return self.rt.range_from_address(r.lstrip("="), None)
            except VBARuntimeErr:
                raise rt_error(424)
        if t & 1:
            n_ = str_to_num_or_none(r)
            if n_ is None:
                raise VBARuntimeErr(13, "Number is not valid")
            return n_
        if t & 4:
            return to_bool(r)
        return r

    def g_intersect(self, args, named):
        rs = [a for a in args if a is not MISSING]
        for r in rs:
            if not isinstance(r, RangeObj):
                raise rt_error(13)
        cur = rs[0]
        r1, c1, r2, c2 = cur.r1, cur.c1, cur.r2, cur.c2
        for o in rs[1:]:
            if o.sheet is not cur.sheet:
                raise rt_error(1004)
            r1, c1, r2, c2 = max(r1, o.r1), max(c1, o.c1), min(r2, o.r2), min(c2, o.c2)
            if r1 > r2 or c1 > c2:
                return NOTHING
        return RangeObj(self.rt, cur.sheet, r1, c1, r2, c2)

    def g_union(self, args, named):
        rs = [a for a in args if a is not MISSING]
        areas = []
        for r in rs:
            if not isinstance(r, RangeObj):
                raise rt_error(13)
            areas.extend(r.all_areas())
        first = rs[0]
        res = RangeObj(self.rt, first.sheet, first.r1, first.c1, first.r2, first.c2)
        res.areas = areas if len(areas) > 1 else None
        return res

    def g_goto(self, args, named):
        ref, scroll = bind(args, named, ["reference", "scroll"])
        if isinstance(ref, RangeObj):
            self.rt.select_range(ref)
        elif ref is not MISSING:
            self.rt.select_range(self.rt.range_from_address(to_str(ref), None))
        return EMPTY

    def g_quit(self, a, n):
        return EMPTY

    def g_getopenfilename(self, args, named):
        h = getattr(self.rt.host, "get_open_filename", None)
        if h is None:
            return False
        r = h()
        return False if not r else r

    def g_getsaveasfilename(self, args, named):
        h = getattr(self.rt.host, "get_save_filename", None)
        if h is None:
            return False
        r = h()
        return False if not r else r

    def g_filedialog(self, a, n):
        raise VBARuntimeErr(5, "Application.FileDialog is not supported in PortableExcel")

    def g_sendkeys(self, a, n):
        raise VBARuntimeErr(5, "SendKeys is not supported in PortableExcel")

    def g_converformula(self, a, n):
        raise rt_error(438)

    def g_convertformula(self, args, named):
        formula, fromst, tost, toabs, rel = bind(args, named, ["formula", "fromreferencestyle",
                                                               "toreferencestyle", "toabsolute",
                                                               "relativeto"], 2)
        f = to_str(formula)
        if index_int(fromst) == -4150:
            return r1c1_to_a1(f, 0, 0)
        return a1_to_r1c1(f, 0, 0)

    def g_max(self, args, named):
        return self.rt.call_worksheet_function("max", args, False)

    def g_min(self, args, named):
        return self.rt.call_worksheet_function("min", args, False)

    def g_commandbars(self, a, n):
        return NoOpObj("CommandBars")

    def g_dialogs(self, a, n):
        return NoOpObj("Dialogs")

    def g_addins(self, a, n):
        return NoOpObj("AddIns", {"count": 0})

    def g_activeprinter(self, a, n):
        return ""

    def g_charts(self, a, n):
        return NoOpObj("Sheets", {"count": 0})

    def g_activechart(self, a, n):
        return NOTHING

    def g_hwnd(self, a, n):
        return 0

    def g_ready(self, a, n):
        return True

    def g_mousepointer(self, a, n):
        return 0


_APP_NOOP = {"interactive": True, "cursor": -4143, "displaystatusbar": True,
             "askto updatelinks".replace(" ", ""): True, "windowstate": -4143, "visible": True,
             "printcommunication": True, "iteration": False, "maxchange": 0.001,
             "maxiterations": 100, "enablecancelkey": 1, "displayformulabar": True,
             "displayfullscreen": False, "displaycommentindicator": -1,
             "displayclipboardwindow": False, "enableanimations": True, "autorecover": None,
             "showwindowsintaskbar": True, "displayscrollbars": True, "width": 1024,
             "height": 768, "left": 0, "top": 0, "usableheight": 700, "usablewidth": 1000,
             "copyobjectswithcells": True, "movedirection": -4121, "moveafterreturn": True,
             "moveafterreturndirection": -4121, "standardfont": "Calibri",
             "standardfontsize": 11, "referencestyle": 1, "displaynoteindicator": True,
             "enablelivepreview": True, "automationsecurity": 1, "editdirectlyincell": True,
             "showstartupdialog": False, "displayexcel4menus": False, "warnonfunctionnameconflict": False,
             "errorcheckingoptions": None, "autocorrect": None, "filevalidation": 0}

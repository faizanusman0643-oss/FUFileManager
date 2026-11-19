"""Worksheet operations used by the Excel object model (mixed into VBARuntime)."""
from __future__ import annotations

import math
import re
from decimal import Decimal

from ..model import (MAX_ROWS, MAX_COLS, XLError, literal_value, parse_range_a1, col_to_name,
                     name_to_col, ERRORS_BY_CODE, value_to_text)
from .errors import rt_error, VBARuntimeErr
from .values import (EMPTY, NULL, NOTHING, MISSING, VDate, VErr, VBAArray, VBAObject, to_str,
                     to_value, to_num, to_bool, index_int, XL_ERR_CODES, fmt_double)
from .objects import bind

_DATE_FMT_CACHE = {}


def _is_date_numfmt(fmt):
    r = _DATE_FMT_CACHE.get(fmt)
    if r is None:
        s = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", fmt)
        r = bool(re.search(r"[dDyYhHsS]|[mM]", s)) and not re.search(r"[0#?]", s) and \
            s.lower() != "general"
        _DATE_FMT_CACHE[fmt] = r
    return r


def _xl_to_vba(v):
    if v is None:
        return EMPTY
    if isinstance(v, XLError):
        return VErr(XL_ERR_CODES.get(v.code, 2015))
    return v


def _vba_to_xl(v):
    """VBA value -> engine argument."""
    from .excel import RangeObj
    if isinstance(v, RangeObj):
        from ..engine import RangeRef
        return RangeRef(v.sheet, v.r1, v.c1, v.r2, v.c2)
    if isinstance(v, VBAObject):
        v = to_value(v)
    t = type(v)
    if v is EMPTY or v is MISSING or v is NULL:
        return None
    if t is VBAArray:
        return [[_vba_to_xl(x) for x in row] for row in v.to_rows()]
    if t is VDate or t is Decimal:
        return float(v)
    if t is VErr:
        return ERRORS_BY_CODE.get(v.xl_text, XLError(v.xl_text))
    if t is int:
        return float(v)
    return v


def _xl_result(v):
    if isinstance(v, list):
        rows = [[_xl_to_vba(x) for x in (r if isinstance(r, list) else [r])] for r in v]
        if len(rows) == 1 and len(rows[0]) == 1:
            return rows[0][0]
        return VBAArray.from_2d(rows, 1)
    return _xl_to_vba(v)


class SheetOpsMixin:
    # ------------------------------------------------------------------
    # engine
    # ------------------------------------------------------------------
    def engine_invalidate(self):
        e = self.engine
        if e is not None:
            try:
                e.invalidate()
            except Exception:
                pass

    def engine_recalc(self):
        e = self.engine
        if e is not None:
            e.recalc()

    def host_refresh_soft(self):
        if self.app.screenupdating:
            try:
                self.host.refresh()
            except Exception:
                pass

    def host_sheets_changed(self):
        h = getattr(self.host, "sheets_changed", None)
        if h is not None:
            try:
                h()
            except Exception:
                pass
        self.host_refresh_soft()

    def save_workbook(self):
        h = getattr(self.host, "save_workbook", None)
        if h is not None:
            h()

    # ------------------------------------------------------------------
    # cells
    # ------------------------------------------------------------------
    def cell_value(self, sheet, r, c, dates=True):
        e = self.engine
        if e is not None:
            v = e.value(sheet, r, c)
        else:
            raw = sheet.cells.get((r, c))
            if raw is None:
                return EMPTY
            v = literal_value(raw) if not raw.startswith("=") else raw
        if v is None:
            return EMPTY
        t = type(v)
        if t is float:
            if dates:
                st = sheet.styles.get((r, c))
                if st:
                    nf = st.get("number_format")
                    if nf and _is_date_numfmt(nf):
                        return VDate(v)
            return v
        if t is XLError or isinstance(v, XLError):
            return VErr(XL_ERR_CODES.get(v.code, 2015))
        if t is int:
            return float(v)
        return v

    def cell_formula(self, sheet, r, c):
        return sheet.cells.get((r, c), "")

    def cell_text(self, sheet, r, c):
        e = self.engine
        if e is not None:
            return e.display_text(sheet, r, c)
        return value_to_text(literal_value(sheet.get_raw(r, c)))

    def value_to_raw(self, v, sheet=None, r=0, c=0):
        if isinstance(v, VBAObject):
            v = to_value(v)
        t = type(v)
        if t is str:
            return v
        if t is bool:
            return "TRUE" if v else "FALSE"
        if t is int:
            return str(v)
        if t is float:
            if v != v or math.isinf(v):
                return "#NUM!"
            if v.is_integer() and abs(v) < 1e15:
                return str(int(v))
            return repr(v)
        if v is EMPTY or v is NULL or v is MISSING:
            return ""
        if t is VDate:
            if sheet is not None:
                st = sheet.styles.get((r, c), {})
                nf = st.get("number_format")
                if not nf or not _is_date_numfmt(nf):
                    fmt = "m/d/yyyy" if float(v).is_integer() else (
                        "h:mm:ss" if 0 <= v < 1 else "m/d/yyyy h:mm")
                    sheet.update_style(r, c, number_format=fmt)
            f = float(v)
            return str(int(f)) if f.is_integer() else repr(f)
        if t is Decimal:
            s = format(v.normalize(), "f")
            return s
        if t is VErr:
            return v.xl_text
        if t is VBAArray:
            raise rt_error(13)
        return to_str(v)

    def sheet_index(self, sheet):
        idx = self._indexes.get(id(sheet))
        if idx is None or idx[0] is not sheet.cells or idx[2] != len(sheet.cells):
            from .excel import SheetIndex
            idx = (sheet.cells, SheetIndex(sheet), len(sheet.cells))
            self._indexes[id(sheet)] = idx
        return idx[1]

    def _drop_index(self, sheet):
        self._indexes.pop(id(sheet), None)

    def write_cell(self, sheet, r, c, raw):
        if self.udf_depth:
            raise VBARuntimeErr(1004, "A worksheet function cannot change cells")
        key = (r, c)
        cells = sheet.cells
        had = key in cells
        if raw is None:
            raw = ""
        self.book.set_cell(sheet, r, c, raw)
        idx = self._indexes.get(id(sheet))
        if idx is not None:
            has = key in sheet.cells
            if idx[0] is not sheet.cells:
                self._drop_index(sheet)
            elif had != has:
                if has:
                    idx[1].add(r, c)
                else:
                    idx[1].remove(r, c)
                self._indexes[id(sheet)] = (idx[0], idx[1], len(sheet.cells))

    def after_write(self, sheet, r1, c1, r2, c2):
        self.engine_invalidate()
        imgs = getattr(sheet, "images", None)
        if imgs and any(im.get("barcode") for im in imgs):
            try:
                from .. import barcodes
                barcodes.refresh_barcodes(self.book, self.engine, sheet)
            except Exception:
                pass
        if self.app.enableevents and self.event_depth < 30:
            self.notify_change(sheet, r1, c1, r2, c2)

    def write_values(self, rng, value, formula=False):
        if isinstance(value, VBAObject) and not formula:
            value = to_value(value)
        elif isinstance(value, VBAObject):
            value = to_value(value)
        sh = rng.sheet
        if id(sh) in self.protected:
            raise VBARuntimeErr(1004, "The cell or chart you're trying to change is on a protected sheet.")
        for (r1, c1, r2, c2) in rng.all_areas():
            if type(value) is VBAArray:
                rows = value.to_rows()
                if r1 == r2 and c1 == c2:
                    first = rows[0][0] if rows and rows[0] else EMPTY
                    self.write_cell(sh, r1, c1, self.value_to_raw(first, sh, r1, c1))
                    continue
                nr = len(rows)
                for i, r in enumerate(range(r1, r2 + 1)):
                    for j, c in enumerate(range(c1, c2 + 1)):
                        if i < nr and j < len(rows[i]):
                            v = rows[i][j]
                            raw = self.value_to_raw(v, sh, r, c)
                        elif nr == 1 and i >= 1 and j < len(rows[0]):
                            raw = self.value_to_raw(rows[0][j], sh, r, c)
                        else:
                            raw = "#N/A"
                        self.write_cell(sh, r, c, raw)
            else:
                if value is EMPTY or value == "" and type(value) is str:
                    for (r, c, _) in list(sh.iter_cells_in(r1, c1, r2, c2)):
                        self.write_cell(sh, r, c, "")
                    continue
                if (r2 - r1 + 1) * (c2 - c1 + 1) > 2_000_000:
                    raise rt_error(7)
                raw = self.value_to_raw(value, sh, r1, c1)
                isdate = type(value) is VDate
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        if isdate:
                            self.value_to_raw(value, sh, r, c)
                        self.write_cell(sh, r, c, raw)
        self.after_write(sh, rng.r1, rng.c1, rng.r2, rng.c2)

    def style_range(self, rng, **kw):
        from .excel import _clamp_rect
        sh = rng.sheet
        for (r1, c1, r2, c2) in rng.all_areas():
            r1, c1, r2, c2 = _clamp_rect(sh, r1, c1, r2, c2)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    sh.update_style(r, c, **kw)
        self.host_refresh_soft() if False else None

    def clear_range(self, rng, contents=True, formats=True):
        sh = rng.sheet
        for (r1, c1, r2, c2) in rng.all_areas():
            if contents:
                for (r, c, _) in list(sh.iter_cells_in(r1, c1, r2, c2)):
                    self.write_cell(sh, r, c, "")
            if formats:
                keys = [k for k in sh.styles if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]
                for k in keys:
                    del sh.styles[k]
        if contents:
            self.after_write(sh, rng.r1, rng.c1, rng.r2, rng.c2)

    # ------------------------------------------------------------------
    # addresses
    # ------------------------------------------------------------------
    def _sheet_for(self, sheet):
        return sheet if sheet is not None else self.book.active

    def range_from_address(self, text, sheet):
        from .excel import RangeObj
        areas = []
        target = None
        for part in _split_commas(text):
            sh, rect = self._parse_one(part.strip(), sheet)
            if target is None:
                target = sh
            areas.append(rect)
        if not areas:
            raise rt_error(1004, "Method 'Range' of object '_Global' failed")
        r = RangeObj(self, target, *areas[0])
        if len(areas) > 1:
            r.areas = areas
        return r

    def _parse_one(self, text, sheet):
        sh = self._sheet_for(sheet)
        t = text
        if "!" in t:
            sname, t = t.rsplit("!", 1)
            sname = sname.strip()
            if sname.startswith("'") and sname.endswith("'"):
                sname = sname[1:-1].replace("''", "'")
            if sname.startswith("["):
                sname = sname.split("]", 1)[1]
            s2 = self.book.sheet_by_name(sname)
            if s2 is None:
                raise rt_error(1004, "Method 'Range' of object '_Global' failed")
            sh = s2
        rect = parse_range_a1(t)
        if rect is None:
            # defined name?
            for k, v in self.book.defined_names.items():
                if k.lower() == t.strip().lower():
                    return self._parse_one(v.lstrip("="), sh)
            if re.fullmatch(r"\$?[A-Za-z]{1,3}", t.strip()):
                c = name_to_col(t.strip().lstrip("$"))
                return sh, (0, c, MAX_ROWS - 1, c)
            if re.fullmatch(r"\$?\d+", t.strip()):
                r = int(t.strip().lstrip("$")) - 1
                return sh, (r, 0, r, MAX_COLS - 1)
            raise rt_error(1004, "Method 'Range' of object '_Global' failed")
        return sh, rect

    def make_range_from_args(self, sheet, c1, c2=MISSING):
        from .excel import RangeObj
        sheet = self._sheet_for(sheet)

        def one(x):
            if isinstance(x, RangeObj):
                return x
            x = to_value(x) if isinstance(x, VBAObject) else x
            if not isinstance(x, str):
                raise rt_error(1004, "Method 'Range' of object '_Global' failed")
            return self.range_from_address(x, sheet)
        a = one(c1)
        if c2 is MISSING:
            return a
        b = one(c2)
        return RangeObj(self, a.sheet, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2),
                        max(a.c2, b.c2))

    def app_global(self, name, args, named, sctx):
        if sctx is not None and name in ("range", "cells", "rows", "columns", "evaluate", "names"):
            ws = self.doc_object_for_module(sctx)
            if name != "names":
                return ws.vba_get(name, args, named)
        return self.app.vba_get(name, args, named)

    def evaluate_text(self, text, sheet):
        t = text.strip()
        if t.startswith("="):
            t = t[1:]
        try:
            return self.range_from_address(t, sheet)
        except VBARuntimeErr:
            pass
        e = self.engine
        if e is None:
            raise rt_error(1004)
        res = e.evaluate(t, self._sheet_for(sheet))
        return _xl_result(res)

    # ------------------------------------------------------------------
    # selection / sheets
    # ------------------------------------------------------------------
    def selection_tuple(self):
        sel = None
        try:
            sel = self.host.get_selection()
        except Exception:
            sel = None
        if not sel:
            sh = self.book.active
            return (sh, 0, 0, 0, 0)
        sh, r1, c1, r2, c2 = sel
        if sh is None:
            sh = self.book.active
        return (sh, r1, c1, r2, c2)

    def active_cell_tuple(self):
        ac = getattr(self.host, "get_active_cell", None)
        if ac is not None:
            try:
                a = ac()
                if a:
                    return (a[0] or self.book.active, a[1], a[2])
            except Exception:
                pass
        s = self.selection_tuple()
        return (s[0], s[1], s[2])

    def select_range(self, rng, activate=False):
        sh = rng.sheet
        if sh is not self.book.active:
            if not activate:
                raise VBARuntimeErr(1004, "Select method of Range class failed")
            self.activate_sheet(sh)
        self.host.select(sh, rng.r1, rng.c1, rng.r2, rng.c2)

    def activate_sheet(self, sheet):
        idx = self.book.sheets.index(sheet)
        old = self.book.active
        self.book.active_index = idx
        try:
            self.host.activate_sheet(sheet)
        except Exception:
            pass
        if old is not sheet and self.app.enableevents and self.running:
            if self.has_handler(sheet.code_name, "worksheet_activate"):
                self.fire_event(sheet.code_name, "Worksheet_Activate", [])

    def rename_sheet(self, sheet, new):
        if not new or len(new) > 31 or re.search(r"[\\/?*\[\]:]", new):
            raise VBARuntimeErr(1004, "Application-defined or object-defined error")
        other = self.book.sheet_by_name(new)
        if other is not None and other is not sheet:
            raise VBARuntimeErr(1004, "That name is already taken. Try a different one.")
        old = sheet.name
        try:
            from ..engine import rename_sheet_refs
            rename_sheet_refs(self.book, old, new)
        except Exception:
            ren = getattr(self.engine, "rename_sheet_refs", None)
            if ren:
                ren(old, new)
        sheet.name = new
        self.engine_invalidate()
        self.host_sheets_changed()

    def add_sheet(self, before=MISSING, after=MISSING):
        from .excel import WorksheetObj
        book = self.book
        if isinstance(before, WorksheetObj):
            pos = book.sheets.index(before.sheet)
        elif isinstance(after, WorksheetObj):
            pos = book.sheets.index(after.sheet) + 1
        else:
            pos = book.active_index
        sh = book.add_sheet(None, pos)
        book.active_index = book.sheets.index(sh)
        try:
            self.host.activate_sheet(sh)
        except Exception:
            pass
        self.host_sheets_changed()
        if self.has_handler("thisworkbook", "workbook_newsheet"):
            self.fire_event("ThisWorkbook", "Workbook_NewSheet", [self.sheet_object(sh)])
        return self.sheet_object(sh)

    def delete_sheet(self, sheet):
        book = self.book
        if len(book.sheets) <= 1 or sum(1 for s in book.sheets if s.visible) <= 1 and sheet.visible:
            raise VBARuntimeErr(1004, "A workbook must contain at least one visible worksheet.")
        if self.app.displayalerts:
            r = self.host.msgbox("Microsoft Excel will permanently delete this sheet. Do you want to continue?",
                                 1 + 48, "Microsoft Excel")
            if r != 1:
                return False
        active = book.active
        book.sheets.remove(sheet)
        try:
            from ..engine import delete_sheet_refs
            delete_sheet_refs(book, sheet.name)
        except Exception:
            pass
        if active is sheet:
            book.active_index = min(book.active_index, len(book.sheets) - 1)
            if not book.sheets[book.active_index].visible:
                book.active_index = next(i for i, s in enumerate(book.sheets) if s.visible)
        else:
            book.active_index = book.sheets.index(active)
        self._sheet_objs.pop(id(sheet), None)
        self.engine_invalidate()
        self.host_sheets_changed()
        return True

    def _target_pos(self, before, after):
        from .excel import WorksheetObj
        if isinstance(before, WorksheetObj):
            return self.book.sheets.index(before.sheet)
        if isinstance(after, WorksheetObj):
            return self.book.sheets.index(after.sheet) + 1
        return None

    def copy_sheet(self, sheet, before, after):
        import copy as _copy
        pos = self._target_pos(before, after)
        if pos is None:
            raise VBARuntimeErr(5, "Copying a sheet to a new workbook is not supported")
        base = sheet.name
        n = 2
        while self.book.sheet_by_name(f"{base} ({n})"):
            n += 1
        new = self.book.add_sheet(f"{base} ({n})", pos)
        new.cells = dict(sheet.cells)
        new.styles = _copy.deepcopy(sheet.styles)
        new.col_widths = dict(sheet.col_widths)
        new.row_heights = dict(sheet.row_heights)
        self.book.active_index = self.book.sheets.index(new)
        self.engine_invalidate()
        self.host_sheets_changed()

    def move_sheet(self, sheet, before, after):
        pos = self._target_pos(before, after)
        if pos is None:
            raise VBARuntimeErr(5, "Moving a sheet to a new workbook is not supported")
        active = self.book.active
        old = self.book.sheets.index(sheet)
        self.book.sheets.pop(old)
        if pos > old:
            pos -= 1
        self.book.sheets.insert(pos, sheet)
        self.book.active_index = self.book.sheets.index(active)
        self.host_sheets_changed()

    # ------------------------------------------------------------------
    # structure
    # ------------------------------------------------------------------
    def _insdel(self, sheet, axis, index, count):
        try:
            from ..engine import insert_delete
        except Exception:
            insert_delete = None
        if insert_delete is not None:
            insert_delete(self.book, sheet, axis, index, count)
        else:
            raise VBARuntimeErr(5, "Insert/Delete requires the formula engine")
        self._drop_index(sheet)
        self.engine_invalidate()

    def _move_block(self, sh, keys_fn, dr, dc):
        """Move cells (and styles) selected by keys_fn by (dr, dc)."""
        cells = [(k, sh.cells[k]) for k in list(sh.cells) if keys_fn(k)]
        styles = [(k, sh.styles[k]) for k in list(sh.styles) if keys_fn(k)]
        for k, _ in cells:
            del sh.cells[k]
        for k, _ in styles:
            del sh.styles[k]
        for (r, c), raw in cells:
            nr, nc = r + dr, c + dc
            if 0 <= nr < MAX_ROWS and 0 <= nc < MAX_COLS:
                sh.cells[(nr, nc)] = raw
        for (r, c), st in styles:
            nr, nc = r + dr, c + dc
            if 0 <= nr < MAX_ROWS and 0 <= nc < MAX_COLS:
                sh.styles[(nr, nc)] = st

    def delete_range(self, rng, shift):
        sh = rng.sheet
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        if c1 == 0 and c2 == MAX_COLS - 1:
            self._insdel(sh, "row", r1, -(r2 - r1 + 1))
        elif r1 == 0 and r2 == MAX_ROWS - 1:
            self._insdel(sh, "col", c1, -(c2 - c1 + 1))
        else:
            if shift is None:
                shift = -4162 if (r2 - r1) >= (c2 - c1) or True else -4159
            for k in [k for k in sh.cells if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]:
                del sh.cells[k]
            for k in [k for k in sh.styles if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]:
                del sh.styles[k]
            if shift == -4159:
                n = c2 - c1 + 1
                self._move_block(sh, lambda k: r1 <= k[0] <= r2 and k[1] > c2, 0, -n)
            else:
                n = r2 - r1 + 1
                self._move_block(sh, lambda k: c1 <= k[1] <= c2 and k[0] > r2, -n, 0)
            self._drop_index(sh)
            self.engine_invalidate()
        self.after_write(sh, r1, c1, r2, c2)

    def insert_range(self, rng, shift):
        sh = rng.sheet
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        if c1 == 0 and c2 == MAX_COLS - 1:
            self._insdel(sh, "row", r1, r2 - r1 + 1)
        elif r1 == 0 and r2 == MAX_ROWS - 1:
            self._insdel(sh, "col", c1, c2 - c1 + 1)
        else:
            if shift == -4161:
                n = c2 - c1 + 1
                self._move_block(sh, lambda k: r1 <= k[0] <= r2 and k[1] >= c1, 0, n)
            else:
                n = r2 - r1 + 1
                self._move_block(sh, lambda k: c1 <= k[1] <= c2 and k[0] >= r1, n, 0)
            self._drop_index(sh)
            self.engine_invalidate()
        self.after_write(sh, r1, c1, r2, c2)

    # ------------------------------------------------------------------
    # copy / paste
    # ------------------------------------------------------------------
    def snapshot(self, rng, cut=False):
        sh = rng.sheet
        cells = []
        for r in range(rng.r1, min(rng.r2, rng.r1 + 200000) + 1):
            pass
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000:
            nr, nc = sh.used_bounds()
            r2 = min(r2, max(nr - 1, r1))
            c2 = min(c2, max(nc - 1, c1))
        data = {}
        for (r, c, raw) in sh.iter_cells_in(r1, c1, r2, c2):
            data[(r - r1, c - c1)] = [raw, self.cell_value(sh, r, c, dates=False), None]
        for k, st in sh.styles.items():
            if r1 <= k[0] <= r2 and c1 <= k[1] <= c2:
                e = data.setdefault((k[0] - r1, k[1] - c1), ["", EMPTY, None])
                e[2] = dict(st)
        widths = {c - c1: sh.col_widths[c] for c in range(c1, min(c2, c1 + 16383) + 1) if c in sh.col_widths}
        return {"sheet": sh, "r1": r1, "c1": c1, "nr": r2 - r1 + 1, "nc": c2 - c1 + 1,
                "data": data, "cut": cut, "widths": widths}

    def paste(self, snap, dest, mode="all", skip_blanks=False, transpose=False, operation=0):
        try:
            from ..engine import shift_formula
        except Exception:
            shift_formula = None
        sh = dest.sheet
        nr, nc = snap["nr"], snap["nc"]
        dr0, dc0 = dest.r1, dest.c1
        if transpose:
            nr, nc = nc, nr
        # tile if destination is a multiple of the source
        reps_r = dest.nrows // nr if dest.nrows % nr == 0 and dest.nrows > nr else 1
        reps_c = dest.ncols // nc if dest.ncols % nc == 0 and dest.ncols > nc else 1
        if snap["cut"]:
            src = snap["sheet"]
            for (i, j) in list(snap["data"]):
                k = (snap["r1"] + i, snap["c1"] + j)
                if k in src.cells:
                    self.write_cell(src, k[0], k[1], "")
                src.styles.pop(k, None)
        for tr in range(reps_r):
            for tc in range(reps_c):
                br, bc = dr0 + tr * nr, dc0 + tc * nc
                if mode in ("all", "values", "formulas", "values+formats") and not skip_blanks:
                    for (r, c, _) in list(sh.iter_cells_in(br, bc, br + nr - 1, bc + nc - 1)):
                        self.write_cell(sh, r, c, "")
                for (i, j), (raw, val, st) in snap["data"].items():
                    if transpose:
                        i, j = j, i
                    r, c = br + i, bc + j
                    if r >= MAX_ROWS or c >= MAX_COLS:
                        continue
                    if skip_blanks and raw == "":
                        continue
                    if mode in ("all", "formulas"):
                        out = raw
                        if raw.startswith("=") and not snap["cut"] and shift_formula is not None:
                            out = shift_formula(raw, r - (snap["r1"] + (j if transpose else i)),
                                                c - (snap["c1"] + (i if transpose else j)))
                        if operation and not raw.startswith("="):
                            out = self._paste_op(sh, r, c, val, operation)
                        self.write_cell(sh, r, c, out)
                    elif mode in ("values", "values+formats"):
                        if operation:
                            self.write_cell(sh, r, c, self._paste_op(sh, r, c, val, operation))
                        else:
                            self.write_cell(sh, r, c, self.value_to_raw(val) if raw else "")
                    if mode in ("all", "formats", "values+formats"):
                        if st:
                            sh.styles[(r, c)] = dict(st)
                        elif mode == "formats" or mode == "all":
                            sh.styles.pop((r, c), None)
                if mode == "widths":
                    for j, w in snap["widths"].items():
                        sh.col_widths[bc + j] = w
        if snap["cut"]:
            self.clipboard = None
        self.after_write(sh, dr0, dc0, min(dr0 + nr * reps_r - 1, MAX_ROWS - 1),
                         min(dc0 + nc * reps_c - 1, MAX_COLS - 1))

    def _paste_op(self, sh, r, c, val, op):
        cur = self.cell_value(sh, r, c, dates=False)
        a = 0.0 if cur is EMPTY else to_num(cur)
        b = 0.0 if val is EMPTY else to_num(val)
        res = {2: a + b, 3: a - b, 4: a * b, 5: (a / b if b else float("nan"))}.get(op, b)
        return self.value_to_raw(float(res)) if res == res else "#DIV/0!"

    # ------------------------------------------------------------------
    # find / replace
    # ------------------------------------------------------------------
    def _cell_matcher(self, what, lookat, matchcase):
        w = what if matchcase else what.lower()
        wild = any(ch in w for ch in "*?~")
        if wild:
            out = []
            i = 0
            while i < len(w):
                ch = w[i]
                if ch == "~" and i + 1 < len(w):
                    out.append(re.escape(w[i + 1]))
                    i += 2
                    continue
                out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
                i += 1
            rx = re.compile("".join(out), re.S)
            if lookat == 1:
                return lambda s: rx.fullmatch(s if matchcase else s.lower()) is not None
            return lambda s: rx.search(s if matchcase else s.lower()) is not None
        if lookat == 1:
            return lambda s: (s if matchcase else s.lower()) == w
        return lambda s: w in (s if matchcase else s.lower())

    def range_find(self, rng, args, named):
        what, after, lookin, lookat, order, direction, matchcase, mbyte, fmt = bind(
            args, named, ["what", "after", "lookin", "lookat", "searchorder", "searchdirection",
                          "matchcase", "matchbyte", "searchformat"], 1)
        st = {"what": to_str(what),
              "lookin": -4123 if lookin is MISSING else index_int(lookin),
              "lookat": 2 if lookat is MISSING else index_int(lookat),
              "order": 1 if order is MISSING else index_int(order),
              "matchcase": False if matchcase is MISSING else to_bool(matchcase),
              "range": rng}
        self.find_state = st
        fwd = direction is MISSING or index_int(direction) != 2
        return self._find(st, after if after is not MISSING else None, fwd)

    def range_find_next(self, rng, after, forward=True):
        st = self.find_state
        if st is None:
            return NOTHING
        return self._find(st, None if after is MISSING else after, forward)

    def _find(self, st, after, forward):
        from .excel import RangeObj
        rng = st["range"]
        sh = rng.sheet
        match = self._cell_matcher(st["what"], st["lookat"], st["matchcase"])
        cands = list(sh.iter_cells_in(rng.r1, rng.c1, rng.r2, rng.c2))
        if st["order"] == 2:
            cands.sort(key=lambda t: (t[1], t[0]))
        else:
            cands.sort(key=lambda t: (t[0], t[1]))
        if not cands:
            return NOTHING
        if isinstance(after, RangeObj):
            ar, ac = after.r1, after.c1
        else:
            ar, ac = (rng.r1, rng.c1) if forward else (rng.r2, rng.c2)
            if not forward:
                ar, ac = rng.r1, rng.c1
        key = (lambda t: (t[1], t[0])) if st["order"] == 2 else (lambda t: (t[0], t[1]))
        ak = key((ar, ac))
        if forward:
            order = [t for t in cands if key(t) > ak] + [t for t in cands if key(t) <= ak]
        else:
            order = [t for t in reversed(cands) if key(t) < ak] + [t for t in reversed(cands) if key(t) >= ak]
        for (r, c, raw) in order:
            if st["lookin"] == -4123:
                text = raw
            else:
                text = self.cell_text(sh, r, c)
            if match(text):
                return RangeObj(self, sh, r, c, r, c)
        return NOTHING

    def range_replace(self, rng, args, named):
        what, repl, lookat, order, matchcase = bind(
            args, named, ["what", "replacement", "lookat", "searchorder", "matchcase",
                          "matchbyte", "searchformat", "replaceformat"][:5], 2)
        w = to_str(what)
        rp = to_str(repl)
        la = 2 if lookat is MISSING else index_int(lookat)
        mc = False if matchcase is MISSING else to_bool(matchcase)
        sh = rng.sheet
        match = self._cell_matcher(w, la, mc)
        if any(ch in w for ch in "*?"):
            pat = re.compile("".join(".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in w),
                             0 if mc else re.I)
        else:
            pat = re.compile(re.escape(w), 0 if mc else re.I)
        changed = False
        for (r, c, raw) in list(sh.iter_cells_in(rng.r1, rng.c1, rng.r2, rng.c2)):
            if not match(raw):
                continue
            new = rp if la == 1 else pat.sub(lambda m: rp, raw)
            if new != raw:
                self.write_cell(sh, r, c, new)
                changed = True
        if changed:
            self.after_write(sh, rng.r1, rng.c1, rng.r2, rng.c2)
        return True

    # ------------------------------------------------------------------
    # sort
    # ------------------------------------------------------------------
    def range_sort(self, rng, args, named):
        names = ["key1", "order1", "key2", "type", "order2", "key3", "order3", "header",
                 "ordercustom", "matchcase", "orientation", "sortmethod", "dataoption1",
                 "dataoption2", "dataoption3"]
        v = dict(zip(names, bind(args, named, names)))
        keys = []
        for k, o in (("key1", "order1"), ("key2", "order2"), ("key3", "order3")):
            if v[k] is not MISSING:
                keys.append((v[k], 1 if v[o] is MISSING else index_int(v[o])))
        header = 2 if v["header"] is MISSING else index_int(v["header"])
        mc = False if v["matchcase"] is MISSING else to_bool(v["matchcase"])
        bycols = v["orientation"] is not MISSING and index_int(v["orientation"]) == 2
        self.do_sort(rng, keys, header, mc, bycols)
        return True

    def do_sort(self, rng, keys, header, matchcase, bycols=False):
        from .excel import RangeObj
        try:
            from ..engine import shift_formula
        except Exception:
            shift_formula = None
        sh = rng.sheet
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        if rng.nrows == 1 and rng.ncols == 1:
            cr = rng.g_currentregion((), {})
            r1, c1, r2, c2 = cr.r1, cr.c1, cr.r2, cr.c2
        if (r2 - r1) > 100000 or (c2 - c1) > 5000:
            nr, nc = sh.used_bounds()
            r2 = min(r2, max(nr - 1, r1))
            c2 = min(c2, max(nc - 1, c1))
        if header == 0:   # xlGuess
            first = [self.cell_value(sh, r1, c) for c in range(c1, c2 + 1)]
            second = [self.cell_value(sh, r1 + 1, c) for c in range(c1, c2 + 1)] if r1 < r2 else []
            header = 1 if all(isinstance(x, str) for x in first if x is not EMPTY) and \
                any(not isinstance(x, str) and x is not EMPTY for x in second) else 2
        if header == 1:
            if bycols:
                c1 += 1
            else:
                r1 += 1
        if r1 > r2 or c1 > c2:
            return
        if not keys:
            keys = [(RangeObj(self, sh, r1, c1, r1, c1), 1)]
        kcols = []
        for k, o in keys:
            if isinstance(k, str):
                k = self.range_from_address(k, sh)
            if not isinstance(k, RangeObj):
                raise rt_error(1004, "The sort reference is not valid.")
            kcols.append(((k.r1 if bycols else k.c1), o == 2))

        def sortkey_val(v):
            if v is EMPTY:
                return None
            if isinstance(v, bool):
                return (2, v)
            if isinstance(v, (int, float)):
                return (0, float(v))
            if isinstance(v, VErr):
                return (3, v.number)
            s = to_str(v)
            return (1, s if matchcase else s.lower())

        lines = list(range(c1, c2 + 1)) if bycols else list(range(r1, r2 + 1))

        def val(line, k):
            return self.cell_value(sh, k, line) if bycols else self.cell_value(sh, line, k)
        order = list(lines)
        for kc, desc in reversed(kcols):
            blanks = [ln for ln in order if sortkey_val(val(ln, kc)) is None]
            rest = [ln for ln in order if sortkey_val(val(ln, kc)) is not None]
            rest.sort(key=lambda ln: sortkey_val(val(ln, kc)), reverse=desc)
            order = rest + blanks
        if order == lines:
            return
        # snapshot rows
        snap = {}
        for ln in lines:
            if bycols:
                snap[ln] = [(r, sh.cells.get((r, ln)), sh.styles.get((r, ln))) for r in range(r1, r2 + 1)]
            else:
                snap[ln] = [(c, sh.cells.get((ln, c)), sh.styles.get((ln, c))) for c in range(c1, c2 + 1)]
        for new_ln, old_ln in zip(lines, order):
            for pos, raw, st in snap[old_ln]:
                r, c = (pos, new_ln) if bycols else (new_ln, pos)
                if raw and raw.startswith("=") and shift_formula is not None and new_ln != old_ln:
                    raw = shift_formula(raw, 0, new_ln - old_ln) if bycols else shift_formula(raw, new_ln - old_ln, 0)
                self.write_cell(sh, r, c, raw or "")
                if st:
                    sh.styles[(r, c)] = st
                else:
                    sh.styles.pop((r, c), None)
        self.after_write(sh, r1, c1, r2, c2)

    # ------------------------------------------------------------------
    # autofill / duplicates / special cells
    # ------------------------------------------------------------------
    def autofill(self, src, dest, typ=0):
        from .excel import RangeObj
        try:
            from ..engine import shift_formula
        except Exception:
            shift_formula = None
        if not isinstance(dest, RangeObj):
            raise rt_error(1004)
        sh = src.sheet
        down = dest.ncols == src.ncols and dest.nrows >= src.nrows
        if down:
            lines = [(None, c) for c in range(dest.c1, dest.c2 + 1)]
            n_src = src.nrows
            start, end = dest.r1, dest.r2
        else:
            lines = [(r, None) for r in range(dest.r1, dest.r2 + 1)]
            n_src = src.ncols
            start, end = dest.c1, dest.c2
        for fixed_r, fixed_c in lines:
            def rc(p):
                return (p, fixed_c) if down else (fixed_r, p)
            src_pos = list(range(src.r1, src.r2 + 1)) if down else list(range(src.c1, src.c2 + 1))
            raws = [sh.get_raw(*rc(p)) for p in src_pos]
            vals = [literal_value(x) if x and not x.startswith("=") else None for x in raws]
            numeric = all(isinstance(v, float) for v in vals) and raws
            step = None
            if typ != 1 and numeric and len(vals) >= 2:
                step = (vals[-1] - vals[0]) / (len(vals) - 1)
            elif typ == 2 and numeric:
                step = 1.0
            for k, p in enumerate(range(start, end + 1)):
                if src_pos[0] <= p <= src_pos[-1]:
                    continue
                off = p - src_pos[0]
                i = off % n_src
                raw = raws[i]
                if step is not None:
                    v = vals[0] + step * off
                    out = self.value_to_raw(float(v))
                elif raw.startswith("=") and shift_formula is not None:
                    d = p - src_pos[i]
                    out = shift_formula(raw, d, 0) if down else shift_formula(raw, 0, d)
                elif typ != 1 and n_src == 1 and re.search(r"\d+$", raw) and not isinstance(vals[0], float):
                    m = re.search(r"(\d+)$", raw)
                    out = raw[:m.start()] + str(int(m.group(1)) + off)
                else:
                    out = raw
                r, c = rc(p)
                self.write_cell(sh, r, c, out)
                st = sh.styles.get(rc(src_pos[i]))
                if st:
                    sh.styles[(r, c)] = dict(st)
        self.after_write(sh, dest.r1, dest.c1, dest.r2, dest.c2)

    def remove_duplicates(self, rng, cols, header):
        sh = rng.sheet
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        nr, nc = sh.used_bounds()
        r2 = min(r2, max(nr - 1, r1))
        if header is not MISSING and index_int(header) == 1:
            r1 += 1
        if cols is MISSING:
            cidx = list(range(c1, c2 + 1))
        elif type(cols) is VBAArray:
            cidx = [c1 + index_int(x) - 1 for x in cols.data]
        else:
            cidx = [c1 + index_int(cols) - 1]
        seen = set()
        keep = []
        for r in range(r1, r2 + 1):
            key = tuple((lambda v: v.lower() if isinstance(v, str) else v)(self.cell_value(sh, r, c, dates=False))
                        for c in cidx)
            if key in seen:
                continue
            seen.add(key)
            keep.append(r)
        rows = {r: [(sh.cells.get((r, c)), sh.styles.get((r, c))) for c in range(c1, c2 + 1)]
                for r in keep}
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) in sh.cells:
                    self.write_cell(sh, r, c, "")
                sh.styles.pop((r, c), None)
        for i, r in enumerate(keep):
            for j, (raw, st) in enumerate(rows[r]):
                if raw:
                    self.write_cell(sh, r1 + i, c1 + j, raw)
                if st:
                    sh.styles[(r1 + i, c1 + j)] = st
        self.after_write(sh, r1, c1, r2, c2)

    def special_cells(self, rng, typ, value):
        from .excel import RangeObj
        sh = rng.sheet
        if typ == 11:
            nr, nc = sh.used_bounds()
            return RangeObj(self, sh, max(nr - 1, 0), max(nc - 1, 0), max(nr - 1, 0), max(nc - 1, 0))
        r1, c1, r2, c2 = rng.r1, rng.c1, rng.r2, rng.c2
        if rng.nrows * rng.ncols == 1 or (r2 - r1) > 100000 or (c2 - c1) > 5000:
            ur = self.sheet_object(sh).g_usedrange((), {})
            if rng.nrows * rng.ncols == 1:
                r1, c1, r2, c2 = ur.r1, ur.c1, ur.r2, ur.c2
            else:
                r1, c1 = max(r1, ur.r1), max(c1, ur.c1)
                r2, c2 = min(r2, ur.r2), min(c2, ur.c2)
        cells = []
        if typ == 12:
            return rng
        if typ == 4:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) not in sh.cells:
                        cells.append((r, c))
        elif typ in (2, -4123):
            want_formula = typ == -4123
            for (r, c, raw) in sh.iter_cells_in(r1, c1, r2, c2):
                if raw.startswith("=") == want_formula:
                    if value is not MISSING:
                        v = self.cell_value(sh, r, c, dates=False)
                        flags = index_int(value)
                        kind = 1 if isinstance(v, (int, float)) and not isinstance(v, bool) else \
                            4 if isinstance(v, bool) else 16 if isinstance(v, VErr) else 2
                        if not (flags & kind):
                            continue
                    cells.append((r, c))
        else:
            raise rt_error(1004, "No cells were found.")
        if not cells:
            raise VBARuntimeErr(1004, "No cells were found.")
        res = RangeObj(self, sh, cells[0][0], cells[0][1], cells[0][0], cells[0][1])
        if len(cells) > 1:
            res.areas = [(r, c, r, c) for r, c in cells]
        return res

    # ------------------------------------------------------------------
    # worksheet functions
    # ------------------------------------------------------------------
    def _wf_name(self, name):
        e = self.engine
        funcs = getattr(e, "FUNCTIONS", None) or {}
        up = name.upper()
        if up in funcs:
            return up
        if up.replace("_", ".") in funcs:
            return up.replace("_", ".")
        return None

    def has_worksheet_function(self, name):
        return self.engine is not None and self._wf_name(name) is not None

    def call_worksheet_function(self, name, args, raise_errors=True):
        e = self.engine
        fname = self._wf_name(name) if e is not None else None
        disp = _WF_NAMES.get(name.lower()) or (name[:1].upper() + name[1:])
        if fname is None:
            raise rt_error(438)
        xargs = list(args)
        while xargs and xargs[-1] is MISSING:
            xargs.pop()
        conv = [_vba_to_xl(a) for a in xargs]
        try:
            from ..engine import FormulaError
        except Exception:
            FormulaError = None
        try:
            res = e.call_function(fname, conv, self.book.active)
        except Exception as ex:
            if FormulaError is not None and isinstance(ex, FormulaError):
                err = ex.error
            else:
                err = XLError("#VALUE!")
            if raise_errors:
                raise VBARuntimeErr(1004, f"Unable to get the {disp} property of the WorksheetFunction class")
            return VErr(XL_ERR_CODES.get(getattr(err, "code", "#VALUE!"), 2015))
        if isinstance(res, XLError):
            if raise_errors:
                raise VBARuntimeErr(1004, f"Unable to get the {disp} property of the WorksheetFunction class")
            return VErr(XL_ERR_CODES.get(res.code, 2015))
        return _xl_result(res)


def _split_commas(text):
    parts = []
    cur = []
    q = False
    for ch in text:
        if ch == "'":
            q = not q
        if ch == "," and not q:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return [p for p in parts if p.strip()]


_WF_NAMES = {n.lower(): n for n in (
    "VLookup", "HLookup", "XLookup", "Match", "Index", "CountIf", "CountIfs", "SumIf", "SumIfs",
    "AverageIf", "AverageIfs", "IfError", "CountA", "CountBlank", "SumProduct", "RoundUp",
    "RoundDown", "StDev", "StDev_S", "Var_S", "MaxIfs", "MinIfs", "EoMonth", "EDate", "NetworkDays",
    "WorkDay", "WeekNum", "Transpose", "TextJoin", "Proper", "Rept", "Large", "Small", "Rank",
    "Median", "Mode", "Ceiling", "Floor", "Lookup", "Choose", "Sum", "Max", "Min", "Average",
    "Count", "Round", "Trim", "Text", "Substitute", "Find", "Search", "Concat")}

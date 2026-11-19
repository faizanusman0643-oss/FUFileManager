"""Minimal stand-in for portable_excel.engine (GUI tests only)."""
import re
from portable_excel.model import (ERR_DIV0, ERR_NAME, ERR_VALUE, XLError, literal_value,
                                  parse_range_a1, value_to_text, a1, parse_a1, col_to_name,
                                  name_to_col)

FUNCTIONS = {"SUM": None, "AVERAGE": None, "COUNT": None, "MAX": None, "MIN": None,
             "IF": None, "ROUND": None, "CONCAT": None, "LEN": None, "TODAY": None}


class RangeRef:
    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2


_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def shift_formula(raw, dr, dc):
    if not raw.startswith("="):
        return raw

    def rep(m):
        c = m.group(2) if m.group(1) else col_to_name(max(0, name_to_col(m.group(2)) + dc))
        r = m.group(4) if m.group(3) else str(max(1, int(m.group(4)) + dr))
        return f"{m.group(1)}{c}{m.group(3)}{r}"
    return _REF.sub(rep, raw)


def insert_delete(book, sheet, axis, index, count):
    new = {}
    news = {}
    for (r, c), v in sheet.cells.items():
        k = r if axis == "row" else c
        if count < 0 and index <= k < index - count:
            continue
        if k >= index:
            k += count
        key = (k, c) if axis == "row" else (r, k)
        new[key] = v
        if (r, c) in sheet.styles:
            news[key] = sheet.styles[(r, c)]
    sheet.cells = new
    sheet.styles = news


def rename_sheet_refs(book, old, new):
    pass


def delete_sheet_refs(book, name):
    pass


class FormulaEngine:
    def __init__(self, book):
        self.book = book

    def invalidate(self):
        pass

    def recalc(self):
        pass

    def value(self, sheet, r, c, _depth=0):
        raw = sheet.get_raw(r, c)
        if not raw.startswith("="):
            return literal_value(raw)
        return self.evaluate(raw, sheet, _depth)

    def evaluate(self, formula, sheet, _depth=0):
        if _depth > 50:
            return XLError("#CYCLE!")
        expr = formula[1:].upper()

        def rng(m):
            r1, c1, r2, c2 = parse_range_a1(m.group(0))
            vals = [self.value(sheet, r, c, _depth + 1) for r in range(r1, r2 + 1)
                    for c in range(c1, c2 + 1)]
            return "[" + ",".join(str(v) for v in vals if isinstance(v, (int, float))) + "]"
        expr = re.sub(r"\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+", rng, expr)

        def ref(m):
            p = parse_a1(m.group(0))
            v = self.value(sheet, p[0], p[1], _depth + 1)
            return repr(v if v is not None else 0)
        expr = re.sub(r"\$?[A-Z]{1,3}\$?\d+", ref, expr)
        env = {"SUM": lambda *a: sum(_flat(a)), "MAX": lambda *a: max(_flat(a)),
               "MIN": lambda *a: min(_flat(a)), "AVERAGE": lambda *a: sum(_flat(a)) / len(_flat(a)),
               "COUNT": lambda *a: len(_flat(a))}
        try:
            return eval(expr.replace("^", "**"), {"__builtins__": {}}, env)
        except ZeroDivisionError:
            return ERR_DIV0
        except Exception:
            return ERR_NAME

    def display_text(self, sheet, r, c):
        v = self.value(sheet, r, c)
        fmt = sheet.get_style(r, c).get("number_format")
        if isinstance(v, float) and fmt and fmt not in ("General", "@"):
            if fmt.endswith("%"):
                d = len(fmt.split(".")[1]) - 1 if "." in fmt else 0
                return f"{v * 100:.{d}f}%"
            d = len(fmt.split(".")[1]) if "." in fmt else 0
            return f"{v:,.{d}f}" if "," in fmt else f"{v:.{d}f}"
        return value_to_text(v)


def _flat(a):
    out = []
    for x in a:
        if isinstance(x, (list, tuple)):
            out.extend(_flat(x))
        elif isinstance(x, (int, float)):
            out.append(x)
    return out

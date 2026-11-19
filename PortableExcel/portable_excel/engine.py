"""Formula engine for PortableExcel.

A recursive-descent parser and evaluator for Excel-style formulas, plus helpers
to rewrite formulas when cells are copied, rows/columns are inserted or deleted
and sheets are renamed or removed.

Public API (others code against it):
    RangeRef, FormulaEngine, FormulaError, format_value, shift_formula,
    insert_delete, rename_sheet_refs, delete_sheet_refs, FUNCTIONS

Values: None (empty), float, str, bool or model.XLError.
Dates are Excel serial numbers (1900 date system with the 1900 leap-year bug).
"""
from __future__ import annotations

import calendar
import datetime as _dt
import math
import random
import re
import sys
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, InvalidOperation
from functools import lru_cache
from typing import Callable, Dict, List, Optional

from .model import (
    ERR_CYCLE, ERR_DIV0, ERR_NA, ERR_NAME, ERR_NULL, ERR_NUM, ERR_REF,
    ERR_SYNTAX, ERR_VALUE, ERRORS_BY_CODE, MAX_COLS, MAX_ROWS, Workbook,
    XLError, col_to_name, literal_value, name_to_col, value_to_text,
)

__all__ = [
    "RangeRef", "FormulaEngine", "FormulaError", "format_value", "shift_formula",
    "insert_delete", "rename_sheet_refs", "delete_sheet_refs", "FUNCTIONS",
    "date_to_serial", "serial_to_date",
]

# cells deeper than this in a dependency chain are evaluated first by the
# driver loop instead of by deeper Python recursion
_DEPTH_LIMIT = 40


# ==========================================================================
# Errors / references
# ==========================================================================
class FormulaError(Exception):
    """Raised with an XLError; also used internally to propagate errors."""

    def __init__(self, error: XLError):
        super().__init__(error.code if isinstance(error, XLError) else str(error))
        self.error = error


class _Deep(BaseException):
    """Internal: dependency chain too deep, evaluate `key` first."""

    def __init__(self, key):
        super().__init__()
        self.key = key


class RangeRef:
    """A rectangular reference to cells of a Sheet (0-based, inclusive)."""

    __slots__ = ("sheet", "r1", "c1", "r2", "c2")

    def __init__(self, sheet, r1: int, c1: int, r2: int, c2: int):
        self.sheet = sheet
        self.r1, self.r2 = (r1, r2) if r1 <= r2 else (r2, r1)
        self.c1, self.c2 = (c1, c2) if c1 <= c2 else (c2, c1)

    @property
    def nrows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def ncols(self) -> int:
        return self.c2 - self.c1 + 1

    def __eq__(self, other):
        return (isinstance(other, RangeRef) and other.sheet is self.sheet
                and (other.r1, other.c1, other.r2, other.c2)
                == (self.r1, self.c1, self.r2, self.c2))

    def __hash__(self):
        return hash((id(self.sheet), self.r1, self.c1, self.r2, self.c2))

    def __repr__(self):
        name = getattr(self.sheet, "name", "?")
        return (f"RangeRef({name!r}, {col_to_name(self.c1)}{self.r1 + 1}:"
                f"{col_to_name(self.c2)}{self.r2 + 1})")


def _isnum(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ==========================================================================
# Tokenizer
# ==========================================================================
class _Tok:
    __slots__ = ("type", "text", "val")

    def __init__(self, type_, text, val=None):
        self.type = type_
        self.text = text
        self.val = val

    def __repr__(self):
        return f"Tok({self.type},{self.text!r})"


class _Ref:
    """A reference token as written: keeps $-flags and endpoint order."""

    __slots__ = ("sheet", "kind", "r1", "c1", "r2", "c2", "ar1", "ac1", "ar2", "ac2")

    def __init__(self, sheet, kind, r1=None, c1=None, r2=None, c2=None,
                 ar1=False, ac1=False, ar2=False, ac2=False):
        self.sheet = sheet
        self.kind = kind  # cell | area | cols | rows | err
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.ar1, self.ac1, self.ar2, self.ac2 = ar1, ac1, ar2, ac2

    def copy(self) -> "_Ref":
        return _Ref(self.sheet, self.kind, self.r1, self.c1, self.r2, self.c2,
                    self.ar1, self.ac1, self.ar2, self.ac2)

    def normalize(self) -> None:
        if self.r1 is not None and self.r2 is not None and self.r1 > self.r2:
            self.r1, self.r2, self.ar1, self.ar2 = self.r2, self.r1, self.ar2, self.ar1
        if self.c1 is not None and self.c2 is not None and self.c1 > self.c2:
            self.c1, self.c2, self.ac1, self.ac2 = self.c2, self.c1, self.ac2, self.ac1

    def text(self) -> str:
        pre = _quote_sheet(self.sheet) + "!" if self.sheet is not None else ""
        k = self.kind
        if k == "err":
            return pre + "#REF!"
        if k == "cell":
            return pre + _cell_text(self.r1, self.c1, self.ar1, self.ac1)
        if k == "area":
            return (pre + _cell_text(self.r1, self.c1, self.ar1, self.ac1) + ":"
                    + _cell_text(self.r2, self.c2, self.ar2, self.ac2))
        if k == "cols":
            return (pre + ("$" if self.ac1 else "") + col_to_name(self.c1) + ":"
                    + ("$" if self.ac2 else "") + col_to_name(self.c2))
        return (pre + ("$" if self.ar1 else "") + str(self.r1 + 1) + ":"
                + ("$" if self.ar2 else "") + str(self.r2 + 1))


def _cell_text(r, c, ar, ac) -> str:
    return ("$" if ac else "") + col_to_name(c) + ("$" if ar else "") + str(r + 1)


_SIMPLE_SHEET_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*\Z")
_LOOKS_A1_RE = re.compile(r"[A-Za-z]{1,3}\d+\Z")
_LOOKS_R1C1_RE = re.compile(r"[RrCc]\d*([RrCc]\d*)?\Z")


def _quote_sheet(name: str) -> str:
    if (_SIMPLE_SHEET_RE.match(name) and not _LOOKS_A1_RE.match(name)
            and not _LOOKS_R1C1_RE.match(name)
            and name.upper() not in ("TRUE", "FALSE")):
        return name
    return "'" + name.replace("'", "''") + "'"


_END = r"(?![A-Za-z0-9_(.!])"
_SHEET_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*))!")
_AREA_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+):(\$?)([A-Za-z]{1,3})(\$?)(\d+)" + _END)
_CELLR_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)" + _END)
_COLS_RE = re.compile(r"(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})" + _END)
_ROWS_RE = re.compile(r"(\$?)(\d+):(\$?)(\d+)" + _END)
_NUM_RE = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?")
_NAME_RE = re.compile(r"[A-Za-z_\\¡-￿][A-Za-z0-9_.\\¡-￿]*")
_ERR_RE = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A)", re.I)
_WS_RE = re.compile(r"\s+")


def _okrow(r):
    return 0 <= r < MAX_ROWS


def _okcol(c):
    return 0 <= c < MAX_COLS


def _match_body(s: str, i: int):
    m = _AREA_RE.match(s, i)
    if m:
        c1, r1 = name_to_col(m[2]), int(m[4]) - 1
        c2, r2 = name_to_col(m[6]), int(m[8]) - 1
        if _okrow(r1) and _okrow(r2) and _okcol(c1) and _okcol(c2):
            return m.end(), _Ref(None, "area", r1, c1, r2, c2,
                                 bool(m[3]), bool(m[1]), bool(m[7]), bool(m[5]))
    m = _CELLR_RE.match(s, i)
    if m:
        c1, r1 = name_to_col(m[2]), int(m[4]) - 1
        if _okrow(r1) and _okcol(c1):
            return m.end(), _Ref(None, "cell", r1, c1, r1, c1, bool(m[3]), bool(m[1]),
                                 bool(m[3]), bool(m[1]))
    m = _COLS_RE.match(s, i)
    if m:
        c1, c2 = name_to_col(m[2]), name_to_col(m[4])
        if _okcol(c1) and _okcol(c2):
            return m.end(), _Ref(None, "cols", None, c1, None, c2,
                                 False, bool(m[1]), False, bool(m[3]))
    m = _ROWS_RE.match(s, i)
    if m:
        r1, r2 = int(m[2]) - 1, int(m[4]) - 1
        if _okrow(r1) and _okrow(r2):
            return m.end(), _Ref(None, "rows", r1, None, r2, None,
                                 bool(m[1]), False, bool(m[3]), False)
    return None


class _SyntaxErr(Exception):
    pass


def _tokenize(s: str) -> List[_Tok]:
    toks: List[_Tok] = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch.isspace():
            m = _WS_RE.match(s, i)
            toks.append(_Tok("ws", m.group()))
            i = m.end()
            continue
        if ch == '"':
            j = i + 1
            buf = []
            while True:
                if j >= n:
                    raise _SyntaxErr("unterminated string")
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(s[j])
                j += 1
            toks.append(_Tok("str", s[i:j + 1], "".join(buf)))
            i = j + 1
            continue
        if ch == "#":
            m = _ERR_RE.match(s, i)
            if not m:
                raise _SyntaxErr("bad error literal")
            toks.append(_Tok("err", m.group(), ERRORS_BY_CODE[m.group().upper()]))
            i = m.end()
            continue
        if ch == "'" or ch == "$" or ch.isalnum() or ch == "_" or ch > " ":
            # sheet-qualified reference?
            m = _SHEET_RE.match(s, i)
            if m:
                sheet = m[1].replace("''", "'") if m[1] is not None else m[2]
                j = m.end()
                if s[j:j + 5].upper() == "#REF!":
                    toks.append(_Tok("ref", s[i:j + 5], _Ref(sheet, "err")))
                    i = j + 5
                    continue
                body = _match_body(s, j)
                if body is None:
                    raise _SyntaxErr("bad reference")
                end, ref = body
                ref.sheet = sheet
                toks.append(_Tok("ref", s[i:end], ref))
                i = end
                continue
            if ch == "'":
                raise _SyntaxErr("bad quote")
            body = _match_body(s, i)
            if body is not None:
                end, ref = body
                toks.append(_Tok("ref", s[i:end], ref))
                i = end
                continue
            if ch.isdigit():
                m = _NUM_RE.match(s, i)
                toks.append(_Tok("num", m.group(), float(m.group())))
                i = m.end()
                continue
            m = _NAME_RE.match(s, i)
            if m:
                toks.append(_Tok("name", m.group()))
                i = m.end()
                continue
            raise _SyntaxErr("unexpected character")
        if ch == ".":
            m = _NUM_RE.match(s, i)
            if not m:
                raise _SyntaxErr("bad number")
            toks.append(_Tok("num", m.group(), float(m.group())))
            i = m.end()
            continue
        two = s[i:i + 2]
        if two in ("<=", ">=", "<>"):
            toks.append(_Tok("op", two))
            i += 2
            continue
        if ch in "+-*/^&=<>%":
            toks.append(_Tok("op", ch))
            i += 1
            continue
        if ch in "(),;{}:":
            toks.append(_Tok(ch, ch))
            i += 1
            continue
        raise _SyntaxErr("unexpected character")
    return toks


# ==========================================================================
# Parser -> AST (tuples)
# ==========================================================================
#   ('num', x) ('str', s) ('bool', b) ('err', XLError) ('missing',) ('lit', v)
#   ('ref', sheetname|None, r1, c1, r2, c2)  ('name', text)
#   ('func', NAME, [args])  ('bin', op, a, b)  ('neg', a)  ('pct', a)
#   ('rng', a, b)  ('array', [[values]])
_CMP_OPS = ("=", "<>", "<", ">", "<=", ">=")


class _Parser:
    def __init__(self, toks):
        self.t = [k for k in toks if k.type != "ws"]
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else None

    def take(self):
        if self.i >= len(self.t):
            raise _SyntaxErr("unexpected end")
        tok = self.t[self.i]
        self.i += 1
        return tok

    def is_op(self, *ops):
        p = self.peek()
        return p is not None and p.type == "op" and p.text in ops

    def is_type(self, typ):
        p = self.peek()
        return p is not None and p.type == typ

    def expect(self, typ):
        tok = self.take()
        if tok.type != typ:
            raise _SyntaxErr(f"expected {typ}")
        return tok

    def parse(self):
        if not self.t:
            raise _SyntaxErr("empty")
        node = self.comparison()
        if self.peek() is not None:
            raise _SyntaxErr("trailing tokens")
        return node

    def comparison(self):
        n = self.concat()
        while self.is_op(*_CMP_OPS):
            op = self.take().text
            n = ("bin", op, n, self.concat())
        return n

    def concat(self):
        n = self.additive()
        while self.is_op("&"):
            self.take()
            n = ("bin", "&", n, self.additive())
        return n

    def additive(self):
        n = self.mult()
        while self.is_op("+", "-"):
            op = self.take().text
            n = ("bin", op, n, self.mult())
        return n

    def mult(self):
        n = self.power()
        while self.is_op("*", "/"):
            op = self.take().text
            n = ("bin", op, n, self.power())
        return n

    def power(self):
        n = self.unary()
        while self.is_op("^"):
            self.take()
            n = ("bin", "^", n, self.unary())
        return n

    def unary(self):
        if self.is_op("-"):
            self.take()
            return ("neg", self.unary())
        if self.is_op("+"):
            self.take()
            return self.unary()
        return self.postfix()

    def postfix(self):
        n = self.range_op()
        while self.is_op("%"):
            self.take()
            n = ("pct", n)
        return n

    def range_op(self):
        n = self.primary()
        while self.is_type(":"):
            self.take()
            n = ("rng", n, self.primary())
        return n

    def primary(self):
        tok = self.take()
        typ = tok.type
        if typ == "num":
            return ("num", tok.val)
        if typ == "str":
            return ("str", tok.val)
        if typ == "err":
            return ("err", tok.val)
        if typ == "ref":
            return _ref_node(tok.val)
        if typ == "name":
            if self.is_type("("):
                self.take()
                name = tok.text.upper()
                for pre in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
                    if name.startswith(pre):
                        name = name[len(pre):]
                args = []
                if self.is_type(")"):
                    self.take()
                    return ("func", name, args)
                while True:
                    if self.is_type(",") or self.is_type(")"):
                        args.append(("missing",))
                    else:
                        args.append(self.comparison())
                    sep = self.take()
                    if sep.type == ",":
                        continue
                    if sep.type == ")":
                        break
                    raise _SyntaxErr("expected , or )")
                return ("func", name, args)
            up = tok.text.upper()
            if up == "TRUE":
                return ("bool", True)
            if up == "FALSE":
                return ("bool", False)
            return ("name", tok.text)
        if typ == "(":
            n = self.comparison()
            self.expect(")")
            return n
        if typ == "{":
            return self.array()
        raise _SyntaxErr("unexpected token")

    def array(self):
        rows = [[]]
        while True:
            neg = False
            while self.is_op("-", "+"):
                if self.take().text == "-":
                    neg = not neg
            tok = self.take()
            if tok.type == "num":
                v = -tok.val if neg else tok.val
            elif tok.type == "str" and not neg:
                v = tok.val
            elif tok.type == "err" and not neg:
                v = tok.val
            elif tok.type == "name" and tok.text.upper() in ("TRUE", "FALSE") and not neg:
                v = tok.text.upper() == "TRUE"
            else:
                raise _SyntaxErr("bad array constant")
            rows[-1].append(v)
            sep = self.take()
            if sep.type == ",":
                continue
            if sep.type == ";":
                rows.append([])
                continue
            if sep.type == "}":
                break
            raise _SyntaxErr("bad array")
        if len({len(r) for r in rows}) != 1:
            raise _SyntaxErr("ragged array")
        return ("array", rows)


def _ref_node(ref: _Ref):
    k = ref.kind
    if k == "err":
        return ("err", ERR_REF)
    if k == "cell":
        return ("ref", ref.sheet, ref.r1, ref.c1, ref.r1, ref.c1)
    if k == "area":
        return ("ref", ref.sheet, min(ref.r1, ref.r2), min(ref.c1, ref.c2),
                max(ref.r1, ref.r2), max(ref.c1, ref.c2))
    if k == "cols":
        return ("ref", ref.sheet, 0, min(ref.c1, ref.c2), MAX_ROWS - 1, max(ref.c1, ref.c2))
    return ("ref", ref.sheet, min(ref.r1, ref.r2), 0, max(ref.r1, ref.r2), MAX_COLS - 1)


_PARSE_CACHE: Dict[str, object] = {}


def _parse(text: str):
    """Parse formula text (without leading '='). Raises FormulaError(ERR_SYNTAX)."""
    node = _PARSE_CACHE.get(text)
    if node is None:
        try:
            node = _Parser(_tokenize(text)).parse()
        except (_SyntaxErr, IndexError, ValueError, RecursionError):
            node = ERR_SYNTAX
        if len(_PARSE_CACHE) > 50000:
            _PARSE_CACHE.clear()
        _PARSE_CACHE[text] = node
    if node is ERR_SYNTAX:
        raise FormulaError(ERR_SYNTAX)
    return node


# ==========================================================================
# Value conversions
# ==========================================================================
_ISO_DT_RE = re.compile(
    r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d*)?))?)?\Z")
_TIME_RE = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d*)?))?\s*([AaPp][Mm])?\Z")


def _str_to_num(s: str) -> Optional[float]:
    t = s.strip()
    if not t:
        return None
    lv = literal_value(t)
    if _isnum(lv):
        return float(lv)
    if t[0] in "$" and len(t) > 1:
        lv = literal_value(t[1:].strip())
        if _isnum(lv):
            return float(lv)
    m = _ISO_DT_RE.match(t)
    if m:
        try:
            d = _dt.date(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            return None
        v = float(date_to_serial(d))
        if m[4]:
            v += (int(m[4]) * 3600 + int(m[5]) * 60 + float(m[6] or 0)) / 86400.0
        return v
    m = _TIME_RE.match(t)
    if m:
        h = int(m[1])
        if m[4]:
            h = h % 12 + (12 if m[4].upper() == "PM" else 0)
        return (h * 3600 + int(m[2]) * 60 + float(m[3] or 0)) / 86400.0
    return None


def _to_num(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, str):
        x = _str_to_num(v)
        if x is None:
            raise FormulaError(ERR_VALUE)
        return x
    raise FormulaError(ERR_VALUE)


def _num_text(x: float) -> str:
    """Number -> text as Excel does for concatenation (15 significant digits)."""
    if x != x or x in (math.inf, -math.inf):
        return "#NUM!"
    if float(x).is_integer() and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.15g}"
    if "e" in s:
        mant, exp = s.split("e")
        s = f"{mant}E{exp[0]}{exp[1:].lstrip('0').zfill(2)}"
    return s


def _to_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return _num_text(float(v))
    if isinstance(v, XLError):
        raise FormulaError(v)
    return str(v)


def _to_bool(v) -> bool:
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, str):
        u = v.strip().upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
    raise FormulaError(ERR_VALUE)


def _fin(x: float) -> float:
    if isinstance(x, complex) or x != x or x in (math.inf, -math.inf):
        raise FormulaError(ERR_NUM)
    return float(x)


def _rank(v) -> int:
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0


def _cmp(a, b) -> int:
    if isinstance(a, XLError):
        raise FormulaError(a)
    if isinstance(b, XLError):
        raise FormulaError(b)
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
    ra, rb = _rank(a), _rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.lower(), b.lower()
    elif ra == 0:
        a, b = float(a), float(b)
        if a != b and abs(a - b) <= 1e-15 * max(abs(a), abs(b)):
            return 0
    return (a > b) - (a < b)


def _bin_scalar(op, a, b):
    if op == "&":
        return _to_str(a) + _to_str(b)
    if op in _CMP_OPS:
        r = _cmp(a, b)
        if op == "=":
            return r == 0
        if op == "<>":
            return r != 0
        if op == "<":
            return r < 0
        if op == ">":
            return r > 0
        if op == "<=":
            return r <= 0
        return r >= 0
    x = _to_num(a)
    y = _to_num(b)
    if op == "+":
        return _fin(x + y)
    if op == "-":
        return _fin(x - y)
    if op == "*":
        return _fin(x * y)
    if op == "/":
        if y == 0:
            raise FormulaError(ERR_DIV0)
        return _fin(x / y)
    if op == "^":
        if x == 0 and y == 0:
            raise FormulaError(ERR_NUM)
        if x == 0 and y < 0:
            raise FormulaError(ERR_DIV0)
        try:
            return _fin(x ** y)
        except (OverflowError, ZeroDivisionError):
            raise FormulaError(ERR_NUM)
    raise FormulaError(ERR_VALUE)


# --------------------------------------------------------------------------
# Wildcards / criteria
# --------------------------------------------------------------------------
@lru_cache(maxsize=4096)
def _wild_re(pat: str, anchored: bool = True):
    out = []
    i = 0
    while i < len(pat):
        ch = pat[i]
        if ch == "~" and i + 1 < len(pat):
            out.append(re.escape(pat[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    body = "".join(out)
    return re.compile(body + (r"\Z" if anchored else ""), re.I | re.S)


def _has_wild(s: str) -> bool:
    return "*" in s or "?" in s or "~" in s


def _text_eq(pat: str, v: str, wildcards: bool = True) -> bool:
    if wildcards and _has_wild(pat):
        return _wild_re(pat).match(v) is not None
    return pat.lower() == v.lower()


def _criterion(c) -> Callable:
    """Build a predicate for SUMIF/COUNTIF style criteria."""
    if isinstance(c, XLError):
        return lambda v: v == c
    if isinstance(c, bool):
        return lambda v: isinstance(v, bool) and v == c
    if _isnum(c):
        x = float(c)
        return lambda v: _crit_num(v) == x
    s = "" if c is None else str(c)
    op = None
    for o in (">=", "<=", "<>", ">", "<", "="):
        if s.startswith(o):
            op, s = o, s[len(o):]
            break
    if s == "":
        if op in (None, "="):
            return lambda v: v is None or v == ""
        if op == "<>":
            return lambda v: not (v is None or v == "")
        return lambda v: False
    lv = literal_value(s)
    if _isnum(lv):
        x = float(lv)
        if op in (None, "="):
            return lambda v: _crit_num(v) == x
        if op == "<>":
            return lambda v: _crit_num(v) != x
        cmp = {">": lambda a: a > x, "<": lambda a: a < x,
               ">=": lambda a: a >= x, "<=": lambda a: a <= x}[op]
        return lambda v: _isnum(v) and cmp(float(v))
    if isinstance(lv, bool):
        if op in (None, "="):
            return lambda v: isinstance(v, bool) and v == lv
        if op == "<>":
            return lambda v: not (isinstance(v, bool) and v == lv)
        return lambda v: False
    if isinstance(lv, XLError):
        if op in (None, "="):
            return lambda v: v == lv
        if op == "<>":
            return lambda v: v != lv
    if op in (None, "="):
        return lambda v: isinstance(v, str) and _text_eq(s, v)
    if op == "<>":
        return lambda v: not (isinstance(v, str) and _text_eq(s, v))
    low = s.lower()
    cmp = {">": lambda a: a > low, "<": lambda a: a < low,
           ">=": lambda a: a >= low, "<=": lambda a: a <= low}[op]
    return lambda v: isinstance(v, str) and cmp(v.lower())


def _crit_num(v):
    if _isnum(v):
        return float(v)
    if isinstance(v, str):
        lv = literal_value(v.strip()) if v.strip() else None
        if _isnum(lv):
            return float(lv)
    return None


# ==========================================================================
# Dates
# ==========================================================================
_EPOCH = _dt.date(1899, 12, 30)
_MONTHS = ["January", "February", "March", "April", "May", "June", "July",
           "August", "September", "October", "November", "December"]
_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


def date_to_serial(d) -> int:
    """datetime.date -> Excel serial (1900 system, includes fake 1900-02-29)."""
    if isinstance(d, _dt.datetime):
        d = d.date()
    s = (d - _EPOCH).days
    return s - 1 if s < 61 else s


def serial_to_date(n: float):
    """Excel serial -> (year, month, day). Serial 60 -> (1900, 2, 29)."""
    n = int(math.floor(n))
    if n < 0:
        raise FormulaError(ERR_NUM)
    if n == 60:
        return 1900, 2, 29
    d = _EPOCH + _dt.timedelta(days=n + 1 if n < 60 else n)
    return d.year, d.month, d.day


def _pydate(n: float) -> _dt.date:
    y, m, d = serial_to_date(n)
    if (m, d) == (2, 29) and y == 1900:
        d = 28
    return _dt.date(y, m, d)


def _make_date(y: int, m: int, d: int) -> float:
    if 0 <= y < 1900:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    if not 1 <= y <= 9999:
        raise FormulaError(ERR_NUM)
    try:
        base = _dt.date(y, m, 1).toordinal() + d - 1
        dd = _dt.date.fromordinal(base)
    except (ValueError, OverflowError):
        raise FormulaError(ERR_NUM)
    s = date_to_serial(dd)
    if s < 0:
        raise FormulaError(ERR_NUM)
    return float(s)


def _split_time(x: float):
    days = math.floor(x)
    secs = int(round((x - days) * 86400))
    if secs >= 86400:
        days += 1
        secs -= 86400
    return days, secs // 3600, secs % 3600 // 60, secs % 60


# ==========================================================================
# Number formatting
# ==========================================================================
def _split_sections(fmt: str) -> List[str]:
    secs, buf, q, i = [], [], False, 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == '"':
            q = not q
        elif ch == "\\" and not q and i + 1 < len(fmt):
            buf.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == ";" and not q:
            secs.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    secs.append("".join(buf))
    return secs


def _strip_literals(sec: str) -> str:
    s = re.sub(r'"[^"]*"?', "", sec)
    s = re.sub(r"\\.", "", s)
    s = re.sub(r"[_*].", "", s)
    s = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", s, flags=re.I)
    return s


def _is_date_fmt(sec: str) -> bool:
    s = _strip_literals(sec)
    if re.search(r"[0#?]", s):
        return False
    return re.search(r"[ymdhs]", s, re.I) is not None


def _fmt_tokens(sec: str):
    out = []
    i, n = 0, len(sec)
    while i < n:
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            out.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\":
            out.append(("lit", sec[i + 1:i + 2]))
            i += 2
        elif ch == "_":
            out.append(("lit", " "))
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = n if j < 0 else j
            content = sec[i + 1:j]
            if content.startswith("$"):
                out.append(("lit", content[1:].split("-")[0]))
            elif content and content[0].lower() in "hms" and len(set(content.lower())) == 1:
                out.append(("el", content[0].lower(), len(content)))
            i = j + 1
        else:
            out.append(("ch", ch))
            i += 1
    return out


@lru_cache(maxsize=1024)
def _num_layout(sec: str):
    toks = _fmt_tokens(sec)
    items = []  # ('lit', s) | ('ph', ch) | ('pct',) | ('sci', sign)
    k = 0
    while k < len(toks):
        t = toks[k]
        if t[0] == "ch":
            ch = t[1]
            if ch in "0#?.,":
                items.append(("ph", ch))
            elif ch in "Ee" and k + 1 < len(toks) and toks[k + 1] in (("ch", "+"), ("ch", "-")) \
                    and any(x[0] == "ph" for x in items):
                items.append(("sci", toks[k + 1][1]))
                k += 1
            elif ch == "%":
                items.append(("pct",))
            elif ch == "@":
                pass
            else:
                items.append(("lit", ch))
        elif t[0] == "lit":
            items.append(t)
        k += 1
    ph_idx = [i for i, it in enumerate(items) if it[0] in ("ph", "sci")]
    pct = sum(1 for it in items if it[0] == "pct")
    if not ph_idx:
        text = "".join(it[1] if it[0] == "lit" else "%" for it in items if it[0] in ("lit", "pct"))
        return None, text, "", pct
    first, last = ph_idx[0], ph_idx[-1]
    prefix = "".join(it[1] if it[0] == "lit" else "%" for it in items[:first]
                     if it[0] in ("lit", "pct"))
    suffix = "".join(it[1] if it[0] == "lit" else "%" for it in items[first:]
                     if it[0] in ("lit", "pct"))
    pattern = "".join(it[1] if it[0] == "ph" else "E" + it[1] for it in items[first:last + 1]
                      if it[0] in ("ph", "sci"))
    return pattern, prefix, suffix, pct


def _group(digits: str) -> str:
    head = digits.lstrip("0") or ""
    zeros = digits[:len(digits) - len(head)]
    if not head:
        return digits
    out = []
    while len(head) > 3:
        out.insert(0, head[-3:])
        head = head[:-3]
    out.insert(0, head)
    return zeros + ",".join(out)


def _render_fixed(x: float, int_pat: str, frac_pat: str, has_dot: bool, thousands: bool) -> str:
    pats = [ch for ch in frac_pat if ch in "0#?"]
    dec = len(pats)
    q = Decimal(repr(x)).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP)
    s = format(q, "f")
    ip, _, fp = s.partition(".")
    fp = fp.ljust(dec, "0")[:dec]
    fl = list(fp)
    k = len(fl) - 1
    while k >= 0 and pats[k] != "0" and fl[k] == "0":
        fl[k] = " " if pats[k] == "?" else ""
        k -= 1
    fp = "".join(fl)
    min_int = int_pat.count("0")
    if ip == "0" and min_int == 0:
        ip = ""
    ip = ip.zfill(min_int)
    if thousands and ip:
        ip = _group(ip)
    return ip + ("." + fp if has_dot else "")


def _format_number(x: float, sec: str, neg_auto: bool) -> str:
    pattern, prefix, suffix, pct = _num_layout(sec)
    x = x * (100 ** pct)
    if pattern is None:
        return ("-" if neg_auto else "") + prefix
    sci = None
    m = re.search(r"E([+-])", pattern)
    if m:
        sci = m.group(1)
        mant_pat, exp_pat = pattern[:m.start()], pattern[m.end():]
    else:
        mant_pat, exp_pat = pattern, ""
    int_pat, dot, frac_pat = mant_pat.partition(".")
    scale = 0
    while int_pat.endswith(","):
        int_pat = int_pat[:-1]
        scale += 1
    thousands = "," in int_pat
    int_pat = int_pat.replace(",", "")
    frac_pat = frac_pat.replace(",", "")
    x = x / (1000 ** scale)
    if sci:
        n_int = max(1, sum(ch in "0#?" for ch in int_pat))
        exp = 0 if x == 0 else math.floor(math.log10(x))
        if n_int > 1:
            exp -= exp % n_int
        mant = x / (10 ** exp) if x else 0.0
        body = _render_fixed(mant, int_pat or "0", frac_pat, bool(dot), False)
        if x and float(body.replace(" ", "") or 0) >= 10 ** n_int:
            exp += n_int
            body = _render_fixed(x / (10 ** exp), int_pat or "0", frac_pat, bool(dot), False)
        esign = "-" if exp < 0 else ("+" if sci == "+" else "")
        body += "E" + esign + str(abs(exp)).zfill(exp_pat.count("0") or 1)
    else:
        body = _render_fixed(x, int_pat, frac_pat, bool(dot), thousands)
    sign = "-" if neg_auto and re.search(r"[1-9]", body) else ""
    return sign + prefix + body + suffix


def _date_tokens(sec: str):
    toks = []
    i, n = 0, len(sec)
    low = sec.lower()
    while i < n:
        ch, lc = sec[i], low[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            toks.append(("lit", sec[i + 1:j]))
            i = j + 1
        elif ch == "\\":
            toks.append(("lit", sec[i + 1:i + 2]))
            i += 2
        elif ch == "_":
            toks.append(("lit", " "))
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = n if j < 0 else j
            c = low[i + 1:j]
            if c and c[0] in "hms" and len(set(c)) == 1:
                toks.append(("el", c[0], len(c)))
            i = j + 1
        elif low.startswith("am/pm", i):
            toks.append(("ampm", "AMPM"))
            i += 5
        elif low.startswith("a/p", i):
            toks.append(("ampm", "AP"))
            i += 3
        elif lc in "ymdhs":
            j = i
            while j < n and low[j] == lc:
                j += 1
            toks.append(("code", lc, j - i))
            i = j
        else:
            toks.append(("lit", ch))
            i += 1
    codes = [k for k, t in enumerate(toks) if t[0] in ("code", "el")]
    for pos, k in enumerate(codes):
        t = toks[k]
        if t[0] == "code" and t[1] == "m":
            prev = toks[codes[pos - 1]] if pos > 0 else None
            nxt = toks[codes[pos + 1]] if pos + 1 < len(codes) else None
            if (prev is not None and prev[1] == "h") or (nxt is not None and nxt[1] == "s"):
                toks[k] = ("code", "M", t[2])
    return toks


def _format_date(x: float, sec: str) -> str:
    if x < 0:
        return value_to_text(x)
    days, h, mi, s = _split_time(x)
    y, mo, d = serial_to_date(days)
    toks = _date_tokens(sec)
    ampm = any(t[0] == "ampm" for t in toks)
    out = []
    for t in toks:
        kind = t[0]
        if kind == "lit":
            out.append(t[1])
        elif kind == "ampm":
            if t[1] == "AMPM":
                out.append("AM" if h < 12 else "PM")
            else:
                out.append("A" if h < 12 else "P")
        elif kind == "el":
            total = {"h": days * 24 + h, "m": (days * 24 + h) * 60 + mi,
                     "s": ((days * 24 + h) * 60 + mi) * 60 + s}[t[1]]
            out.append(str(total).zfill(t[2]))
        else:
            c, cnt = t[1], t[2]
            if c == "y":
                out.append(f"{y % 100:02d}" if cnt <= 2 else f"{y:04d}")
            elif c == "m":
                if cnt == 1:
                    out.append(str(mo))
                elif cnt == 2:
                    out.append(f"{mo:02d}")
                elif cnt == 3:
                    out.append(_MONTHS[mo - 1][:3])
                elif cnt == 5:
                    out.append(_MONTHS[mo - 1][0])
                else:
                    out.append(_MONTHS[mo - 1])
            elif c == "d":
                if cnt == 1:
                    out.append(str(d))
                elif cnt == 2:
                    out.append(f"{d:02d}")
                else:
                    name = _DAYS[(days - 1) % 7] if days > 0 else _DAYS[6]
                    out.append(name[:3] if cnt == 3 else name)
            elif c == "h":
                hh = (h % 12 or 12) if ampm else h
                out.append(f"{hh:02d}" if cnt >= 2 else str(hh))
            elif c == "M":
                out.append(f"{mi:02d}" if cnt >= 2 else str(mi))
            elif c == "s":
                out.append(f"{s:02d}" if cnt >= 2 else str(s))
    return "".join(out)


def format_value(v, number_format: Optional[str]) -> str:
    """Format a cell value with an Excel number format string."""
    if v is None:
        return ""
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    fmt = number_format
    if not fmt or fmt.strip().lower() == "general":
        return value_to_text(v)
    try:
        secs = _split_sections(fmt)
        if isinstance(v, str):
            sec = secs[3] if len(secs) >= 4 else (secs[0] if "@" in secs[0] else None)
            if sec is None:
                return v
            out = []
            for t in _fmt_tokens(sec):
                if t[0] == "lit":
                    out.append(t[1])
                elif t[0] == "ch":
                    out.append(v if t[1] == "@" else t[1])
            return "".join(out)
        if not _isnum(v):
            return value_to_text(v)
        x = float(v)
        neg_auto = False
        if len(secs) >= 2 and x < 0:
            sec, x = secs[1], -x
        elif len(secs) >= 3 and x == 0:
            sec = secs[2]
        else:
            sec = secs[0]
            if x < 0:
                neg_auto, x = True, -x
        if sec.strip().lower() == "general":
            return ("-" if neg_auto else "") + value_to_text(x)
        if _is_date_fmt(sec):
            return _format_date(-x if neg_auto else x, sec)
        return _format_number(x, sec, neg_auto)
    except Exception:
        return value_to_text(v)


# ==========================================================================
# Evaluation
# ==========================================================================
FUNCTIONS: Dict[str, Callable] = {}
_LAZY = set()


def _fn(*names, lazy=False):
    def deco(f):
        for n in names:
            FUNCTIONS[n] = f
            if lazy:
                _LAZY.add(n)
        return f
    return deco


class _Eval:
    """Evaluation context: engine + sheet + (optional) the formula's cell."""

    __slots__ = ("eng", "sheet", "r", "c", "names")

    def __init__(self, eng, sheet, r, c):
        self.eng = eng
        self.sheet = sheet
        self.r = r
        self.c = c
        self.names = set()

    # -- cells ------------------------------------------------------------
    def cell(self, sheet, r, c):
        if not (0 <= r < MAX_ROWS and 0 <= c < MAX_COLS):
            raise FormulaError(ERR_REF)
        return self.eng.value(sheet, r, c)

    def range_values(self, rr: RangeRef):
        """Values of the non-empty cells in rr, row-major."""
        eng = self.eng
        for r, c in eng._existing(rr.sheet, rr.r1, rr.c1, rr.r2, rr.c2):
            v = eng.value(rr.sheet, r, c)
            if v is not None:
                yield v

    def dims(self, rr: RangeRef):
        nr, nc = rr.nrows, rr.ncols
        if nr == MAX_ROWS or nc == MAX_COLS:
            ur, uc = self.eng._used(rr.sheet)
            if nr == MAX_ROWS:
                nr = max(0, ur - rr.r1)
            if nc == MAX_COLS:
                nc = max(0, uc - rr.c1)
        return nr, nc

    def grid(self, v):
        if isinstance(v, RangeRef):
            nr, nc = self.dims(v)
            sh = v.sheet
            val = self.eng.value
            return [[val(sh, v.r1 + i, v.c1 + j) for j in range(nc)] for i in range(nr)]
        if isinstance(v, list):
            return v
        return [[v]]

    # -- nodes ------------------------------------------------------------
    def ev(self, n):
        t = n[0]
        if t == "num" or t == "str" or t == "bool" or t == "err" or t == "lit":
            return n[1]
        if t == "ref":
            if n[1] is None:
                sh = self.sheet
            else:
                sh = self.eng.book.sheet_by_name(n[1])
                if sh is None:
                    raise FormulaError(ERR_REF)
            return RangeRef(sh, n[2], n[3], n[4], n[5])
        if t == "bin":
            return self.binop(n[1], self.ev(n[2]), self.ev(n[3]))
        if t == "func":
            return self.call(n[1], n[2])
        if t == "neg":
            return self.unop(lambda x: -_to_num(x) if x is not None else -0.0, self.ev(n[1]))
        if t == "pct":
            return self.unop(lambda x: _to_num(x) / 100.0, self.ev(n[1]))
        if t == "missing":
            return None
        if t == "name":
            return self.name(n[1])
        if t == "array":
            return [list(r) for r in n[1]]
        if t == "rng":
            a, b = self.ev(n[1]), self.ev(n[2])
            if not (isinstance(a, RangeRef) and isinstance(b, RangeRef)) or a.sheet is not b.sheet:
                raise FormulaError(ERR_VALUE)
            return RangeRef(a.sheet, min(a.r1, b.r1), min(a.c1, b.c1),
                            max(a.r2, b.r2), max(a.c2, b.c2))
        raise FormulaError(ERR_VALUE)

    def name(self, name):
        dn = self.eng.book.defined_names
        low = name.lower()
        text = None
        for k, v in dn.items():
            if k.lower() == low:
                text = v
                break
        if text is None:
            raise FormulaError(ERR_NAME)
        if low in self.names:
            raise FormulaError(ERR_CYCLE)
        text = str(text).strip()
        if text.startswith("="):
            text = text[1:]
        self.names.add(low)
        try:
            return self.ev(_parse(text))
        finally:
            self.names.discard(low)

    def call(self, name, nodes):
        f = FUNCTIONS.get(name)
        if f is None:
            udf = getattr(self.eng, "udf_resolver", None)
            if udf is None:
                raise FormulaError(ERR_NAME)
            args = []
            for nd in nodes:
                try:
                    args.append(self.ev(nd))
                except FormulaError as e:
                    args.append(e.error)
            res = udf(name, args, self.sheet)
            if isinstance(res, list):
                res = res[0][0] if res and isinstance(res[0], list) and res[0] else \
                    (res[0] if res else None)
            if isinstance(res, XLError):
                raise FormulaError(res)
            if isinstance(res, bool) or isinstance(res, str) or res is None:
                return "" if res is None else res
            if isinstance(res, (int, float)):
                return float(res)
            return str(res)
        if name in _LAZY:
            return f(self, nodes)
        args = []
        for nd in nodes:
            try:
                args.append(self.ev(nd))
            except FormulaError as e:
                args.append(e.error)
        return f(self, args)

    @staticmethod
    def _multi(v):
        return isinstance(v, list) or (isinstance(v, RangeRef) and (v.r1 != v.r2 or v.c1 != v.c2))

    def binop(self, op, a, b):
        if self._multi(a) or self._multi(b):
            ga, gb = self.grid(a), self.grid(b)
            R = max(len(ga), len(gb))
            C = max(len(ga[0]) if ga else 0, len(gb[0]) if gb else 0)
            out = []
            for i in range(R):
                row = []
                for j in range(C):
                    try:
                        x = _bget(ga, i, j)
                        y = _bget(gb, i, j)
                        row.append(_bin_scalar(op, x, y))
                    except FormulaError as e:
                        row.append(e.error)
                out.append(row)
            return out
        return _bin_scalar(op, _sc(self, a), _sc(self, b))

    def unop(self, f, a):
        if self._multi(a):
            out = []
            for row in self.grid(a):
                r2 = []
                for x in row:
                    try:
                        r2.append(_fin(f(x)))
                    except FormulaError as e:
                        r2.append(e.error)
                out.append(r2)
            return out
        return _fin(f(_sc(self, a)))


def _bget(g, i, j):
    if not g:
        return None
    if len(g) == 1:
        i = 0
    if len(g[0]) == 1:
        j = 0
    if i >= len(g) or j >= len(g[i]):
        raise FormulaError(ERR_NA)
    return g[i][j]


def _sc(E: _Eval, v):
    """Reduce a function argument / operand to a scalar value."""
    if isinstance(v, RangeRef):
        if v.r1 == v.r2 and v.c1 == v.c2:
            return E.cell(v.sheet, v.r1, v.c1)
        if E.r is not None and v.sheet is E.sheet:
            if v.c1 == v.c2 and v.r1 <= E.r <= v.r2:
                return E.cell(v.sheet, E.r, v.c1)
            if v.r1 == v.r2 and v.c1 <= E.c <= v.c2:
                return E.cell(v.sheet, v.r1, E.c)
        raise FormulaError(ERR_VALUE)
    if isinstance(v, list):
        if v and v[0]:
            return _sc(E, v[0][0])
        return None
    return v


# ==========================================================================
# The engine
# ==========================================================================
_NOVAL = object()


class FormulaEngine:
    FUNCTIONS = FUNCTIONS

    def __init__(self, book: Workbook):
        self.book = book
        self._cache: Dict[tuple, object] = {}
        self._active = set()
        self._depth = 0
        self._used_cache: Dict[int, tuple] = {}
        self._index_cache: Dict[int, dict] = {}
        # callable(name, args, sheet) -> value for functions the engine does
        # not know (VBA user-defined functions). args: scalars, RangeRef,
        # nested lists. Return ERR_NAME for unknown names.
        self.udf_resolver = None
        if sys.getrecursionlimit() < 4000:
            sys.setrecursionlimit(4000)

    # -- cache ------------------------------------------------------------
    def invalidate(self) -> None:
        self._cache.clear()
        self._used_cache.clear()
        self._index_cache.clear()

    def recalc(self) -> None:
        self.invalidate()
        for sh in list(self.book.sheets):
            for key in sorted(sh.cells):
                raw = sh.cells.get(key)
                if raw and raw.startswith("=") and len(raw) > 1:
                    self.value(sh, key[0], key[1])

    def _used(self, sheet):
        u = self._used_cache.get(id(sheet))
        if u is None:
            u = sheet.used_bounds()
            self._used_cache[id(sheet)] = u
        return u

    def _index(self, sheet):
        idx = self._index_cache.get(id(sheet))
        if idx is None:
            idx = {}
            for (r, c) in sheet.cells:
                idx.setdefault(c, []).append(r)
            for lst in idx.values():
                lst.sort()
            self._index_cache[id(sheet)] = idx
        return idx

    def _existing(self, sheet, r1, c1, r2, c2):
        """(r, c) of non-empty cells inside the rectangle, row-major."""
        area = (r2 - r1 + 1) * (c2 - c1 + 1)
        cells = sheet.cells
        if area <= 4096 or area <= len(cells) // 4:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) in cells:
                        yield r, c
            return
        import bisect
        idx = self._index(sheet)
        if c2 - c1 + 1 > len(idx):
            cols = sorted(c for c in idx if c1 <= c <= c2)
        else:
            cols = [c for c in range(c1, c2 + 1) if c in idx]
        if len(cols) == 1:
            c = cols[0]
            rows = idx[c]
            lo, hi = bisect.bisect_left(rows, r1), bisect.bisect_right(rows, r2)
            for k in range(lo, hi):
                if (rows[k], c) in cells:
                    yield rows[k], c
            return
        pairs = []
        for c in cols:
            rows = idx[c]
            lo, hi = bisect.bisect_left(rows, r1), bisect.bisect_right(rows, r2)
            pairs.extend((rows[k], c) for k in range(lo, hi))
        pairs.sort()
        for p in pairs:
            if p in cells:
                yield p

    # -- values -----------------------------------------------------------
    def value(self, sheet, r: int, c: int):
        key = (sheet, r, c)
        v = self._cache.get(key, _NOVAL)
        if v is not _NOVAL:
            return v
        raw = sheet.cells.get((r, c))
        if raw is None or raw == "":
            return None
        if not (raw.startswith("=") and len(raw) > 1):
            v = literal_value(raw)
            if isinstance(v, int) and not isinstance(v, bool):
                v = float(v)
            self._cache[key] = v
            return v
        if self._depth == 0:
            return self._drive(key)
        if self._depth >= _DEPTH_LIMIT and key not in self._active:
            raise _Deep(key)
        return self._compute(key, raw)

    def _drive(self, key):
        stack = [key]
        seen = {key}
        while stack:
            k = stack[-1]
            if k in self._cache:
                stack.pop()
                continue
            raw = k[0].cells.get((k[1], k[2]))
            if not raw or not raw.startswith("=") or len(raw) < 2:
                stack.pop()
                continue
            try:
                self._compute(k, raw)
                stack.pop()
            except _Deep as d:
                if d.key in seen:
                    self._cache[d.key] = ERR_CYCLE
                else:
                    seen.add(d.key)
                    stack.append(d.key)
        return self._cache.get(key, ERR_CYCLE)

    def _compute(self, key, raw):
        if key in self._active:
            return ERR_CYCLE
        sheet, r, c = key
        self._active.add(key)
        self._depth += 1
        try:
            try:
                ast = _parse(raw[1:])
                E = _Eval(self, sheet, r, c)
                v = self._cell_result(E, E.ev(ast))
            except FormulaError as e:
                v = e.error
            except RecursionError:
                v = ERR_CYCLE
            except Exception:
                v = ERR_VALUE
        finally:
            self._depth -= 1
            self._active.discard(key)
        self._cache[key] = v
        return v

    @staticmethod
    def _cell_result(E, res):
        if isinstance(res, (RangeRef, list)):
            res = _sc(E, res)
        if res is None:
            return 0.0
        if isinstance(res, bool) or isinstance(res, (str, XLError)):
            return res
        if isinstance(res, (int, float)):
            return float(res)
        return ERR_VALUE

    def display_text(self, sheet, r: int, c: int) -> str:
        v = self.value(sheet, r, c)
        return format_value(v, sheet.get_style(r, c).get("number_format"))

    def _export(self, E, res):
        if isinstance(res, RangeRef):
            if res.nrows == 1 and res.ncols == 1:
                return self.value(res.sheet, res.r1, res.c1)
            return E.grid(res)
        if isinstance(res, list):
            return [[_plain(x) for x in row] for row in res]
        return _plain(res)

    def evaluate(self, formula: str, sheet=None):
        sheet = sheet if sheet is not None else self.book.active
        text = formula.strip()
        if text.startswith("="):
            text = text[1:]
        try:
            ast = _parse(text)
            E = _Eval(self, sheet, None, None)
            return self._export(E, E.ev(ast))
        except FormulaError as e:
            return e.error
        except RecursionError:
            return ERR_CYCLE
        except Exception:
            return ERR_VALUE

    def call_function(self, name: str, args: list, sheet=None):
        name = name.upper()
        for pre in ("_XLFN.", "_XLWS."):
            if name.startswith(pre):
                name = name[len(pre):]
        f = FUNCTIONS.get(name)
        if f is None:
            raise FormulaError(ERR_NAME)
        sheet = sheet if sheet is not None else self.book.active
        E = _Eval(self, sheet, None, None)
        conv = [_from_python(a) for a in args]
        try:
            if name in _LAZY:
                res = f(E, [("lit", a) for a in conv])
            else:
                res = f(E, conv)
            res = self._export(E, res)
        except FormulaError:
            raise
        except Exception:
            raise FormulaError(ERR_VALUE)
        if isinstance(res, XLError):
            raise FormulaError(res)
        return res


def _plain(v):
    if isinstance(v, int) and not isinstance(v, bool):
        return float(v)
    return v


def _from_python(a):
    if isinstance(a, RangeRef) or a is None or isinstance(a, (str, bool, XLError)):
        return a
    if isinstance(a, (int, float)):
        return float(a)
    if isinstance(a, (list, tuple)):
        if not a:
            return [[None]]
        if any(isinstance(x, (list, tuple)) for x in a):
            rows = [[_from_python(x) for x in (row if isinstance(row, (list, tuple)) else [row])]
                    for row in a]
            w = max(len(r) for r in rows)
            return [r + [None] * (w - len(r)) for r in rows]
        return [[_from_python(x) for x in a]]
    if isinstance(a, (_dt.datetime, _dt.date)):
        v = float(date_to_serial(a))
        if isinstance(a, _dt.datetime):
            v += (a.hour * 3600 + a.minute * 60 + a.second) / 86400.0
        return v
    return str(a)


# ==========================================================================
# Function helpers
# ==========================================================================
def _num(E, v) -> float:
    return _to_num(_sc(E, v))


def _int(E, v) -> int:
    x = _num(E, v)
    return int(math.floor(x + 1e-12)) if x >= 0 else int(math.ceil(x - 1e-12))


def _txt(E, v) -> str:
    return _to_str(_sc(E, v))


def _bool(E, v) -> bool:
    return _to_bool(_sc(E, v))


def _opt(args, i, default=None):
    return args[i] if i < len(args) else default


def _err(e=ERR_VALUE):
    raise FormulaError(e)


def _need(args, lo, hi=None):
    if len(args) < lo or (hi is not None and len(args) > hi):
        raise FormulaError(ERR_VALUE)


def _nums(E, args) -> List[float]:
    out = []
    for a in args:
        if isinstance(a, RangeRef):
            for v in E.range_values(a):
                if _isnum(v):
                    out.append(float(v))
                elif isinstance(v, XLError):
                    raise FormulaError(v)
        elif isinstance(a, list):
            for row in a:
                for v in row:
                    if _isnum(v):
                        out.append(float(v))
                    elif isinstance(v, XLError):
                        raise FormulaError(v)
        else:
            out.append(_to_num(a))
    return out


def _all_values(E, args, include_empty=False):
    """Flatten args to scalar values (ranges: existing cells unless include_empty)."""
    for a in args:
        if isinstance(a, RangeRef):
            if include_empty:
                for row in E.grid(a):
                    yield from row
            else:
                yield from E.range_values(a)
        elif isinstance(a, list):
            for row in a:
                yield from row
        else:
            yield a


class _Area:
    """2-D view over a RangeRef or an array, used by lookups and *IF(S)."""

    def __init__(self, E, v):
        self.E = E
        if isinstance(v, RangeRef):
            self.rr, self.grid = v, None
            self.nrows, self.ncols = v.nrows, v.ncols
        else:
            g = v if isinstance(v, list) else [[_sc(E, v)]]
            self.rr, self.grid = None, g
            self.nrows, self.ncols = len(g), (len(g[0]) if g else 0)

    def get(self, i, j):
        if self.grid is not None:
            if 0 <= i < self.nrows and 0 <= j < len(self.grid[i]):
                return self.grid[i][j]
            return None
        r, c = self.rr.r1 + i, self.rr.c1 + j
        if r >= MAX_ROWS or c >= MAX_COLS:
            return None
        return self.E.eng.value(self.rr.sheet, r, c)

    def existing(self):
        if self.grid is not None:
            for i, row in enumerate(self.grid):
                for j, v in enumerate(row):
                    if v is not None:
                        yield i, j, v
            return
        rr = self.rr
        eng = self.E.eng
        for r, c in eng._existing(rr.sheet, rr.r1, rr.c1, rr.r2, rr.c2):
            v = eng.value(rr.sheet, r, c)
            if v is not None:
                yield r - rr.r1, c - rr.c1, v

    def dense(self):
        if self.grid is not None:
            return self.nrows, self.ncols
        return self.E.dims(self.rr)

    def sub_col(self, j):
        if self.grid is not None:
            return _Area(self.E, [[row[j] if j < len(row) else None] for row in self.grid])
        rr = self.rr
        return _Area(self.E, RangeRef(rr.sheet, rr.r1, rr.c1 + j, rr.r2, rr.c1 + j))

    def sub_row(self, i):
        if self.grid is not None:
            return _Area(self.E, [list(self.grid[i])])
        rr = self.rr
        return _Area(self.E, RangeRef(rr.sheet, rr.r1 + i, rr.c1, rr.r1 + i, rr.c2))

    # vector view (single row or column)
    @property
    def is_vector(self):
        return self.nrows == 1 or self.ncols == 1

    def vlen(self):
        nr, nc = self.dense()
        return nr if self.ncols == 1 else nc

    def vget(self, k):
        return self.get(k, 0) if self.ncols == 1 else self.get(0, k)

    def vexisting(self):
        col = self.ncols == 1
        for i, j, v in self.existing():
            yield (i if col else j), v

    def result_at(self, k, along_rows):
        """Value / sub-range at position k of a lookup result vector."""
        if self.rr is not None:
            rr = self.rr
            if along_rows:
                return RangeRef(rr.sheet, rr.r1 + k, rr.c1, rr.r1 + k, rr.c2)
            return RangeRef(rr.sheet, rr.r1, rr.c1 + k, rr.r2, rr.c1 + k)
        if along_rows:
            return [list(self.grid[k])] if k < self.nrows else _err(ERR_REF)
        return [[row[k]] for row in self.grid]


def _lookup_eq(val, v, wildcards=True) -> bool:
    if isinstance(val, str):
        return isinstance(v, str) and _text_eq(val, v, wildcards)
    if isinstance(val, bool):
        return isinstance(v, bool) and v == val
    if _isnum(val):
        return _isnum(v) and _cmp(float(val), float(v)) == 0
    if val is None:
        return v is None
    return False


def _same_type(a, b) -> bool:
    if a is None or b is None or isinstance(a, XLError) or isinstance(b, XLError):
        return False
    return _rank(a) == _rank(b)


def _vsearch(area: _Area, val, mode: int, reverse=False, wildcards=True):
    """Return index of match in vector area or None.
    mode 0 exact, 1 = largest <= val (sorted asc), -1 = smallest >= val (sorted desc),
    'next_smaller' / 'next_larger' for XLOOKUP (unsorted)."""
    if isinstance(val, XLError):
        raise FormulaError(val)
    if mode == 0:
        if reverse:
            hits = [k for k, v in area.vexisting() if _lookup_eq(val, v, wildcards)]
            return hits[-1] if hits else None
        if val is None or val == "":
            for k in range(area.vlen()):
                if _lookup_eq(val, area.vget(k), wildcards) or (val == "" and area.vget(k) is None):
                    return k
            return None
        for k, v in area.vexisting():
            if _lookup_eq(val, v, wildcards):
                return k
        return None
    if mode in (1, -1):
        best = None
        for k in range(area.vlen()):
            v = area.vget(k)
            if not _same_type(val, v):
                continue
            r = _cmp(v, val)
            if mode == 1:
                if r <= 0:
                    best = k
                else:
                    break
            else:
                if r >= 0:
                    best = k
                else:
                    break
        return best
    # XLOOKUP approximate (unsorted): exact or next smaller/larger
    best, bestv = None, None
    seq = list(area.vexisting())
    if reverse:
        seq.reverse()
    for k, v in seq:
        if not _same_type(val, v):
            continue
        r = _cmp(v, val)
        if r == 0:
            return k
        if mode == "next_smaller" and r < 0 and (bestv is None or _cmp(v, bestv) > 0):
            best, bestv = k, v
        if mode == "next_larger" and r > 0 and (bestv is None or _cmp(v, bestv) < 0):
            best, bestv = k, v
    return best


def _dround(x: float, n: int, mode) -> float:
    try:
        return float(Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-n), rounding=mode))
    except InvalidOperation:
        return x


# ==========================================================================
# Functions: math & aggregates
# ==========================================================================
@_fn("SUM")
def _f_sum(E, a):
    return math.fsum(_nums(E, a))


@_fn("PRODUCT")
def _f_product(E, a):
    xs = _nums(E, a)
    if not xs:
        return 0.0
    p = 1.0
    for x in xs:
        p *= x
    return _fin(p)


@_fn("SUMSQ")
def _f_sumsq(E, a):
    return math.fsum(x * x for x in _nums(E, a))


@_fn("AVERAGE")
def _f_average(E, a):
    xs = _nums(E, a)
    if not xs:
        _err(ERR_DIV0)
    return math.fsum(xs) / len(xs)


@_fn("AVERAGEA")
def _f_averagea(E, a):
    xs = []
    for v in _all_values(E, a):
        if isinstance(v, XLError):
            raise FormulaError(v)
        if _isnum(v):
            xs.append(float(v))
        elif isinstance(v, bool):
            xs.append(1.0 if v else 0.0)
        elif isinstance(v, str):
            xs.append(0.0)
    if not xs:
        _err(ERR_DIV0)
    return math.fsum(xs) / len(xs)


@_fn("MIN")
def _f_min(E, a):
    xs = _nums(E, a)
    return min(xs) if xs else 0.0


@_fn("MAX")
def _f_max(E, a):
    xs = _nums(E, a)
    return max(xs) if xs else 0.0


@_fn("COUNT")
def _f_count(E, a):
    n = 0
    for x in a:
        if isinstance(x, (RangeRef, list)):
            n += sum(1 for v in _all_values(E, [x]) if _isnum(v))
        elif _isnum(x) or isinstance(x, bool):
            n += 1
        elif isinstance(x, str) and _str_to_num(x) is not None:
            n += 1
    return float(n)


@_fn("COUNTA")
def _f_counta(E, a):
    n = 0
    for x in a:
        if isinstance(x, (RangeRef, list)):
            n += sum(1 for v in _all_values(E, [x]) if v is not None)
        elif x is not None:
            n += 1
    return float(n)


@_fn("COUNTBLANK")
def _f_countblank(E, a):
    _need(a, 1, 1)
    x = a[0]
    if isinstance(x, RangeRef):
        filled = sum(1 for v in E.range_values(x) if v != "")
        return float(x.nrows * x.ncols - filled)
    g = E.grid(x)
    return float(sum(1 for row in g for v in row if v is None or v == ""))


@_fn("MEDIAN")
def _f_median(E, a):
    xs = sorted(_nums(E, a))
    if not xs:
        _err(ERR_NUM)
    n = len(xs)
    return xs[n // 2] if n % 2 else (xs[n // 2 - 1] + xs[n // 2]) / 2.0


@_fn("MODE", "MODE.SNGL")
def _f_mode(E, a):
    xs = _nums(E, a)
    counts = {}
    for x in xs:
        counts[x] = counts.get(x, 0) + 1
    best, bc = None, 1
    for x in xs:
        if counts[x] > bc:
            best, bc = x, counts[x]
    if best is None:
        _err(ERR_NA)
    return best


def _var(xs, sample):
    n = len(xs)
    if n < (2 if sample else 1):
        _err(ERR_DIV0)
    m = math.fsum(xs) / n
    return math.fsum((x - m) ** 2 for x in xs) / (n - 1 if sample else n)


@_fn("VAR", "VAR.S")
def _f_var(E, a):
    return _var(_nums(E, a), True)


@_fn("VARP", "VAR.P")
def _f_varp(E, a):
    return _var(_nums(E, a), False)


@_fn("STDEV", "STDEV.S")
def _f_stdev(E, a):
    return math.sqrt(_var(_nums(E, a), True))


@_fn("STDEVP", "STDEV.P")
def _f_stdevp(E, a):
    return math.sqrt(_var(_nums(E, a), False))


@_fn("LARGE")
def _f_large(E, a):
    _need(a, 2, 2)
    xs = sorted(_nums(E, [a[0]]), reverse=True)
    k = _int(E, a[1])
    if not 1 <= k <= len(xs):
        _err(ERR_NUM)
    return xs[k - 1]


@_fn("SMALL")
def _f_small(E, a):
    _need(a, 2, 2)
    xs = sorted(_nums(E, [a[0]]))
    k = _int(E, a[1])
    if not 1 <= k <= len(xs):
        _err(ERR_NUM)
    return xs[k - 1]


@_fn("RANK", "RANK.EQ")
def _f_rank(E, a):
    _need(a, 2, 3)
    x = _num(E, a[0])
    xs = _nums(E, [a[1]])
    order = _num(E, _opt(a, 2, 0.0))
    if not any(_cmp(x, v) == 0 for v in xs):
        _err(ERR_NA)
    if order == 0:
        return float(1 + sum(1 for v in xs if v > x and _cmp(v, x) != 0))
    return float(1 + sum(1 for v in xs if v < x and _cmp(v, x) != 0))


@_fn("SUMPRODUCT")
def _f_sumproduct(E, a):
    if not a:
        _err(ERR_VALUE)
    grids = [E.grid(x) if isinstance(x, (RangeRef, list)) else [[_sc(E, x)]] for x in a]
    R = len(grids[0])
    C = len(grids[0][0]) if R else 0
    for g in grids:
        if len(g) != R or (R and len(g[0]) != C):
            _err(ERR_VALUE)
    total = 0.0
    for i in range(R):
        for j in range(C):
            p = 1.0
            for g in grids:
                v = g[i][j]
                if isinstance(v, XLError):
                    raise FormulaError(v)
                p *= float(v) if _isnum(v) else 0.0
            total += p
    return total


@_fn("ABS")
def _f_abs(E, a):
    return abs(_num(E, a[0]))


@_fn("SQRT")
def _f_sqrt(E, a):
    x = _num(E, a[0])
    if x < 0:
        _err(ERR_NUM)
    return math.sqrt(x)


@_fn("POWER")
def _f_power(E, a):
    _need(a, 2, 2)
    return _bin_scalar("^", _num(E, a[0]), _num(E, a[1]))


@_fn("EXP")
def _f_exp(E, a):
    try:
        return _fin(math.exp(_num(E, a[0])))
    except OverflowError:
        _err(ERR_NUM)


@_fn("LN")
def _f_ln(E, a):
    x = _num(E, a[0])
    if x <= 0:
        _err(ERR_NUM)
    return math.log(x)


@_fn("LOG")
def _f_log(E, a):
    x = _num(E, a[0])
    b = _num(E, _opt(a, 1, 10.0))
    if x <= 0 or b <= 0:
        _err(ERR_NUM)
    if b == 1:
        _err(ERR_DIV0)
    return math.log(x) / math.log(b) if b != 10 else math.log10(x)


@_fn("LOG10")
def _f_log10(E, a):
    x = _num(E, a[0])
    if x <= 0:
        _err(ERR_NUM)
    return math.log10(x)


@_fn("PI")
def _f_pi(E, a):
    return math.pi


@_fn("RAND")
def _f_rand(E, a):
    return random.random()


@_fn("RANDBETWEEN")
def _f_randbetween(E, a):
    lo, hi = math.ceil(_num(E, a[0])), math.floor(_num(E, a[1]))
    if lo > hi:
        _err(ERR_NUM)
    return float(random.randint(lo, hi))


@_fn("SIGN")
def _f_sign(E, a):
    x = _num(E, a[0])
    return float((x > 0) - (x < 0))


@_fn("INT")
def _f_int(E, a):
    return float(math.floor(_num(E, a[0])))


@_fn("TRUNC")
def _f_trunc(E, a):
    return _dround(_num(E, a[0]), _int(E, _opt(a, 1, 0.0)), ROUND_DOWN)


@_fn("MOD")
def _f_mod(E, a):
    _need(a, 2, 2)
    x, y = _num(E, a[0]), _num(E, a[1])
    if y == 0:
        _err(ERR_DIV0)
    r = x - y * math.floor(x / y)
    return _dround(r, 12, ROUND_HALF_UP) if abs(r) < 1e15 else r


@_fn("QUOTIENT")
def _f_quotient(E, a):
    x, y = _num(E, a[0]), _num(E, a[1])
    if y == 0:
        _err(ERR_DIV0)
    return float(math.trunc(x / y))


@_fn("ROUND")
def _f_round(E, a):
    _need(a, 2, 2)
    return _dround(_num(E, a[0]), _int(E, a[1]), ROUND_HALF_UP)


@_fn("ROUNDUP")
def _f_roundup(E, a):
    _need(a, 2, 2)
    return _dround(_num(E, a[0]), _int(E, a[1]), ROUND_UP)


@_fn("ROUNDDOWN")
def _f_rounddown(E, a):
    _need(a, 2, 2)
    return _dround(_num(E, a[0]), _int(E, a[1]), ROUND_DOWN)


@_fn("MROUND")
def _f_mround(E, a):
    x, m = _num(E, a[0]), _num(E, a[1])
    if m == 0:
        return 0.0
    if x * m < 0:
        _err(ERR_NUM)
    q = _dround(x / m, 0, ROUND_HALF_UP)
    return _dround(q * m, 12, ROUND_HALF_UP)


def _ceil_floor(x, sig, up):
    if sig == 0:
        if up:
            return 0.0
        _err(ERR_DIV0)
    if x > 0 and sig < 0:
        _err(ERR_NUM)
    if x < 0 and sig < 0:
        q = round(abs(x) / abs(sig), 9)
        q = math.ceil(q) if up else math.floor(q)
        return -_dround(q * abs(sig), 12, ROUND_HALF_UP)
    q = round(x / sig, 9)
    q = math.ceil(q) if up else math.floor(q)
    return _dround(q * sig, 12, ROUND_HALF_UP)


@_fn("CEILING")
def _f_ceiling(E, a):
    x = _num(E, a[0])
    sig = _num(E, _opt(a, 1, 1.0 if x >= 0 else -1.0))
    return _ceil_floor(x, sig, True)


@_fn("FLOOR")
def _f_floor(E, a):
    x = _num(E, a[0])
    sig = _num(E, _opt(a, 1, 1.0 if x >= 0 else -1.0))
    return _ceil_floor(x, sig, False)


@_fn("CEILING.MATH")
def _f_ceiling_math(E, a):
    x = _num(E, a[0])
    sig = abs(_num(E, _opt(a, 1, 1.0))) or 1.0
    mode = _num(E, _opt(a, 2, 0.0))
    if x < 0 and mode:
        return -_ceil_floor(-x, sig, True)
    return _dround(math.ceil(round(x / sig, 9)) * sig, 12, ROUND_HALF_UP)


@_fn("FLOOR.MATH")
def _f_floor_math(E, a):
    x = _num(E, a[0])
    sig = abs(_num(E, _opt(a, 1, 1.0))) or 1.0
    mode = _num(E, _opt(a, 2, 0.0))
    if x < 0 and mode:
        return -_ceil_floor(-x, sig, False)
    return _dround(math.floor(round(x / sig, 9)) * sig, 12, ROUND_HALF_UP)


@_fn("EVEN")
def _f_even(E, a):
    x = _num(E, a[0])
    v = math.ceil(abs(x) / 2 - 1e-12) * 2
    return float(v if x >= 0 else -v)


@_fn("ODD")
def _f_odd(E, a):
    x = _num(E, a[0])
    v = math.ceil(abs(x))
    if v % 2 == 0:
        v += 1
    return float(v if x >= 0 else -v)


@_fn("FACT")
def _f_fact(E, a):
    x = _num(E, a[0])
    if x < 0:
        _err(ERR_NUM)
    try:
        return _fin(float(math.factorial(int(x))))
    except OverflowError:
        _err(ERR_NUM)


@_fn("GCD")
def _f_gcd(E, a):
    g = 0
    for x in _nums(E, a):
        if x < 0:
            _err(ERR_NUM)
        g = math.gcd(g, int(x))
    return float(g)


@_fn("LCM")
def _f_lcm(E, a):
    l = 1
    for x in _nums(E, a):
        if x < 0:
            _err(ERR_NUM)
        x = int(x)
        if x == 0:
            return 0.0
        l = l * x // math.gcd(l, x)
    return float(l)


def _unary_math(name, f, domain=None):
    def impl(E, a):
        x = _num(E, a[0])
        if domain and not domain(x):
            _err(ERR_NUM)
        try:
            return _fin(f(x))
        except (ValueError, OverflowError):
            _err(ERR_NUM)
    FUNCTIONS[name] = impl


_unary_math("SIN", math.sin)
_unary_math("COS", math.cos)
_unary_math("TAN", math.tan)
_unary_math("ASIN", math.asin, lambda x: -1 <= x <= 1)
_unary_math("ACOS", math.acos, lambda x: -1 <= x <= 1)
_unary_math("ATAN", math.atan)
_unary_math("SINH", math.sinh)
_unary_math("COSH", math.cosh)
_unary_math("TANH", math.tanh)
_unary_math("DEGREES", math.degrees)
_unary_math("RADIANS", math.radians)


@_fn("ATAN2")
def _f_atan2(E, a):
    x, y = _num(E, a[0]), _num(E, a[1])
    if x == 0 and y == 0:
        _err(ERR_DIV0)
    return math.atan2(y, x)


# ==========================================================================
# Conditional aggregates
# ==========================================================================
def _ifs_positions(E, pairs, target: Optional[_Area] = None):
    areas = [_Area(E, r) for r, _ in pairs]
    preds = [_criterion(_sc(E, c)) for _, c in pairs]
    a0 = areas[0]
    for ar in areas[1:]:
        if (ar.nrows, ar.ncols) != (a0.nrows, a0.ncols):
            _err(ERR_VALUE)
    extra = 0
    if not preds[0](None):
        cand = [(i, j) for i, j, v in a0.existing() if preds[0](v)]
        others = range(1, len(areas))
    else:
        dims = [ar.dense() for ar in areas]
        if target is not None:
            dims.append(target.dense())
        R = min(a0.nrows, max(d[0] for d in dims))
        C = min(a0.ncols, max(d[1] for d in dims))
        cand = [(i, j) for i in range(R) for j in range(C)]
        others = range(len(areas))
        if all(p(None) for p in preds):
            extra = a0.nrows * a0.ncols - R * C
    res = []
    for i, j in cand:
        ok = True
        for k in others:
            if not preds[k](areas[k].get(i, j)):
                ok = False
                break
        if ok:
            res.append((i, j))
    return res, extra


def _pairs(args, start):
    rest = args[start:]
    if not rest or len(rest) % 2:
        _err(ERR_VALUE)
    return [(rest[k], rest[k + 1]) for k in range(0, len(rest), 2)]


def _target_nums(target: _Area, pos):
    out = []
    for i, j in pos:
        v = target.get(i, j)
        if _isnum(v):
            out.append(float(v))
        elif isinstance(v, XLError):
            raise FormulaError(v)
    return out


@_fn("COUNTIF")
def _f_countif(E, a):
    _need(a, 2, 2)
    pos, extra = _ifs_positions(E, [(a[0], a[1])])
    return float(len(pos) + extra)


@_fn("COUNTIFS")
def _f_countifs(E, a):
    pos, extra = _ifs_positions(E, _pairs(a, 0))
    return float(len(pos) + extra)


@_fn("SUMIF")
def _f_sumif(E, a):
    _need(a, 2, 3)
    target = _Area(E, a[2] if len(a) > 2 and a[2] is not None else a[0])
    pos, _ = _ifs_positions(E, [(a[0], a[1])], target)
    return math.fsum(_target_nums(target, pos))


@_fn("SUMIFS")
def _f_sumifs(E, a):
    target = _Area(E, a[0])
    pos, _ = _ifs_positions(E, _pairs(a, 1), target)
    return math.fsum(_target_nums(target, pos))


@_fn("AVERAGEIF")
def _f_averageif(E, a):
    _need(a, 2, 3)
    target = _Area(E, a[2] if len(a) > 2 and a[2] is not None else a[0])
    pos, _ = _ifs_positions(E, [(a[0], a[1])], target)
    xs = _target_nums(target, pos)
    if not xs:
        _err(ERR_DIV0)
    return math.fsum(xs) / len(xs)


@_fn("AVERAGEIFS")
def _f_averageifs(E, a):
    target = _Area(E, a[0])
    pos, _ = _ifs_positions(E, _pairs(a, 1), target)
    xs = _target_nums(target, pos)
    if not xs:
        _err(ERR_DIV0)
    return math.fsum(xs) / len(xs)


@_fn("MAXIFS")
def _f_maxifs(E, a):
    target = _Area(E, a[0])
    pos, _ = _ifs_positions(E, _pairs(a, 1), target)
    xs = _target_nums(target, pos)
    return max(xs) if xs else 0.0


@_fn("MINIFS")
def _f_minifs(E, a):
    target = _Area(E, a[0])
    pos, _ = _ifs_positions(E, _pairs(a, 1), target)
    xs = _target_nums(target, pos)
    return min(xs) if xs else 0.0


# ==========================================================================
# Logical
# ==========================================================================
@_fn("IF", lazy=True)
def _f_if(E, n):
    _need(n, 1, 3)
    cond = _bool(E, E.ev(n[0]))
    if cond:
        return E.ev(n[1]) if len(n) > 1 else True
    return E.ev(n[2]) if len(n) > 2 else False


@_fn("IFS", lazy=True)
def _f_ifs(E, n):
    if not n or len(n) % 2:
        _err(ERR_VALUE)
    for k in range(0, len(n), 2):
        if _bool(E, E.ev(n[k])):
            return E.ev(n[k + 1])
    _err(ERR_NA)


def _safe_scalar(E, node):
    try:
        v = E.ev(node)
        if isinstance(v, RangeRef) and v.nrows == 1 and v.ncols == 1:
            return v, E.cell(v.sheet, v.r1, v.c1)
        if isinstance(v, (RangeRef, list)):
            return v, None
        return v, v
    except FormulaError as e:
        return e.error, e.error


@_fn("IFERROR", lazy=True)
def _f_iferror(E, n):
    _need(n, 2, 2)
    v, s = _safe_scalar(E, n[0])
    if isinstance(s, XLError):
        return E.ev(n[1])
    return v


@_fn("IFNA", lazy=True)
def _f_ifna(E, n):
    _need(n, 2, 2)
    v, s = _safe_scalar(E, n[0])
    if s == ERR_NA:
        return E.ev(n[1])
    return v


def _logicals(E, a):
    out = []
    for x in a:
        if isinstance(x, (RangeRef, list)):
            for v in _all_values(E, [x]):
                if isinstance(v, XLError):
                    raise FormulaError(v)
                if isinstance(v, bool):
                    out.append(v)
                elif _isnum(v):
                    out.append(v != 0)
        else:
            out.append(_to_bool(x))
    if not out:
        _err(ERR_VALUE)
    return out


@_fn("AND")
def _f_and(E, a):
    return all(_logicals(E, a))


@_fn("OR")
def _f_or(E, a):
    return any(_logicals(E, a))


@_fn("XOR")
def _f_xor(E, a):
    return sum(_logicals(E, a)) % 2 == 1


@_fn("NOT")
def _f_not(E, a):
    _need(a, 1, 1)
    return not _bool(E, a[0])


@_fn("TRUE")
def _f_true(E, a):
    return True


@_fn("FALSE")
def _f_false(E, a):
    return False


@_fn("SWITCH", lazy=True)
def _f_switch(E, n):
    if len(n) < 3:
        _err(ERR_VALUE)
    v = _sc(E, E.ev(n[0]))
    if isinstance(v, XLError):
        raise FormulaError(v)
    k = 1
    while k + 1 < len(n):
        cv = _sc(E, E.ev(n[k]))
        if _cmp(v, cv) == 0 and _rank(v) == _rank(cv):
            return E.ev(n[k + 1])
        k += 2
    if k < len(n):
        return E.ev(n[k])
    _err(ERR_NA)


@_fn("CHOOSE", lazy=True)
def _f_choose(E, n):
    if len(n) < 2:
        _err(ERR_VALUE)
    i = _int(E, E.ev(n[0]))
    if not 1 <= i < len(n):
        _err(ERR_VALUE)
    return E.ev(n[i])


# ==========================================================================
# Text
# ==========================================================================
@_fn("LEN")
def _f_len(E, a):
    return float(len(_txt(E, a[0])))


@_fn("LEFT")
def _f_left(E, a):
    s = _txt(E, a[0])
    n = _int(E, _opt(a, 1, 1.0))
    if n < 0:
        _err(ERR_VALUE)
    return s[:n]


@_fn("RIGHT")
def _f_right(E, a):
    s = _txt(E, a[0])
    n = _int(E, _opt(a, 1, 1.0))
    if n < 0:
        _err(ERR_VALUE)
    return s[len(s) - n:] if n else ""


@_fn("MID")
def _f_mid(E, a):
    _need(a, 3, 3)
    s = _txt(E, a[0])
    st, n = _int(E, a[1]), _int(E, a[2])
    if st < 1 or n < 0:
        _err(ERR_VALUE)
    return s[st - 1:st - 1 + n]


@_fn("UPPER")
def _f_upper(E, a):
    return _txt(E, a[0]).upper()


@_fn("LOWER")
def _f_lower(E, a):
    return _txt(E, a[0]).lower()


@_fn("PROPER")
def _f_proper(E, a):
    out, prev = [], False
    for ch in _txt(E, a[0]):
        out.append(ch.lower() if prev else ch.upper())
        prev = ch.isalpha()
    return "".join(out)


@_fn("TRIM")
def _f_trim(E, a):
    return re.sub(" +", " ", _txt(E, a[0]).strip(" "))


@_fn("CLEAN")
def _f_clean(E, a):
    return "".join(ch for ch in _txt(E, a[0]) if ord(ch) >= 32)


@_fn("CONCAT")
def _f_concat(E, a):
    return "".join(_to_str(v) for v in _all_values(E, a))


@_fn("CONCATENATE")
def _f_concatenate(E, a):
    return "".join(_txt(E, x) for x in a)


@_fn("TEXTJOIN")
def _f_textjoin(E, a):
    if len(a) < 3:
        _err(ERR_VALUE)
    delim = _txt(E, a[0])
    ignore = _bool(E, a[1])
    parts = []
    for v in _all_values(E, a[2:], include_empty=True):
        s = _to_str(v)
        if ignore and s == "":
            continue
        parts.append(s)
    return delim.join(parts)


@_fn("SUBSTITUTE")
def _f_substitute(E, a):
    _need(a, 3, 4)
    s, old, new = _txt(E, a[0]), _txt(E, a[1]), _txt(E, a[2])
    if old == "":
        return s
    if len(a) > 3:
        n = _int(E, a[3])
        if n < 1:
            _err(ERR_VALUE)
        idx = -1
        for _ in range(n):
            idx = s.find(old, idx + 1)
            if idx < 0:
                return s
        return s[:idx] + new + s[idx + len(old):]
    return s.replace(old, new)


@_fn("REPLACE")
def _f_replace(E, a):
    _need(a, 4, 4)
    s = _txt(E, a[0])
    st, n = _int(E, a[1]), _int(E, a[2])
    if st < 1 or n < 0:
        _err(ERR_VALUE)
    return s[:st - 1] + _txt(E, a[3]) + s[st - 1 + n:]


@_fn("FIND")
def _f_find(E, a):
    _need(a, 2, 3)
    f, s = _txt(E, a[0]), _txt(E, a[1])
    st = _int(E, _opt(a, 2, 1.0))
    if st < 1 or st > len(s) + 1:
        _err(ERR_VALUE)
    k = s.find(f, st - 1)
    if k < 0:
        _err(ERR_VALUE)
    return float(k + 1)


@_fn("SEARCH")
def _f_search(E, a):
    _need(a, 2, 3)
    f, s = _txt(E, a[0]), _txt(E, a[1])
    st = _int(E, _opt(a, 2, 1.0))
    if st < 1 or st > len(s) + 1:
        _err(ERR_VALUE)
    m = _wild_re(f, False).search(s, st - 1)
    if not m:
        _err(ERR_VALUE)
    return float(m.start() + 1)


@_fn("REPT")
def _f_rept(E, a):
    n = _int(E, a[1])
    if n < 0:
        _err(ERR_VALUE)
    return _txt(E, a[0]) * n


@_fn("EXACT")
def _f_exact(E, a):
    return _txt(E, a[0]) == _txt(E, a[1])


@_fn("VALUE", "NUMBERVALUE")
def _f_value(E, a):
    v = _sc(E, a[0])
    if isinstance(v, XLError):
        raise FormulaError(v)
    if _isnum(v):
        return float(v)
    if v is None:
        return 0.0
    if isinstance(v, bool):
        _err(ERR_VALUE)
    x = _str_to_num(str(v))
    if x is None:
        _err(ERR_VALUE)
    return x


@_fn("TEXT")
def _f_text(E, a):
    _need(a, 2, 2)
    v = _sc(E, a[0])
    if isinstance(v, XLError):
        raise FormulaError(v)
    fmt = _txt(E, a[1])
    if isinstance(v, str):
        x = _str_to_num(v)
        if x is not None:
            v = x
    if v is None:
        v = 0.0
    return format_value(v, fmt)


@_fn("FIXED")
def _f_fixed(E, a):
    x = _num(E, a[0])
    d = _int(E, _opt(a, 1, 2.0))
    no_commas = _bool(E, _opt(a, 2, False))
    x = _dround(x, d, ROUND_HALF_UP)
    fmt = ("0" if no_commas else "#,##0") + ("." + "0" * d if d > 0 else "")
    return format_value(x, fmt)


@_fn("CHAR")
def _f_char(E, a):
    n = _int(E, a[0])
    if not 1 <= n <= 255:
        _err(ERR_VALUE)
    try:
        return bytes([n]).decode("cp1252")
    except UnicodeDecodeError:
        return chr(n)


@_fn("UNICHAR")
def _f_unichar(E, a):
    n = _int(E, a[0])
    if n < 1:
        _err(ERR_VALUE)
    try:
        return chr(n)
    except (ValueError, OverflowError):
        _err(ERR_VALUE)


@_fn("CODE")
def _f_code(E, a):
    s = _txt(E, a[0])
    if not s:
        _err(ERR_VALUE)
    try:
        return float(s[0].encode("cp1252")[0])
    except UnicodeEncodeError:
        return float(ord(s[0]))


@_fn("UNICODE")
def _f_unicode(E, a):
    s = _txt(E, a[0])
    if not s:
        _err(ERR_VALUE)
    return float(ord(s[0]))


@_fn("T")
def _f_t(E, a):
    v = _sc(E, a[0])
    if isinstance(v, XLError):
        raise FormulaError(v)
    return v if isinstance(v, str) else ""


@_fn("N")
def _f_n(E, a):
    v = _sc(E, a[0])
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    return float(v) if _isnum(v) else 0.0


# ==========================================================================
# Lookup & reference
# ==========================================================================
@_fn("VLOOKUP")
def _f_vlookup(E, a):
    _need(a, 3, 4)
    val = _sc(E, a[0])
    table = _Area(E, a[1])
    col = _int(E, a[2])
    approx = _bool(E, _opt(a, 3, True))
    if col < 1:
        _err(ERR_VALUE)
    if col > table.ncols:
        _err(ERR_REF)
    k = _vsearch(table.sub_col(0), val, 1 if approx else 0)
    if k is None:
        _err(ERR_NA)
    return table.get(k, col - 1)


@_fn("HLOOKUP")
def _f_hlookup(E, a):
    _need(a, 3, 4)
    val = _sc(E, a[0])
    table = _Area(E, a[1])
    row = _int(E, a[2])
    approx = _bool(E, _opt(a, 3, True))
    if row < 1:
        _err(ERR_VALUE)
    if row > table.nrows:
        _err(ERR_REF)
    k = _vsearch(table.sub_row(0), val, 1 if approx else 0)
    if k is None:
        _err(ERR_NA)
    return table.get(row - 1, k)


@_fn("MATCH")
def _f_match(E, a):
    _need(a, 2, 3)
    val = _sc(E, a[0])
    arr = _Area(E, a[1])
    mt = _num(E, _opt(a, 2, 1.0))
    if not arr.is_vector:
        _err(ERR_NA)
    mode = 0 if mt == 0 else (1 if mt > 0 else -1)
    k = _vsearch(arr, val, mode)
    if k is None:
        _err(ERR_NA)
    return float(k + 1)


@_fn("XLOOKUP")
def _f_xlookup(E, a):
    _need(a, 3, 6)
    val = _sc(E, a[0])
    look = _Area(E, a[1])
    ret = _Area(E, a[2])
    not_found = _opt(a, 3, None)
    mm = _int(E, _opt(a, 4, 0.0)) if _opt(a, 4) is not None else 0
    sm = _int(E, _opt(a, 5, 1.0)) if _opt(a, 5) is not None else 1
    if not look.is_vector:
        _err(ERR_VALUE)
    mode = {0: 0, 2: 0, -1: "next_smaller", 1: "next_larger"}.get(mm)
    if mode is None:
        _err(ERR_VALUE)
    k = _vsearch(look, val, mode, reverse=sm < 0, wildcards=(mm == 2))
    if k is None:
        if len(a) > 3 and not_found is not None:
            return not_found
        _err(ERR_NA)
    along_rows = look.ncols == 1 and look.nrows > 1 or (look.nrows == look.ncols == 1)
    if along_rows:
        if ret.ncols == 1:
            return ret.get(k, 0)
        return ret.result_at(k, True)
    if ret.nrows == 1:
        return ret.get(0, k)
    return ret.result_at(k, False)


@_fn("INDEX")
def _f_index(E, a):
    _need(a, 2, 3)
    src = a[0]
    area = _Area(E, src)
    r = _int(E, a[1]) if a[1] is not None else 0
    c = _int(E, a[2]) if len(a) > 2 and a[2] is not None else None
    if c is None:
        if area.nrows == 1 and area.ncols > 1:
            r, c = 1, r
        else:
            c = 1 if area.ncols == 1 else 0
    if r < 0 or c < 0 or r > area.nrows or c > area.ncols:
        _err(ERR_REF)
    if isinstance(src, RangeRef):
        rr = src
        r1, r2 = (rr.r1, rr.r2) if r == 0 else (rr.r1 + r - 1,) * 2
        c1, c2 = (rr.c1, rr.c2) if c == 0 else (rr.c1 + c - 1,) * 2
        return RangeRef(rr.sheet, r1, c1, r2, c2)
    g = area.grid
    if r == 0 and c == 0:
        return g
    if r == 0:
        return [[row[c - 1]] for row in g]
    if c == 0:
        return [list(g[r - 1])]
    return g[r - 1][c - 1]


@_fn("ROW")
def _f_row(E, a):
    if not a or a[0] is None:
        if E.r is None:
            _err(ERR_VALUE)
        return float(E.r + 1)
    if not isinstance(a[0], RangeRef):
        _err(ERR_VALUE)
    return float(a[0].r1 + 1)


@_fn("COLUMN")
def _f_column(E, a):
    if not a or a[0] is None:
        if E.c is None:
            _err(ERR_VALUE)
        return float(E.c + 1)
    if not isinstance(a[0], RangeRef):
        _err(ERR_VALUE)
    return float(a[0].c1 + 1)


@_fn("ROWS")
def _f_rows(E, a):
    x = a[0]
    if isinstance(x, XLError):
        raise FormulaError(x)
    if isinstance(x, RangeRef):
        return float(x.nrows)
    if isinstance(x, list):
        return float(len(x))
    return 1.0


@_fn("COLUMNS")
def _f_columns(E, a):
    x = a[0]
    if isinstance(x, XLError):
        raise FormulaError(x)
    if isinstance(x, RangeRef):
        return float(x.ncols)
    if isinstance(x, list):
        return float(len(x[0]) if x else 0)
    return 1.0


@_fn("OFFSET")
def _f_offset(E, a):
    _need(a, 3, 5)
    base = a[0]
    if isinstance(base, XLError):
        raise FormulaError(base)
    if not isinstance(base, RangeRef):
        _err(ERR_VALUE)
    dr, dc = _int(E, a[1]), _int(E, a[2])
    h = _int(E, a[3]) if len(a) > 3 and a[3] is not None else base.nrows
    w = _int(E, a[4]) if len(a) > 4 and a[4] is not None else base.ncols
    if h < 1 or w < 1:
        _err(ERR_REF)
    r1, c1 = base.r1 + dr, base.c1 + dc
    r2, c2 = r1 + h - 1, c1 + w - 1
    if not (_okrow(r1) and _okrow(r2) and _okcol(c1) and _okcol(c2)):
        _err(ERR_REF)
    return RangeRef(base.sheet, r1, c1, r2, c2)


@_fn("INDIRECT")
def _f_indirect(E, a):
    text = _txt(E, a[0]).strip()
    try:
        toks = [t for t in _tokenize(text) if t.type != "ws"]
    except _SyntaxErr:
        toks = []
    if len(toks) == 1 and toks[0].type == "ref":
        return E.ev(_ref_node(toks[0].val))
    if len(toks) == 1 and toks[0].type == "name":
        try:
            v = E.name(toks[0].text)
            if isinstance(v, RangeRef):
                return v
        except FormulaError:
            pass
    _err(ERR_REF)


@_fn("CHOOSECOLS")
def _f_choosecols(E, a):
    g = E.grid(a[0])
    out = []
    for row in g:
        out.append([row[_int(E, k) - 1] for k in a[1:]])
    return out


@_fn("TRANSPOSE")
def _f_transpose(E, a):
    g = E.grid(a[0])
    return [list(r) for r in zip(*g)] if g else [[None]]


# ==========================================================================
# Information
# ==========================================================================
def _is_scalar(E, x):
    try:
        return _sc(E, x)
    except FormulaError as e:
        return e.error


@_fn("ISBLANK")
def _f_isblank(E, a):
    return _is_scalar(E, a[0]) is None


@_fn("ISNUMBER")
def _f_isnumber(E, a):
    return _isnum(_is_scalar(E, a[0]))


@_fn("ISTEXT")
def _f_istext(E, a):
    return isinstance(_is_scalar(E, a[0]), str)


@_fn("ISNONTEXT")
def _f_isnontext(E, a):
    return not isinstance(_is_scalar(E, a[0]), str)


@_fn("ISLOGICAL")
def _f_islogical(E, a):
    return isinstance(_is_scalar(E, a[0]), bool)


@_fn("ISERROR")
def _f_iserror(E, a):
    return isinstance(_is_scalar(E, a[0]), XLError)


@_fn("ISERR")
def _f_iserr(E, a):
    v = _is_scalar(E, a[0])
    return isinstance(v, XLError) and v != ERR_NA


@_fn("ISNA")
def _f_isna(E, a):
    return _is_scalar(E, a[0]) == ERR_NA


@_fn("ISEVEN")
def _f_iseven(E, a):
    return int(math.trunc(_num(E, a[0]))) % 2 == 0


@_fn("ISODD")
def _f_isodd(E, a):
    return int(math.trunc(_num(E, a[0]))) % 2 == 1


@_fn("ISREF")
def _f_isref(E, a):
    return isinstance(a[0], RangeRef)


@_fn("NA")
def _f_na(E, a):
    return ERR_NA


@_fn("ERROR.TYPE")
def _f_error_type(E, a):
    v = _is_scalar(E, a[0])
    codes = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"]
    if isinstance(v, XLError) and v.code in codes:
        return float(codes.index(v.code) + 1)
    _err(ERR_NA)


# ==========================================================================
# Date & time
# ==========================================================================
@_fn("TODAY")
def _f_today(E, a):
    return float(date_to_serial(_dt.date.today()))


@_fn("NOW")
def _f_now(E, a):
    now = _dt.datetime.now()
    return date_to_serial(now.date()) + (now.hour * 3600 + now.minute * 60 + now.second) / 86400.0


@_fn("DATE")
def _f_date(E, a):
    _need(a, 3, 3)
    return _make_date(_int(E, a[0]), _int(E, a[1]), _int(E, a[2]))


@_fn("TIME")
def _f_time(E, a):
    _need(a, 3, 3)
    secs = _int(E, a[0]) * 3600 + _int(E, a[1]) * 60 + _int(E, a[2])
    if secs < 0:
        _err(ERR_NUM)
    return (secs % 86400) / 86400.0


@_fn("DATEVALUE")
def _f_datevalue(E, a):
    v = _sc(E, a[0])
    if not isinstance(v, str):
        _err(ERR_VALUE)
    x = _str_to_num(v)
    if x is None:
        _err(ERR_VALUE)
    return float(math.floor(x))


@_fn("TIMEVALUE")
def _f_timevalue(E, a):
    v = _sc(E, a[0])
    if not isinstance(v, str):
        _err(ERR_VALUE)
    x = _str_to_num(v)
    if x is None:
        _err(ERR_VALUE)
    return x - math.floor(x)


def _serial_arg(E, v) -> float:
    x = _num(E, v)
    if x < 0:
        _err(ERR_NUM)
    return x


@_fn("YEAR")
def _f_year(E, a):
    return float(serial_to_date(_serial_arg(E, a[0]))[0])


@_fn("MONTH")
def _f_month(E, a):
    return float(serial_to_date(_serial_arg(E, a[0]))[1])


@_fn("DAY")
def _f_day(E, a):
    return float(serial_to_date(_serial_arg(E, a[0]))[2])


@_fn("HOUR")
def _f_hour(E, a):
    return float(_split_time(_serial_arg(E, a[0]))[1])


@_fn("MINUTE")
def _f_minute(E, a):
    return float(_split_time(_serial_arg(E, a[0]))[2])


@_fn("SECOND")
def _f_second(E, a):
    return float(_split_time(_serial_arg(E, a[0]))[3])


@_fn("WEEKDAY")
def _f_weekday(E, a):
    n = int(math.floor(_serial_arg(E, a[0])))
    t = _int(E, _opt(a, 1, 1.0))
    if t == 1 or t == 17:
        return float((n - 1) % 7 + 1)
    if t == 2 or t == 11:
        return float((n - 2) % 7 + 1)
    if t == 3:
        return float((n - 2) % 7)
    if 12 <= t <= 16:  # week starts Tue..Sat
        return float((n - (t - 9)) % 7 + 1)
    _err(ERR_NUM)


def _add_months(serial, months, eom=False):
    y, m, d = serial_to_date(serial)
    total = y * 12 + (m - 1) + months
    ny, nm = divmod(total, 12)
    nm += 1
    if not 1900 <= ny <= 9999:
        _err(ERR_NUM)
    last = calendar.monthrange(ny, nm)[1]
    return float(date_to_serial(_dt.date(ny, nm, last if eom else min(d, last))))


@_fn("EDATE")
def _f_edate(E, a):
    _need(a, 2, 2)
    return _add_months(_serial_arg(E, a[0]), _int(E, a[1]))


@_fn("EOMONTH")
def _f_eomonth(E, a):
    _need(a, 2, 2)
    return _add_months(_serial_arg(E, a[0]), _int(E, a[1]), eom=True)


@_fn("DATEDIF")
def _f_datedif(E, a):
    _need(a, 3, 3)
    s, e = _serial_arg(E, a[0]), _serial_arg(E, a[1])
    unit = _txt(E, a[2]).upper()
    if s > e:
        _err(ERR_NUM)
    ds, de = _pydate(s), _pydate(e)
    if unit == "D":
        return float(math.floor(e) - math.floor(s))
    months = (de.year - ds.year) * 12 + de.month - ds.month - (1 if de.day < ds.day else 0)
    if unit == "M":
        return float(months)
    if unit == "Y":
        return float(months // 12)
    if unit == "YM":
        return float(months % 12)
    if unit == "MD":
        if de.day >= ds.day:
            return float(de.day - ds.day)
        py, pm = (de.year, de.month - 1) if de.month > 1 else (de.year - 1, 12)
        return float(de.day - ds.day + calendar.monthrange(py, pm)[1])
    if unit == "YD":
        try:
            shifted = ds.replace(year=de.year)
        except ValueError:
            shifted = _dt.date(de.year, 3, 1)
        if shifted > de:
            try:
                shifted = ds.replace(year=de.year - 1)
            except ValueError:
                shifted = _dt.date(de.year - 1, 3, 1)
        return float((de - shifted).days)
    _err(ERR_NUM)


@_fn("DAYS")
def _f_days(E, a):
    return float(math.floor(_num(E, a[0])) - math.floor(_num(E, a[1])))


@_fn("NETWORKDAYS")
def _f_networkdays(E, a):
    s, e = int(math.floor(_num(E, a[0]))), int(math.floor(_num(E, a[1])))
    hol = set(int(x) for x in _nums(E, [a[2]])) if len(a) > 2 else set()
    sign = 1
    if s > e:
        s, e, sign = e, s, -1
    n = 0
    for d in range(s, e + 1):
        if (d - 1) % 7 not in (0, 6) and d not in hol:
            n += 1
    return float(n * sign)


# ==========================================================================
# Formula rewriting (copy/fill, insert/delete, sheet rename/delete)
# ==========================================================================
def _rewrite(raw, fn):
    """Apply fn(_Ref) -> new text | None to every reference in a formula."""
    if not isinstance(raw, str) or not raw.startswith("=") or len(raw) < 2:
        return raw
    try:
        toks = _tokenize(raw[1:])
    except _SyntaxErr:
        return raw
    changed = False
    out = []
    for t in toks:
        if t.type == "ref":
            new = fn(t.val)
            if new is not None and new != t.text:
                out.append(new)
                changed = True
                continue
        out.append(t.text)
    return "=" + "".join(out) if changed else raw


def _rewrite_name_value(text: str, fn) -> str:
    s = str(text)
    if s.startswith("="):
        return _rewrite(s, fn)
    new = _rewrite("=" + s, fn)
    return new[1:]


def shift_formula(raw: str, dr: int, dc: int) -> str:
    """Shift relative references (not $-absolute) by dr rows / dc cols."""
    def f(ref: _Ref):
        if ref.kind == "err":
            return None
        n = ref.copy()
        if n.kind in ("cell", "area", "rows"):
            if not n.ar1:
                n.r1 += dr
            if n.r2 is not None and not n.ar2:
                n.r2 += dr
            if n.kind == "cell":
                n.r2 = n.r1
            if not (_okrow(n.r1) and _okrow(n.r2)):
                return "#REF!"
        if n.kind in ("cell", "area", "cols"):
            if not n.ac1:
                n.c1 += dc
            if n.c2 is not None and not n.ac2:
                n.c2 += dc
            if n.kind == "cell":
                n.c2 = n.c1
            if not (_okcol(n.c1) and _okcol(n.c2)):
                return "#REF!"
        return n.text()
    if dr == 0 and dc == 0:
        return raw
    return _rewrite(raw, f)


def _adj_single(p, index, count, limit):
    if count > 0:
        np_ = p + count if p >= index else p
        return np_ if np_ < limit else None
    n = -count
    if p < index:
        return p
    if p <= index + n - 1:
        return None
    return p - n


def _adj_span(a, b, index, count, limit):
    if count > 0:
        na = a + count if a >= index else a
        nb = b + count if b >= index else b
        if na >= limit:
            return None
        return na, min(nb, limit - 1)
    n = -count
    d1, d2 = index, index + n - 1
    if a >= d1 and b <= d2:
        return None
    na = a if a < d1 else (d1 if a <= d2 else a - n)
    nb = b if b < d1 else (d1 - 1 if b <= d2 else b - n)
    return na, nb


def _move_keys(d: dict, pos_of, make, index, count, limit):
    out = {}
    n = -count
    for k, v in d.items():
        p = pos_of(k)
        if count > 0:
            if p >= index:
                p += count
                if p >= limit:
                    continue
        else:
            if index <= p < index + n:
                continue
            if p >= index + n:
                p -= n
        out[make(k, p)] = v
    return out


def insert_delete(book, sheet, axis: str, index: int, count: int) -> None:
    """Insert (count>0) or delete (count<0) rows/cols and fix all formulas."""
    if count == 0:
        return
    rows = axis.lower().startswith("row")
    limit = MAX_ROWS if rows else MAX_COLS
    # 1. move cell contents & formatting
    if rows:
        pos_of = lambda k: k[0]
        make = lambda k, p: (p, k[1])
    else:
        pos_of = lambda k: k[1]
        make = lambda k, p: (k[0], p)
    sheet.cells = _move_keys(sheet.cells, pos_of, make, index, count, limit)
    sheet.styles = _move_keys(sheet.styles, pos_of, make, index, count, limit)
    sizes_name = "row_heights" if rows else "col_widths"
    sizes = getattr(sheet, sizes_name)
    setattr(sheet, sizes_name, _move_keys(sizes, lambda k: k, lambda k, p: p, index, count, limit))

    # 2. rewrite references
    def make_fn(home):
        def f(ref: _Ref):
            if ref.kind == "err":
                return None
            target = home if ref.sheet is None else book.sheet_by_name(ref.sheet)
            if target is not sheet:
                return None
            n = ref.copy()
            n.normalize()
            if rows:
                if n.kind == "cols":
                    return None
                if n.kind == "cell":
                    p = _adj_single(n.r1, index, count, limit)
                    if p is None:
                        return _err_text(n)
                    n.r1 = n.r2 = p
                else:
                    sp = _adj_span(n.r1, n.r2, index, count, limit)
                    if sp is None:
                        return _err_text(n)
                    n.r1, n.r2 = sp
            else:
                if n.kind == "rows":
                    return None
                if n.kind == "cell":
                    p = _adj_single(n.c1, index, count, limit)
                    if p is None:
                        return _err_text(n)
                    n.c1 = n.c2 = p
                else:
                    sp = _adj_span(n.c1, n.c2, index, count, limit)
                    if sp is None:
                        return _err_text(n)
                    n.c1, n.c2 = sp
            if n.kind == "cell":
                n.r2, n.c2 = n.r1, n.c1
            return n.text()
        return f

    for sh in book.sheets:
        fn = make_fn(sh)
        for key, raw in list(sh.cells.items()):
            if raw.startswith("="):
                new = _rewrite(raw, fn)
                if new != raw:
                    sh.cells[key] = new
    fn = make_fn(None)
    for k, v in list(book.defined_names.items()):
        book.defined_names[k] = _rewrite_name_value(v, fn)
    _shift_images(sheet, rows, index, count, limit)


def _shift_images(sheet, rows, index, count, limit):
    """Move floating images (Sheet.images) anchored after the insert/delete
    point; images on deleted rows/cols move to the first surviving one.
    Barcode source references shift like formula references."""
    import re as _re
    key = "row" if rows else "col"
    for img in getattr(sheet, "images", None) or []:
        p = int(img.get(key, 0))
        np_ = _adj_single(p, index, count, limit)
        if np_ is None:
            np_ = min(index, limit - 1) if count < 0 else limit - 1
        img[key] = np_
        bc = img.get("barcode")
        src = bc.get("source") if isinstance(bc, dict) else None
        if not src or "#REF" in src.upper():
            continue
        m = _re.match(r"^((?:'[^']+'|[^!]+)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$", src.strip())
        if not m:
            continue
        if m.group(1):
            ref_sheet = m.group(1)[:-1].strip("'")
            if ref_sheet.lower() != sheet.name.lower():
                continue
        r = int(m.group(5)) - 1
        c = name_to_col(m.group(3))
        q = _adj_single(r if rows else c, index, count, limit)
        if q is None:
            bc["source"] = "#REF!"
            continue
        if rows:
            r = q
        else:
            c = q
        bc["source"] = f"{m.group(1) or ''}{m.group(2)}{col_to_name(c)}{m.group(4)}{r + 1}"


def _err_text(ref: _Ref) -> str:
    return "#REF!"


def rename_sheet_refs(book, old_name: str, new_name: str) -> None:
    low = old_name.lower()

    def f(ref: _Ref):
        if ref.sheet is not None and ref.sheet.lower() == low:
            n = ref.copy()
            n.sheet = new_name
            return n.text()
        return None

    for sh in book.sheets:
        for key, raw in list(sh.cells.items()):
            if raw.startswith("="):
                new = _rewrite(raw, f)
                if new != raw:
                    sh.cells[key] = new
    for k, v in list(book.defined_names.items()):
        book.defined_names[k] = _rewrite_name_value(v, f)


def delete_sheet_refs(book, sheet_name: str) -> None:
    low = sheet_name.lower()

    def f(ref: _Ref):
        if ref.sheet is not None and ref.sheet.lower() == low:
            return "#REF!"
        return None

    for sh in book.sheets:
        if sh.name.lower() == low:
            continue
        for key, raw in list(sh.cells.items()):
            if raw.startswith("="):
                new = _rewrite(raw, f)
                if new != raw:
                    sh.cells[key] = new
    for k, v in list(book.defined_names.items()):
        book.defined_names[k] = _rewrite_name_value(v, f)

"""VBA value model: Variant special values, arrays, UDTs, coercion and operators."""
from __future__ import annotations

import math
import re
from decimal import Decimal, ROUND_HALF_EVEN, ROUND_HALF_UP, InvalidOperation

from .errors import rt_error, VBARuntimeErr
from . import dates as _dates


# --------------------------------------------------------------------------
# Special values
# --------------------------------------------------------------------------
class _Special:
    __slots__ = ("name",)

    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return self.name

    def __bool__(self):
        return False

    def __copy__(self):
        return self

    def __deepcopy__(self, memo):
        return self


EMPTY = _Special("Empty")
NULL = _Special("Null")
NOTHING = _Special("Nothing")
MISSING = _Special("Missing")   # omitted optional argument


class VDate(float):
    """A VBA Date (serial number, days since 1899-12-30)."""
    __slots__ = ()

    def __repr__(self):
        return f"VDate({float(self)!r})"


XL_ERR_CODES = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023,
                "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042, "#GETTING_DATA": 2043,
                "#CYCLE!": 2036, "#ERROR!": 2015}
XL_CODE_TO_TEXT = {v: k for k, v in XL_ERR_CODES.items() if k not in ("#CYCLE!", "#ERROR!")}


class VErr:
    """A Variant of subtype Error (CVErr / error cell value)."""
    __slots__ = ("number",)

    def __init__(self, number: int):
        self.number = int(number)

    def __eq__(self, other):
        return isinstance(other, VErr) and other.number == self.number

    def __hash__(self):
        return hash(("verr", self.number))

    def __repr__(self):
        return f"Error {self.number}"

    @property
    def xl_text(self):
        return XL_CODE_TO_TEXT.get(self.number, "#VALUE!")


# --------------------------------------------------------------------------
# Objects
# --------------------------------------------------------------------------
class VBAObject:
    """Base class for every object visible to VBA code.

    Subclasses implement members as methods ``g_<name>(self, args, named)``
    (property get / method call) and ``s_<name>(self, args, named, value)``
    (property let/set).  The default member uses the name ``_default``.
    """
    vba_typename = "Object"

    def vba_get(self, name, args, named):
        m = getattr(self, "g_" + (name or "_default"), None)
        if m is None:
            raise rt_error(438)
        return m(args, named)

    def vba_let(self, name, args, named, value, is_set=False):
        m = getattr(self, "s_" + (name or "_default"), None)
        if m is not None:
            return m(args, named, value)
        if name:
            target = self.vba_get(name, args, named)
            if isinstance(target, VBAObject) and not is_set:
                return target.vba_let("", [], {}, value)
        raise rt_error(438)

    def vba_iter(self):
        raise rt_error(438)

    def vba_same(self, other):
        return self is other

    def default_value(self):
        return self.vba_get("", [], {})


def is_object(v):
    return isinstance(v, VBAObject) or v is NOTHING


def to_value(v):
    """Let-coercion: replace an object by its default member value."""
    n = 0
    while isinstance(v, VBAObject):
        v = v.default_value()
        n += 1
        if n > 20:
            raise rt_error(13)
    if v is NOTHING:
        raise rt_error(91)
    return v


# --------------------------------------------------------------------------
# Type specifications
# --------------------------------------------------------------------------
class TypeSpec:
    """Declared type of a variable, array element, field or parameter."""
    __slots__ = ("name", "kind", "size", "udt", "cls", "new")

    def __init__(self, name="variant", kind="variant", size=0, udt=None, cls=None, new=False):
        self.name = name      # lower-case type name
        self.kind = kind      # variant | scalar | object | udt | fixedstr
        self.size = size      # String * n
        self.udt = udt        # UDTDef for kind == udt
        self.cls = cls        # class name for objects (lower) / factory
        self.new = new        # As New

    def __repr__(self):
        return f"TypeSpec({self.name})"

    def default(self):
        k = self.kind
        if k == "variant":
            return EMPTY
        if k == "scalar":
            return SCALAR_DEFAULTS[self.name]
        if k == "fixedstr":
            return " " * self.size
        if k == "udt":
            return self.udt.instantiate()
        return NOTHING

    def coerce(self, v):
        k = self.kind
        if k == "variant":
            if isinstance(v, VBAObject) or v is NOTHING:
                return v
            if type(v) is VBAArray or type(v) is UDTValue:
                return v.copy()
            return v
        if k == "scalar":
            return SCALAR_COERCE[self.name](v)
        if k == "fixedstr":
            s = to_str(v)
            return s[:self.size].ljust(self.size)
        if k == "udt":
            if type(v) is UDTValue:
                return v.copy()
            raise rt_error(13)
        # object
        if isinstance(v, VBAObject) or v is NOTHING:
            return v
        raise rt_error(91) if v is EMPTY else rt_error(13)

    @property
    def display(self):
        if self.kind == "udt":
            return self.udt.name
        if self.kind == "fixedstr":
            return "String"
        return TYPE_DISPLAY.get(self.name, self.cls or self.name.title())


VARIANT = TypeSpec()

TYPE_DISPLAY = {"variant": "Variant", "integer": "Integer", "long": "Long", "longlong": "LongLong",
                "longptr": "Long", "byte": "Byte", "boolean": "Boolean", "single": "Single",
                "double": "Double", "currency": "Currency", "decimal": "Decimal", "date": "Date",
                "string": "String", "object": "Object"}


# --------------------------------------------------------------------------
# Arrays
# --------------------------------------------------------------------------
class VBAArray:
    __slots__ = ("dims", "data", "etype", "fixed")

    def __init__(self, dims, etype=None, data=None, fixed=False):
        self.dims = dims           # list of (lb, ub) or None (unallocated dynamic array)
        self.etype = etype or VARIANT
        self.fixed = fixed
        if data is not None:
            self.data = data
        elif dims is None:
            self.data = []
        else:
            n = 1
            for lb, ub in dims:
                n *= max(0, ub - lb + 1)
            if self.etype.kind in ("udt",):
                self.data = [self.etype.default() for _ in range(n)]
            else:
                d = self.etype.default()
                self.data = [d] * n

    @classmethod
    def from_list(cls, items, lb=0, etype=None):
        items = list(items)
        return cls([(lb, lb + len(items) - 1)], etype, items)

    @classmethod
    def from_2d(cls, rows, lb=1):
        nr = len(rows)
        nc = max((len(r) for r in rows), default=0)
        data = []
        for r in rows:
            data.extend(r)
            data.extend([EMPTY] * (nc - len(r)))
        return cls([(lb, lb + nr - 1), (lb, lb + nc - 1)], None, data)

    def copy(self):
        d = self.data
        if any(type(x) is VBAArray or type(x) is UDTValue for x in d):
            d = [x.copy() if (type(x) is VBAArray or type(x) is UDTValue) else x for x in d]
        else:
            d = list(d)
        return VBAArray(None if self.dims is None else list(self.dims), self.etype, d, False)

    @property
    def ndims(self):
        return 0 if self.dims is None else len(self.dims)

    def _offset(self, idx):
        dims = self.dims
        if dims is None or len(idx) != len(dims):
            raise rt_error(9)
        off = 0
        for (lb, ub), i in zip(dims, idx):
            if type(i) is not int:
                i = index_int(i)
            if i < lb or i > ub:
                raise rt_error(9)
            off = off * (ub - lb + 1) + (i - lb)
        return off

    def get(self, idx):
        if len(idx) == 1:
            dims = self.dims
            if dims is not None and len(dims) == 1:
                i = idx[0]
                if type(i) is not int:
                    i = index_int(i)
                lb, ub = dims[0]
                if lb <= i <= ub:
                    return self.data[i - lb]
                raise rt_error(9)
        return self.data[self._offset(idx)]

    def set(self, idx, value):
        et = self.etype
        if et.kind != "variant" or type(value) is VBAArray or type(value) is UDTValue:
            value = et.coerce(value)
        if len(idx) == 1:
            dims = self.dims
            if dims is not None and len(dims) == 1:
                i = idx[0]
                if type(i) is not int:
                    i = index_int(i)
                lb, ub = dims[0]
                if lb <= i <= ub:
                    self.data[i - lb] = value
                    return
                raise rt_error(9)
        self.data[self._offset(idx)] = value

    def lbound(self, dim=1):
        if self.dims is None or not (1 <= dim <= len(self.dims)):
            raise rt_error(9)
        return self.dims[dim - 1][0]

    def ubound(self, dim=1):
        if self.dims is None or not (1 <= dim <= len(self.dims)):
            raise rt_error(9)
        return self.dims[dim - 1][1]

    def redim(self, dims, preserve=False):
        if self.fixed:
            raise rt_error(10)
        if not preserve or self.dims is None:
            new = VBAArray(dims, self.etype)
            self.dims, self.data = new.dims, new.data
            return
        old = self.dims
        if len(old) != len(dims):
            raise rt_error(9)
        new = VBAArray(dims, self.etype)
        if len(dims) == 1:
            (olb, oub), (nlb, nub) = old[0], dims[0]
            for i in range(max(olb, nlb), min(oub, nub) + 1):
                new.data[i - nlb] = self.data[i - olb]
        else:
            for idx in self.iter_indices():
                if all(dims[k][0] <= idx[k] <= dims[k][1] for k in range(len(dims))):
                    new.data[new._offset(idx)] = self.data[self._offset(idx)]
        self.dims, self.data = new.dims, new.data

    def iter_indices(self):
        """Row-major list of index tuples."""
        if self.dims is None:
            return []
        out = [()]
        for lb, ub in self.dims:
            out = [t + (i,) for t in out for i in range(lb, ub + 1)]
        return out

    def iter_values_colmajor(self):
        """For Each order (first dimension varies fastest)."""
        if self.dims is None:
            return []
        if len(self.dims) == 1:
            return list(self.data)
        idxs = [()]
        for lb, ub in reversed(self.dims):
            idxs = [(i,) + t for i in range(lb, ub + 1) for t in idxs]
        # idxs now has last dim varying slowest? build explicitly instead
        res = []
        dims = self.dims

        def rec(k, cur):
            if k < 0:
                res.append(self.data[self._offset(tuple(cur))])
                return
            lb, ub = dims[k]
            for i in range(lb, ub + 1):
                cur[k] = i
                rec(k - 1, cur)
        rec(len(dims) - 1, [0] * len(dims))
        return res

    def to_rows(self):
        """2-D -> list of rows; 1-D -> one row; for worksheet transfer."""
        if self.dims is None:
            return []
        if len(self.dims) == 1:
            return [list(self.data)]
        if len(self.dims) == 2:
            (l1, u1), (l2, u2) = self.dims
            w = u2 - l2 + 1
            return [self.data[i * w:(i + 1) * w] for i in range(u1 - l1 + 1)]
        raise rt_error(13)

    def __repr__(self):
        return f"VBAArray({self.dims}, {self.data[:10]})"


# --------------------------------------------------------------------------
# User defined types
# --------------------------------------------------------------------------
class UDTDef:
    def __init__(self, name, fields):
        self.name = name
        self.fields = fields   # list of (lower name, display name, TypeSpec, dims or None)

    def instantiate(self):
        vals = {}
        for low, disp, ts, dims in self.fields:
            if dims is not None:
                vals[low] = VBAArray(dims, ts, fixed=bool(dims)) if dims else VBAArray(None, ts)
            else:
                vals[low] = ts.default()
        return UDTValue(self, vals)


class UDTValue:
    __slots__ = ("tdef", "fields")

    def __init__(self, tdef, fields):
        self.tdef = tdef
        self.fields = fields

    def copy(self):
        f = {}
        for k, v in self.fields.items():
            f[k] = v.copy() if (type(v) is VBAArray or type(v) is UDTValue) else v
        return UDTValue(self.tdef, f)

    def get_field(self, name):
        try:
            return self.fields[name]
        except KeyError:
            raise rt_error(438, f"Method or data member not found: {name}")

    def set_field(self, name, value):
        for low, disp, ts, dims in self.tdef.fields:
            if low == name:
                if dims is not None:
                    if type(value) is not VBAArray:
                        raise rt_error(13)
                    self.fields[name] = value.copy()
                else:
                    self.fields[name] = ts.coerce(value)
                return
        raise rt_error(438, f"Method or data member not found: {name}")


# --------------------------------------------------------------------------
# Number parsing / formatting
# --------------------------------------------------------------------------
_NUM_RE = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?$")


def str_to_num_or_none(s: str):
    t = s.strip()
    if not t:
        return None
    if t[0] == "&" and len(t) > 2:
        k = t[1].upper()
        try:
            if k == "H":
                return int(t[2:].rstrip("&"), 16)
            if k == "O":
                return int(t[2:].rstrip("&"), 8)
        except ValueError:
            return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        t = t[1:-1].strip()
        neg = True
    if t.startswith("$"):
        t = t[1:].strip()
    elif t.startswith("-$"):
        t = "-" + t[2:].strip()
    if "," in t:
        if re.fullmatch(r"[+-]?\d{1,3}(,\d{3})+(\.\d*)?", t):
            t = t.replace(",", "")
        else:
            return None
    if _NUM_RE.match(t):
        t = t.replace("d", "e").replace("D", "e")
        try:
            v = float(t)
        except ValueError:
            return None
        return -v if neg else v
    return None


def str_to_num(s: str):
    v = str_to_num_or_none(s)
    if v is None:
        raise rt_error(13)
    return v


def fmt_double(x: float) -> str:
    if x != x:
        return "NaN"
    if math.isinf(x):
        return "1.#INF" if x > 0 else "-1.#INF"
    if x == 0:
        return "0"
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    s = "%.15g" % x
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


def fmt_single(x: float) -> str:
    if x == 0:
        return "0"
    s = "%.7g" % x
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


def fmt_currency(d: Decimal) -> str:
    s = format(d.quantize(Decimal("0.0001")), "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    if s == "-0":
        s = "0"
    return s


def fmt_date_default(v: float) -> str:
    dt = _dates.serial_to_dt(v)
    ip = math.floor(v)
    frac = v - ip
    datepart = f"{dt.month}/{dt.day}/{dt.year}"
    h = dt.hour % 12 or 12
    timepart = f"{h}:{dt.minute:02d}:{dt.second:02d} {'AM' if dt.hour < 12 else 'PM'}"
    if abs(frac) < 1e-9:
        return datepart
    if ip == 0:
        return timepart
    return datepart + " " + timepart


# --------------------------------------------------------------------------
# Coercions
# --------------------------------------------------------------------------
def index_int(v) -> int:
    t = type(v)
    if t is int:
        return v
    if t is float or t is VDate:
        return int(round(v))
    if t is bool:
        return -1 if v else 0
    if t is Decimal:
        return int(v.quantize(Decimal(1), ROUND_HALF_EVEN))
    if v is EMPTY:
        return 0
    if t is str:
        return int(round(str_to_num(v)))
    if isinstance(v, VBAObject):
        return index_int(to_value(v))
    if v is NULL:
        raise rt_error(94)
    raise rt_error(13)


def to_num(v):
    t = type(v)
    if t is int or t is float:
        return v
    if t is bool:
        return -1 if v else 0
    if v is EMPTY:
        return 0
    if t is str:
        return str_to_num(v)
    if t is VDate:
        return float(v)
    if t is Decimal:
        return v
    if v is NULL:
        raise rt_error(94)
    if isinstance(v, VBAObject):
        return to_num(to_value(v))
    if v is NOTHING:
        raise rt_error(91)
    raise rt_error(13)


def to_float(v) -> float:
    t = type(v)
    if t is float:
        return v
    if t is int:
        return float(v)
    return float(to_num(v))


def _round_int(v) -> int:
    t = type(v)
    if t is int:
        return v
    if t is bool:
        return -1 if v else 0
    if t is Decimal:
        return int(v.quantize(Decimal(1), ROUND_HALF_EVEN))
    if v != v or math.isinf(v):
        raise rt_error(6)
    return int(round(v))


def to_int_range(v, lo, hi) -> int:
    if type(v) is int and lo <= v <= hi:
        return v
    if type(v) is str:
        s = v.strip().lower()
        if s == "true":
            n = -1
        elif s == "false":
            n = 0
        else:
            n = _round_int(to_num(v))
    else:
        n = _round_int(to_num(v))
    if n < lo or n > hi:
        raise rt_error(6)
    return n


def to_integer(v):
    return to_int_range(v, -32768, 32767)


def to_long(v):
    return to_int_range(v, -2147483648, 2147483647)


def to_longlong(v):
    return to_int_range(v, -9223372036854775808, 9223372036854775807)


def to_byte(v):
    return to_int_range(v, 0, 255)


def to_double(v) -> float:
    t = type(v)
    if t is float:
        return v
    if t is int:
        return float(v)
    if t is str:
        s = v.strip().lower()
        if s in ("true", "false"):
            return -1.0 if s == "true" else 0.0
        n = str_to_num_or_none(v)
        if n is None:
            d = _dates.parse_date_string(v)
            if d is not None:
                return float(d)
            raise rt_error(13)
        return float(n)
    return float(to_num(v))


def to_single(v) -> float:
    x = to_double(v)
    if abs(x) > 3.402823e38:
        raise rt_error(6)
    return x


def to_currency(v) -> Decimal:
    if type(v) is Decimal:
        d = v
    else:
        n = to_num(v) if type(v) is not str else to_double(v)
        if type(n) is Decimal:
            d = n
        else:
            d = Decimal(repr(n)) if type(n) is float else Decimal(n)
    d = d.quantize(Decimal("0.0001"), ROUND_HALF_EVEN)
    if abs(d) > Decimal("922337203685477.5807"):
        raise rt_error(6)
    return d


def to_bool(v) -> bool:
    t = type(v)
    if t is bool:
        return v
    if t is int or t is float or t is VDate or t is Decimal:
        return v != 0
    if v is EMPTY:
        return False
    if t is str:
        s = v.strip().lower()
        if s == "true":
            return True
        if s == "false":
            return False
        n = str_to_num_or_none(v)
        if n is None:
            raise rt_error(13)
        return n != 0
    if v is NULL:
        raise rt_error(94)
    if isinstance(v, VBAObject):
        return to_bool(to_value(v))
    if v is NOTHING:
        raise rt_error(91)
    raise rt_error(13)


def truthy(v) -> bool:
    if v is True:
        return True
    if v is False:
        return False
    t = type(v)
    if t is int or t is float:
        return v != 0
    if v is NULL or v is EMPTY:
        return False
    return to_bool(v)


def to_date(v) -> VDate:
    t = type(v)
    if t is VDate:
        return v
    if t is int or t is float or t is Decimal:
        return VDate(float(v))
    if t is str:
        d = _dates.parse_date_string(v)
        if d is None:
            n = str_to_num_or_none(v)
            if n is None:
                raise rt_error(13)
            return VDate(float(n))
        return VDate(d)
    if t is bool:
        return VDate(-1.0 if v else 0.0)
    if v is EMPTY:
        return VDate(0.0)
    if v is NULL:
        raise rt_error(94)
    if isinstance(v, VBAObject):
        return to_date(to_value(v))
    raise rt_error(13)


def to_str(v) -> str:
    t = type(v)
    if t is str:
        return v
    if t is int:
        return str(v)
    if t is float:
        return fmt_double(v)
    if t is bool:
        return "True" if v else "False"
    if v is EMPTY:
        return ""
    if t is VDate:
        return fmt_date_default(v)
    if t is Decimal:
        return fmt_currency(v)
    if v is NULL:
        raise rt_error(94)
    if t is VErr:
        return f"Error {v.number}"
    if isinstance(v, VBAObject):
        return to_str(to_value(v))
    if v is NOTHING:
        raise rt_error(91)
    raise rt_error(13)


def concat_str(v) -> str:
    t = type(v)
    if t is str:
        return v
    if v is NULL:
        return ""
    return to_str(v)


def to_variant(v):
    return v


SCALAR_DEFAULTS = {"integer": 0, "long": 0, "longlong": 0, "longptr": 0, "byte": 0,
                   "boolean": False, "single": 0.0, "double": 0.0,
                   "currency": Decimal("0"), "decimal": Decimal("0"),
                   "date": VDate(0.0), "string": ""}


def _to_decimal(v):
    if type(v) is Decimal:
        return v
    n = to_num(v) if type(v) is not str else to_double(v)
    return n if type(n) is Decimal else Decimal(repr(n)) if type(n) is float else Decimal(n)


def _to_string_assign(v):
    t = type(v)
    if t is str:
        return v
    if type(v) is VBAArray:
        raise rt_error(13)
    return to_str(v)


SCALAR_COERCE = {"integer": to_integer, "long": to_long, "longlong": to_longlong,
                 "longptr": to_longlong, "byte": to_byte, "boolean": to_bool,
                 "single": to_single, "double": to_double, "currency": to_currency,
                 "decimal": _to_decimal, "date": to_date, "string": _to_string_assign}


# --------------------------------------------------------------------------
# Type names
# --------------------------------------------------------------------------
def type_name(v) -> str:
    t = type(v)
    if t is bool:
        return "Boolean"
    if t is int:
        return "Integer" if -32768 <= v <= 32767 else "Long"
    if t is float:
        return "Double"
    if t is str:
        return "String"
    if t is VDate:
        return "Date"
    if t is Decimal:
        return "Currency"
    if v is EMPTY:
        return "Empty"
    if v is NULL:
        return "Null"
    if v is NOTHING:
        return "Nothing"
    if t is VErr:
        return "Error"
    if t is VBAArray:
        return v.etype.display + "()"
    if t is UDTValue:
        return v.tdef.name
    if isinstance(v, VBAObject):
        return v.vba_typename
    if v is MISSING:
        return "Error"
    return "Variant"


VARTYPE_BY_NAME = {"Empty": 0, "Null": 1, "Integer": 2, "Long": 3, "Single": 4, "Double": 5,
                   "Currency": 6, "Date": 7, "String": 8, "Object": 9, "Error": 10,
                   "Boolean": 11, "Variant": 12, "Decimal": 14, "Byte": 17, "LongLong": 20}


def var_type(v) -> int:
    t = type(v)
    if t is VBAArray:
        base = VARTYPE_BY_NAME.get(v.etype.display, 12)
        return 8192 + base
    if t is UDTValue:
        return 36
    if isinstance(v, VBAObject):
        try:
            dv = v.default_value()
            if not isinstance(dv, VBAObject):
                return var_type(dv)
        except VBARuntimeErr:
            pass
        return 9
    if v is NOTHING:
        return 9
    return VARTYPE_BY_NAME.get(type_name(v), 12)


# --------------------------------------------------------------------------
# Arithmetic operators
# --------------------------------------------------------------------------
def _num_pair(a, b):
    """Convert both operands to numbers; returns (a, b, kind) kind: date/cur/plain."""
    a = to_num(a) if type(a) is not VDate else a
    b = to_num(b) if type(b) is not VDate else b
    if type(a) is Decimal or type(b) is Decimal:
        a = _to_decimal(a)
        b = _to_decimal(b)
    return a, b


def _wrap(res, a, b):
    return res


def _unwrap_obj(v):
    if isinstance(v, VBAObject) or v is NOTHING:
        return to_value(v)
    return v


def op_add(a, b):
    ta = type(a)
    tb = type(b)
    if (ta is int or ta is float) and (tb is int or tb is float):
        return a + b
    return _add_slow(a, b)


def _add_slow(a, b):
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    ta, tb = type(a), type(b)
    if ta is str and tb is str:
        return a + b
    if ta is str and b is EMPTY:
        return a
    if tb is str and a is EMPTY:
        return b
    if type(a) is VBAArray or type(b) is VBAArray:
        raise rt_error(13)
    isdate = ta is VDate or tb is VDate
    x, y = _num_pair(a, b)
    if type(x) is Decimal:
        r = x + y
        return VDate(float(r)) if isdate else r
    r = x + y
    if isdate:
        return VDate(float(r))
    return r


def op_sub(a, b):
    ta = type(a)
    tb = type(b)
    if (ta is int or ta is float) and (tb is int or tb is float):
        return a - b
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    da = type(a) is VDate
    db = type(b) is VDate
    x, y = _num_pair(a, b)
    r = x - y
    if da and not db:
        return VDate(float(r))
    if da and db:
        return float(r)
    return r


def op_mul(a, b):
    ta = type(a)
    tb = type(b)
    if (ta is int or ta is float) and (tb is int or tb is float):
        return a * b
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = _num_pair(a, b)
    r = x * y
    if type(r) is VDate:
        return float(r)
    return r


def op_div(a, b):
    ta = type(a)
    tb = type(b)
    if (ta is int or ta is float) and (tb is int or tb is float):
        if b == 0:
            raise rt_error(6 if a == 0 else 11)
        return a / b
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = _num_pair(a, b)
    if y == 0:
        raise rt_error(6 if x == 0 else 11)
    if type(x) is Decimal:
        return float(x) / float(y)
    return float(x) / y


def op_intdiv(a, b):
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    x = _round_int(to_num(a))
    y = _round_int(to_num(b))
    if y == 0:
        raise rt_error(11)
    q = abs(x) // abs(y)
    return q if (x >= 0) == (y >= 0) else -q


def op_mod(a, b):
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    x = _round_int(to_num(a))
    y = _round_int(to_num(b))
    if y == 0:
        raise rt_error(11)
    r = abs(x) % abs(y)
    return -r if x < 0 else r


def op_pow(a, b):
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    x = to_float(a)
    y = to_float(b)
    if x == 0 and y < 0:
        raise rt_error(5)
    if x < 0 and not y.is_integer():
        raise rt_error(5)
    try:
        return x ** y
    except OverflowError:
        raise rt_error(6)


def op_neg(a):
    t = type(a)
    if t is int or t is float:
        return -a
    a = _unwrap_obj(a)
    if a is NULL:
        return NULL
    if type(a) is VDate:
        return VDate(-float(a))
    n = to_num(a)
    return -n


def op_concat(a, b):
    if type(a) is str and type(b) is str:
        return a + b
    if a is NULL and b is NULL:
        return NULL
    return concat_str(_unwrap_obj(a)) + concat_str(_unwrap_obj(b))


# --------------------------------------------------------------------------
# Comparison
# --------------------------------------------------------------------------
def _cmp_prepare(v):
    t = type(v)
    if t is bool:
        return -1 if v else 0
    if t is VDate:
        return float(v)
    return v


def vcompare(a, b, text=False):
    """-1/0/1, or None when either side is Null."""
    ta = type(a)
    tb = type(b)
    if (ta is int or ta is float) and (tb is int or tb is float):
        return (a > b) - (a < b)
    if ta is str and tb is str:
        if text:
            a = a.lower()
            b = b.lower()
        return (a > b) - (a < b)
    if isinstance(a, VBAObject) or a is NOTHING:
        a = to_value(a)
    if isinstance(b, VBAObject) or b is NOTHING:
        b = to_value(b)
    if a is NULL or b is NULL:
        return None
    if a is EMPTY and b is EMPTY:
        return 0
    if a is EMPTY:
        a = "" if type(b) is str else 0
    if b is EMPTY:
        b = "" if type(a) is str else 0
    ta, tb = type(a), type(b)
    if ta is VErr or tb is VErr:
        if ta is VErr and tb is VErr:
            return (a.number > b.number) - (a.number < b.number)
        raise rt_error(13)
    if ta is VBAArray or tb is VBAArray or ta is UDTValue or tb is UDTValue:
        raise rt_error(13)
    if ta is str and tb is str:
        if text:
            a, b = a.lower(), b.lower()
        return (a > b) - (a < b)
    if ta is str or tb is str:
        if ta is str:
            s, n, sign = a, b, 1
        else:
            s, n, sign = b, a, -1
        if type(n) is VDate:
            d = _dates.parse_date_string(s)
            if d is not None:
                x, y = (d, float(n)) if sign == 1 else (float(n), d)
                return (x > y) - (x < y)
        if type(n) is bool:
            ls = s.strip().lower()
            if ls in ("true", "false"):
                x = -1 if ls == "true" else 0
                y = -1 if n else 0
                x, y = (x, y) if sign == 1 else (y, x)
                return (x > y) - (x < y)
        num = str_to_num_or_none(s)
        if num is None:
            # numbers sort before strings
            return sign
        n = _cmp_prepare(n)
        if type(n) is Decimal:
            num = Decimal(repr(num))
        x, y = (num, n) if sign == 1 else (n, num)
        return (x > y) - (x < y)
    a = _cmp_prepare(a)
    b = _cmp_prepare(b)
    if type(a) is Decimal or type(b) is Decimal:
        a = _to_decimal(a)
        b = _to_decimal(b)
    return (a > b) - (a < b)


def make_cmp_ops(text):
    def eq(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float or (ta is str and not text)):
            return a == b
        r = vcompare(a, b, text)
        return NULL if r is None else r == 0

    def ne(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float or (ta is str and not text)):
            return a != b
        r = vcompare(a, b, text)
        return NULL if r is None else r != 0

    def lt(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float):
            return a < b
        r = vcompare(a, b, text)
        return NULL if r is None else r < 0

    def le(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float):
            return a <= b
        r = vcompare(a, b, text)
        return NULL if r is None else r <= 0

    def gt(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float):
            return a > b
        r = vcompare(a, b, text)
        return NULL if r is None else r > 0

    def ge(a, b):
        ta = type(a)
        if ta is type(b) and (ta is int or ta is float):
            return a >= b
        r = vcompare(a, b, text)
        return NULL if r is None else r >= 0

    return {"=": eq, "<>": ne, "<": lt, "<=": le, ">": gt, ">=": ge}


CMP_BINARY = make_cmp_ops(False)
CMP_TEXT = make_cmp_ops(True)

_like_cache = {}


def like_to_regex(pattern: str, text: bool):
    key = (pattern, text)
    rx = _like_cache.get(key)
    if rx is not None:
        return rx
    out = []
    i = 0
    n = len(pattern)
    while i < n:
        ch = pattern[i]
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        elif ch == "#":
            out.append("[0-9]")
        elif ch == "[":
            j = pattern.find("]", i + 1)
            if j < 0:
                raise rt_error(93)
            body = pattern[i + 1:j]
            if body.startswith("!"):
                neg = "^"
                body = body[1:]
            else:
                neg = ""
            if body == "":
                out.append("\\[\\]" if not neg else ".")
            else:
                esc = "".join("\\" + c if c in "\\^]" else c for c in body)
                out.append(f"[{neg}{esc}]")
            i = j
        else:
            out.append(re.escape(ch))
        i += 1
    rx = re.compile("".join(out) + r"\Z", re.S | (re.I if text else 0))
    _like_cache[key] = rx
    return rx


def op_like(a, b, text=False):
    a = _unwrap_obj(a)
    b = _unwrap_obj(b)
    if a is NULL or b is NULL:
        return NULL
    return like_to_regex(to_str(b), text).match(to_str(a)) is not None


def op_is(a, b):
    if a is NOTHING or b is NOTHING:
        return a is b
    if isinstance(a, VBAObject) and isinstance(b, VBAObject):
        return a.vba_same(b)
    if not (isinstance(a, VBAObject) or a is NOTHING) or not (isinstance(b, VBAObject) or b is NOTHING):
        raise rt_error(424)
    return a is b


# --------------------------------------------------------------------------
# Logical / bitwise
# --------------------------------------------------------------------------
def _logic_operand(v):
    t = type(v)
    if t is bool:
        return v, True
    if t is int:
        return v, False
    v = _unwrap_obj(v)
    t = type(v)
    if t is bool:
        return v, True
    if v is NULL:
        return NULL, False
    if v is EMPTY:
        return 0, False
    if t is str:
        s = v.strip().lower()
        if s == "true":
            return True, True
        if s == "false":
            return False, True
    return _round_int(to_num(v)), False


def _logic(a, b, fn, nullfn):
    x, xb = _logic_operand(a)
    y, yb = _logic_operand(b)
    if x is NULL or y is NULL:
        return nullfn(x, y)
    if xb and yb:
        return bool(fn(-1 if x else 0, -1 if y else 0) & 1)
    xi = -1 if x is True else 0 if x is False else x
    yi = -1 if y is True else 0 if y is False else y
    return fn(xi, yi)


def _null_and(x, y):
    other = y if x is NULL else x
    if other is NULL:
        return NULL
    if other is False or other == 0:
        return False if type(other) is bool else 0
    return NULL


def _null_or(x, y):
    other = y if x is NULL else x
    if other is NULL:
        return NULL
    if other is True or (type(other) is int and other == -1):
        return other
    return NULL


def op_and(a, b):
    if type(a) is bool and type(b) is bool:
        return a and b
    return _logic(a, b, lambda x, y: x & y, _null_and)


def op_or(a, b):
    if type(a) is bool and type(b) is bool:
        return a or b
    return _logic(a, b, lambda x, y: x | y, _null_or)


def op_xor(a, b):
    if type(a) is bool and type(b) is bool:
        return a != b
    return _logic(a, b, lambda x, y: x ^ y, lambda x, y: NULL)


def op_eqv(a, b):
    return _logic(a, b, lambda x, y: ~(x ^ y), lambda x, y: NULL)


def op_imp(a, b):
    return _logic(a, b, lambda x, y: (~x) | y, lambda x, y: NULL)


def op_not(a):
    if type(a) is bool:
        return not a
    x, xb = _logic_operand(a)
    if x is NULL:
        return NULL
    if xb:
        return not x
    return ~x


# --------------------------------------------------------------------------
# Rounding helpers
# --------------------------------------------------------------------------
def vba_round(x, digits=0):
    if type(x) is Decimal:
        return x.quantize(Decimal(1).scaleb(-digits), ROUND_HALF_EVEN)
    if digits == 0:
        return float(round(x))
    # use decimal on the shortest repr for VBA-like results
    try:
        d = Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-digits), ROUND_HALF_EVEN)
        return float(d)
    except InvalidOperation:
        return round(x, digits)


def round_half_up(x: float, digits: int) -> Decimal:
    d = Decimal(repr(float(x)))
    return d.quantize(Decimal(1).scaleb(-digits), ROUND_HALF_UP)

"""Tokenizer and preprocessor for VBA source code."""
from __future__ import annotations

import re

from .errors import VBACompileErr
from . import dates as _dates
from .values import VDate


class Tok:
    __slots__ = ("kind", "val", "text", "line", "space", "suffix")

    def __init__(self, kind, val, text, line, space=False, suffix=""):
        self.kind = kind      # id num str date op nl brk eof
        self.val = val        # lower-case name for ids; python value for literals; op text
        self.text = text      # original text
        self.line = line      # 1-based line
        self.space = space    # whitespace before the token
        self.suffix = suffix  # type suffix on identifiers / numbers

    def __repr__(self):
        return f"Tok({self.kind},{self.val!r},L{self.line})"


TYPE_SUFFIX = {"%": "integer", "&": "long", "!": "single", "#": "double", "@": "currency",
               "$": "string", "^": "longlong"}

_DATE_CHARS = re.compile(r"^[0-9A-Za-z/\-:., ]+$")


# --------------------------------------------------------------------------
# Header stripping & conditional compilation
# --------------------------------------------------------------------------
def strip_header(lines):
    """Blank out VERSION / BEGIN...END blocks of exported module text."""
    i = 0
    n = len(lines)
    while i < n and not lines[i].strip():
        i += 1
    if i < n and lines[i].strip().upper().startswith("VERSION "):
        lines[i] = ""
        i += 1
        while i < n and not lines[i].strip():
            i += 1
        if i < n and lines[i].strip().upper().startswith("BEGIN"):
            depth = 0
            while i < n:
                s = lines[i].strip().upper()
                if s.startswith("BEGIN"):
                    depth += 1
                elif s == "END" or s.startswith("END "):
                    depth -= 1
                lines[i] = ""
                i += 1
                if depth <= 0:
                    break
            # Object = ... lines may follow in form modules
    return lines


PP_DEFAULTS = {"vba7": True, "vba6": True, "win64": True, "win32": True, "win16": False,
               "mac": False}


class _PPEval:
    def __init__(self, toks, consts):
        self.t = [t for t in toks if t.kind != "nl"]
        self.i = 0
        self.c = consts

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else None

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def expr(self):
        return self.imp()

    def imp(self):
        v = self.or_()
        while self.peek() is not None and self.peek().kind == "id" and self.peek().val in ("xor", "eqv", "imp"):
            op = self.take().val
            w = self.or_()
            v = (bool(v) != bool(w)) if op == "xor" else (bool(v) == bool(w)) if op == "eqv" else ((not v) or bool(w))
        return v

    def or_(self):
        v = self.and_()
        while self.peek() is not None and self.peek().kind == "id" and self.peek().val == "or":
            self.take()
            w = self.and_()
            v = bool(v) or bool(w)
        return v

    def and_(self):
        v = self.not_()
        while self.peek() is not None and self.peek().kind == "id" and self.peek().val == "and":
            self.take()
            w = self.not_()
            v = bool(v) and bool(w)
        return v

    def not_(self):
        p = self.peek()
        if p is not None and p.kind == "id" and p.val == "not":
            self.take()
            return not self.not_()
        return self.cmp()

    def cmp(self):
        v = self.atom()
        p = self.peek()
        if p is not None and p.kind == "op" and p.val in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take().val
            w = self.atom()
            try:
                return {"=": v == w, "<>": v != w, "<": v < w, ">": v > w, "<=": v <= w, ">=": v >= w}[op]
            except TypeError:
                return False
        return v

    def atom(self):
        t = self.take()
        if t is None:
            return False
        if t.kind == "op" and t.val == "(":
            v = self.expr()
            self.take()
            return v
        if t.kind == "op" and t.val == "-":
            return -self.atom()
        if t.kind in ("num", "str"):
            return t.val
        if t.kind == "id":
            if t.val == "true":
                return True
            if t.val == "false":
                return False
            v = self.c.get(t.val, 0)
            return v
        return False


def preprocess(code: str):
    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = strip_header(lines)
    consts = dict(PP_DEFAULTS)
    stack = []   # (parent_active, this_branch_active, any_taken)
    active = True
    for i, raw in enumerate(lines):
        s = raw.strip()
        if s.startswith("#"):
            low = s.lower()
            m = re.match(r"#\s*(if|elseif|else|end\s*if|endif|const)\b(.*)$", low)
            if m:
                kw = re.sub(r"\s+", " ", m.group(1))
                rest = s[len(s) - len(m.group(2)):] if m.group(2) else ""
                rest = re.sub(r"(?i)\bthen\s*$", "", rest.split("'")[0]).strip()
                if kw == "if":
                    cond = active and bool(_pp_eval(rest, consts, i + 1))
                    stack.append((active, cond, cond))
                    active = cond
                elif kw == "elseif":
                    if stack:
                        parent, _, taken = stack[-1]
                        cond = parent and not taken and bool(_pp_eval(rest, consts, i + 1))
                        stack[-1] = (parent, cond, taken or cond)
                        active = cond
                elif kw == "else":
                    if stack:
                        parent, _, taken = stack[-1]
                        cond = parent and not taken
                        stack[-1] = (parent, cond, True)
                        active = cond
                elif kw in ("end if", "endif"):
                    if stack:
                        active = stack.pop()[0]
                elif kw == "const":
                    if active and "=" in rest:
                        name, val = rest.split("=", 1)
                        consts[name.strip().lower()] = _pp_eval(val, consts, i + 1)
                lines[i] = ""
                continue
        if not active:
            lines[i] = ""
    return lines


def _pp_eval(text, consts, line):
    try:
        toks = tokenize_lines([text], first_line=line)
        return _PPEval(toks, consts).expr()
    except VBACompileErr:
        return False


# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------
def _hex_value(digits: str, long_suffix: bool, base: int):
    v = int(digits, base)
    if not long_suffix and v <= 0xFFFF and len(digits) <= (4 if base == 16 else 6):
        if v >= 0x8000:
            v -= 0x10000
    elif v <= 0xFFFFFFFF:
        if v >= 0x80000000:
            v -= 0x100000000
    return v


def tokenize_lines(lines, first_line=1):
    toks = []
    append = toks.append
    cont = False
    for li, line in enumerate(lines):
        lineno = li + first_line
        i = 0
        n = len(line)
        stmt_start = not cont
        space = not cont
        cont = False
        while i < n:
            ch = line[i]
            if ch in " \t\x0c\xa0":
                i += 1
                space = True
                continue
            if ch == "'":
                break
            if ch == "_" and space and line[i + 1:].strip() == "":
                cont = True
                break
            if ch == '"':
                j = i + 1
                buf = []
                while True:
                    k = line.find('"', j)
                    if k < 0:
                        raise VBACompileErr("Syntax error: unterminated string", lineno)
                    buf.append(line[j:k])
                    if k + 1 < n and line[k + 1] == '"':
                        buf.append('"')
                        j = k + 2
                        continue
                    j = k + 1
                    break
                s = "".join(buf)
                append(Tok("str", s, line[i:j], lineno, space))
                i = j
                space = False
                stmt_start = False
                continue
            if ch.isdigit() or (ch == "." and i + 1 < n and line[i + 1].isdigit()
                                and not (toks and not space and toks[-1].kind in ("id", "op") and toks[-1].val in (")",) )
                                and not (toks and not space and toks[-1].kind == "id")):
                m = re.match(r"(\d*\.?\d*)([eEdD][+-]?\d+)?", line[i:])
                txt = m.group(0)
                j = i + len(txt)
                suffix = ""
                if j < n and line[j] in "%&!#@^" and not (j + 1 < n and (line[j + 1].isalnum() or line[j + 1] == '"')):
                    suffix = line[j]
                    j += 1
                body = txt.replace("d", "e").replace("D", "e")
                if "." in body or "e" in body.lower() or suffix in ("!", "#", "@"):
                    try:
                        val = float(body)
                    except ValueError:
                        raise VBACompileErr("Syntax error: bad number", lineno)
                    if suffix == "@":
                        from decimal import Decimal
                        val = Decimal(body)
                else:
                    val = int(body)
                append(Tok("num", val, line[i:j], lineno, space, suffix))
                i = j
                space = False
                stmt_start = False
                continue
            if ch == "&" and i + 1 < n and line[i + 1] in "hHoO0123456789":
                k = line[i + 1]
                if k in "hH":
                    m = re.match(r"[0-9A-Fa-f]+", line[i + 2:])
                    base, start = 16, i + 2
                elif k in "oO":
                    m = re.match(r"[0-7]+", line[i + 2:])
                    base, start = 8, i + 2
                else:
                    m = re.match(r"[0-7]+", line[i + 1:])
                    base, start = 8, i + 1
                if m:
                    j = start + len(m.group(0))
                    lsuf = False
                    if j < n and line[j] in "&%^":
                        lsuf = line[j] == "&"
                        j += 1
                    append(Tok("num", _hex_value(m.group(0), lsuf, base), line[i:j], lineno, space))
                    i = j
                    space = False
                    stmt_start = False
                    continue
            if ch.isalpha():
                j = i + 1
                while j < n and (line[j].isalnum() or line[j] == "_"):
                    j += 1
                name = line[i:j]
                suffix = ""
                if j < n and line[j] in "$%&!#@^":
                    c = line[j]
                    nxt = line[j + 1] if j + 1 < n else ""
                    if c == "!" and (nxt.isalpha() or nxt == "["):
                        pass  # bang operator
                    elif c == "&" and (nxt.isalnum() or nxt in '"_'):
                        pass
                    elif c == "#" and nxt.isdigit():
                        pass
                    else:
                        suffix = c
                        j += 1
                low = name.lower()
                if stmt_start and low == "rem" and suffix == "":
                    break
                append(Tok("id", low, name, lineno, space, suffix))
                i = j
                space = False
                stmt_start = False
                continue
            if ch == "#":
                k = line.find("#", i + 1)
                if k > i + 1:
                    inner = line[i + 1:k]
                    if _DATE_CHARS.match(inner) and re.search(r"[/:\-]|[A-Za-z]{3}", inner) \
                            and not inner.strip().isdigit():
                        d = _dates.parse_date_string(inner)
                        if d is not None:
                            append(Tok("date", VDate(d), line[i:k + 1], lineno, space))
                            i = k + 1
                            space = False
                            stmt_start = False
                            continue
                append(Tok("op", "#", "#", lineno, space))
                i += 1
                space = False
                stmt_start = False
                continue
            if ch == "[":
                k = line.find("]", i + 1)
                if k < 0:
                    raise VBACompileErr("Syntax error: missing ]", lineno)
                append(Tok("brk", line[i + 1:k], line[i:k + 1], lineno, space))
                i = k + 1
                space = False
                stmt_start = False
                continue
            two = line[i:i + 2]
            if two in ("<>", "<=", ">=", ":=", "=<", "=>", "><"):
                v = {"=<": "<=", "=>": ">=", "><": "<>"}.get(two, two)
                append(Tok("op", v, two, lineno, space))
                i += 2
                space = False
                stmt_start = False
                continue
            if ch in "+-*/\\^&=<>(),.!;:":
                append(Tok("op", ch, ch, lineno, space))
                i += 1
                stmt_start = ch == ":"
                space = False
                continue
            if ch == "?":
                append(Tok("id", "print", "?", lineno, space))
                i += 1
                space = False
                stmt_start = False
                continue
            raise VBACompileErr(f"Syntax error: unexpected character {ch!r}", lineno)
        if not cont:
            append(Tok("nl", "\n", "\n", lineno, True))
    last = toks[-1].line if toks else first_line
    toks.append(Tok("eof", None, "", last, True))
    return toks


def tokenize(code: str):
    return tokenize_lines(preprocess(code))

"""VBA classic file I/O (Open/Print #/Write #/Input #/Line Input #/Close)."""
from __future__ import annotations

import os

from .errors import rt_error, VBARuntimeErr
from .values import (EMPTY, NULL, VDate, VErr, VBAObject, to_str, to_value, index_int,
                     str_to_num_or_none, fmt_double)


class FileHandle:
    def __init__(self, path, mode):
        self.path = path
        self.mode = mode
        self.pos = 0
        if mode == "input":
            with open(path, "r", encoding="utf-8", errors="replace", newline="") as fh:
                self.buf = fh.read()
            self.fh = None
        else:
            m = {"output": "w", "append": "a"}.get(mode, "a")
            self.fh = open(path, m, encoding="utf-8", newline="")
            self.buf = None

    def eof(self):
        if self.mode != "input":
            return True
        return self.pos >= len(self.buf)

    def lof(self):
        if self.buf is not None:
            return len(self.buf.encode("utf-8"))
        self.fh.flush()
        return os.path.getsize(self.path)

    def write(self, s):
        if self.fh is None:
            raise rt_error(54)
        self.fh.write(s)

    def read_line(self):
        if self.buf is None:
            raise rt_error(54)
        b = self.buf
        if self.pos >= len(b):
            raise rt_error(62)
        j = self.pos
        while j < len(b) and b[j] not in "\r\n":
            j += 1
        line = b[self.pos:j]
        if j < len(b) and b[j] == "\r":
            j += 1
        if j < len(b) and b[j] == "\n":
            j += 1
        self.pos = j
        return line

    def read_chars(self, n):
        if self.buf is None:
            raise rt_error(54)
        if self.pos >= len(self.buf):
            raise rt_error(62)
        s = self.buf[self.pos:self.pos + n]
        self.pos += len(s)
        return s

    def read_field(self):
        b = self.buf
        if b is None:
            raise rt_error(54)
        n = len(b)
        i = self.pos
        while i < n and b[i] in " \t":
            i += 1
        if i >= n:
            raise rt_error(62)
        if b[i] == '"':
            j = b.find('"', i + 1)
            if j < 0:
                j = n
            val = b[i + 1:j]
            i = j + 1
            while i < n and b[i] in " \t":
                i += 1
            quoted = True
        else:
            j = i
            while j < n and b[j] not in ",\r\n":
                j += 1
            val = b[i:j].strip()
            i = j
            quoted = False
        if i < n and b[i] == ",":
            i += 1
        elif i < n and b[i] in "\r\n":
            if b[i] == "\r":
                i += 1
            if i < n and b[i] == "\n":
                i += 1
        self.pos = i
        if quoted:
            return val
        u = val.upper()
        if u == "#TRUE#":
            return True
        if u == "#FALSE#":
            return False
        if u == "#NULL#":
            return NULL
        if u.startswith("#") and u.endswith("#") and len(u) > 2:
            from . import dates as D
            d = D.parse_date_string(val[1:-1])
            if d is not None:
                return VDate(d)
        num = str_to_num_or_none(val)
        if num is not None:
            return int(num) if float(num).is_integer() and abs(num) < 2 ** 31 else num
        return val if val else EMPTY

    def close(self):
        if self.fh is not None:
            try:
                self.fh.close()
            except Exception:
                pass
            self.fh = None


class FileIOMixin:
    def cwd(self):
        if self.cur_dir:
            return self.cur_dir
        p = getattr(self.book, "path", None)
        if p:
            return os.path.dirname(os.path.abspath(p))
        return os.getcwd()

    def resolve_path(self, p):
        p = p.strip()
        if os.sep == "/" and "\\" in p and not os.path.exists(p):
            p = p.replace("\\", "/")
        p = os.path.expanduser(p)
        if not os.path.isabs(p):
            p = os.path.join(self.cwd(), p)
        return p

    def _fnum(self, n):
        if isinstance(n, VBAObject):
            n = to_value(n)
        if isinstance(n, str):
            n = n.lstrip("#")
        return index_int(n)

    def file_handle(self, n):
        k = self._fnum(n)
        h = self.files.get(k)
        if h is None:
            raise rt_error(52)
        return h

    def file_open(self, path, mode, n):
        k = self._fnum(n)
        if k in self.files:
            raise rt_error(55)
        if k < 1 or k > 511:
            raise rt_error(52)
        mode = (mode or "random").lower()
        if mode in ("binary", "random"):
            raise VBARuntimeErr(5, f"Open ... For {mode.title()} is not supported in PortableExcel")
        p = self.resolve_path(path)
        if mode == "input" and not os.path.isfile(p):
            raise rt_error(53)
        try:
            self.files[k] = FileHandle(p, mode)
        except FileNotFoundError:
            raise rt_error(76)
        except PermissionError:
            raise rt_error(70)

    def file_close(self, nums):
        if not nums:
            self.files_close_all()
            return
        for n in nums:
            k = self._fnum(n)
            h = self.files.pop(k, None)
            if h is not None:
                h.close()

    def files_close_all(self):
        for h in self.files.values():
            h.close()
        self.files.clear()

    def file_print(self, n, text, newline):
        self.file_handle(n).write(text + ("\r\n" if newline else ""))

    def file_write(self, n, values):
        parts = []
        for v in values:
            if isinstance(v, VBAObject):
                v = to_value(v)
            t = type(v)
            if t is str:
                parts.append('"' + v + '"')
            elif t is bool:
                parts.append("#TRUE#" if v else "#FALSE#")
            elif v is EMPTY:
                parts.append("")
            elif v is NULL:
                parts.append("#NULL#")
            elif t is VDate:
                from . import dates as D
                d = D.serial_to_dt(v)
                parts.append("#" + d.strftime("%Y-%m-%d") + ("" if float(v).is_integer() else d.strftime(" %H:%M:%S")) + "#")
            elif t is VErr:
                parts.append(f"#ERROR {v.number}#")
            else:
                parts.append(to_str(v))
        self.file_handle(n).write(",".join(parts) + "\r\n")

    def file_input_field(self, n):
        return self.file_handle(n).read_field()

    def file_line_input(self, n):
        return self.file_handle(n).read_line()

    def file_rename(self, old, new):
        a = self.resolve_path(old)
        b = self.resolve_path(new)
        if not os.path.exists(a):
            raise rt_error(53)
        if os.path.exists(b):
            raise rt_error(58)
        os.rename(a, b)

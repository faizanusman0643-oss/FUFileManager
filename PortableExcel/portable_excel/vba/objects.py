"""Built-in COM-like objects: Collection, Dictionary, Err, RegExp, FileSystemObject,
user class instances."""
from __future__ import annotations

import os
import re
import shutil
import glob as _glob
from datetime import datetime

from .errors import rt_error, VBARuntimeErr
from .values import (VBAObject, VBAArray, UDTValue, EMPTY, NULL, NOTHING, MISSING, VDate,
                     to_str, to_num, to_bool, to_long, index_int, to_value, type_name)
from . import dates as _dates


def bind(args, named, names, required=0):
    """Map positional + named arguments onto parameter `names` (lower-case).

    Returns a list (MISSING for not supplied)."""
    out = list(args[:len(names)])
    if len(args) > len(names):
        raise rt_error(450)
    out.extend([MISSING] * (len(names) - len(out)))
    if named:
        for k, v in named.items():
            try:
                out[names.index(k)] = v
            except ValueError:
                raise rt_error(448, f"Named argument not found: {k}")
    for i in range(required):
        if out[i] is MISSING:
            raise rt_error(449)
    return out


def opt(v, default):
    return default if v is MISSING else v


def index_value(v, args):
    """Apply an argument list to a value: array element or default member."""
    if type(v) is VBAArray:
        return v.get(args)
    if isinstance(v, VBAObject):
        return v.vba_get("", args, {})
    if v is NOTHING:
        raise rt_error(91)
    if not args:
        return v
    raise rt_error(13)


# --------------------------------------------------------------------------
# Collection
# --------------------------------------------------------------------------
class Collection(VBAObject):
    vba_typename = "Collection"

    def __init__(self):
        self.items = []    # list of [key_lower or None, value]
        self.keys = {}

    def _pos(self, ix):
        if isinstance(ix, str):
            k = ix.lower()
            if k not in self.keys:
                raise rt_error(5)
            ent = self.keys[k]
            for i, e in enumerate(self.items):
                if e is ent:
                    return i
            raise rt_error(5)
        i = index_int(ix)
        if i < 1 or i > len(self.items):
            raise rt_error(9)
        return i - 1

    def g_add(self, args, named):
        item, key, before, after = bind(args, named, ["item", "key", "before", "after"], 1)
        if isinstance(item, (VBAArray, UDTValue)):
            item = item.copy()
        ent = [None, item]
        if key is not MISSING:
            k = to_str(key)
            kl = k.lower()
            if kl in self.keys:
                raise rt_error(457)
            ent[0] = kl
        if before is not MISSING:
            pos = self._pos(to_value(before))
        elif after is not MISSING:
            pos = self._pos(to_value(after)) + 1
        else:
            pos = len(self.items)
        self.items.insert(pos, ent)
        if ent[0] is not None:
            self.keys[ent[0]] = ent
        return EMPTY

    def g_item(self, args, named):
        (ix,) = bind(args, named, ["index"], 1)
        ix = to_value(ix)
        if isinstance(ix, str):
            ent = self.keys.get(ix.lower())
            if ent is None:
                raise rt_error(5)
            return ent[1]
        return self.items[self._pos(ix)][1]

    g__default = g_item

    def g_count(self, args, named):
        return len(self.items)

    def g_remove(self, args, named):
        (ix,) = bind(args, named, ["index"], 1)
        p = self._pos(to_value(ix))
        ent = self.items.pop(p)
        if ent[0] is not None:
            self.keys.pop(ent[0], None)
        return EMPTY

    def vba_iter(self):
        return [e[1] for e in self.items]


# --------------------------------------------------------------------------
# Scripting.Dictionary
# --------------------------------------------------------------------------
class Dictionary(VBAObject):
    vba_typename = "Dictionary"

    def __init__(self):
        self.d = {}
        self.compare = 0

    def _norm(self, k):
        if isinstance(k, VBAObject):
            return ("o", id(k))
        t = type(k)
        if t is str:
            return ("s", k.lower() if self.compare else k)
        if t is bool:
            return ("n", -1.0 if k else 0.0)
        if t in (int, float, VDate):
            return ("n", float(k))
        if k is EMPTY:
            return ("e", None)
        try:
            return ("n", float(to_num(k)))
        except VBARuntimeErr:
            return ("x", repr(k))

    def g_add(self, args, named):
        key, item = bind(args, named, ["key", "item"], 2)
        nk = self._norm(key)
        if nk in self.d:
            raise rt_error(457)
        if isinstance(item, (VBAArray, UDTValue)):
            item = item.copy()
        self.d[nk] = [key, item]
        return EMPTY

    def g_item(self, args, named):
        (key,) = bind(args, named, ["key"], 1)
        nk = self._norm(key)
        ent = self.d.get(nk)
        if ent is None:
            ent = self.d[nk] = [key, EMPTY]
        return ent[1]

    g__default = g_item

    def s_item(self, args, named, value):
        (key,) = bind(args, named, ["key"], 1)
        if isinstance(value, (VBAArray, UDTValue)):
            value = value.copy()
        nk = self._norm(key)
        ent = self.d.get(nk)
        if ent is None:
            self.d[nk] = [key, value]
        else:
            ent[1] = value

    s__default = s_item

    def s_key(self, args, named, value):
        (key,) = bind(args, named, ["key"], 1)
        nk = self._norm(key)
        if nk not in self.d:
            raise rt_error(32811)
        nn = self._norm(value)
        if nn in self.d and nn != nk:
            raise rt_error(457)
        items = list(self.d.items())
        self.d = {}
        for k, ent in items:
            if k == nk:
                self.d[nn] = [value, ent[1]]
            else:
                self.d[k] = ent

    def g_exists(self, args, named):
        (key,) = bind(args, named, ["key"], 1)
        return self._norm(key) in self.d

    def g_remove(self, args, named):
        (key,) = bind(args, named, ["key"], 1)
        nk = self._norm(key)
        if nk not in self.d:
            raise rt_error(32811)
        del self.d[nk]
        return EMPTY

    def g_removeall(self, args, named):
        self.d.clear()
        return EMPTY

    def g_keys(self, args, named):
        return VBAArray.from_list([e[0] for e in self.d.values()], 0)

    def g_items(self, args, named):
        return VBAArray.from_list([e[1] for e in self.d.values()], 0)

    def g_count(self, args, named):
        return len(self.d)

    def g_comparemode(self, args, named):
        return self.compare

    def s_comparemode(self, args, named, value):
        if self.d:
            raise rt_error(5)
        self.compare = index_int(value)

    def vba_iter(self):
        return [e[0] for e in self.d.values()]


# --------------------------------------------------------------------------
# Err object
# --------------------------------------------------------------------------
class ErrObject(VBAObject):
    vba_typename = "ErrObject"

    def __init__(self):
        self.clear()

    def clear(self):
        self.number = 0
        self.description = ""
        self.source = ""
        self.helpfile = ""
        self.helpcontext = 0
        self.erl = 0

    def set_from(self, e: VBARuntimeErr, source_default="", erl=0):
        self.number = e.number
        self.description = e.description
        self.source = e.source or source_default
        self.erl = erl or 0

    def g_number(self, args, named):
        return self.number

    g__default = g_number

    def s_number(self, args, named, v):
        self.number = to_long(v)
        from .errors import error_message
        if self.number and not self.description:
            self.description = error_message(self.number)

    s__default = s_number

    def g_description(self, args, named):
        return self.description

    def s_description(self, args, named, v):
        self.description = to_str(v)

    def g_source(self, args, named):
        return self.source

    def s_source(self, args, named, v):
        self.source = to_str(v)

    def g_helpfile(self, args, named):
        return self.helpfile

    def s_helpfile(self, args, named, v):
        self.helpfile = to_str(v)

    def g_helpcontext(self, args, named):
        return self.helpcontext

    def s_helpcontext(self, args, named, v):
        self.helpcontext = v

    def g_lastdllerror(self, args, named):
        return 0

    def g_clear(self, args, named):
        self.clear()
        return EMPTY

    def g_raise(self, args, named):
        number, source, desc, hf, hc = bind(args, named, ["number", "source", "description",
                                                          "helpfile", "helpcontext"], 1)
        n = to_long(number)
        e = VBARuntimeErr(n, None if desc is MISSING else to_str(desc),
                          "" if source is MISSING else to_str(source))
        raise e


# --------------------------------------------------------------------------
# User class instances
# --------------------------------------------------------------------------
class ClassInstance(VBAObject):
    def __init__(self, rt, module):
        self.rt = rt
        self.module = module
        self.vars = module.new_instance_vars()
        self.terminated = False
        self.vba_typename = module.text
        if getattr(module, "we_vars", None):
            rt.we_holders.add(self)

    def vba_get(self, name, args, named):
        return self.rt.class_get(self, name, args, named)

    def vba_let(self, name, args, named, value, is_set=False):
        return self.rt.class_let(self, name, args, named, value, is_set)

    def vba_iter(self):
        p = self.module.find_enum_proc()
        if p is None:
            raise rt_error(438)
        res = self.rt.invoke(p, [], {}, self)[0]
        if isinstance(res, VBAObject):
            return res.vba_iter()
        raise rt_error(438)

    def __del__(self):
        try:
            if not self.terminated and self.module.terminate_proc is not None and self.rt.alive:
                self.rt.pending_term.append(self)
        except Exception:
            pass


class UnsupportedFormObj(VBAObject):
    vba_typename = "UserForm"

    def __init__(self, name):
        self.name = name

    def _fail(self):
        e = VBARuntimeErr(5, f"UserForms are not supported yet ({self.name})")
        e.fatal = True
        return e

    def vba_get(self, name, args, named):
        raise self._fail()

    def vba_let(self, name, args, named, value, is_set=False):
        raise self._fail()


class NoOpObj(VBAObject):
    """Accepts any property get/set; used for Borders, PageSetup etc."""
    vba_typename = "Object"

    def __init__(self, typename="Object", values=None):
        self.vba_typename = typename
        self.values = values or {}

    def vba_get(self, name, args, named):
        if name in self.values:
            return self.values[name]
        if name == "":
            return self
        return NoOpObj(self.vba_typename)

    def vba_let(self, name, args, named, value, is_set=False):
        self.values[name] = value

    def vba_iter(self):
        return []


# --------------------------------------------------------------------------
# VBScript.RegExp
# --------------------------------------------------------------------------
class RegExpObj(VBAObject):
    vba_typename = "IRegExp2"

    def __init__(self):
        self.pattern = ""
        self.glob = False
        self.ignorecase = False
        self.multiline = False

    def _rx(self):
        flags = 0
        if self.ignorecase:
            flags |= re.I
        if self.multiline:
            flags |= re.M
        try:
            return re.compile(self.pattern, flags)
        except re.error:
            raise rt_error(5017, "Syntax error in regular expression")

    def g_pattern(self, a, n):
        return self.pattern

    def s_pattern(self, a, n, v):
        self.pattern = to_str(v)

    def g_global(self, a, n):
        return self.glob

    def s_global(self, a, n, v):
        self.glob = to_bool(v)

    def g_ignorecase(self, a, n):
        return self.ignorecase

    def s_ignorecase(self, a, n, v):
        self.ignorecase = to_bool(v)

    def g_multiline(self, a, n):
        return self.multiline

    def s_multiline(self, a, n, v):
        self.multiline = to_bool(v)

    def g_test(self, args, named):
        (s,) = bind(args, named, ["sourcestring"], 1)
        return self._rx().search(to_str(s)) is not None

    def g_execute(self, args, named):
        (s,) = bind(args, named, ["sourcestring"], 1)
        s = to_str(s)
        rx = self._rx()
        ms = list(rx.finditer(s)) if self.glob else [m for m in [rx.search(s)] if m]
        return MatchCollection([MatchObj(m) for m in ms])

    def g_replace(self, args, named):
        s, rep = bind(args, named, ["sourcestring", "replacevar"], 2)
        s = to_str(s)
        rep = to_str(rep)
        out = []
        i = 0
        while i < len(rep):
            ch = rep[i]
            if ch == "$" and i + 1 < len(rep):
                nx = rep[i + 1]
                if nx.isdigit():
                    j = i + 1
                    while j < len(rep) and rep[j].isdigit() and j < i + 3:
                        j += 1
                    out.append("\\g<" + rep[i + 1:j] + ">")
                    i = j
                    continue
                if nx == "&":
                    out.append("\\g<0>")
                    i += 2
                    continue
                if nx == "$":
                    out.append("$")
                    i += 2
                    continue
            if ch == "\\":
                out.append("\\\\")
            else:
                out.append(ch)
            i += 1
        try:
            return self._rx().sub("".join(out), s, count=0 if self.glob else 1)
        except (re.error, IndexError):
            raise rt_error(5)


class MatchObj(VBAObject):
    vba_typename = "IMatch2"

    def __init__(self, m):
        self.m = m

    def g_value(self, a, n):
        return self.m.group(0)

    g__default = g_value

    def g_firstindex(self, a, n):
        return self.m.start()

    def g_length(self, a, n):
        return self.m.end() - self.m.start()

    def g_submatches(self, a, n):
        sub = [g if g is not None else EMPTY for g in self.m.groups()]
        c = MatchCollection(sub)
        c.vba_typename = "ISubMatches"
        if a:
            return c.g_item(a, n)
        return c


class MatchCollection(VBAObject):
    vba_typename = "IMatchCollection2"

    def __init__(self, items):
        self.items = items

    def g_count(self, a, n):
        return len(self.items)

    def g_item(self, args, named):
        i = index_int(to_value(args[0]))
        if i < 0 or i >= len(self.items):
            raise rt_error(5)
        return self.items[i]

    g__default = g_item

    def vba_iter(self):
        return list(self.items)


# --------------------------------------------------------------------------
# Scripting.FileSystemObject (basic)
# --------------------------------------------------------------------------
class FSO(VBAObject):
    vba_typename = "FileSystemObject"

    def __init__(self, rt):
        self.rt = rt

    def _p(self, v):
        return self.rt.resolve_path(to_str(v))

    def g_fileexists(self, a, n):
        return os.path.isfile(self._p(a[0]))

    def g_folderexists(self, a, n):
        return os.path.isdir(self._p(a[0]))

    def g_createfolder(self, a, n):
        p = self._p(a[0])
        if os.path.exists(p):
            raise rt_error(58)
        os.makedirs(p)
        return FolderObj(self.rt, p)

    def g_deletefile(self, a, n):
        pat = self._p(a[0])
        files = _glob.glob(pat)
        if not files:
            raise rt_error(53)
        for f in files:
            os.remove(f)
        return EMPTY

    def g_deletefolder(self, a, n):
        p = self._p(a[0])
        if not os.path.isdir(p):
            raise rt_error(76)
        shutil.rmtree(p)
        return EMPTY

    def g_copyfile(self, a, n):
        src, dst, over = bind(a, n, ["source", "destination", "overwritefiles"], 2)
        s, d = self._p(src), self._p(dst)
        if os.path.exists(d) and not os.path.isdir(d) and over is not MISSING and not to_bool(over):
            raise rt_error(58)
        shutil.copy(s, d)
        return EMPTY

    def g_movefile(self, a, n):
        shutil.move(self._p(a[0]), self._p(a[1]))
        return EMPTY

    g_movefolder = g_movefile

    def g_copyfolder(self, a, n):
        shutil.copytree(self._p(a[0]), self._p(a[1]), dirs_exist_ok=True)
        return EMPTY

    def g_getfilename(self, a, n):
        return os.path.basename(to_str(a[0]).replace("\\", "/"))

    def g_getbasename(self, a, n):
        return os.path.splitext(os.path.basename(to_str(a[0]).replace("\\", "/")))[0]

    def g_getextensionname(self, a, n):
        e = os.path.splitext(to_str(a[0]))[1]
        return e[1:] if e else ""

    def g_getparentfoldername(self, a, n):
        return os.path.dirname(to_str(a[0]).rstrip("/\\"))

    def g_buildpath(self, a, n):
        p, name = to_str(a[0]), to_str(a[1])
        if not p:
            return name
        sep = "\\" if "\\" in p and "/" not in p else os.sep
        return p.rstrip("/\\") + sep + name

    def g_getabsolutepathname(self, a, n):
        return os.path.abspath(self._p(a[0]))

    def g_getspecialfolder(self, a, n):
        import tempfile
        return FolderObj(self.rt, tempfile.gettempdir())

    def g_gettempname(self, a, n):
        import random
        return "rad%05X.tmp" % random.randint(0, 0xFFFFF)

    def g_getfile(self, a, n):
        p = self._p(a[0])
        if not os.path.isfile(p):
            raise rt_error(53)
        return FileObj(self.rt, p)

    def g_getfolder(self, a, n):
        p = self._p(a[0])
        if not os.path.isdir(p):
            raise rt_error(76)
        return FolderObj(self.rt, p)

    def g_opentextfile(self, a, n):
        path, iomode, create, fmt = bind(a, n, ["filename", "iomode", "create", "format"], 1)
        p = self._p(path)
        mode = 1 if iomode is MISSING else index_int(iomode)
        create = False if create is MISSING else to_bool(create)
        if mode == 1:
            if not os.path.isfile(p):
                raise rt_error(53)
            return TextStream(open(p, "r", encoding="utf-8", errors="replace", newline=""), "r")
        if not os.path.exists(p) and not create:
            raise rt_error(53)
        return TextStream(open(p, "a" if mode == 8 else "w", encoding="utf-8", newline=""), "w")

    def g_createtextfile(self, a, n):
        path, over, uni = bind(a, n, ["filename", "overwrite", "unicode"], 1)
        p = self._p(path)
        if os.path.exists(p) and over is not MISSING and not to_bool(over):
            raise rt_error(58)
        return TextStream(open(p, "w", encoding="utf-8", newline=""), "w")

    def g_drives(self, a, n):
        return Collection()


class FileObj(VBAObject):
    vba_typename = "File"

    def __init__(self, rt, path):
        self.rt = rt
        self.path = path

    def g_name(self, a, n):
        return os.path.basename(self.path)

    g__default = g_name

    def g_path(self, a, n):
        return self.path

    def g_size(self, a, n):
        return os.path.getsize(self.path)

    def g_datelastmodified(self, a, n):
        return VDate(_dates.dt_to_serial(datetime.fromtimestamp(os.path.getmtime(self.path))))

    g_datecreated = g_datelastmodified

    def g_parentfolder(self, a, n):
        return FolderObj(self.rt, os.path.dirname(self.path))

    def g_type(self, a, n):
        return os.path.splitext(self.path)[1].lstrip(".").upper() + " File"

    def g_delete(self, a, n):
        os.remove(self.path)
        return EMPTY

    def g_openastextstream(self, a, n):
        mode = 1 if not a or a[0] is MISSING else index_int(a[0])
        if mode == 1:
            return TextStream(open(self.path, "r", encoding="utf-8", newline=""), "r")
        return TextStream(open(self.path, "a" if mode == 8 else "w", encoding="utf-8", newline=""), "w")


class FolderObj(FileObj):
    vba_typename = "Folder"

    def g_files(self, a, n):
        c = Collection()
        for name in sorted(os.listdir(self.path)):
            p = os.path.join(self.path, name)
            if os.path.isfile(p):
                c.items.append([name.lower(), FileObj(self.rt, p)])
                c.keys[name.lower()] = c.items[-1]
        return c

    def g_subfolders(self, a, n):
        c = Collection()
        for name in sorted(os.listdir(self.path)):
            p = os.path.join(self.path, name)
            if os.path.isdir(p):
                c.items.append([name.lower(), FolderObj(self.rt, p)])
                c.keys[name.lower()] = c.items[-1]
        return c

    def g_size(self, a, n):
        total = 0
        for root, _, files in os.walk(self.path):
            for f in files:
                try:
                    total += os.path.getsize(os.path.join(root, f))
                except OSError:
                    pass
        return total


class TextStream(VBAObject):
    vba_typename = "TextStream"

    def __init__(self, fh, mode):
        self.fh = fh
        self.mode = mode
        self.line = 1
        self._buf = None
        if mode == "r":
            self._buf = fh.read()
            self._pos = 0

    def g_atendofstream(self, a, n):
        return self._pos >= len(self._buf) if self.mode == "r" else True

    def g_line(self, a, n):
        return self.line

    def g_readline(self, a, n):
        if self.mode != "r":
            raise rt_error(54)
        if self._pos >= len(self._buf):
            raise rt_error(62)
        b = self._buf
        j = self._pos
        while j < len(b) and b[j] not in "\r\n":
            j += 1
        line = b[self._pos:j]
        if j < len(b) and b[j] == "\r":
            j += 1
        if j < len(b) and b[j] == "\n":
            j += 1
        self._pos = j
        self.line += 1
        return line

    def g_readall(self, a, n):
        s = self._buf[self._pos:]
        self._pos = len(self._buf)
        return s

    def g_read(self, a, n):
        k = index_int(a[0])
        s = self._buf[self._pos:self._pos + k]
        self._pos += len(s)
        return s

    def g_skipline(self, a, n):
        self.g_readline(a, n)
        return EMPTY

    def g_write(self, a, n):
        self.fh.write(to_str(a[0]))
        return EMPTY

    def g_writeline(self, a, n):
        self.fh.write((to_str(a[0]) if a and a[0] is not MISSING else "") + "\r\n")
        self.line += 1
        return EMPTY

    def g_writeblanklines(self, a, n):
        self.fh.write("\r\n" * index_int(a[0]))
        return EMPTY

    def g_close(self, a, n):
        try:
            self.fh.close()
        except Exception:
            pass
        return EMPTY

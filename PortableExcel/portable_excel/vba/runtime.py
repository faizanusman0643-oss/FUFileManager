"""VBARuntime: project loading, procedure invocation, name resolution, events, files."""
from __future__ import annotations

import os
import random
import re
import sys

from .errors import (VBAError, VBARuntimeErr, VBACompileErr, VBAControl, VBAEnd, VBAInterrupt,
                     VBAStop, rt_error, error_message)
from .values import (EMPTY, NULL, NOTHING, MISSING, VDate, VErr, VBAArray, UDTValue, UDTDef,
                     VBAObject, TypeSpec, VARIANT, to_value, to_str, to_num, type_name,
                     fmt_double)
from .objects import (Collection, Dictionary, ErrObject, ClassInstance, UnsupportedFormObj,
                      RegExpObj, FSO, NoOpObj, index_value)
from .builtins import BUILTINS, CONSTANTS
from .parser import parse_module, parse_statements, Parser
from . import ast as A
from .compiler import Compiler, Frame, VarInfo, execute, APP_GLOBALS, SHEET_SCOPED, RESUME_NEXT
from .sheetops import SheetOpsMixin
from .fileio import FileIOMixin

SCALAR_TYPES = {"integer", "long", "longlong", "longptr", "byte", "boolean", "single", "double",
                "currency", "decimal", "date", "string"}
KNOWN_OBJECT_TYPES = {"object", "range", "worksheet", "workbook", "collection", "dictionary",
                      "scripting.dictionary", "application", "sheets", "worksheets", "name",
                      "font", "interior", "chart", "shape", "window", "workbooks", "errobject",
                      "regexp", "vbscript.regexp", "filesystemobject", "scripting.filesystemobject",
                      "textstream", "file", "folder", "match", "matchcollection", "areas",
                      "excel.range", "excel.worksheet", "excel.workbook", "excel.application",
                      "msforms.userform", "control", "listobject", "pivottable", "variant"}

FORM_MEMBERS = {"controls", "caption", "hide", "show", "repaint", "activecontrol", "width",
                "height", "left", "top", "backcolor", "forecolor", "font", "tag",
                "startupposition", "insidewidth", "insideheight", "enabled", "visible",
                "setfocus", "move"}

EVENT_NAMES = re.compile(r"^(worksheet|workbook|class|userform|app|chart)_", re.I)


class ParamInfo:
    __slots__ = ("name", "text", "slot", "byref", "optional", "default_node", "default_value",
                 "paramarray", "ts", "is_array")

    def __init__(self, p, ts, slot):
        self.name = p.name
        self.text = p.text
        self.slot = slot
        self.byref = p.byref
        self.optional = p.optional
        self.default_node = p.default
        self.default_value = MISSING
        self.paramarray = p.paramarray
        self.ts = ts
        self.is_array = p.is_array


class Proc:
    def __init__(self, module, pdef, rt):
        self.module = module
        self.pdef = pdef
        self.lname = pdef.name
        self.text = pdef.text
        self.kind = pdef.kind
        self.private = pdef.private
        self.static = pdef.static
        self.ret_ts = rt.type_spec(pdef.ret, module.mdef_ctx)
        self.ret_is_array = False
        self.params = [ParamInfo(p, rt.type_spec(p.type, module.mdef_ctx), i + 1)
                       for i, p in enumerate(pdef.params)]
        self.code = None
        self.lines = None
        self.nslots = 0
        self.template = None
        self.factories = None
        self.max_with = 0
        self.statics = []
        self.static_index = {}
        self.compile_error = None

    @property
    def full_name(self):
        return f"{self.module.name}.{self.text}"


class Module:
    def __init__(self, rt, mdict):
        self.rt = rt
        self.name = mdict.get("name") or "Module"
        self.lname = self.name.lower()
        t = (mdict.get("type") or "standard").lower()
        if t in ("std", "module", "bas"):
            t = "standard"
        if t in ("doc", "sheet", "workbook"):
            t = "document"
        self.type = t
        self.code = mdict.get("code") or ""
        self.form_def = mdict.get("form") if t == "form" else None
        self.mdef = parse_module(self.name, t, self.code)
        self.mdef_ctx = self
        self.explicit = self.mdef.option_explicit
        self.option_base = self.mdef.option_base
        self.text_compare = self.mdef.option_compare_text
        self.text = self.name
        self.procs = {}
        self.prop_get = {}
        self.prop_let = {}
        self.prop_set = {}
        self.var_index = {}
        self.var_ts = []
        self.var_is_array = []
        self.var_private = []
        self.var_decls = []
        self.vars = []
        self.consts = {}           # lname -> (VarDecl, private)
        self.const_values = {}
        self.udts = {}
        self.enums = {}            # lname -> EnumDef
        self.enum_values = {}      # enum lname -> {member: value}
        self.enum_members = {}     # member -> enum lname
        self.declares = dict(self.mdef.declares)
        self.terminate_proc = None
        self.initialize_proc = None
        self._evaluating = set()

    # setup in two phases: types first (all modules), then the rest
    def setup_types(self):
        for td in self.mdef.types:
            self.udts[td.name] = (td, None)
        for ed in self.mdef.enums:
            self.enums[ed.name] = ed
            for mname, mtext, e in ed.members:
                self.enum_members[mname] = ed.name

    def setup(self):
        rt = self.rt
        for name, (td, d) in list(self.udts.items()):
            self.udt_def(name)
        for d, private in self.mdef.vars:
            ts = rt.type_spec(d.type, self)
            self.var_index[d.name] = len(self.vars)
            self.var_ts.append(ts)
            self.var_is_array.append(d.is_array)
            self.var_private.append(private)
            self.var_decls.append(d)
            self.vars.append(EMPTY)
        for d, private in self.mdef.consts:
            self.consts[d.name] = (d, private)
        for pdef in self.mdef.procs:
            p = Proc(self, pdef, rt)
            if p.kind in ("sub", "function"):
                self.procs[p.lname] = p
            elif p.kind == "get":
                self.prop_get[p.lname] = p
            elif p.kind == "let":
                self.prop_let[p.lname] = p
            else:
                self.prop_set[p.lname] = p
        self.we_vars = [(i, d.name) for i, d in enumerate(self.var_decls)
                        if getattr(d, "withevents", False)]
        if self.type == "class":
            self.terminate_proc = self.procs.get("class_terminate")
            self.initialize_proc = self.procs.get("class_initialize")

    def udt_def(self, name):
        td, d = self.udts[name]
        if d is None:
            d = UDTDef(td.text, [])
            self.udts[name] = (td, d)
            fields = []
            for f in td.fields:
                ts = self.rt.type_spec(f.type, self)
                dims = None
                if f.is_array:
                    if f.dims:
                        dims = []
                        for lb, ub in f.dims:
                            lo = self.const_eval(lb) if lb is not None else self.option_base
                            dims.append((int(to_num(lo)), int(to_num(self.const_eval(ub)))))
                    else:
                        dims = []
                fields.append((f.name, f.text, ts, dims if f.is_array else None))
            d.fields = fields
        return d

    def const_eval(self, expr):
        from .compiler import compile_module_expr
        return compile_module_expr(self.rt, self, expr)

    def init_vars(self):
        for i, d in enumerate(self.var_decls):
            if self.type in ("class", "form"):
                continue
            self.vars[i] = self.make_var_default(i)

    def make_var_default(self, i):
        d = self.var_decls[i]
        ts = self.var_ts[i]
        if d.is_array:
            if d.dims:
                dims = []
                for lb, ub in d.dims:
                    lo = self.const_eval(lb) if lb is not None else self.option_base
                    dims.append((int(to_num(lo)), int(to_num(self.const_eval(ub)))))
                return VBAArray(dims, ts, fixed=True)
            return VBAArray(None, ts)
        if ts.kind == "object" and ts.new:
            return self.rt.new_object(ts.cls, self)
        return ts.default()

    def new_instance_vars(self):
        return [self.make_var_default(i) for i in range(len(self.var_decls))]

    def find_enum_proc(self):
        for p in list(self.procs.values()) + list(self.prop_get.values()):
            if re.search(r"VB_UserMemId\s*=\s*-4", self.code, re.I) and \
                    re.search(rf"Attribute\s+{re.escape(p.text)}\.VB_UserMemId\s*=\s*-4", self.code, re.I):
                return p
        return None

    def const_value(self, name):
        if name in self.const_values:
            return self.const_values[name]
        if name in self._evaluating:
            raise VBACompileErr("Circular constant reference", None, self.name)
        self._evaluating.add(name)
        try:
            d, _ = self.consts[name]
            v = self.const_eval(d.init)
            ts = self.rt.type_spec(d.type, self)
            if ts.kind == "scalar":
                v = ts.coerce(v)
            self.const_values[name] = v
            return v
        finally:
            self._evaluating.discard(name)

    def enum_dict(self, ename):
        if ename in self.enum_values:
            return self.enum_values[ename]
        ed = self.enums[ename]
        vals = {}
        self.enum_values[ename] = vals
        nxt = 0
        for mname, mtext, e in ed.members:
            if e is not None:
                v = int(to_num(self.const_eval(e)))
            else:
                v = nxt
            vals[mname] = v
            nxt = v + 1
        return vals

    def resolve_member(self, name, rt, inside=True):
        i = self.var_index.get(name)
        if i is not None and (inside or not self.var_private[i]):
            kind = "cls" if self.type in ("class", "form") else "mod"
            return ("var", VarInfo(kind, i, self.var_ts[i], self.var_is_array[i], name,
                                   None if kind == "cls" else self.vars))
        c = self.consts.get(name)
        if c is not None and (inside or not c[1]):
            return ("const", self.const_value(name))
        p = self.procs.get(name) or self.prop_get.get(name) or self.prop_let.get(name) \
            or self.prop_set.get(name)
        if p is not None and (inside or not p.private):
            if self.type in ("class", "form") and not inside:
                return None
            return ("proc", p)
        if self.type == "form" and inside:
            for c in (self.form_def or {}).get("controls") or []:
                if (c.get("name") or "").lower() == name:
                    return ("formctrl", c["name"])
            if name in FORM_MEMBERS:
                return ("memember", name)
        e = self.enum_members.get(name)
        if e is not None:
            return ("const", self.enum_dict(e)[name])
        if name in self.enums:
            return ("enum", self.enum_dict(name))
        if name in self.declares and inside:
            return ("declare", self.declares[name])
        return None


class VBAHost:
    """Default host: console output, auto-answers dialogs. GUI subclasses it."""

    def __init__(self):
        self._selection = None
        self.output = []

    def msgbox(self, prompt, buttons=0, title=""):
        print(f"[MsgBox] {title}: {prompt}")
        b = buttons & 15
        return {0: 1, 1: 1, 2: 3, 3: 6, 4: 6, 5: 4}.get(b, 1)

    def inputbox(self, prompt, title="", default=""):
        return default

    def debug_print(self, text):
        self.output.append(text)
        print(text)

    def select(self, sheet, r1, c1, r2, c2):
        self._selection = (sheet, r1, c1, r2, c2)

    def get_selection(self):
        return self._selection

    def activate_sheet(self, sheet):
        pass

    def refresh(self):
        pass

    def status_bar(self, text):
        pass

    def process_events(self):
        pass


def _proc_name_ok_for_macro(p):
    return p.kind == "sub" and not p.private and not p.params


class VBARuntime(SheetOpsMixin, FileIOMixin):
    def __init__(self, book, engine, host=None):
        self.book = book
        self.engine = engine
        self.host = host if host is not None else VBAHost()
        self.modules = {}
        self.module_list = []
        self.err = ErrObject()
        self.project_name = "VBAProject"
        self.budget = 2000
        self.stop_requested = False
        self.depth = 0
        self.max_depth = 1500
        self.pending_term = []
        self.alive = True
        self.rng = random.Random()
        self.last_rnd = 0.0
        self.cur_mod = None
        self.cur_dir = None
        self.dir_iter = None
        self.files = {}
        self._types = {}
        self._sheet_objs = {}
        self._indexes = {}
        self.protected = set()
        self.sort_objects = {}
        self.clipboard = None
        self.find_state = None
        self.event_depth = 0
        self.running = 0
        self.udf_depth = 0
        self.trap_hook = None      # debugging: fn(proc, line, err) for trapped errors
        from .apiemu import ApiEmulator
        self.api = ApiEmulator(self)
        self.addressof_procs = {}
        self._clip_text = None
        import weakref
        self.we_holders = weakref.WeakSet()
        self.udf_steps = 0
        self._udf_cache = {}
        self.live_forms = []
        self.form_defaults = {}
        self.modal_depth = 0
        self.end_pending = False
        self._globals_cache = {}
        self._wb = None
        from .excel import ApplicationObj, WorkbookObj, WindowObj
        self.app = ApplicationObj(self)
        self.workbook = WorkbookObj(self)
        self.window = WindowObj(self)
        if sys.getrecursionlimit() < 20000:
            sys.setrecursionlimit(20000)

    # ------------------------------------------------------------------
    # loading
    # ------------------------------------------------------------------
    def load_modules(self, modules):
        self.modules = {}
        self.module_list = []
        self._types = {}
        self._globals_cache = {}
        errors = []
        for md in modules or []:
            try:
                m = Module(self, md)
            except VBACompileErr as e:
                errors.append(VBAError("Compile error: " + e.message, md.get("name"), e.line))
                continue
            self.modules[m.lname] = m
            self.module_list.append(m)
            for e in m.mdef.errors:
                errors.append(VBAError("Compile error: " + e.message.replace("Syntax error: ", "Syntax error: ")
                                       if not e.message.startswith("Syntax") else "Compile error: " + e.message,
                                       m.name, e.line))
        for m in self.module_list:
            m.setup_types()
        for m in self.module_list:
            try:
                m.setup()
            except VBACompileErr as e:
                errors.append(VBAError("Compile error: " + e.message, m.name, e.line))
        self._initialized = False
        return errors

    def ensure_init(self):
        if getattr(self, "_initialized", False):
            return
        self._initialized = True
        for m in self.module_list:
            try:
                m.init_vars()
            except VBACompileErr as e:
                raise VBAError("Compile error: " + e.message, m.name, e.line)

    def reset_state(self):
        """End statement: reset module-level variables."""
        for m in self.module_list:
            for p in list(m.procs.values()) + list(m.prop_get.values()):
                p.statics[:] = []
                p.static_index.clear()
                p.code = None
            m.vars[:] = [EMPTY] * len(m.vars)
        self._initialized = False
        self.files_close_all()
        self.close_all_forms()

    def close_all_forms(self):
        for inst in list(self.live_forms):
            inst.close_view()
        self.live_forms[:] = []
        self.form_defaults.clear()

    def form_default(self, m):
        inst = self.form_defaults.get(m.lname)
        if inst is None or not inst.loaded:
            from .userforms import FormInstance
            self.ensure_init()
            inst = FormInstance(self, m, default=True)
            self.form_defaults[m.lname] = inst
            inst.initialize()
        return inst

    # ------------------------------------------------------------------
    # types
    # ------------------------------------------------------------------
    def base_type(self, name):
        ts = self._types.get(name)
        if ts is None:
            if name in SCALAR_TYPES:
                ts = TypeSpec(name, "scalar")
            elif name in ("variant", "any"):
                ts = VARIANT
            else:
                ts = TypeSpec(name, "object", cls=name)
            self._types[name] = ts
        return ts

    def type_spec(self, tref, module=None):
        if tref is None:
            return VARIANT
        if isinstance(tref, str):
            return self.base_type(tref.lower())
        name = tref.name
        if tref.size is not None:
            n = int(to_num(module.const_eval(tref.size))) if module is not None else 0
            return TypeSpec("string", "fixedstr", size=n)
        if name in SCALAR_TYPES or name in ("variant", "any"):
            return self.base_type(name)
        # user defined type?
        if module is not None and isinstance(module, Module):
            short = name.split(".")[-1]
            if short in module.udts:
                return TypeSpec(short, "udt", udt=module.udt_def(short))
            for m in self.module_list:
                if short in m.udts and m.type == "standard":
                    return TypeSpec(short, "udt", udt=m.udt_def(short))
            if short in module.enums or any(short in m.enums for m in self.module_list):
                return self.base_type("long")
        if tref.new:
            return TypeSpec(name, "object", cls=name, new=True)
        if name.startswith(("vb", "xl", "mso")) and name not in KNOWN_OBJECT_TYPES and \
                name not in self.modules:
            return self.base_type("long")
        return self.base_type(name)

    # ------------------------------------------------------------------
    # resolution
    # ------------------------------------------------------------------
    def resolve_global_user(self, name):
        for m in self.module_list:
            if m.type != "standard":
                continue
            r = m.resolve_member(name, self, inside=False)
            if r is not None and r[0] != "declare":
                return r
        return None

    def resolve_global(self, name):
        r = self.resolve_global_user(name)
        if r is not None:
            return r
        m = self.modules.get(name)
        if m is not None:
            if m.type == "standard":
                return ("module", m)
            if m.type == "class":
                return ("class", m)
            if m.type == "form":
                return ("form", m)
        if name == "thisworkbook" or (m is not None and m.type == "document"):
            return ("docobj", self._docobj_getter(name))
        for sh in self.book.sheets:
            if sh.code_name.lower() == name:
                return ("docobj", self._docobj_getter(name))
        if name in CONSTANTS:
            return ("const", CONSTANTS[name])
        if name == "err":
            return ("const", self.err)
        if name == "debug":
            return ("const", DebugObj(self))
        if name == "vba":
            return ("vbalib",)
        if name in ("excel",):
            return ("const", ExcelLib(self))
        if name in BUILTINS:
            return ("builtin", BUILTINS[name])
        if name in APP_GLOBALS:
            return ("global", name)
        if name in ("scripting", "vbscript"):
            return ("const", NoOpObj(name))
        return None

    def _docobj_getter(self, lname):
        def get():
            return self.doc_object(lname)
        return get

    def doc_object(self, lname):
        lname = lname.lower()
        if lname == "thisworkbook":
            return self.workbook
        for sh in self.book.sheets:
            if sh.code_name.lower() == lname:
                return self.sheet_object(sh)
        raise rt_error(424)

    def doc_object_for_module(self, m):
        if m.lname == "thisworkbook":
            return self.workbook
        for sh in self.book.sheets:
            if sh.code_name.lower() == m.lname:
                return self.sheet_object(sh)
        return self.workbook

    def doc_module(self, code_name):
        m = self.modules.get(code_name.lower())
        if m is not None and m.type in ("document", "class") and m.type == "document":
            return m
        return None

    def sheet_object(self, sheet):
        o = self._sheet_objs.get(id(sheet))
        if o is None or o.sheet is not sheet:
            from .excel import WorksheetObj
            o = self._sheet_objs[id(sheet)] = WorksheetObj(self, sheet)
        return o

    def doc_member_get(self, mod, obj, name, args, named):
        from .excel import _NOTFOUND
        r = mod.resolve_member(name, self, inside=False) if name else None
        if r is None:
            return _NOTFOUND
        k = r[0]
        if k == "proc":
            p = r[1]
            if p.kind in ("let", "set"):
                p = mod.prop_get.get(p.lname)
                if p is None:
                    return _NOTFOUND
            return self.invoke(p, list(args), named, obj)[0]
        if k == "var":
            v = mod.vars[r[1].idx]
            return index_value(v, list(args)) if args else v
        if k == "const":
            return r[1]
        return _NOTFOUND

    def doc_member_let(self, mod, obj, name, args, named, value, is_set):
        i = mod.var_index.get(name)
        if i is not None and not mod.var_private[i]:
            if args:
                cont = mod.vars[i]
                if type(cont) is VBAArray:
                    cont.set(list(args), value)
                    return True
            ts = mod.var_ts[i]
            mod.vars[i] = value if is_set else ts.coerce(value if not isinstance(value, VBAObject) else to_value(value))
            return True
        p = (mod.prop_set.get(name) if is_set else None) or mod.prop_let.get(name)
        if p is not None and not p.private:
            self.invoke(p, list(args) + [value], named, obj)
            return True
        return False

    # ------------------------------------------------------------------
    # objects
    # ------------------------------------------------------------------
    def check_class_name(self, cls, mod, line):
        return True

    def new_object(self, cls, mod=None):
        name = cls.lower()
        short = name.split(".")[-1]
        m = self.modules.get(short)
        if m is not None and m.type == "class":
            inst = ClassInstance(self, m)
            if m.initialize_proc is not None:
                self.invoke(m.initialize_proc, [], {}, inst)
            return inst
        if m is not None and m.type == "form":
            from .userforms import FormInstance
            inst = FormInstance(self, m)
            inst.initialize()
            return inst
        if short == "collection":
            return Collection()
        if short == "dictionary":
            return Dictionary()
        if short in ("regexp", "regexp55"):
            return RegExpObj()
        if short == "filesystemobject":
            return FSO(self)
        if short == "dataobject":
            from .objects import DataObj
            return DataObj(self)
        raise VBARuntimeErr(429, f"ActiveX component can't create object ({cls})")

    def create_com(self, progid):
        p = progid.strip().lower()
        if p == "scripting.dictionary":
            return Dictionary()
        if p == "scripting.filesystemobject":
            return FSO(self)
        if p in ("new:{1c3b4210-f441-11ce-b9ea-00aa006b1a69}", "forms.dataobject",
                 "msforms.dataobject", "{1c3b4210-f441-11ce-b9ea-00aa006b1a69}"):
            from .objects import DataObj
            return DataObj(self)
        if p in ("vbscript.regexp",):
            return RegExpObj()
        if p in ("excel.application",):
            return self.app
        if p in ("system.collections.arraylist",):
            raise VBARuntimeErr(429, "ActiveX component can't create object (System.Collections.ArrayList)")
        if p.startswith("wscript.") or p.startswith("shell."):
            raise VBARuntimeErr(429, f"{progid} is not supported in PortableExcel (no shell access)")
        raise VBARuntimeErr(429, f"ActiveX component can't create object ({progid})")

    def typeof_is(self, v, tn):
        tn = tn.lower()
        if not isinstance(v, VBAObject):
            if v is NOTHING:
                return False
            raise rt_error(424)
        if tn == "object":
            return True
        names = {v.vba_typename.lower()}
        if tn == "sheet" or tn == "worksheet":
            pass
        return tn in names or (tn == "dictionary" and isinstance(v, Dictionary))

    def class_get(self, inst, name, args, named):
        m = inst.module
        p = m.prop_get.get(name) or m.procs.get(name)
        if p is not None and not p.private:
            return self.invoke(p, list(args), named, inst)[0]
        i = m.var_index.get(name)
        if i is not None and not m.var_private[i]:
            v = inst.vars[i]
            return index_value(v, list(args)) if args else v
        if name == "":
            raise rt_error(438)
        raise rt_error(438)

    def class_let(self, inst, name, args, named, value, is_set):
        m = inst.module
        p = None
        if is_set:
            p = m.prop_set.get(name)
        if p is None:
            p = m.prop_let.get(name)
            if p is None and not is_set and isinstance(value, VBAObject):
                p = m.prop_set.get(name)
        if p is not None and not p.private:
            self.invoke(p, list(args) + [value], named, inst)
            return
        i = m.var_index.get(name)
        if i is not None and not m.var_private[i]:
            if args:
                cont = inst.vars[i]
                if type(cont) is VBAArray:
                    cont.set(list(args), value)
                    return
                raise rt_error(13)
            ts = m.var_ts[i]
            if is_set:
                inst.vars[i] = value
            else:
                if isinstance(value, VBAObject):
                    value = to_value(value)
                inst.vars[i] = ts.coerce(value)
            return
        raise rt_error(438)

    def run_terminates(self):
        while self.pending_term:
            inst = self.pending_term.pop()
            if inst.terminated:
                continue
            inst.terminated = True
            try:
                self.invoke(inst.module.terminate_proc, [], {}, inst)
            except VBARuntimeErr:
                pass

    # ------------------------------------------------------------------
    # invocation
    # ------------------------------------------------------------------
    def compile(self, proc):
        if proc.code is not None:
            return
        if proc.compile_error is not None:
            raise proc.compile_error
        try:
            Compiler(self, proc.module, proc).compile_proc()
        except VBACompileErr as e:
            msg = e.message
            if not msg.startswith("Compile error"):
                msg = "Compile error: " + msg
            err = VBAError(msg, proc.module.name, e.line or proc.pdef.line, 0, e.message, proc.text)
            proc.compile_error = err
            proc.code = None
            raise err

    def invoke(self, proc, args, named, me):
        if proc.code is None:
            self.compile(proc)
        v = proc.template[:]
        for slot, fac in proc.factories:
            v[slot] = fac()
        f = Frame(v, me, proc.max_with, proc)
        params = proc.params
        na = len(args)
        if na > len(params) and not (params and params[-1].paramarray):
            raise rt_error(450)
        for i, p in enumerate(params):
            if p.paramarray:
                rest = list(args[i:])
                v[p.slot] = VBAArray.from_list(rest, 0)
                break
            val = args[i] if i < na else MISSING
            if val is MISSING and named and p.name in named:
                val = named[p.name]
            if val is MISSING:
                if not p.optional:
                    raise rt_error(449)
                val = p.default_value
                if val is MISSING:
                    if p.ts.kind != "variant" or p.is_array:
                        val = p.ts.default() if not p.is_array else VBAArray(None, p.ts)
                    v[p.slot] = val
                    continue
            if p.is_array:
                if type(val) is not VBAArray:
                    if isinstance(val, VBAObject):
                        val = to_value(val)
                    if type(val) is not VBAArray:
                        raise rt_error(13)
                if not p.byref:
                    val = val.copy()
            else:
                k = p.ts.kind
                if k == "variant":
                    if not p.byref and (type(val) is VBAArray or type(val) is UDTValue):
                        val = val.copy()
                elif k == "object":
                    if not (isinstance(val, VBAObject) or val is NOTHING):
                        raise rt_error(13)
                elif k == "udt":
                    if type(val) is not UDTValue:
                        raise rt_error(13)
                    if not p.byref:
                        val = val.copy()
                else:
                    if isinstance(val, VBAObject):
                        val = to_value(val)
                    val = p.ts.coerce(val)
            v[p.slot] = val
        if named:
            pn = {p.name for p in params}
            for k in named:
                if k not in pn:
                    raise rt_error(448)
        self.depth += 1
        if self.depth > self.max_depth:
            self.depth -= 1
            raise rt_error(28)
        prev_mod = self.cur_mod
        self.cur_mod = proc.module
        try:
            execute(self, proc, f)
        finally:
            self.depth -= 1
            self.cur_mod = prev_mod
        if f.err_mode is not None or f.in_handler:
            self.err.clear()
        if self.pending_term:
            self.run_terminates()
        return v[0], f

    def cur_text_compare(self):
        m = self.cur_mod
        return bool(m and m.text_compare)

    def cur_option_base(self):
        m = self.cur_mod
        return m.option_base if m else 0

    def as_vba_error(self, ex):
        if isinstance(ex, VBARuntimeErr):
            return ex
        if isinstance(ex, VBAError):
            e = VBARuntimeErr(ex.vba_number or 5, ex.description)
            e.fatal = True
            e.module, e.line = ex.module, ex.line
            if ex.vba_number == 0:
                e.compile_error = ex
            return e
        if isinstance(ex, VBACompileErr):
            e = VBARuntimeErr(0, ex.message)
            e.fatal = True
            e.compile_error = VBAError("Compile error: " + ex.message, ex.module, ex.line)
            return e
        if isinstance(ex, ZeroDivisionError):
            return VBARuntimeErr(11)
        if isinstance(ex, OverflowError):
            return VBARuntimeErr(6)
        if isinstance(ex, RecursionError):
            return VBARuntimeErr(28)
        if isinstance(ex, MemoryError):
            return VBARuntimeErr(7)
        if isinstance(ex, FileNotFoundError):
            return VBARuntimeErr(53)
        if isinstance(ex, PermissionError):
            return VBARuntimeErr(70)
        if isinstance(ex, IsADirectoryError):
            return VBARuntimeErr(75)
        if isinstance(ex, OSError):
            return VBARuntimeErr(75, f"Path/File access error ({ex})")
        return VBARuntimeErr(51, f"Internal error: {type(ex).__name__}: {ex}")

    # ------------------------------------------------------------------
    # control
    # ------------------------------------------------------------------
    def poll(self):
        self.budget = 2000
        if self.udf_depth:
            self.udf_steps += 2000
            if self.udf_steps > 5_000_000:
                raise VBARuntimeErr(28, "User-defined function ran too long")
            return
        try:
            self.host.process_events()
        except Exception:
            pass
        if self.stop_requested:
            self.stop_requested = False
            raise VBAInterrupt()

    def check_stop(self):
        if self.stop_requested:
            self.stop_requested = False
            raise VBAInterrupt()

    def request_stop(self):
        self.stop_requested = True
        self.budget = -1

    def _to_vba_error(self, e, proc=None):
        if isinstance(e, VBAError):
            return e
        if isinstance(e, VBARuntimeErr):
            ce = getattr(e, "compile_error", None)
            if ce is not None:
                return ce
            if e.number == 0 and e.fatal:
                return VBAError(e.description, e.module, e.line, 0)
            num = e.number
            if num < 0:
                code = f"{num} ({num & 0xFFFFFFFF:08X})"
            else:
                code = str(num)
            return VBAError(f"Run-time error '{code}': {e.description}", e.module, e.line, num,
                            e.description, e.procedure)
        if isinstance(e, VBAInterrupt):
            return VBAError("Execution interrupted", self.cur_mod.name if self.cur_mod else None,
                            None, 18, "Execution interrupted")
        if isinstance(e, VBAStop):
            return VBAError(e.message, e.module, e.line, 0, e.message)
        if isinstance(e, VBACompileErr):
            return VBAError("Compile error: " + e.message, e.module, e.line, 0)
        return VBAError(f"Internal error: {type(e).__name__}: {e}", None, None, 51)

    def _top(self, fn):
        """Run fn() as a top-level VBA entry point."""
        self.ensure_init()
        if self.running == 0:
            self.stop_requested = False
            self._indexes.clear()
            self.budget = 2000
        self.running += 1
        try:
            return fn()
        except VBAEnd:
            if self.running > 1:
                # End inside an event handler while a macro waits in a modal Show:
                # close everything; the suspended macro stops when Show returns.
                self.end_pending = True
                self.close_all_forms()
                return None
            self.reset_state()
            self.end_pending = False
            return None
        except (VBARuntimeErr, VBAControl, VBACompileErr) as e:
            raise self._to_vba_error(e) from None
        except VBAError:
            raise
        except RecursionError:
            raise VBAError("Run-time error '28': Out of stack space", None, None, 28) from None
        finally:
            self.running -= 1
            self.depth = 0 if self.running == 0 else self.depth
            if self.running == 0:
                self.cur_mod = None
                self.app.reset()
                try:
                    self.engine_invalidate()
                    self.host.refresh()
                except Exception:
                    pass

    # ------------------------------------------------------------------
    # public API
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    # worksheet user-defined functions
    # ------------------------------------------------------------------
    def attach_udfs(self, engine=None):
        """Let formulas call Public Functions of standard modules."""
        eng = engine if engine is not None else self.engine
        if eng is not None:
            eng.udf_resolver = self.udf_call
            try:
                eng.invalidate()
            except Exception:
                pass

    def detach_udfs(self, engine=None):
        eng = engine if engine is not None else self.engine
        if eng is not None and getattr(eng, "udf_resolver", None) == self.udf_call:
            eng.udf_resolver = None

    def find_udf(self, name):
        low = name.lower()
        for pre in ("_xlfn.", "_xludf.", "_xlws."):
            if low.startswith(pre):
                low = low[len(pre):]
        key = low
        if key in self._udf_cache:
            return self._udf_cache[key]
        p = None
        if "." in low:
            mname, fname = low.rsplit(".", 1)
            m = self.modules.get(mname)
            if m is not None and m.type == "standard":
                p = m.procs.get(fname)
        else:
            for m in self.module_list:
                if m.type == "standard" and not m.mdef.option_private:
                    q = m.procs.get(low)
                    if q is not None and q.kind == "function" and not q.private:
                        p = q
                        break
        if p is not None and (p.kind != "function" or p.private):
            p = None
        self._udf_cache[key] = p
        return p

    def udf_call(self, name, args, sheet):
        from ..model import ERR_NAME, ERR_VALUE, XLError
        p = self.find_udf(name)
        if p is None:
            return ERR_NAME
        if self.udf_depth > 20:
            return ERR_VALUE
        from .sheetops import _xl_to_vba, _vba_to_xl
        from .excel import RangeObj
        vargs = []
        for a in args:
            if hasattr(a, "r1") and hasattr(a, "sheet") and hasattr(a, "c2"):
                vargs.append(RangeObj(self, a.sheet, a.r1, a.c1, a.r2, a.c2))
            elif isinstance(a, list):
                rows = [[_xl_to_vba(x) for x in (r if isinstance(r, list) else [r])] for r in a]
                vargs.append(VBAArray.from_2d(rows, 1))
            elif isinstance(a, float) and a.is_integer() and abs(a) < 2 ** 31:
                vargs.append(a)
            else:
                vargs.append(_xl_to_vba(a))
        if len(vargs) > len(p.params) and not (p.params and p.params[-1].paramarray):
            return ERR_VALUE
        self.udf_depth += 1
        if self.udf_depth == 1:
            self.udf_steps = 0
        prev_mod, prev_depth = self.cur_mod, self.depth
        try:
            self.ensure_init()
            res = self.invoke(p, vargs, {}, None)[0]
        except (VBAError, VBARuntimeErr, VBAControl, VBACompileErr, RecursionError):
            return ERR_VALUE
        except Exception:
            return ERR_VALUE
        finally:
            self.udf_depth -= 1
            self.cur_mod, self.depth = prev_mod, prev_depth
        try:
            if isinstance(res, VBAObject):
                res = to_value(res)
            if type(res) is VBAArray:
                res = res.data[0] if res.data else EMPTY
                if isinstance(res, VBAObject):
                    res = to_value(res)
        except (VBARuntimeErr, Exception):
            return ERR_VALUE
        if res is EMPTY or res is MISSING:
            return 0.0
        if res is NULL or res is NOTHING:
            return ERR_VALUE
        if type(res) is VErr:
            from .values import XL_CODE_TO_TEXT
            from ..model import ERRORS_BY_CODE
            return ERRORS_BY_CODE.get(XL_CODE_TO_TEXT.get(res.number, "#VALUE!"), ERR_VALUE)
        if isinstance(res, bool) or isinstance(res, str):
            return res
        if isinstance(res, (int, float)):
            return float(res)
        try:
            return float(res)
        except Exception:
            return to_str(res)

    def addressof(self, proc):
        for k, p in self.addressof_procs.items():
            if p is proc:
                return k
        addr = 0x7F100000 + 16 * len(self.addressof_procs)
        self.addressof_procs[addr] = proc
        return addr

    def mouse_wheel(self, delta, x=None, y=None, instance=None):
        """GUI -> runtime: the mouse wheel turned (delta: +120 per notch up).

        Calls VBA low-level mouse hooks registered with SetWindowsHookEx
        (WH_MOUSE_LL / WH_MOUSE) as the real hook would. Returns True when a
        hook handled it (the GUI should then not scroll by itself)."""
        from .apiemu import WM_MOUSEWHEEL, HOOK_LPARAM
        handled = False
        hooks = [h for h in self.api.hooks.values() if h[1] is not None and h[0] in (7, 14)]
        if not hooks:
            return False
        if x is None or y is None:
            inst = instance or next((i for i in reversed(self.live_forms) if i.visible), None)
            if inst is not None:
                l, t, r, b = self.api.form_rect(inst)
                x, y = (l + r) // 2, (t + b) // 2
            else:
                x, y = 0, 0
        self.api.cursor = (int(x), int(y))
        self.api.wheel_delta = int(delta)
        try:
            for idhook, proc in hooks:
                out = self._top(lambda proc=proc: self.invoke(proc, [0, WM_MOUSEWHEEL, HOOK_LPARAM], {}, None))
                if out is not None and out[0] not in (0, EMPTY, False):
                    handled = True
        finally:
            self.api.cursor = None
            self.api.wheel_delta = 0
        return handled

    def list_macros(self):
        out = []
        for m in self.module_list:
            if m.type not in ("standard", "document"):
                continue
            for p in m.procs.values():
                if not _proc_name_ok_for_macro(p):
                    continue
                if m.type == "document" and EVENT_NAMES.match(p.text):
                    continue
                if m.mdef.option_private and m.type == "standard":
                    continue
                out.append(f"{m.name}.{p.text}")
        return out

    def find_proc(self, name):
        if "!" in name:
            name = name.split("!", 1)[1]
        name = name.strip().strip("'")
        if "." in name:
            mname, pname = name.rsplit(".", 1)
            m = self.modules.get(mname.lower())
            if m is None:
                return None
            return m.procs.get(pname.lower())
        low = name.lower()
        for m in self.module_list:
            if m.type == "standard" and low in m.procs:
                return m.procs[low]
        for m in self.module_list:
            if low in m.procs and m.type != "class":
                return m.procs[low]
        return None

    def run(self, name, args=()):
        def go():
            p = self.find_proc(name)
            if p is None:
                raise VBAError(f"Cannot run the macro '{name}'. The macro may not be available in this workbook.",
                               None, None, 1004)
            me = self.doc_object_for_module(p.module) if p.module.type == "document" else None
            return self.invoke(p, list(args), {}, me)[0]
        return self._top(go)

    def run_nested(self, name, args):
        if name.strip().lower().split("!")[-1] == "pe_addbarcode":
            return self.pe_add_barcode(list(args))
        p = self.find_proc(name)
        if p is None:
            raise VBARuntimeErr(1004, f"Cannot run the macro '{name}'. The macro may not be available in this workbook.")
        me = self.doc_object_for_module(p.module) if p.module.type == "document" else None
        return self.invoke(p, list(args), {}, me)[0]

    def pe_add_barcode(self, args):
        """PortableExcel extension:
        Application.Run "PE_AddBarcode", Source, Symbology, Anchor[, Text]
        Source: cell address whose displayed value is encoded ("" = use Text);
        Anchor: cell where the picture's top-left goes. Returns the shape name."""
        from .. import barcodes
        from .excel import RangeObj
        args = list(args) + [MISSING] * 4
        src, sym, anchor, text = args[:4]
        if sym is MISSING or anchor is MISSING:
            raise rt_error(449)
        sheet = None
        src_addr = None
        if isinstance(src, RangeObj):
            sheet = src.sheet
            src_addr = src.address(False, False).split(":")[0]
        elif src is not MISSING and to_str(src):
            rng = self.range_from_address(to_str(src), None)
            sheet = rng.sheet
            src_addr = rng.address(False, False).split(":")[0]
        a = anchor if isinstance(anchor, RangeObj) else \
            self.range_from_address(to_str(anchor), sheet)
        sheet = a.sheet if sheet is None else sheet
        if a.sheet is not sheet:
            raise VBARuntimeErr(1004, "PE_AddBarcode: source and anchor must be on the same sheet")
        try:
            img = barcodes.make_barcode_image(
                sheet, src_addr, None if text is MISSING else to_str(text), to_str(sym),
                engine=self.engine, row=a.r1, col=a.c1)
        except ValueError as e:
            raise VBARuntimeErr(5, f"PE_AddBarcode: {e}")
        sheet.images.append(img)
        h = getattr(self.host, "images_changed", None)
        if h is not None:
            try:
                h(sheet)
            except Exception:
                pass
        self.host_refresh_soft()
        return img["name"]

    def _wrap_event_arg(self, a):
        from ..model import Sheet
        if isinstance(a, Sheet):
            return self.sheet_object(a)
        if a is None:
            return EMPTY
        if isinstance(a, int) and not isinstance(a, bool):
            return a
        return a

    # ------------------------------------------------------------------
    # WithEvents
    # ------------------------------------------------------------------
    def we_targets(self, match, event):
        """(proc, me) pairs for WithEvents variables whose value satisfies
        match(value): handler <VarName>_<Event> in the holder's module."""
        out = []
        ev = event.lower()
        for holder in list(self.we_holders):
            m = holder.module
            for idx, vname in m.we_vars:
                try:
                    v = holder.vars[idx]
                except IndexError:
                    continue
                if isinstance(v, VBAObject) and match(v):
                    p = m.procs.get(f"{vname}_{ev}")
                    if p is not None:
                        out.append((p, holder))
        for m in self.module_list:
            if m.type in ("document", "standard") and getattr(m, "we_vars", None):
                for idx, vname in m.we_vars:
                    v = m.vars[idx] if idx < len(m.vars) else None
                    if isinstance(v, VBAObject) and match(v):
                        p = m.procs.get(f"{vname}_{ev}")
                        if p is not None:
                            me = self.doc_object_for_module(m) if m.type == "document" else None
                            out.append((p, me))
        return out

    def has_we(self):
        return bool(len(self.we_holders)) or any(getattr(m, "we_vars", None) and m.type == "document"
                                                 for m in self.module_list)

    @staticmethod
    def fit_args(p, args):
        vargs = list(args)
        n = len(p.params)
        if p.params and p.params[-1].paramarray:
            return vargs
        while len(vargs) < n:
            pi = p.params[len(vargs)]
            vargs.append(False if pi.ts.name == "boolean" or pi.name == "cancel" else
                         0 if pi.ts.kind == "scalar" and pi.ts.name != "string" else EMPTY)
        return vargs[:n]

    def run_targets(self, targets, args, nested=False):
        """Run event handlers in order, threading ByRef args through.
        Returns the final argument values."""
        args = list(args)
        seen = set()
        for p, me in targets:
            key = (id(p), id(me))
            if key in seen:
                continue
            seen.add(key)
            vargs = self.fit_args(p, args)
            if nested:
                res, f = self.invoke(p, vargs, {}, me)
            else:
                out = self._top(lambda p=p, vargs=vargs, me=me: self.invoke(p, vargs, {}, me))
                if out is None:
                    return args
                res, f = out
            new = [f.v[q.slot] for q in p.params]
            args = new + args[len(new):]
        return args

    _APP_EVENT = {"open": "WorkbookOpen", "beforeclose": "WorkbookBeforeClose",
                  "beforesave": "WorkbookBeforeSave", "activate": "WorkbookActivate",
                  "deactivate": "WorkbookDeactivate", "newsheet": "WorkbookNewSheet"}

    def fire_event(self, module_name, proc, args=()):
        if not self.app.enableevents:
            return None
        m = self.modules.get((module_name or "").lower())
        targets = []
        p = m.procs.get(proc.lower()) if m is not None else None
        if p is not None:
            me = self.doc_object_for_module(m) if m.type == "document" else None
            targets.append((p, me))
        low = proc.lower()
        if self.has_we() and "_" in low:
            from .excel import WorkbookObj, WorksheetObj, ApplicationObj
            ev = low.split("_", 1)[1]
            if low.startswith("workbook_"):
                targets += self.we_targets(lambda v: isinstance(v, WorkbookObj), ev)
                app_ev = self._APP_EVENT.get(ev, ev if ev.startswith("sheet") else "Workbook" + ev)
                targets += self.we_targets(lambda v: isinstance(v, ApplicationObj), app_ev)
            elif low.startswith("worksheet_") and m is not None:
                sh = next((s for s in self.book.sheets if s.code_name.lower() == m.lname), None)
                if sh is not None:
                    targets += self.we_targets(lambda v: isinstance(v, WorksheetObj) and v.sheet is sh, ev)
        if not targets:
            return None
        vargs = [self._wrap_event_arg(a) for a in args]
        if self.event_depth > 50:
            return None
        self.event_depth += 1
        try:
            return self.run_targets(targets, vargs)
        finally:
            self.event_depth -= 1

    def has_handler(self, module_name, proc):
        m = self.modules.get(module_name.lower())
        return m is not None and proc.lower() in m.procs

    def notify_change(self, sheet, r1, c1, r2, c2):
        if not self.app.enableevents:
            return
        cn = sheet.code_name
        h1 = self.has_handler(cn, "worksheet_change") or self.has_we()
        h2 = self.has_handler("thisworkbook", "workbook_sheetchange") or self.has_we()
        if not (h1 or h2):
            return
        from .excel import RangeObj
        rng = RangeObj(self, sheet, r1, c1, r2, c2)
        if h1:
            self.fire_event(cn, "Worksheet_Change", [rng])
        if h2 and self.app.enableevents:
            self.fire_event("ThisWorkbook", "Workbook_SheetChange", [self.sheet_object(sheet), rng])

    def make_range(self, sheet, r1, c1, r2, c2):
        from .excel import RangeObj
        return RangeObj(self, sheet, r1, c1, r2, c2)

    def evaluate_immediate(self, text):
        t = text.strip()
        if not t:
            return ""
        out = []
        old = self.host.debug_print

        def capture(s):
            out.append(s)
        is_print = t.startswith("?") or re.match(r"(?i)print\s", t)
        if is_print:
            t = "Debug.Print " + (t[1:] if t.startswith("?") else t[5:].lstrip())
        imm = self._immediate_module()
        try:
            body = parse_statements(t, "Immediate")
        except VBACompileErr as e:
            return "Compile error: " + e.message
        from .ast import ProcDef, TypeRef
        pdef = ProcDef("__immediate__", "Immediate", "sub", [], TypeRef(), body, True, False, 1, imm.mdef)
        pdef.end_line = 1
        p = Proc(imm, pdef, self)
        if is_print:
            self.host.debug_print = capture
        try:
            self._top(lambda: self.invoke(p, [], {}, None))
        except VBAError as e:
            return str(e.message)
        finally:
            if is_print:
                self.host.debug_print = old
        return "\n".join(out)

    def _immediate_module(self):
        m = getattr(self, "_imm", None)
        if m is None:
            m = Module(self, {"name": "Immediate", "type": "standard", "code": ""})
            m.setup_types()
            m.setup()
            self._imm = m
        return m

    # ------------------------------------------------------------------
    # printing
    # ------------------------------------------------------------------
    def print_repr(self, v):
        if isinstance(v, VBAObject):
            v = to_value(v)
        t = type(v)
        if t in (int, float) or t.__name__ == "Decimal":
            s = to_str(v)
            return (s if s.startswith("-") else " " + s) + " "
        if v is NULL:
            return "Null"
        if t is VErr:
            return f"Error {v.number}"
        if v is MISSING:
            return "Error 448"
        return to_str(v)

    def workbook_name(self):
        p = self.book.path
        if p:
            return os.path.basename(p)
        return getattr(self.book, "title", None) or "Book1"


class DebugObj(VBAObject):
    vba_typename = "Debug"

    def __init__(self, rt):
        self.rt = rt

    def g_print(self, args, named):
        self.rt.host.debug_print(" ".join(self.rt.print_repr(a) for a in args))
        return EMPTY

    def g_assert(self, args, named):
        from .values import truthy
        if args and not truthy(args[0]):
            raise VBAStop("Debug.Assert failed")
        return EMPTY


class ExcelLib(VBAObject):
    vba_typename = "Excel"

    def __init__(self, rt):
        self.rt = rt

    def vba_get(self, name, args, named):
        if name == "application":
            return self.rt.app
        return self.rt.app_global(name, args, named, None)

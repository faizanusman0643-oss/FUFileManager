"""Compiles procedure ASTs into flat lists of Python closures and executes them."""
from __future__ import annotations

from .errors import (VBARuntimeErr, VBACompileErr, VBAControl, VBAEnd, VBAStop, rt_error)
from .values import (EMPTY, NULL, NOTHING, MISSING, VDate, VErr, VBAArray, UDTValue, VBAObject,
                     TypeSpec, VARIANT, to_value, to_str, to_num, index_int, truthy, type_name,
                     var_type, op_add, op_sub, op_mul, op_div, op_intdiv, op_mod, op_pow, op_neg,
                     op_concat, op_like, op_is, op_and, op_or, op_xor, op_eqv, op_imp, op_not,
                     CMP_BINARY, CMP_TEXT, to_float)
from .objects import index_value, UnsupportedFormObj
from .builtins import BUILTINS, CONSTANTS
from . import ast as A

RESUME_NEXT = "resume-next"
NO_ARGS = ()
NO_NAMED = {}

ARITH = {"+": op_add, "-": op_sub, "*": op_mul, "/": op_div, "\\": op_intdiv, "mod": op_mod,
         "^": op_pow, "&": op_concat, "and": op_and, "or": op_or, "xor": op_xor,
         "eqv": op_eqv, "imp": op_imp, "is": op_is}

APP_GLOBALS = {"range", "cells", "rows", "columns", "activesheet", "activecell", "activeworkbook",
               "thisworkbook", "selection", "sheets", "worksheets", "workbooks", "application",
               "worksheetfunction", "names", "intersect", "union", "evaluate", "calculate",
               "activewindow", "charts", "activechart", "windows", "run", "screenupdating",
               "statusbar", "displayalerts", "enableevents", "calculation", "cutcopymode",
               "username", "wait", "ontime", "volatile", "addins", "commandbars", "dialogs",
               "filedialog", "getopenfilename", "getsaveasfilename", "activeprinter",
               "displaystatusbar", "version", "caption", "workbook_dummy"}
SHEET_SCOPED = {"range", "cells", "rows", "columns", "evaluate", "names"}


class Label:
    __slots__ = ("pc", "name")

    def __init__(self, name=""):
        self.pc = -1
        self.name = name


class Frame:
    __slots__ = ("v", "me", "withs", "err_mode", "in_handler", "err_pc", "gosub", "erl", "proc")

    def __init__(self, v, me, nwith, proc):
        self.v = v
        self.me = me
        self.withs = [None] * nwith if nwith else None
        self.err_mode = None
        self.in_handler = False
        self.err_pc = 0
        self.gosub = None
        self.erl = 0
        self.proc = proc


class VarInfo:
    __slots__ = ("kind", "idx", "ts", "is_array", "name", "container", "retval")

    def __init__(self, kind, idx, ts, is_array=False, name="", container=None, retval=False):
        self.kind = kind          # local | mod | cls | static
        self.idx = idx
        self.ts = ts
        self.is_array = is_array
        self.name = name
        self.container = container
        self.retval = retval


def member_get(obj, name, args, named):
    if isinstance(obj, VBAObject):
        return obj.vba_get(name, args, named)
    if type(obj) is UDTValue:
        v = obj.get_field(name)
        return index_value(v, args) if args else v
    if obj is NOTHING:
        raise rt_error(91)
    raise rt_error(424)


def let_value(v):
    if isinstance(v, VBAObject) or v is NOTHING:
        return to_value(v)
    return v


def iter_items(v):
    if type(v) is VBAArray:
        return v.iter_values_colmajor()
    if isinstance(v, VBAObject):
        return v.vba_iter()
    if v is NOTHING:
        raise rt_error(91)
    raise rt_error(92 if v is EMPTY else 424)


# --------------------------------------------------------------------------
# Execution
# --------------------------------------------------------------------------
def execute(rt, proc, f):
    code = proc.code
    lines = proc.lines
    n = len(code)
    pc = 0
    while True:
        try:
            while pc < n:
                r = code[pc](f)
                if r is None:
                    pc += 1
                else:
                    pc = r
            return
        except VBAControl as c:
            if isinstance(c, VBAStop) and c.line is None:
                c.module = proc.module.name
                c.line = lines[pc] if pc < n else proc.pdef.end_line
            raise
        except Exception as ex:
            e = rt.as_vba_error(ex)
            mode = f.err_mode
            if mode is None or f.in_handler or e.fatal:
                if e.line is None:
                    e.module = proc.module.name
                    e.line = lines[pc] if pc < n else proc.pdef.end_line
                    e.procedure = proc.text
                if e is ex:
                    raise
                raise e from None
            rt.err.set_from(e, rt.project_name, f.erl)
            if mode is RESUME_NEXT:
                pc += 1
            else:
                f.in_handler = True
                f.err_pc = pc
                pc = mode.pc


# --------------------------------------------------------------------------
# Compiler
# --------------------------------------------------------------------------
class Compiler:
    def __init__(self, rt, module, proc=None):
        self.rt = rt
        self.module = module
        self.proc = proc
        self.locals = {}
        self.nslots = 1 if proc is not None else 0
        self.code = []
        self.lines = []
        self.labels = {}
        self.loops = []
        self.with_depth = 0
        self.max_with = 0
        self.factories = []       # (slot, factory)
        self.template = [EMPTY] * self.nslots
        self.end_label = Label("<end>")
        self.cmp = CMP_TEXT if module.text_compare else CMP_BINARY
        self.explicit = module.explicit
        self.cur_line = proc.pdef.line if proc is not None else 0

    # -- errors -----------------------------------------------------------
    def error(self, msg, line=None):
        raise VBACompileErr(msg, line or self.cur_line, self.module.name)

    # -- slots ----------------------------------------------------------------
    def new_slot(self, default=EMPTY):
        s = self.nslots
        self.nslots += 1
        self.template.append(default)
        return s

    def declare_local(self, name, ts, is_array=False, dims=None, static=False, line=0):
        if name in self.locals and not self.locals[name].retval:
            vi = self.locals[name]
            if vi.kind == "localconst":
                self.error("Duplicate declaration in current scope", line)
            return vi
        rt = self.rt
        if is_array:
            if dims:
                bounds = [self.const_eval(lb) if lb is not None else self.module.option_base
                          for lb, _ in dims]
                ubs = [self.const_eval(ub) for _, ub in dims]
                dd = [(index_int(lb), index_int(ub)) for lb, ub in zip(bounds, ubs)]

                def fac(dd=dd, ts=ts):
                    return VBAArray(list(dd), ts, fixed=True)
            else:
                def fac(ts=ts):
                    return VBAArray(None, ts)
        elif ts.kind == "udt":
            fac = ts.default
        elif ts.kind == "object" and ts.new:
            cls = ts.cls

            def fac(cls=cls, mod=self.module):
                return rt.new_object(cls, mod)
        else:
            fac = None
        if static or (self.proc is not None and self.proc.static):
            store = self.proc.statics
            key = name
            if key not in self.proc.static_index:
                self.proc.static_index[key] = len(store)
                store.append(fac() if fac else ts.default())
            vi = VarInfo("static", self.proc.static_index[key], ts, is_array, name, store)
        else:
            slot = self.new_slot(ts.default() if fac is None else EMPTY)
            if fac is not None:
                self.factories.append((slot, fac))
            vi = VarInfo("local", slot, ts, is_array, name)
        self.locals[name] = vi
        return vi

    def const_eval(self, expr):
        fn = self.c_expr(expr)
        dummy = Frame([EMPTY] * max(self.nslots, 1), None, 0, None)
        try:
            return fn(dummy)
        except VBARuntimeErr as e:
            self.error("Constant expression required", getattr(expr, "line", None))

    # -- procedure compilation -------------------------------------------------
    def compile_proc(self):
        proc = self.proc
        pdef = proc.pdef
        if pdef.error is not None:
            raise pdef.error
        rt = self.rt
        # return value
        if proc.kind in ("function", "get"):
            ts = proc.ret_ts
            self.template[0] = ts.default() if ts.kind != "udt" else EMPTY
            if ts.kind == "udt":
                self.factories.append((0, ts.default))
            self.locals[proc.lname] = VarInfo("local", 0, ts, proc.ret_is_array, proc.lname,
                                              retval=True)
        # params
        for p in proc.params:
            slot = self.new_slot()
            p.slot = slot
            self.locals[p.name] = VarInfo("local", slot, p.ts, p.is_array or p.paramarray, p.name)
        for p in proc.params:
            if p.default_node is not None:
                p.default_value = self.const_eval(p.default_node)
        # declarations (procedure-wide scope)
        self.collect_decls(pdef.body)
        self.collect_implicit(pdef.body)
        self.c_block(pdef.body)
        self.end_label.pc = len(self.code)
        for name, lab in self.labels.items():
            if lab.pc < 0:
                self.error(f"Label not defined: {lab.name}")
        proc.code = self.code
        proc.lines = self.lines
        proc.nslots = self.nslots
        proc.template = self.template
        proc.factories = self.factories
        proc.max_with = self.max_with

    def collect_decls(self, body):
        for st in body:
            t = type(st)
            if t is A.SDim:
                self.cur_line = st.line
                for d in st.decls:
                    if st.kind == "const":
                        if d.name in self.locals:
                            self.error("Duplicate declaration in current scope", st.line)
                        ts = self.rt.type_spec(d.type, self.module)
                        val = self.const_eval(d.init)
                        if ts.kind not in ("variant",):
                            val = ts.coerce(val)
                        vi = VarInfo("localconst", val, ts, False, d.name)
                        self.locals[d.name] = vi
                    else:
                        if d.name in self.locals and not self.locals[d.name].retval:
                            if self.locals[d.name].kind != "localconst":
                                self.error("Duplicate declaration in current scope", st.line)
                        ts = self.rt.type_spec(d.type, self.module)
                        self.declare_local(d.name, ts, d.is_array, d.dims, st.kind == "static", st.line)
            else:
                for sub in self._sub_bodies(st):
                    self.collect_decls(sub)

    @staticmethod
    def _sub_bodies(st):
        t = type(st)
        if t is A.SIf:
            out = [b for _, b in st.branches]
            if st.else_:
                out.append(st.else_)
            return out
        if t in (A.SFor, A.SForEach, A.SDo, A.SWith):
            return [st.body]
        if t is A.SSelect:
            out = [b for _, b, _ in st.cases]
            if st.else_:
                out.append(st.else_)
            return out
        return []

    def collect_implicit(self, body):
        for st in body:
            t = type(st)
            names = []
            if t is A.SAssign and type(st.target) is A.Ident:
                names.append(st.target)
            elif t in (A.SFor, A.SForEach) and type(st.var) is A.Ident:
                names.append(st.var)
            elif t is A.SInput:
                names.extend(x for x in st.targets if type(x) is A.Ident)
            elif t is A.SReDim:
                for target, dims, tref in st.items:
                    if type(target) is A.Ident and self.lookup_var(target.name) is None:
                        ts = self.rt.type_spec(tref, self.module) if tref else VARIANT
                        self.declare_local(target.name, ts, True, None, False, st.line)
            for node in names:
                if node.name in self.locals:
                    continue
                b = self.resolve(node.name, implicit=False)
                if b is None or b[0] in ("builtin", "global") and node.name not in ("activecell", "selection"):
                    if b is not None and b[0] == "global":
                        continue
                    if self.explicit:
                        if b is None:
                            self.error("Variable not defined", st.line)
                        continue
                    self.declare_local(node.name, self.suffix_ts(node), False, None, False, st.line)
            for sub in self._sub_bodies(st):
                self.collect_implicit(sub)

    def suffix_ts(self, node):
        from .lexer import TYPE_SUFFIX
        sfx = getattr(node, "suffix", "")
        if sfx and sfx in TYPE_SUFFIX:
            return self.rt.base_type(TYPE_SUFFIX[sfx])
        return VARIANT

    def lookup_var(self, name):
        if name in self.locals:
            return self.locals[name]
        b = self.resolve(name, implicit=False)
        if b is not None and b[0] == "var":
            return b[1]
        return None

    # -- name resolution -----------------------------------------------------------
    def resolve(self, name, implicit=True, line=None):
        vi = self.locals.get(name)
        if vi is not None:
            if vi.kind == "localconst":
                return ("const", vi.idx)
            return ("var", vi)
        rt = self.rt
        m = self.module
        r = m.resolve_member(name, rt, inside=True)
        if r is not None:
            return r
        r = rt.resolve_global(name)
        if r is not None:
            return r
        if implicit:
            if self.explicit:
                self.error("Variable not defined", line)
            vi = self.declare_local(name, VARIANT)
            return ("var", vi)
        return None

    # -- code emission ---------------------------------------------------------
    def emit(self, fn, line=None):
        self.code.append(fn)
        self.lines.append(line if line is not None else self.cur_line)
        return len(self.code) - 1

    def label(self, name):
        lab = self.labels.get(name)
        if lab is None:
            lab = self.labels[name] = Label(name)
        return lab

    def c_block(self, body):
        for st in body:
            self.cur_line = st.line
            self.c_stmt(st)

    def c_stmt(self, st):
        m = getattr(self, "s_" + type(st).__name__)
        m(st)

    # statements ---------------------------------------------------------------
    def s_SNop(self, st):
        pass

    def s_SDim(self, st):
        pass

    def s_SLabel(self, st):
        lab = self.label(st.name)
        if lab.pc >= 0:
            self.error("Duplicate label", st.line)
        lab.pc = len(self.code)
        if st.number is not None:
            n = st.number

            def ins(f):
                f.erl = n
            self.emit(ins, st.line)

    def s_SAssign(self, st):
        fn = self.c_assign(st.target, st.expr, st.is_set, st.kind)
        self.emit(fn, st.line)

    def s_SCall(self, st):
        fn = self.c_callstmt(st)
        self.emit(fn, st.line)

    def s_SIf(self, st):
        end = Label()
        for cond, body in st.branches:
            nxt = Label()
            c = self.c_expr(cond)

            def test(f, c=c, nxt=nxt):
                v = c(f)
                if v is True:
                    return None
                if v is False:
                    return nxt.pc
                if not truthy(v):
                    return nxt.pc
            self.emit(test, cond.line or st.line)
            self.c_block(body)
            self.emit(lambda f, end=end: end.pc, st.line)
            nxt.pc = len(self.code)
        if st.else_:
            self.c_block(st.else_)
        end.pc = len(self.code)

    def s_SFor(self, st):
        rt = self.rt
        load = self.make_load(st.var)
        store = self.make_store(st.var, False)
        if store is None:
            self.error("Invalid For loop control variable", st.line)
        start = self.c_expr(st.start)
        endf = self.c_expr(st.end)
        stepf = self.c_expr(st.step) if st.step is not None else None
        s_end = self.new_slot()
        s_step = self.new_slot()
        exit_lab = Label()
        body_lab = Label()

        def init(f):
            a = start(f)
            if type(a) is not int and type(a) is not float:
                a = to_num(let_value(a))
            e = endf(f)
            if type(e) is not int and type(e) is not float:
                e = to_num(let_value(e))
            if stepf is None:
                s = 1
            else:
                s = stepf(f)
                if type(s) is not int and type(s) is not float:
                    s = to_num(let_value(s))
            f.v[s_end] = e
            f.v[s_step] = s
            store(f, a)
            a = load(f)
            if (s >= 0 and a > e) or (s < 0 and a < e):
                return exit_lab.pc
        self.emit(init, st.line)
        body_lab.pc = len(self.code)
        self.loops.append(("for", exit_lab))
        self.c_block(st.body)
        self.loops.pop()

        def nxt(f):
            v = f.v
            s = v[s_step]
            x = load(f)
            if type(x) is not int and type(x) is not float:
                x = to_num(x)
            x = x + s
            store(f, x)
            rt.budget -= 1
            if rt.budget < 0:
                rt.poll()
            e = v[s_end]
            if s >= 0:
                if x <= e:
                    return body_lab.pc
            elif x >= e:
                return body_lab.pc
        self.emit(nxt, getattr(st, "end_line", None) or st.line)
        exit_lab.pc = len(self.code)

    def s_SForEach(self, st):
        rt = self.rt
        store = self.make_store(st.var, False)
        rawstore = self.make_store(st.var, None)
        coll = self.c_expr(st.coll)
        s_it = self.new_slot()
        exit_lab = Label()
        body_lab = Label()

        def assign(f, x):
            if isinstance(x, VBAObject) or x is NOTHING:
                rawstore(f, x)
            else:
                store(f, x)

        def init(f):
            items = iter_items(coll(f))
            it = iter(items)
            f.v[s_it] = it
            for x in it:
                assign(f, x)
                return None
            return exit_lab.pc
        self.emit(init, st.line)
        body_lab.pc = len(self.code)
        self.loops.append(("for", exit_lab))
        self.c_block(st.body)
        self.loops.pop()

        def nxt(f):
            rt.budget -= 1
            if rt.budget < 0:
                rt.poll()
            for x in f.v[s_it]:
                assign(f, x)
                return body_lab.pc
            f.v[s_it] = None
        self.emit(nxt, st.line)
        exit_lab.pc = len(self.code)

    def s_SDo(self, st):
        rt = self.rt
        top = Label()
        exit_lab = Label()
        top.pc = len(self.code)
        if st.pre:
            kind, cond = st.pre
            c = self.c_expr(cond)
            if kind == "while":
                def test(f):
                    if not truthy(c(f)):
                        return exit_lab.pc
            else:
                def test(f):
                    if truthy(c(f)):
                        return exit_lab.pc
            self.emit(test, st.line)
        self.loops.append(("do", exit_lab))
        self.c_block(st.body)
        self.loops.pop()
        if st.post:
            kind, cond = st.post
            c = self.c_expr(cond)
            want = kind == "while"

            def back(f):
                rt.budget -= 1
                if rt.budget < 0:
                    rt.poll()
                if truthy(c(f)) == want:
                    return top.pc
        else:
            def back(f):
                rt.budget -= 1
                if rt.budget < 0:
                    rt.poll()
                return top.pc
        self.emit(back, st.end_line or st.line)
        exit_lab.pc = len(self.code)

    def s_SSelect(self, st):
        e = self.c_expr(st.expr)
        slot = self.new_slot()
        end = Label()

        def setv(f):
            v = e(f)
            if isinstance(v, VBAObject):
                v = to_value(v)
            f.v[slot] = v
        self.emit(setv, st.line)
        cmp = self.cmp
        for conds, body, cl in st.cases:
            tests = []
            for c in conds:
                if c[0] == "eq":
                    tests.append(("eq", self.c_expr(c[1]), None))
                elif c[0] == "range":
                    tests.append(("range", self.c_expr(c[1]), self.c_expr(c[2])))
                else:
                    tests.append((c[1], self.c_expr(c[2]), None))
            nxt = Label()

            def test(f, tests=tests, nxt=nxt):
                x = f.v[slot]
                for kind, a, b in tests:
                    if kind == "eq":
                        r = cmp["="](x, a(f))
                    elif kind == "range":
                        r = cmp[">="](x, a(f))
                        if r is True:
                            r = cmp["<="](x, b(f))
                    else:
                        r = cmp[kind](x, a(f))
                    if r is True:
                        return None
                return nxt.pc
            self.emit(test, cl)
            self.c_block(body)
            self.emit(lambda f: end.pc, cl)
            nxt.pc = len(self.code)
        if st.else_:
            self.c_block(st.else_)
        end.pc = len(self.code)

    def s_SWith(self, st):
        e = self.c_expr(st.expr)
        d = self.with_depth

        def ins(f):
            v = e(f)
            if v is NOTHING:
                raise rt_error(91)
            f.withs[d] = v
        self.emit(ins, st.line)
        self.with_depth += 1
        self.max_with = max(self.max_with, self.with_depth)
        self.c_block(st.body)
        self.with_depth -= 1

        def clear(f):
            f.withs[d] = None
        self.emit(clear, st.line)

    def s_SExit(self, st):
        k = st.kind
        rt = self.rt
        if k in ("for", "do"):
            for kind, lab in reversed(self.loops):
                if kind == k:
                    self.emit(lambda f, lab=lab: lab.pc, st.line)
                    return
            self.error(f"Exit {k.title()} not within {k.title()}", st.line)
        end = self.end_label

        def ins(f):
            if f.err_mode is not None or f.in_handler:
                rt.err.clear()
            return end.pc
        self.emit(ins, st.line)

    def s_SGoto(self, st):
        lab = self.label(st.label)
        rt = self.rt
        if st.gosub:
            idx = len(self.code)

            def ins(f):
                if f.gosub is None:
                    f.gosub = []
                f.gosub.append(idx + 1)
                return lab.pc
        else:
            def ins(f):
                rt.budget -= 1
                if rt.budget < 0:
                    rt.poll()
                return lab.pc
        self.emit(ins, st.line)

    def s_SOnGoto(self, st):
        e = self.c_expr(st.expr)
        labs = [self.label(n) for n in st.labels]
        idx = len(self.code)
        gosub = st.gosub

        def ins(f):
            i = index_int(e(f))
            if i < 0 or i > 255:
                raise rt_error(5)
            if 1 <= i <= len(labs):
                if gosub:
                    if f.gosub is None:
                        f.gosub = []
                    f.gosub.append(idx + 1)
                return labs[i - 1].pc
        self.emit(ins, st.line)

    def s_SReturn(self, st):
        def ins(f):
            if not f.gosub:
                raise rt_error(3)
            return f.gosub.pop()
        self.emit(ins, st.line)

    def s_SOnError(self, st):
        rt = self.rt
        mode = st.mode
        if mode == "next":
            def ins(f):
                f.err_mode = RESUME_NEXT
                rt.err.clear()
        elif mode == "zero":
            def ins(f):
                f.err_mode = None
                rt.err.clear()
        elif mode == "minus1":
            def ins(f):
                f.in_handler = False
                rt.err.clear()
        else:
            lab = self.label(st.label)

            def ins(f):
                f.err_mode = lab
                rt.err.clear()
        self.emit(ins, st.line)

    def s_SResume(self, st):
        rt = self.rt
        mode = st.mode
        lab = self.label(st.label) if mode == "label" else None

        def ins(f):
            if not f.in_handler:
                raise rt_error(20)
            f.in_handler = False
            rt.err.clear()
            if mode == "":
                return f.err_pc
            if mode == "next":
                return f.err_pc + 1
            return lab.pc
        self.emit(ins, st.line)

    def s_SEnd(self, st):
        def ins(f):
            raise VBAEnd()
        self.emit(ins, st.line)

    def s_SStop(self, st):
        def ins(f):
            raise VBAStop("Stop statement")
        self.emit(ins, st.line)

    def s_SAssert(self, st):
        e = self.c_expr(st.expr)

        def ins(f):
            v = e(f)
            if v is not NULL and not truthy(v):
                raise VBAStop("Debug.Assert failed")
        self.emit(ins, st.line)

    def s_SError(self, st):
        e = self.c_expr(st.expr)

        def ins(f):
            raise VBARuntimeErr(index_int(e(f)))
        self.emit(ins, st.line)

    def s_SUnsupported(self, st):
        what = st.what

        def ins(f):
            raise VBARuntimeErr(5, f"{what} is not supported in PortableExcel")
        self.emit(ins, st.line)

    def s_SMid(self, st):
        load = self.make_load(st.target)
        store = self.make_store(st.target, False)
        sf = self.c_expr(st.start)
        lf = self.c_expr(st.length) if st.length is not None else None
        vf = self.c_expr(st.expr)

        def ins(f):
            s = to_str(load(f))
            start = index_int(sf(f))
            if start < 1 or start > len(s) + 1:
                raise rt_error(5)
            rep = to_str(vf(f))
            n = len(rep) if lf is None else min(index_int(lf(f)), len(rep))
            n = min(n, len(s) - start + 1)
            store(f, s[:start - 1] + rep[:n] + s[start - 1 + n:])
        self.emit(ins, st.line)

    def s_SName(self, st):
        a = self.c_expr(st.old)
        b = self.c_expr(st.new)
        rt = self.rt

        def ins(f):
            rt.file_rename(to_str(a(f)), to_str(b(f)))
        self.emit(ins, st.line)

    def s_SErase(self, st):
        for t in st.targets:
            load = self.make_load(t)
            store = self.make_store(t, None)

            def ins(f, load=load, store=store):
                v = load(f)
                if type(v) is not VBAArray:
                    raise rt_error(13)
                if v.fixed:
                    fresh = VBAArray(list(v.dims), v.etype, fixed=True)
                    v.data = fresh.data
                else:
                    v.dims = None
                    v.data = []
            self.emit(ins, st.line)

    def s_SReDim(self, st):
        rt = self.rt
        preserve = st.preserve
        for target, dims, tref in st.items:
            dimfns = [(self.c_expr(lb) if lb is not None else None, self.c_expr(ub)) for lb, ub in dims]
            base = self.module.option_base
            load = self.make_load(target)
            rawstore = self.make_store(target, None)
            ets = rt.type_spec(tref, self.module) if tref is not None else None
            if ets is None and type(target) is A.Ident:
                vi = self.lookup_var(target.name)
                if vi is not None and vi.is_array:
                    ets = vi.ts

            def ins(f, dimfns=dimfns, load=load, rawstore=rawstore, ets=ets):
                dd = []
                for lb, ub in dimfns:
                    lo = base if lb is None else index_int(lb(f))
                    hi = index_int(ub(f))
                    if hi < lo - 1:
                        raise rt_error(9)
                    dd.append((lo, hi))
                try:
                    cur = load(f)
                except VBARuntimeErr:
                    cur = EMPTY
                if type(cur) is VBAArray:
                    if ets is not None and not preserve:
                        cur.etype = ets
                    cur.redim(dd, preserve)
                else:
                    rawstore(f, VBAArray(dd, ets))
            self.emit(ins, st.line)

    def s_SPrint(self, st):
        rt = self.rt
        items = [(self.c_expr(e) if e is not None else None, sep) for e, sep in st.items]
        if st.kind == "write":
            ff = self.c_expr(st.file)
            vals = [e for e, _ in items]

            def ins(f):
                rt.file_write(ff(f), [v(f) for v in vals])
            self.emit(ins, st.line)
            return
        ff = self.c_expr(st.file) if st.file is not None else None
        kind = st.kind

        def ins(f):
            out = []
            col = 0
            for e, sep in items:
                if e is not None:
                    s = rt.print_repr(e(f))
                    out.append(s)
                    col += len(s)
                if sep == ",":
                    pad = 14 - (col % 14)
                    out.append(" " * pad)
                    col += pad
            newline = not items or items[-1][1] == ""
            text = "".join(out)
            if kind == "debug":
                rt.host.debug_print(text)
            else:
                rt.file_print(ff(f), text, newline)
        self.emit(ins, st.line)

    def s_SInput(self, st):
        rt = self.rt
        ff = self.c_expr(st.file)
        stores = [self.make_store(t, False) for t in st.targets]
        if st.line_input:
            def ins(f):
                stores[0](f, rt.file_line_input(ff(f)))
        else:
            def ins(f):
                fn = ff(f)
                for s in stores:
                    s(f, rt.file_input_field(fn))
        self.emit(ins, st.line)

    def s_SOpen(self, st):
        rt = self.rt
        pf = self.c_expr(st.path)
        ff = self.c_expr(st.filenum)
        mode = st.mode

        def ins(f):
            rt.file_open(to_str(pf(f)), mode, ff(f))
        self.emit(ins, st.line)

    def s_SClose(self, st):
        rt = self.rt
        fs = [self.c_expr(x) for x in st.files]

        def ins(f):
            rt.file_close([x(f) for x in fs])
        self.emit(ins, st.line)

    # -- lvalues -----------------------------------------------------------------
    def var_load(self, vi):
        k = vi.kind
        i = vi.idx
        if k == "local":
            return lambda f: f.v[i]
        if k in ("mod", "static"):
            c = vi.container
            return lambda f: c[i]
        if k == "cls":
            return lambda f: f.me.vars[i]
        if k == "localconst":
            return lambda f: i
        raise AssertionError(k)

    def coercer(self, vi, mode):
        """mode: False = Let, True = Set, None = raw (keep objects)."""
        ts = vi.ts
        if mode is True:
            def c(v):
                if isinstance(v, VBAObject) or v is NOTHING:
                    return v
                if ts.kind == "variant" and not vi.is_array:
                    raise rt_error(424)
                raise rt_error(424)
            return c
        if vi.is_array:
            def c(v):
                if type(v) is VBAArray:
                    return v.copy() if mode is False else v
                if isinstance(v, VBAObject):
                    v = to_value(v)
                    if type(v) is VBAArray:
                        return v
                raise rt_error(13)
            return c
        k = ts.kind
        if k == "variant":
            if mode is None:
                return lambda v: v
            def c(v):
                t = type(v)
                if t is int or t is float or t is str or t is bool:
                    return v
                if isinstance(v, VBAObject) or v is NOTHING:
                    return to_value(v)
                if t is VBAArray or t is UDTValue:
                    return v.copy()
                return v
            return c
        if k == "object":
            def c(v):
                if isinstance(v, VBAObject) or v is NOTHING:
                    if mode is None:
                        return v
                    raise rt_error(91)
                raise rt_error(91)
            return c
        tc = ts.coerce
        if ts.name == "long":
            def c(v):
                if type(v) is int and -2147483648 <= v <= 2147483647:
                    return v
                if isinstance(v, VBAObject):
                    v = to_value(v)
                return tc(v)
            return c
        if ts.name == "double":
            def c(v):
                if type(v) is float:
                    return v
                if isinstance(v, VBAObject):
                    v = to_value(v)
                return tc(v)
            return c
        if ts.name == "string":
            def c(v):
                if type(v) is str:
                    return v
                if isinstance(v, VBAObject):
                    v = to_value(v)
                return tc(v)
            return c

        def c(v):
            if isinstance(v, VBAObject):
                v = to_value(v)
            return tc(v)
        return c

    def var_store(self, vi, mode):
        if vi.kind == "localconst":
            self.error("Assignment to constant not permitted")
        c = self.coercer(vi, mode)
        i = vi.idx
        k = vi.kind
        rt = self.rt
        if mode is True:
            if k == "local":
                def st(f, v):
                    f.v[i] = c(v)
                    if rt.pending_term:
                        rt.run_terminates()
            elif k == "cls":
                def st(f, v):
                    f.me.vars[i] = c(v)
                    if rt.pending_term:
                        rt.run_terminates()
            else:
                cont = vi.container

                def st(f, v):
                    cont[i] = c(v)
                    if rt.pending_term:
                        rt.run_terminates()
            return st
        if k == "local":
            def st(f, v):
                f.v[i] = c(v)
        elif k == "cls":
            def st(f, v):
                f.me.vars[i] = c(v)
        else:
            cont = vi.container

            def st(f, v):
                cont[i] = c(v)
        return st

    def make_load(self, node):
        return self.c_expr(node)

    def make_store(self, node, mode):
        """fn(f, value) assigning to an lvalue node. mode False=Let, True=Set, None=raw."""
        t = type(node)
        if t is A.Ident:
            b = self.resolve(node.name, line=node.line)
            if b[0] == "var":
                return self.var_store(b[1], mode)
            gen = self.c_assign_value_target(node, mode)
            return gen
        return self.c_assign_value_target(node, mode)

    def c_assign_value_target(self, node, mode):
        """Generic store for non-variable targets; returns fn(f, value)."""
        is_set = mode is True
        t = type(node)
        rt = self.rt
        if t is A.Ident:
            b = self.resolve(node.name, line=node.line)
            k = b[0]
            if k == "proc":
                proc = b[1]
                me_fn = self.me_for(proc)
                lp = proc.module.prop_set.get(proc.lname) if is_set else None
                lp = lp or proc.module.prop_let.get(proc.lname) or proc.module.prop_set.get(proc.lname)
                if lp is None:
                    if proc is self.proc:
                        self.error("Function call on left-hand side of assignment must return Variant or Object")
                    self.error("Invalid use of property")

                def st(f, v):
                    rt.invoke(lp, [v], NO_NAMED, me_fn(f))
                return st
            if k == "global":
                name = node.name
                sctx = self.sheet_ctx()
                if name in ("activecell", "selection", "range", "cells"):
                    def st(f, v):
                        o = rt.app_global(name, NO_ARGS, NO_NAMED, sctx)
                        o.vba_let("", NO_ARGS, NO_NAMED, v, is_set)
                    return st

                def st(f, v):
                    rt.app.vba_let(name, NO_ARGS, NO_NAMED, v, is_set)
                return st
            if k == "const":
                self.error("Assignment to constant not permitted", node.line)
            if k == "member":
                obj, name = b[1], b[2]

                def st(f, v):
                    obj.vba_let(name, NO_ARGS, NO_NAMED, v, is_set)
                return st
            self.error("Variable not defined" if self.explicit else "Invalid assignment target", node.line)
        if t is A.Index:
            base = node.base
            pos, named = self.c_args(node.args)
            if type(base) is A.Ident:
                b = self.resolve(base.name, implicit=False, line=node.line)
                if b is None:
                    self.error("Sub or Function not defined", node.line)
                k = b[0]
                if k == "var" and not b[1].retval:
                    load = self.var_load(b[1])

                    def st(f, v):
                        cont = load(f)
                        args = self.ev_args(pos, f)
                        if type(cont) is VBAArray:
                            cont.set(args, v)
                        elif isinstance(cont, VBAObject):
                            cont.vba_let("", args, self.ev_named(named, f), v, is_set)
                        elif cont is NOTHING:
                            raise rt_error(91)
                        else:
                            raise rt_error(13)
                    return st
                if k == "proc":
                    proc = b[1]
                    me_fn = self.me_for(proc)
                    lp = (proc.module.prop_set.get(proc.lname) if is_set else None) or \
                        proc.module.prop_let.get(proc.lname) or proc.module.prop_set.get(proc.lname)
                    if lp is None:
                        self.error("Invalid use of property", node.line)

                    def st(f, v):
                        rt.invoke(lp, self.ev_args(pos, f) + [v], NO_NAMED, me_fn(f))
                    return st
                if k == "global":
                    name = base.name
                    sctx = self.sheet_ctx()

                    def st(f, v):
                        o = rt.app_global(name, self.ev_args(pos, f), self.ev_named(named, f), sctx)
                        if isinstance(o, VBAObject):
                            o.vba_let("", NO_ARGS, NO_NAMED, v, is_set)
                        else:
                            raise rt_error(424)
                    return st
                if k == "docobj":
                    getobj = b[1]

                    def st(f, v):
                        getobj().vba_let("", self.ev_args(pos, f), self.ev_named(named, f), v, is_set)
                    return st
                self.error("Invalid assignment target", node.line)
            if type(base) is A.Member and not base.bang:
                objf = self.c_member_base(base)
                if isinstance(objf, tuple):
                    # module-qualified
                    return self.c_assign_value_target(A.Index(objf[1], node.args, node.line), mode)
                name = base.name

                def st(f, v):
                    o = objf(f)
                    args = self.ev_args(pos, f)
                    if isinstance(o, VBAObject):
                        o.vba_let(name, args, self.ev_named(named, f), v, is_set)
                    elif type(o) is UDTValue:
                        fld = o.get_field(name)
                        if type(fld) is VBAArray:
                            fld.set(args, v)
                        elif isinstance(fld, VBAObject):
                            fld.vba_let("", args, NO_NAMED, v, is_set)
                        else:
                            raise rt_error(13)
                    elif o is NOTHING:
                        raise rt_error(91)
                    else:
                        raise rt_error(424)
                return st
            contf = self.c_expr(base)

            def st(f, v):
                cont = contf(f)
                args = self.ev_args(pos, f)
                if type(cont) is VBAArray:
                    cont.set(args, v)
                elif isinstance(cont, VBAObject):
                    cont.vba_let("", args, self.ev_named(named, f), v, is_set)
                else:
                    raise rt_error(13)
            return st
        if t is A.Member:
            if node.bang:
                objf = self.c_member_base(node)
                key = node.text

                def st(f, v):
                    o = objf(f)
                    if not isinstance(o, VBAObject):
                        raise rt_error(424)
                    o.vba_let("", [key], NO_NAMED, v, is_set)
                return st
            objf = self.c_member_base(node)
            if isinstance(objf, tuple):
                return self.c_assign_value_target(objf[1], mode)
            name = node.name

            def st(f, v):
                o = objf(f)
                if isinstance(o, VBAObject):
                    o.vba_let(name, NO_ARGS, NO_NAMED, v, is_set)
                elif type(o) is UDTValue:
                    o.set_field(name, v)
                elif o is NOTHING:
                    raise rt_error(91)
                else:
                    raise rt_error(424)
            return st
        if t is A.BracketEval:
            ev = self.c_expr(node)

            def st(f, v):
                o = ev(f)
                if not isinstance(o, VBAObject):
                    raise rt_error(424)
                o.vba_let("", NO_ARGS, NO_NAMED, v, is_set)
            return st
        self.error("Invalid assignment target", getattr(node, "line", None))

    def c_assign(self, target, expr, is_set, kind=""):
        vf = self.c_expr(expr)
        if kind in ("lset", "rset"):
            load = self.make_load(target)
            store = self.make_store(target, False)
            right = kind == "rset"

            def ins(f):
                old = load(f)
                if type(old) is UDTValue:
                    store(f, vf(f))
                    return
                n = len(to_str(old))
                s = to_str(let_value(vf(f)))[:n]
                store(f, s.rjust(n) if right else s.ljust(n))
            return ins
        store = self.make_store(target, True if is_set else False)
        if is_set:
            def ins(f):
                v = vf(f)
                if not (isinstance(v, VBAObject) or v is NOTHING):
                    raise rt_error(424)
                store(f, v)
            return ins
        if type(target) is A.Ident:
            def ins(f):
                store(f, vf(f))
            return ins

        def ins(f):
            v = vf(f)
            if isinstance(v, VBAObject) or v is NOTHING:
                v = to_value(v)
            store(f, v)
        return ins

    # -- calls ------------------------------------------------------------------
    def c_args(self, args):
        pos = []
        named = []
        for a in args:
            if a.name:
                named.append((a.name, self.c_expr(a.expr)))
            elif a.expr is None:
                pos.append(None)
            else:
                pos.append(self.c_expr(a.expr))
        # f(a, ) -> keep explicit trailing missing
        return pos, named

    @staticmethod
    def ev_args(pos, f):
        return [MISSING if p is None else p(f) for p in pos]

    @staticmethod
    def ev_named(named, f):
        if not named:
            return NO_NAMED
        return {k: fn(f) for k, fn in named}

    def make_args_fn(self, arg_nodes):
        pos, named = self.c_args(arg_nodes)
        if named:
            return lambda f: ([MISSING if p is None else p(f) for p in pos],
                              {k: fn(f) for k, fn in named})
        if not pos:
            return lambda f: ([], NO_NAMED)
        if None in pos:
            return lambda f: ([MISSING if p is None else p(f) for p in pos], NO_NAMED)
        if len(pos) == 1:
            p0 = pos[0]
            return lambda f: ([p0(f)], NO_NAMED)
        if len(pos) == 2:
            p0, p1 = pos
            return lambda f: ([p0(f), p1(f)], NO_NAMED)
        return lambda f: ([p(f) for p in pos], NO_NAMED)

    def me_for(self, proc):
        m = proc.module
        rt = self.rt
        if m.type == "class":
            if m is self.module:
                return lambda f: f.me
            return lambda f: None
        if m.type == "document":
            return lambda f: rt.doc_object_for_module(m)
        return lambda f: None

    def c_proccall(self, proc, arg_nodes):
        rt = self.rt
        argsfn = self.make_args_fn(arg_nodes)
        me_fn = self.me_for(proc)
        params = proc.params
        wbs = []
        i = 0
        for a in arg_nodes:
            if a.name:
                p = next((p for p in params if p.name == a.name), None)
            else:
                p = params[i] if i < len(params) else None
                i += 1
            if p is None or p.paramarray or not p.byref or a.expr is None or a.byval:
                continue
            e = a.expr
            if type(e) is A.Ident:
                b = self.resolve(e.name, implicit=False)
                if b is not None and b[0] == "var":
                    wbs.append((p, self.var_store(b[1], None)))
            elif type(e) is A.Index and type(e.base) is A.Ident:
                b = self.resolve(e.base.name, implicit=False)
                if b is not None and b[0] == "var" and not b[1].retval:
                    load = self.var_load(b[1])
                    pos, _ = self.c_args(e.args)

                    def st(f, v, load=load, pos=pos):
                        cont = load(f)
                        if type(cont) is VBAArray:
                            cont.set(self.ev_args(pos, f), v)
                    wbs.append((p, st))
        if wbs:
            def call(f):
                args, named = argsfn(f)
                res, cf = rt.invoke(proc, args, named, me_fn(f))
                for p, st in wbs:
                    st(f, cf.v[p.slot])
                return res
        else:
            def call(f):
                args, named = argsfn(f)
                return rt.invoke(proc, args, named, me_fn(f))[0]
        return call

    def sheet_ctx(self):
        m = self.module
        if m.type == "document" and m.lname != "thisworkbook":
            return m
        return None

    def c_callstmt(self, st):
        e = st.expr
        args = st.args
        if args is None:
            if type(e) is A.Index:
                callee, arg_nodes = e.base, e.args
            else:
                callee, arg_nodes = e, []
        else:
            callee, arg_nodes = e, args
        fn = self.c_call(callee, arg_nodes, stmt=True)
        return fn

    def c_call(self, callee, arg_nodes, stmt=False):
        """Compile `callee(args)` (call or index)."""
        rt = self.rt
        t = type(callee)
        if t is A.Ident:
            name = callee.name
            special = self.c_special_builtin(name, arg_nodes)
            if special is not None:
                return special
            b = self.resolve(name, implicit=False, line=callee.line)
            if b is None:
                if not arg_nodes and not stmt:
                    b = self.resolve(name, line=callee.line)
                else:
                    self.error("Sub or Function not defined", callee.line)
            k = b[0]
            if k == "var":
                vi = b[1]
                if vi.retval and arg_nodes:
                    return self.c_proccall(self.proc, arg_nodes)
                load = self.var_load(vi)
                if not arg_nodes:
                    if stmt:
                        def call(f):
                            v = load(f)
                            if isinstance(v, VBAObject):
                                return v.vba_get("", NO_ARGS, NO_NAMED)
                            return v
                        return call
                    return load
                argsfn = self.make_args_fn(arg_nodes)
                if len(arg_nodes) == 1 and not arg_nodes[0].name and arg_nodes[0].expr is not None:
                    a0 = self.c_expr(arg_nodes[0].expr)

                    def call(f):
                        v = load(f)
                        if type(v) is VBAArray:
                            return v.get((a0(f),))
                        return index_value(v, [a0(f)])
                    return call

                def call(f):
                    v = load(f)
                    args, named = argsfn(f)
                    if type(v) is VBAArray:
                        return v.get(args)
                    if isinstance(v, VBAObject):
                        return v.vba_get("", args, named)
                    return index_value(v, args)
                return call
            if k == "proc":
                return self.c_proccall(b[1], arg_nodes)
            if k == "builtin":
                bi = b[1]
                argsfn = self.make_args_fn(arg_nodes)

                def call(f):
                    args, named = argsfn(f)
                    return bi.call(rt, args, named)
                return call
            if k == "global":
                sctx = self.sheet_ctx()
                argsfn = self.make_args_fn(arg_nodes)

                def call(f):
                    args, named = argsfn(f)
                    return rt.app_global(name, args, named, sctx)
                return call
            if k == "const":
                v = b[1]
                if arg_nodes:
                    argsfn = self.make_args_fn(arg_nodes)
                    return lambda f: index_value(v, argsfn(f)[0])
                return lambda f: v
            if k == "docobj":
                getobj = b[1]
                if not arg_nodes:
                    return lambda f: getobj()
                argsfn = self.make_args_fn(arg_nodes)

                def call(f):
                    args, named = argsfn(f)
                    return getobj().vba_get("", args, named)
                return call
            if k == "member":
                obj, mname = b[1], b[2]
                argsfn = self.make_args_fn(arg_nodes)

                def call(f):
                    args, named = argsfn(f)
                    return obj.vba_get(mname, args, named)
                return call
            if k == "declare":
                dname = b[1]

                def call(f):
                    e = VBARuntimeErr(5, f"Windows API calls are not supported ({dname})")
                    e.fatal = True
                    raise e
                return call
            if k == "form":
                fname = b[1]
                return lambda f: UnsupportedFormObj(fname)
            if k in ("module", "enum", "class", "type"):
                self.error("Expected variable or procedure, not module", callee.line)
            self.error("Sub or Function not defined", callee.line)
        if t is A.Member:
            return self.c_member(callee, arg_nodes)
        base = self.c_expr(callee)
        if not arg_nodes:
            if stmt:
                def call(f):
                    v = base(f)
                    if isinstance(v, VBAObject):
                        return v.vba_get("", NO_ARGS, NO_NAMED)
                    return v
                return call
            return base
        argsfn = self.make_args_fn(arg_nodes)

        def call(f):
            v = base(f)
            args, named = argsfn(f)
            if type(v) is VBAArray:
                return v.get(args)
            if isinstance(v, VBAObject):
                return v.vba_get("", args, named)
            return index_value(v, args)
        return call

    def c_special_builtin(self, name, arg_nodes):
        if name in self.locals:
            return None
        r = self.module.resolve_member(name, self.rt, inside=True)
        if r is not None:
            return None
        if self.rt.resolve_global_user(name) is not None:
            return None
        if name == "array":
            lb = self.module.option_base
            fns = [self.c_expr(a.expr) if a.expr is not None else (lambda f: MISSING) for a in arg_nodes]
            return lambda f: VBAArray.from_list([fn(f) for fn in fns], lb)
        if name in ("typename", "vartype") and len(arg_nodes) == 1 and type(arg_nodes[0].expr) is A.Ident:
            vi = self.lookup_var(arg_nodes[0].expr.name)
            if vi is not None and (vi.ts.kind in ("scalar", "fixedstr") or vi.is_array):
                ts = vi.ts
                load = self.var_load(vi)
                if name == "typename":
                    disp = ts.display + ("()" if vi.is_array else "")
                    if vi.is_array:
                        return lambda f: disp
                    return lambda f: disp
                from .values import VARTYPE_BY_NAME
                code = VARTYPE_BY_NAME.get(ts.display, 12) + (8192 if vi.is_array else 0)
                return lambda f: code
        if name == "ismissing" and len(arg_nodes) == 1:
            return None
        return None

    # -- expressions -------------------------------------------------------------
    def c_expr(self, node):
        t = type(node)
        if t is A.Lit:
            v = node.value
            return lambda f: v
        if t is A.Ident:
            return self.c_ident(node)
        if t is A.Index:
            return self.c_call(node.base, node.args)
        if t is A.Member:
            return self.c_member(node, None)
        if t is A.BinOp:
            return self.c_binop(node)
        if t is A.UnOp:
            a = self.c_expr(node.a)
            if node.op == "-":
                if type(node.a) is A.Lit and type(node.a.value) in (int, float):
                    v = -node.a.value
                    return lambda f: v
                return lambda f: op_neg(a(f))
            return lambda f: op_not(a(f))
        if t is A.NewObj:
            rt = self.rt
            cls = node.cls
            mod = self.module
            rt.check_class_name(cls, mod, node.line)
            return lambda f: rt.new_object(cls, mod)
        if t is A.TypeOfIs:
            e = self.c_expr(node.expr)
            tn = node.tname.split(".")[-1]
            rt = self.rt
            return lambda f: rt.typeof_is(e(f), tn)
        if t is A.BracketEval:
            rt = self.rt
            text = node.text
            if node.base is not None:
                if type(node.base) is A.Member and node.base.base is None and node.base.name == "":
                    d = self.with_depth - 1
                    if d < 0:
                        self.error("Invalid or unqualified reference", node.line)
                    return lambda f: member_get(f.withs[d], "evaluate", [text], NO_NAMED)
                b = self.c_expr(node.base)
                return lambda f: member_get(b(f), "evaluate", [text], NO_NAMED)
            sctx = self.sheet_ctx()
            return lambda f: rt.app_global("evaluate", [text], NO_NAMED, sctx)
        if t is A.MeRef:
            if self.module.type == "standard":
                self.error("Invalid use of Me keyword", node.line)
            if self.module.type == "document":
                rt = self.rt
                m = self.module
                return lambda f: rt.doc_object_for_module(m)
            return lambda f: f.me
        if t is A.AddressOf:
            def fail(f):
                raise VBARuntimeErr(5, "AddressOf is not supported")
            return fail
        self.error("Syntax error in expression", getattr(node, "line", None))

    def c_ident(self, node):
        return self.c_call(node, [], stmt=False)

    def c_binop(self, node):
        op = node.op
        a = self.c_expr(node.a)
        b = self.c_expr(node.b)
        if op in self.cmp:
            fn = self.cmp[op]
            if op == "=":
                text = self.cmp is CMP_TEXT

                def eq(f):
                    x = a(f)
                    y = b(f)
                    tx = type(x)
                    if tx is type(y) and (tx is int or (tx is str and not text)):
                        return x == y
                    return fn(x, y)
                return eq
            if op in ("<", ">", "<=", ">="):
                return lambda f: fn(a(f), b(f))
            return lambda f: fn(a(f), b(f))
        if op == "like":
            text = self.cmp is CMP_TEXT
            return lambda f: op_like(a(f), b(f), text)
        fn = ARITH[op]
        if op == "+":
            def add(f):
                x = a(f)
                y = b(f)
                if type(x) is int and type(y) is int:
                    return x + y
                return op_add(x, y)
            return add
        if op == "-":
            def sub(f):
                x = a(f)
                y = b(f)
                if type(x) is int and type(y) is int:
                    return x - y
                return op_sub(x, y)
            return sub
        return lambda f: fn(a(f), b(f))

    def c_member_base(self, node):
        """Compile the object part of a Member node.

        Returns fn(f)->object, or ("redirect", node) for module-qualified names."""
        base = node.base
        if base is None:
            d = self.with_depth - 1
            if d < 0:
                self.error("Invalid or unqualified reference", node.line)
            return lambda f: f.withs[d]
        if type(base) is A.Ident and base.name not in self.locals:
            r = self.module.resolve_member(base.name, self.rt, inside=True)
            if r is None:
                r = self.rt.resolve_global(base.name)
            if r is not None and r[0] == "module":
                mod = r[1]
                return ("redirect", A.Ident(node.name, node.text, "", node.line), mod)
        return self.c_expr(base)

    def c_member(self, node, arg_nodes):
        rt = self.rt
        base = node.base
        name = node.name
        if type(base) is A.Ident and base.name not in self.locals:
            r = self.module.resolve_member(base.name, rt, inside=True)
            if r is None:
                r = rt.resolve_global(base.name)
            if r is not None:
                k = r[0]
                if k == "module":
                    mod = r[1]
                    rr = mod.resolve_member(name, rt, inside=(mod is self.module))
                    if rr is None:
                        self.error(f"Method or data member not found: {node.text}", node.line)
                    return self.c_resolved(rr, name, arg_nodes or [], node.line)
                if k == "enum":
                    vals = r[1]
                    if name not in vals:
                        self.error(f"Method or data member not found: {node.text}", node.line)
                    v = vals[name]
                    return lambda f: v
                if k == "vbalib":
                    if name in BUILTINS:
                        return self.c_resolved(("builtin", BUILTINS[name]), name, arg_nodes or [], node.line)
                    if name in CONSTANTS:
                        v = CONSTANTS[name]
                        return lambda f: v
                    self.error(f"Method or data member not found: {node.text}", node.line)
        if node.bang:
            objf = self.c_member_base(node)
            key = node.text
            if arg_nodes:
                argsfn = self.make_args_fn(arg_nodes)
                return lambda f: index_value(member_get(objf(f), "", [key], NO_NAMED), argsfn(f)[0])
            return lambda f: member_get(objf(f), "", [key], NO_NAMED)
        objf = self.c_member_base(node)
        if isinstance(objf, tuple):
            return self.c_call(objf[1], arg_nodes or [])
        if not arg_nodes:
            def get(f):
                o = objf(f)
                if isinstance(o, VBAObject):
                    return o.vba_get(name, NO_ARGS, NO_NAMED)
                return member_get(o, name, NO_ARGS, NO_NAMED)
            return get
        argsfn = self.make_args_fn(arg_nodes)

        def get(f):
            o = objf(f)
            args, named = argsfn(f)
            if isinstance(o, VBAObject):
                return o.vba_get(name, args, named)
            return member_get(o, name, args, named)
        return get

    def c_resolved(self, r, name, arg_nodes, line):
        k = r[0]
        if k == "var":
            load = self.var_load(r[1])
            if arg_nodes:
                argsfn = self.make_args_fn(arg_nodes)
                return lambda f: index_value(load(f), argsfn(f)[0])
            return load
        if k == "const":
            v = r[1]
            return lambda f: v
        if k == "proc":
            return self.c_proccall(r[1], arg_nodes)
        if k == "builtin":
            bi = r[1]
            rt = self.rt
            argsfn = self.make_args_fn(arg_nodes)

            def call(f):
                args, named = argsfn(f)
                return bi.call(rt, args, named)
            return call
        if k == "declare":
            dname = r[1]

            def call(f):
                e = VBARuntimeErr(5, f"Windows API calls are not supported ({dname})")
                e.fatal = True
                raise e
            return call
        self.error(f"Method or data member not found: {name}", line)


def compile_module_expr(rt, module, expr):
    """Evaluate a constant expression in module context (Const/Enum values)."""
    c = Compiler(rt, module, None)
    return c.const_eval(expr)

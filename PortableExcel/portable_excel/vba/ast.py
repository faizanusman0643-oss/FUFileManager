"""AST node classes for the VBA parser."""
from __future__ import annotations


class Node:
    __slots__ = ("line",)


# ---- expressions -----------------------------------------------------------
class Lit(Node):
    __slots__ = ("value",)

    def __init__(self, value, line=0):
        self.value = value
        self.line = line


class Ident(Node):
    __slots__ = ("name", "text", "suffix")

    def __init__(self, name, text, suffix="", line=0):
        self.name = name      # lower case
        self.text = text
        self.suffix = suffix
        self.line = line


class Member(Node):
    __slots__ = ("base", "name", "text", "bang")

    def __init__(self, base, name, text, bang=False, line=0):
        self.base = base      # None -> With object
        self.name = name
        self.text = text
        self.bang = bang
        self.line = line


class Arg(Node):
    __slots__ = ("expr", "name", "byval")

    def __init__(self, expr, name=None, byval=False, line=0):
        self.expr = expr      # None = omitted
        self.name = name      # named argument (lower)
        self.byval = byval
        self.line = line


class Index(Node):
    __slots__ = ("base", "args")

    def __init__(self, base, args, line=0):
        self.base = base
        self.args = args      # list[Arg]
        self.line = line


class BinOp(Node):
    __slots__ = ("op", "a", "b")

    def __init__(self, op, a, b, line=0):
        self.op = op
        self.a = a
        self.b = b
        self.line = line


class UnOp(Node):
    __slots__ = ("op", "a")

    def __init__(self, op, a, line=0):
        self.op = op
        self.a = a
        self.line = line


class NewObj(Node):
    __slots__ = ("cls",)

    def __init__(self, cls, line=0):
        self.cls = cls        # lower-case possibly dotted name
        self.line = line


class TypeOfIs(Node):
    __slots__ = ("expr", "tname")

    def __init__(self, expr, tname, line=0):
        self.expr = expr
        self.tname = tname
        self.line = line


class BracketEval(Node):
    __slots__ = ("text", "base")

    def __init__(self, text, base=None, line=0):
        self.text = text
        self.base = base
        self.line = line


class MeRef(Node):
    __slots__ = ()

    def __init__(self, line=0):
        self.line = line


class AddressOf(Node):
    __slots__ = ("name",)

    def __init__(self, name, line=0):
        self.name = name
        self.line = line


# ---- declarations -----------------------------------------------------------
class TypeRef:
    __slots__ = ("name", "size", "new", "text")

    def __init__(self, name="variant", size=None, new=False, text="Variant"):
        self.name = name      # lower, may be dotted ("scripting.dictionary")
        self.size = size      # expr for String * n
        self.new = new
        self.text = text


class VarDecl:
    __slots__ = ("name", "text", "type", "dims", "is_array", "init", "line", "withevents")

    def __init__(self, name, text, type_, dims=None, is_array=False, init=None, line=0):
        self.name = name
        self.text = text
        self.type = type_         # TypeRef
        self.dims = dims          # list of (lb expr|None, ub expr) ; [] dynamic
        self.is_array = is_array
        self.init = init          # Const value expr
        self.line = line
        self.withevents = False


class Param:
    __slots__ = ("name", "text", "byref", "optional", "default", "paramarray", "type", "is_array")

    def __init__(self, name, text, byref=True, optional=False, default=None, paramarray=False,
                 type_=None, is_array=False):
        self.name = name
        self.text = text
        self.byref = byref
        self.optional = optional
        self.default = default
        self.paramarray = paramarray
        self.type = type_ or TypeRef()
        self.is_array = is_array


# ---- statements ----------------------------------------------------------------
class Stmt(Node):
    __slots__ = ()


class SAssign(Stmt):
    __slots__ = ("target", "expr", "is_set", "kind")

    def __init__(self, target, expr, is_set=False, line=0, kind=""):
        self.target = target
        self.expr = expr
        self.is_set = is_set
        self.kind = kind          # "" | "lset" | "rset"
        self.line = line


class SCall(Stmt):
    __slots__ = ("expr", "args", "explicit")

    def __init__(self, expr, args, line=0, explicit=False):
        self.expr = expr
        self.args = args          # list[Arg] or None (use Index args of expr)
        self.explicit = explicit
        self.line = line


class SIf(Stmt):
    __slots__ = ("branches", "else_")

    def __init__(self, branches, else_, line=0):
        self.branches = branches  # list of (cond, body)
        self.else_ = else_
        self.line = line


class SFor(Stmt):
    __slots__ = ("var", "start", "end", "step", "body")

    def __init__(self, var, start, end, step, body, line=0):
        self.var = var
        self.start = start
        self.end = end
        self.step = step
        self.body = body
        self.line = line


class SForEach(Stmt):
    __slots__ = ("var", "coll", "body")

    def __init__(self, var, coll, body, line=0):
        self.var = var
        self.coll = coll
        self.body = body
        self.line = line


class SDo(Stmt):
    __slots__ = ("pre", "post", "body", "end_line")

    def __init__(self, pre, post, body, line=0, end_line=0):
        self.pre = pre            # None or (kind, cond)
        self.post = post
        self.body = body
        self.line = line
        self.end_line = end_line


class SSelect(Stmt):
    __slots__ = ("expr", "cases", "else_")

    def __init__(self, expr, cases, else_, line=0):
        self.expr = expr
        self.cases = cases        # list of (conds, body, line)
        self.else_ = else_
        self.line = line


class SWith(Stmt):
    __slots__ = ("expr", "body")

    def __init__(self, expr, body, line=0):
        self.expr = expr
        self.body = body
        self.line = line


class SExit(Stmt):
    __slots__ = ("kind",)

    def __init__(self, kind, line=0):
        self.kind = kind
        self.line = line


class SGoto(Stmt):
    __slots__ = ("label", "gosub")

    def __init__(self, label, gosub=False, line=0):
        self.label = label
        self.gosub = gosub
        self.line = line


class SOnGoto(Stmt):
    __slots__ = ("expr", "labels", "gosub")

    def __init__(self, expr, labels, gosub, line=0):
        self.expr = expr
        self.labels = labels
        self.gosub = gosub
        self.line = line


class SReturn(Stmt):
    __slots__ = ()

    def __init__(self, line=0):
        self.line = line


class SOnError(Stmt):
    __slots__ = ("mode", "label")

    def __init__(self, mode, label=None, line=0):
        self.mode = mode          # next | goto | zero | minus1
        self.label = label
        self.line = line


class SResume(Stmt):
    __slots__ = ("mode", "label")

    def __init__(self, mode, label=None, line=0):
        self.mode = mode          # "" | next | label
        self.label = label
        self.line = line


class SLabel(Stmt):
    __slots__ = ("name", "number")

    def __init__(self, name, number=None, line=0):
        self.name = name
        self.number = number
        self.line = line


class SDim(Stmt):
    __slots__ = ("decls", "kind")

    def __init__(self, decls, kind, line=0):
        self.decls = decls
        self.kind = kind          # dim | static | const
        self.line = line


class SReDim(Stmt):
    __slots__ = ("items", "preserve")

    def __init__(self, items, preserve, line=0):
        self.items = items        # list of (target expr, dims [(lb, ub)], TypeRef|None)
        self.preserve = preserve
        self.line = line


class SErase(Stmt):
    __slots__ = ("targets",)

    def __init__(self, targets, line=0):
        self.targets = targets
        self.line = line


class SPrint(Stmt):
    __slots__ = ("file", "items", "kind")

    def __init__(self, kind, file, items, line=0):
        self.kind = kind          # debug | file | write
        self.file = file
        self.items = items        # list of (expr|None, sep) sep in ("", ";", ",")
        self.line = line


class SInput(Stmt):
    __slots__ = ("file", "targets", "line_input")

    def __init__(self, file, targets, line_input=False, line=0):
        self.file = file
        self.targets = targets
        self.line_input = line_input
        self.line = line


class SOpen(Stmt):
    __slots__ = ("path", "mode", "filenum", "reclen")

    def __init__(self, path, mode, filenum, reclen=None, line=0):
        self.path = path
        self.mode = mode
        self.filenum = filenum
        self.reclen = reclen
        self.line = line


class SClose(Stmt):
    __slots__ = ("files",)

    def __init__(self, files, line=0):
        self.files = files
        self.line = line


class SEnd(Stmt):
    __slots__ = ()

    def __init__(self, line=0):
        self.line = line


class SStop(Stmt):
    __slots__ = ()

    def __init__(self, line=0):
        self.line = line


class SAssert(Stmt):
    __slots__ = ("expr",)

    def __init__(self, expr, line=0):
        self.expr = expr
        self.line = line


class SMid(Stmt):
    __slots__ = ("target", "start", "length", "expr")

    def __init__(self, target, start, length, expr, line=0):
        self.target = target
        self.start = start
        self.length = length
        self.expr = expr
        self.line = line


class SError(Stmt):
    __slots__ = ("expr",)

    def __init__(self, expr, line=0):
        self.expr = expr
        self.line = line


class SName(Stmt):
    __slots__ = ("old", "new")

    def __init__(self, old, new, line=0):
        self.old = old
        self.new = new
        self.line = line


class SUnsupported(Stmt):
    __slots__ = ("what",)

    def __init__(self, what, line=0):
        self.what = what
        self.line = line


class SNop(Stmt):
    __slots__ = ()

    def __init__(self, line=0):
        self.line = line


# ---- module level ---------------------------------------------------------------
class ProcDef:
    def __init__(self, name, text, kind, params, ret, body, private, static, line, module):
        self.name = name          # lower
        self.text = text
        self.kind = kind          # sub | function | get | let | set
        self.params = params
        self.ret = ret            # TypeRef
        self.body = body
        self.private = private
        self.static = static
        self.line = line
        self.end_line = line
        self.module = module
        self.error = None         # parse error (VBACompileErr) if body failed to parse


class TypeDef:
    def __init__(self, name, text, fields, private, line):
        self.name = name
        self.text = text
        self.fields = fields      # list of VarDecl
        self.private = private
        self.line = line


class EnumDef:
    def __init__(self, name, text, members, private, line):
        self.name = name
        self.text = text
        self.members = members    # list of (lname, text, expr|None)
        self.private = private
        self.line = line


class ModuleDef:
    def __init__(self, name, mtype):
        self.name = name
        self.type = mtype
        self.option_explicit = False
        self.option_base = 0
        self.option_compare_text = False
        self.option_private = False
        self.vars = []            # list of (VarDecl, private)
        self.consts = []          # list of (VarDecl, private)
        self.types = []
        self.enums = []
        self.declares = {}        # lower name -> text
        self.procs = []           # list of ProcDef
        self.errors = []

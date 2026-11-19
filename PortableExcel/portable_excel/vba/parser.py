"""Recursive-descent parser: tokens -> AST (portable_excel.vba.ast)."""
from __future__ import annotations

from .errors import VBACompileErr
from .lexer import Tok, tokenize, TYPE_SUFFIX
from .ast import *  # noqa: F401,F403
from . import ast as A

RESERVED = {
    "and", "or", "not", "xor", "eqv", "imp", "mod", "is", "like", "to", "step", "then",
    "else", "elseif", "end", "if", "for", "next", "do", "loop", "while", "wend", "until",
    "select", "case", "with", "sub", "function", "property", "dim", "redim", "as", "new",
    "set", "let", "call", "exit", "goto", "gosub", "return", "on", "resume", "true", "false",
    "nothing", "empty", "null", "me", "byval", "byref", "optional", "paramarray", "private",
    "public", "global", "static", "const", "type", "enum", "each", "in", "typeof", "option",
    "declare", "friend", "stop", "addressof", "preserve", "erase", "withevents",
}

STMT_END_IDS = {"else"}


class Parser:
    def __init__(self, toks, module_name="", in_immediate=False):
        self.t = toks
        self.i = 0
        self.mod = module_name
        self.single_if_depth = 0

    # -- token helpers ------------------------------------------------------
    @property
    def tok(self) -> Tok:
        return self.t[self.i]

    def peek(self, k=1) -> Tok:
        j = self.i + k
        return self.t[j] if j < len(self.t) else self.t[-1]

    def next(self) -> Tok:
        t = self.t[self.i]
        if t.kind != "eof":
            self.i += 1
        return t

    def error(self, msg, tok=None):
        tok = tok or self.tok
        raise VBACompileErr(msg, tok.line, self.mod)

    def is_id(self, *names, tok=None):
        t = tok or self.tok
        return t.kind == "id" and (not names or t.val in names)

    def is_op(self, *ops, tok=None):
        t = tok or self.tok
        return t.kind == "op" and t.val in ops

    def accept_id(self, *names):
        if self.is_id(*names):
            return self.next()
        return None

    def accept_op(self, *ops):
        if self.is_op(*ops):
            return self.next()
        return None

    def expect_id(self, *names):
        if not self.is_id(*names):
            self.error(f"Syntax error: expected {'/'.join(names) if names else 'identifier'}")
        return self.next()

    def expect_op(self, op):
        if not self.is_op(op):
            self.error(f"Syntax error: expected '{op}'")
        return self.next()

    def at_stmt_end(self):
        t = self.tok
        return t.kind in ("nl", "eof") or (t.kind == "op" and t.val == ":") or \
            (t.kind == "id" and t.val in STMT_END_IDS)

    def skip_to_eol(self):
        while self.tok.kind not in ("nl", "eof"):
            self.next()

    def skip_seps(self):
        while self.tok.kind == "nl" or self.is_op(":"):
            self.next()

    def name_tok(self):
        """Identifier, also accepting reserved words (for member names)."""
        t = self.tok
        if t.kind != "id":
            self.error("Syntax error: expected identifier")
        return self.next()

    # -- types --------------------------------------------------------------
    def parse_type(self):
        """After 'As'."""
        new = bool(self.accept_id("new"))
        t = self.name_tok()
        name = t.val
        text = t.text
        while self.is_op("."):
            self.next()
            t2 = self.name_tok()
            name += "." + t2.val
            text += "." + t2.text
        size = None
        if self.is_op("*"):
            self.next()
            size = self.parse_unary_operand()
        return TypeRef(name, size, new, text)

    def suffix_type(self, suffix):
        if suffix and suffix in TYPE_SUFFIX:
            n = TYPE_SUFFIX[suffix]
            return TypeRef(n, None, False, n.title())
        return None

    # -- expressions --------------------------------------------------------
    def parse_expr(self):
        return self.p_imp()

    def _binloop(self, sub, ops_id=(), ops_op=()):
        left = sub()
        while True:
            t = self.tok
            if t.kind == "id" and t.val in ops_id:
                op = t.val
            elif t.kind == "op" and t.val in ops_op:
                op = t.val
            else:
                return left
            self.next()
            right = sub()
            left = BinOp(op, left, right, t.line)

    def p_imp(self):
        return self._binloop(self.p_eqv, ("imp",))

    def p_eqv(self):
        return self._binloop(self.p_xor, ("eqv",))

    def p_xor(self):
        return self._binloop(self.p_or, ("xor",))

    def p_or(self):
        return self._binloop(self.p_and, ("or",))

    def p_and(self):
        return self._binloop(self.p_not, ("and",))

    def p_not(self):
        if self.is_id("not"):
            t = self.next()
            return UnOp("not", self.p_not(), t.line)
        return self.p_cmp()

    def p_cmp(self):
        return self._binloop(self.p_concat, ("is", "like"), ("=", "<>", "<", ">", "<=", ">="))

    def p_concat(self):
        return self._binloop(self.p_add, (), ("&",))

    def p_add(self):
        return self._binloop(self.p_mod, (), ("+", "-"))

    def p_mod(self):
        return self._binloop(self.p_idiv, ("mod",))

    def p_idiv(self):
        return self._binloop(self.p_mul, (), ("\\",))

    def p_mul(self):
        return self._binloop(self.p_neg, (), ("*", "/"))

    def p_neg(self):
        if self.is_op("-"):
            t = self.next()
            return UnOp("-", self.p_neg(), t.line)
        if self.is_op("+"):
            self.next()
            return self.p_neg()
        return self.p_pow()

    def p_pow(self):
        left = self.parse_postfix()
        while self.is_op("^"):
            t = self.next()
            right = self.parse_unary_operand()
            left = BinOp("^", left, right, t.line)
        return left

    def parse_unary_operand(self):
        if self.is_op("-"):
            t = self.next()
            return UnOp("-", self.parse_unary_operand(), t.line)
        if self.is_op("+"):
            self.next()
            return self.parse_unary_operand()
        return self.parse_postfix()

    def parse_primary(self):
        t = self.tok
        k = t.kind
        if k in ("num", "str", "date"):
            self.next()
            v = t.val
            if k == "num" and t.suffix in ("!", "#") and isinstance(v, int):
                v = float(v)
            return Lit(v, t.line)
        if k == "brk":
            self.next()
            return BracketEval(t.val, None, t.line)
        if k == "op":
            if t.val == "(":
                self.next()
                e = self.parse_expr()
                self.expect_op(")")
                return e
            if t.val in (".", "!"):
                self.next()
                if self.tok.kind == "brk":
                    b = self.next()
                    return BracketEval(b.val, Member(None, "", "", False, t.line), t.line)
                n = self.name_tok()
                return Member(None, n.val, n.text, t.val == "!", t.line)
            if t.val == "-":
                self.next()
                return UnOp("-", self.parse_unary_operand(), t.line)
            self.error(f"Syntax error: unexpected '{t.text}'")
        if k == "id":
            v = t.val
            if v == "true":
                self.next()
                return Lit(True, t.line)
            if v == "false":
                self.next()
                return Lit(False, t.line)
            if v in ("nothing", "empty", "null"):
                self.next()
                from .values import NOTHING, EMPTY, NULL
                return Lit({"nothing": NOTHING, "empty": EMPTY, "null": NULL}[v], t.line)
            if v == "me":
                self.next()
                return MeRef(t.line)
            if v == "new":
                self.next()
                n = self.name_tok()
                name = n.val
                while self.is_op("."):
                    self.next()
                    name += "." + self.name_tok().val
                return NewObj(name, t.line)
            if v == "typeof":
                self.next()
                e = self.parse_postfix()
                self.expect_id("is")
                n = self.name_tok()
                name = n.val
                while self.is_op("."):
                    self.next()
                    name += "." + self.name_tok().val
                return TypeOfIs(e, name, t.line)
            if v == "not":
                self.next()
                return UnOp("not", self.p_not(), t.line)
            if v == "addressof":
                self.next()
                n = self.name_tok()
                return AddressOf(n.val, t.line)
            if v in RESERVED and v not in ("me",):
                self.error(f"Syntax error: unexpected '{t.text}'")
            self.next()
            return Ident(v, t.text, t.suffix, t.line)
        if k == "nl" or k == "eof":
            self.error("Syntax error: expected expression")
        self.error(f"Syntax error: unexpected '{t.text}'")

    def parse_postfix(self, stmt_level=False):
        e = self.parse_primary()
        return self.parse_postfix_tail(e, stmt_level)

    def parse_postfix_tail(self, e, stmt_level=False):
        while True:
            t = self.tok
            if t.kind == "op":
                if t.val == "(":
                    if stmt_level and t.space:
                        return e
                    self.next()
                    args = self.parse_args(")")
                    self.expect_op(")")
                    e = Index(e, args, t.line)
                    continue
                if t.val == "." and not t.space:
                    nt = self.peek()
                    if nt.kind == "brk":
                        self.next()
                        b = self.next()
                        e = BracketEval(b.val, e, t.line)
                        continue
                    if nt.kind != "id":
                        return e
                    self.next()
                    n = self.next()
                    e = Member(e, n.val, n.text, False, t.line)
                    continue
                if t.val == "!" and self.peek().kind == "id":
                    self.next()
                    n = self.next()
                    e = Member(e, n.val, n.text, True, t.line)
                    continue
            return e

    def parse_args(self, closer):
        """Comma separated argument list up to (not including) closer.

        closer: ")" or None (statement end)."""
        args = []
        if closer == ")" and self.is_op(")"):
            return args
        if closer is None and self.at_stmt_end():
            return args
        while True:
            t = self.tok
            if (closer == ")" and self.is_op(")")) or (closer is None and self.at_stmt_end()):
                args.append(Arg(None, None, False, t.line))
                break
            if self.is_op(","):
                args.append(Arg(None, None, False, t.line))
                self.next()
                continue
            name = None
            byval = False
            if t.kind == "id" and self.is_op(":=", tok=self.peek()):
                name = t.val
                self.next()
                self.next()
            elif t.kind == "id" and t.val == "byval" and self.peek().kind == "id":
                self.next()
                byval = True
            expr = self.parse_expr()
            args.append(Arg(expr, name, byval, t.line))
            if self.is_op(","):
                self.next()
                continue
            break
        # drop a trailing omitted arg produced by "f(a, )" only if it is the only one
        return args

    # -- statements -----------------------------------------------------------
    def parse_block(self, enders):
        """Parse statements until one of the `enders` predicates matches."""
        body = []
        while True:
            self.skip_seps()
            if self.tok.kind == "eof":
                return body
            if self._at_ender(enders):
                return body
            st = self.parse_statement(True)
            if st is None:
                continue
            if isinstance(st, list):
                body.extend(st)
            else:
                body.append(st)

    def _at_ender(self, enders):
        t = self.tok
        if t.kind != "id":
            return False
        for e in enders:
            if isinstance(e, tuple):
                if t.val == e[0] and self.peek().kind == "id" and self.peek().val == e[1]:
                    return True
                if len(e) == 2 and t.val == e[0] + e[1]:
                    return True
            elif t.val == e:
                return True
        return False

    def expect_end(self, what):
        if self.is_id("end") and self.is_id(what, tok=self.peek()):
            self.next()
            self.next()
            return
        if self.is_id("end" + what):
            self.next()
            return
        self.error(f"Syntax error: expected End {what.title()}")

    def parse_statement(self, at_line_start=False):
        t = self.tok
        line = t.line
        prev = self.t[self.i - 1] if self.i > 0 else None
        line_start = prev is None or prev.kind == "nl"
        # labels
        if line_start and t.kind == "num" and isinstance(t.val, int):
            self.next()
            self.accept_op(":")
            return SLabel(str(t.val), t.val, line)
        if line_start and t.kind == "id" and self.is_op(":", tok=self.peek()) \
                and t.val not in RESERVED and t.suffix == "":
            self.next()
            self.next()
            return SLabel(t.val, None, line)
        if t.kind == "id":
            h = getattr(self, "st_" + t.val, None)
            if h is not None:
                r = h()
                if r is not NotImplemented:
                    return r
        return self.parse_simple()

    def parse_simple(self):
        t = self.tok
        line = t.line
        target = self.parse_postfix(stmt_level=True)
        if self.is_op("="):
            self.next()
            expr = self.parse_expr()
            if isinstance(target, Index) and isinstance(target.base, Ident) and target.base.name == "mid":
                a = target.args
                return SMid(a[0].expr, a[1].expr, a[2].expr if len(a) > 2 else None, expr, line)
            return SAssign(target, expr, False, line)
        if isinstance(target, (Lit,)) or isinstance(target, BinOp):
            self.error("Syntax error")
        args = None
        if not self.at_stmt_end():
            args = self.parse_args(None)
        return SCall(target, args, line)

    # individual statements ---------------------------------------------------
    def st_call(self):
        line = self.next().line
        e = self.parse_postfix()
        return SCall(e, None, line, explicit=True)

    def st_let(self):
        line = self.next().line
        target = self.parse_postfix()
        self.expect_op("=")
        return SAssign(target, self.parse_expr(), False, line)

    def st_set(self):
        if self.is_op("=", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        target = self.parse_postfix()
        self.expect_op("=")
        return SAssign(target, self.parse_expr(), True, line)

    def st_lset(self):
        if self.is_op("=", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        target = self.parse_postfix()
        self.expect_op("=")
        return SAssign(target, self.parse_expr(), False, line, "lset")

    def st_rset(self):
        if self.is_op("=", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        target = self.parse_postfix()
        self.expect_op("=")
        return SAssign(target, self.parse_expr(), False, line, "rset")

    def parse_var_decl(self, const=False):
        we = bool(self.accept_id("withevents"))
        n = self.tok
        if n.kind != "id":
            self.error("Syntax error: expected variable name")
        self.next()
        dims = None
        is_array = False
        if self.is_op("("):
            self.next()
            is_array = True
            dims = []
            while not self.is_op(")"):
                a = self.parse_expr()
                if self.accept_id("to"):
                    b = self.parse_expr()
                    dims.append((a, b))
                else:
                    dims.append((None, a))
                if not self.accept_op(","):
                    break
            self.expect_op(")")
        tref = self.suffix_type(n.suffix)
        if self.accept_id("as"):
            tref = self.parse_type()
        init = None
        if const:
            self.expect_op("=")
            init = self.parse_expr()
        d = VarDecl(n.val, n.text, tref or TypeRef(), dims, is_array, init, n.line)
        d.withevents = we
        return d

    def parse_decl_list(self, const=False):
        decls = [self.parse_var_decl(const)]
        while self.accept_op(","):
            decls.append(self.parse_var_decl(const))
        return decls

    def st_dim(self):
        line = self.next().line
        return SDim(self.parse_decl_list(), "dim", line)

    def st_static(self):
        line = self.next().line
        if self.is_id("sub", "function", "property"):
            self.error("Syntax error: nested procedure")
        return SDim(self.parse_decl_list(), "static", line)

    def st_const(self):
        line = self.next().line
        return SDim(self.parse_decl_list(True), "const", line)

    def st_redim(self):
        line = self.next().line
        preserve = bool(self.accept_id("preserve"))
        items = []
        while True:
            # target: name or member, followed by (dims)
            t = self.tok
            base = self.parse_primary()
            while self.is_op(".") and self.peek().kind == "id":
                self.next()
                n = self.next()
                base = Member(base, n.val, n.text, False, n.line)
            self.expect_op("(")
            dims = []
            while True:
                a = self.parse_expr()
                if self.accept_id("to"):
                    b = self.parse_expr()
                    dims.append((a, b))
                else:
                    dims.append((None, a))
                if not self.accept_op(","):
                    break
            self.expect_op(")")
            tref = None
            if self.accept_id("as"):
                tref = self.parse_type()
            items.append((base, dims, tref))
            if not self.accept_op(","):
                break
        return SReDim(items, preserve, line)

    def st_erase(self):
        line = self.next().line
        targets = [self.parse_postfix()]
        while self.accept_op(","):
            targets.append(self.parse_postfix())
        return SErase(targets, line)

    def st_if(self):
        line = self.next().line
        cond = self.parse_expr()
        self.expect_id("then")
        if self.tok.kind in ("nl", "eof"):
            branches = []
            else_ = None
            body = self.parse_block(("elseif", "else", ("end", "if")))
            branches.append((cond, body))
            while True:
                if self.accept_id("elseif"):
                    c = self.parse_expr()
                    self.expect_id("then")
                    b = self.parse_block(("elseif", "else", ("end", "if")))
                    branches.append((c, b))
                    continue
                if self.accept_id("else"):
                    else_ = self.parse_block((("end", "if"),))
                    continue
                break
            self.expect_end("if")
            return SIf(branches, else_, line)
        # single-line If
        then_body = self.parse_inline_stmts()
        else_ = None
        if self.accept_id("else"):
            else_ = self.parse_inline_stmts()
        # optional trailing "End If" on the same line (seen in some code)
        if self.is_id("end") and self.is_id("if", tok=self.peek()):
            self.next()
            self.next()
        return SIf([(cond, then_body)], else_, line)

    def parse_inline_stmts(self):
        body = []
        if self.tok.kind == "num" and isinstance(self.tok.val, int):
            t = self.next()
            return [SGoto(str(t.val), False, t.line)]
        while True:
            while self.is_op(":"):
                self.next()
            if self.tok.kind in ("nl", "eof") or self.is_id("else"):
                break
            if self.is_id("end") and self.is_id("if", tok=self.peek()):
                break
            st = self.parse_statement()
            if isinstance(st, list):
                body.extend(st)
            elif st is not None:
                body.append(st)
            if self.is_op(":"):
                continue
            if self.tok.kind in ("nl", "eof") or self.is_id("else"):
                break
            if not self.is_op(":"):
                break
        return body

    def st_for(self):
        line = self.next().line
        if self.accept_id("each"):
            var = self.parse_postfix()
            self.expect_id("in")
            coll = self.parse_expr()
            body = self.parse_block(("next",))
            self._consume_next()
            return SForEach(var, coll, body, line)
        var = self.parse_postfix()
        self.expect_op("=")
        start = self.parse_expr()
        self.expect_id("to")
        end = self.parse_expr()
        step = None
        if self.accept_id("step"):
            step = self.parse_expr()
        body = self.parse_block(("next",))
        self._consume_next()
        return SFor(var, start, end, step, body, line)

    def _consume_next(self):
        if not self.is_id("next"):
            self.error("Syntax error: For without Next")
        self.next()
        if self.tok.kind == "id" and not self.at_stmt_end():
            self.parse_postfix()
            if self.is_op(","):
                # "Next j, i": turn the comma into a synthetic Next for the outer loop
                c = self.tok
                self.t[self.i] = Tok("id", "next", "Next", c.line, True)

    def st_do(self):
        line = self.next().line
        pre = None
        if self.is_id("while", "until"):
            kind = self.next().val
            pre = (kind, self.parse_expr())
        body = self.parse_block(("loop",))
        end_line = self.tok.line
        self.expect_id("loop")
        post = None
        if self.is_id("while", "until"):
            kind = self.next().val
            post = (kind, self.parse_expr())
        return SDo(pre, post, body, line, end_line)

    def st_while(self):
        line = self.next().line
        cond = self.parse_expr()
        body = self.parse_block(("wend",))
        end_line = self.tok.line
        self.expect_id("wend")
        return SDo(("while", cond), None, body, line, end_line)

    def st_select(self):
        line = self.next().line
        self.expect_id("case")
        expr = self.parse_expr()
        cases = []
        else_ = None
        self.skip_seps()
        while self.is_id("case"):
            cl = self.next().line
            if self.accept_id("else"):
                else_ = self.parse_block(("case", ("end", "select")))
                continue
            conds = []
            while True:
                if self.accept_id("is"):
                    op = self.tok
                    if op.kind != "op" or op.val not in ("=", "<>", "<", ">", "<=", ">="):
                        self.error("Syntax error in Case Is")
                    self.next()
                    conds.append(("is", op.val, self.parse_expr()))
                elif self.tok.kind == "op" and self.tok.val in ("<", ">", "<=", ">=", "<>") :
                    op = self.next()
                    conds.append(("is", op.val, self.parse_expr()))
                else:
                    a = self.parse_expr()
                    if self.accept_id("to"):
                        conds.append(("range", a, self.parse_expr()))
                    else:
                        conds.append(("eq", a))
                if not self.accept_op(","):
                    break
            body = self.parse_block(("case", ("end", "select")))
            cases.append((conds, body, cl))
        self.expect_end("select")
        return SSelect(expr, cases, else_, line)

    def st_with(self):
        line = self.next().line
        e = self.parse_expr()
        body = self.parse_block((("end", "with"),))
        self.expect_end("with")
        return SWith(e, body, line)

    def st_exit(self):
        line = self.next().line
        k = self.expect_id("sub", "function", "property", "for", "do").val
        return SExit(k, line)

    def st_goto(self):
        line = self.next().line
        t = self.next()
        return SGoto(str(t.val), False, line)

    def st_gosub(self):
        line = self.next().line
        t = self.next()
        return SGoto(str(t.val), True, line)

    def st_return(self):
        return SReturn(self.next().line)

    def st_on(self):
        line = self.next().line
        self.accept_id("local")
        if self.accept_id("error"):
            if self.accept_id("resume"):
                self.expect_id("next")
                return SOnError("next", None, line)
            self.expect_id("goto")
            if self.is_op("-"):
                self.next()
                self.next()
                return SOnError("minus1", None, line)
            t = self.next()
            if t.kind == "num" and t.val == 0:
                return SOnError("zero", None, line)
            return SOnError("goto", str(t.val), line)
        e = self.parse_expr()
        gosub = bool(self.accept_id("gosub"))
        if not gosub:
            self.expect_id("goto")
        labels = [str(self.next().val)]
        while self.accept_op(","):
            labels.append(str(self.next().val))
        return SOnGoto(e, labels, gosub, line)

    def st_resume(self):
        line = self.next().line
        if self.accept_id("next"):
            return SResume("next", None, line)
        if self.at_stmt_end():
            return SResume("", None, line)
        t = self.next()
        if t.kind == "num" and t.val == 0:
            return SResume("", None, line)
        return SResume("label", str(t.val), line)

    def st_end(self):
        nt = self.peek()
        if nt.kind == "id" and nt.val in ("if", "sub", "function", "property", "select", "with",
                                          "type", "enum"):
            self.error(f"Syntax error: End {nt.text} without block")
        line = self.next().line
        return SEnd(line)

    def st_stop(self):
        return SStop(self.next().line)

    def st_debug(self):
        if not self.is_op(".", tok=self.peek()):
            return NotImplemented
        line = self.tok.line
        m = self.peek(2)
        if m.kind == "id" and m.val == "print":
            self.next()
            self.next()
            self.next()
            return SPrint("debug", None, self.parse_print_items(), line)
        if m.kind == "id" and m.val == "assert":
            self.next()
            self.next()
            self.next()
            return SAssert(self.parse_expr(), line)
        return NotImplemented

    def parse_print_items(self):
        items = []
        while not self.at_stmt_end():
            if self.is_op(";", ","):
                items.append((None, self.next().val))
                continue
            t = self.tok
            if t.kind == "id" and t.val in ("spc", "tab") and self.is_op("(", tok=self.peek()):
                pass
            e = self.parse_expr()
            sep = ""
            if self.is_op(";", ","):
                sep = self.next().val
            items.append((e, sep))
        return items

    def parse_filenum(self):
        self.accept_op("#")
        return self.parse_expr()

    def st_print(self):
        line = self.tok.line
        if self.tok.text == "?":
            self.next()
            return SPrint("debug", None, self.parse_print_items(), line)
        if not self.is_op("#", tok=self.peek()):
            return NotImplemented
        self.next()
        f = self.parse_filenum()
        self.accept_op(",")
        return SPrint("file", f, self.parse_print_items(), line)

    def st_write(self):
        if not self.is_op("#", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        f = self.parse_filenum()
        self.accept_op(",")
        items = []
        while not self.at_stmt_end():
            if self.is_op(",", ";"):
                self.next()
                continue
            items.append((self.parse_expr(), ","))
        return SPrint("write", f, items, line)

    def st_input(self):
        if not self.is_op("#", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        f = self.parse_filenum()
        self.expect_op(",")
        targets = [self.parse_postfix()]
        while self.accept_op(","):
            targets.append(self.parse_postfix())
        return SInput(f, targets, False, line)

    def st_line(self):
        if not self.is_id("input", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        self.next()
        f = self.parse_filenum()
        self.expect_op(",")
        return SInput(f, [self.parse_postfix()], True, line)

    def st_open(self):
        if self.is_op("=", tok=self.peek()) or self.is_op(".", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        path = self.parse_expr()
        mode = "random"
        if self.accept_id("for"):
            mode = self.next().val
        if self.accept_id("access"):
            self.next()
            if self.is_id("write"):
                self.next()
        while self.is_id("shared", "lock"):
            self.next()
            if self.is_id("read", "write"):
                self.next()
                if self.is_id("write"):
                    self.next()
        self.expect_id("as")
        f = self.parse_filenum()
        reclen = None
        if self.accept_id("len"):
            self.expect_op("=")
            reclen = self.parse_expr()
        return SOpen(path, mode, f, reclen, line)

    def st_close(self):
        if self.is_op("=", tok=self.peek()) or self.is_op(".", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        files = []
        while not self.at_stmt_end():
            if self.accept_op(","):
                continue
            files.append(self.parse_filenum())
        return SClose(files, line)

    def st_reset(self):
        if not self.at_stmt_end_after(1):
            return NotImplemented
        line = self.next().line
        return SClose([], line)

    def at_stmt_end_after(self, k):
        t = self.peek(k)
        return t.kind in ("nl", "eof") or (t.kind == "op" and t.val == ":")

    def st_get(self):
        if self.is_op("#", tok=self.peek()):
            line = self.next().line
            self.skip_to_eol()
            return SUnsupported("Binary file access (Get #)", line)
        return NotImplemented

    def st_put(self):
        if self.is_op("#", tok=self.peek()):
            line = self.next().line
            self.skip_to_eol()
            return SUnsupported("Binary file access (Put #)", line)
        return NotImplemented

    def st_error(self):
        nt = self.peek()
        if nt.kind in ("nl", "eof") or (nt.kind == "op" and nt.val in ("=", ".", "(")) and not nt.space:
            return NotImplemented
        if nt.kind == "op" and nt.val in ("=", "."):
            return NotImplemented
        line = self.next().line
        return SError(self.parse_expr(), line)

    def st_name(self):
        nt = self.peek()
        if nt.kind == "op" and nt.val in ("=", ".", "(", ":"):
            return NotImplemented
        if nt.kind in ("nl", "eof"):
            return NotImplemented
        line = self.next().line
        old = self.parse_expr()
        self.expect_id("as")
        new = self.parse_expr()
        return SName(old, new, line)

    def st_raiseevent(self):
        line = self.next().line
        self.skip_to_eol()
        return SNop(line)

    def st_attribute(self):
        line = self.next().line
        self.skip_to_eol()
        return SNop(line)

    def st_option(self):
        line = self.next().line
        self.skip_to_eol()
        return SNop(line)

    def st_appactivate(self):
        line = self.next().line
        self.skip_to_eol()
        return SUnsupported("AppActivate", line)

    def st_sendkeys(self):
        if self.is_op("=", tok=self.peek()):
            return NotImplemented
        line = self.next().line
        self.skip_to_eol()
        return SUnsupported("SendKeys", line)

    def st_mid(self):
        return NotImplemented

    def st_private(self):
        self.error("Syntax error: Private not allowed inside a procedure")

    st_public = st_private
    st_global = st_private

    def st_sub(self):
        self.error("Syntax error: nested procedure definition (missing End Sub?)")

    st_function = st_sub

    def st_property(self):
        if self.is_id("get", "let", "set", tok=self.peek()):
            self.st_sub()
        return NotImplemented

    def st_wend(self):
        self.error("Syntax error: Wend without While")

    def st_loop(self):
        self.error("Syntax error: Loop without Do")

    def st_next(self):
        self.error("Syntax error: Next without For")

    def st_else(self):
        self.error("Syntax error: Else without If")

    st_elseif = st_else

    def st_case(self):
        self.error("Syntax error: Case without Select Case")

    # -- module level ----------------------------------------------------------
    def parse_params(self):
        params = []
        if not self.accept_op("("):
            return params
        while not self.is_op(")"):
            optional = byref = paramarray = False
            byref = True
            while self.is_id("optional", "byval", "byref", "paramarray"):
                k = self.next().val
                if k == "optional":
                    optional = True
                elif k == "byval":
                    byref = False
                elif k == "byref":
                    byref = True
                else:
                    paramarray = True
            n = self.tok
            if n.kind != "id":
                self.error("Syntax error in parameter list")
            self.next()
            is_array = False
            if self.is_op("("):
                self.next()
                self.expect_op(")")
                is_array = True
            tref = self.suffix_type(n.suffix)
            if self.accept_id("as"):
                tref = self.parse_type()
            default = None
            if self.accept_op("="):
                default = self.parse_expr()
            params.append(Param(n.val, n.text, byref, optional, default, paramarray, tref, is_array))
            if not self.accept_op(","):
                break
        self.expect_op(")")
        return params

    def parse_proc(self, mod, private, static):
        kw = self.next()
        line = kw.line
        kind = kw.val
        if kind == "property":
            kind = self.expect_id("get", "let", "set").val
        n = self.tok
        if n.kind != "id":
            self.error("Syntax error: expected procedure name")
        self.next()
        params = self.parse_params()
        ret = self.suffix_type(n.suffix)
        if self.accept_id("as"):
            ret = self.parse_type()
        if self.is_id("static"):
            self.next()
        endword = "sub" if kind == "sub" else "function" if kind == "function" else "property"
        proc = ProcDef(n.val, n.text, kind, params, ret or TypeRef(), [], private, static, line, mod)
        try:
            proc.body = self.parse_block(((("end", endword)),))
            if self.tok.kind == "eof":
                raise VBACompileErr(f"Syntax error: expected End {endword.title()}", line, self.mod)
            proc.end_line = self.tok.line
            self.expect_end(endword)
        except VBACompileErr as e:
            proc.error = e
            if e.module is None:
                e.module = self.mod
            mod.errors.append(e)
            # recover: skip to End <endword>
            while self.tok.kind != "eof":
                if self._at_ender(((("end", endword)),)):
                    proc.end_line = self.tok.line
                    self.next()
                    self.next()
                    break
                self.next()
        return proc

    def parse_module(self, mod):
        while True:
            self.skip_seps()
            t = self.tok
            if t.kind == "eof":
                break
            start = self.i
            try:
                self.parse_module_item(mod)
            except VBACompileErr as e:
                if e.module is None:
                    e.module = self.mod
                mod.errors.append(e)
                if self.i == start:
                    self.next()
                self.skip_to_eol()
        return mod

    def parse_module_item(self, mod):
        t = self.tok
        if t.kind != "id":
            self.error("Syntax error: invalid outside procedure")
        v = t.val
        if v == "option":
            self.next()
            w = self.next().val
            if w == "explicit":
                mod.option_explicit = True
            elif w == "base":
                mod.option_base = int(self.next().val)
            elif w == "compare":
                mod.option_compare_text = self.next().val == "text"
            elif w == "private":
                mod.option_private = True
                self.skip_to_eol()
            return
        if v in ("attribute", "implements", "event"):
            self.skip_to_eol()
            return
        private = False
        static = False
        scope_kw = False
        while self.is_id("public", "private", "global", "friend", "static", "dim"):
            k = self.next().val
            scope_kw = True
            if k == "private":
                private = True
            elif k == "static":
                static = True
            elif k == "dim":
                private = True
                break
        t = self.tok
        if self.is_id("sub", "function") or (self.is_id("property") and self.is_id("get", "let", "set", tok=self.peek())):
            mod.procs.append(self.parse_proc(mod, private, static))
            return
        if self.is_id("const"):
            self.next()
            for d in self.parse_decl_list(True):
                mod.consts.append((d, private))
            return
        if self.is_id("type"):
            self.parse_type_block(mod, private)
            return
        if self.is_id("enum"):
            self.parse_enum_block(mod, private)
            return
        if self.is_id("declare"):
            self.next()
            self.accept_id("ptrsafe")
            self.expect_id("sub", "function")
            n = self.name_tok()
            mod.declares[n.val] = n.text
            self.skip_to_eol()
            return
        if self.is_id("event"):
            self.skip_to_eol()
            return
        if scope_kw:
            for d in self.parse_decl_list():
                mod.vars.append((d, private))
            return
        self.error(f"Syntax error: invalid outside procedure ('{t.text}')")

    def parse_type_block(self, mod, private):
        line = self.next().line
        n = self.name_tok()
        fields = []
        while True:
            self.skip_seps()
            if self.is_id("end") and self.is_id("type", tok=self.peek()):
                self.next()
                self.next()
                break
            if self.tok.kind == "eof":
                self.error("Syntax error: expected End Type")
            fields.append(self.parse_var_decl())
        mod.types.append(TypeDef(n.val, n.text, fields, private, line))

    def parse_enum_block(self, mod, private):
        line = self.next().line
        n = self.name_tok()
        members = []
        while True:
            self.skip_seps()
            if self.is_id("end") and self.is_id("enum", tok=self.peek()):
                self.next()
                self.next()
                break
            if self.tok.kind == "eof":
                self.error("Syntax error: expected End Enum")
            if self.tok.kind == "brk":
                m = self.next()
                mname, mtext = m.val.lower(), m.val
            else:
                m = self.name_tok()
                mname, mtext = m.val, m.text
            e = None
            if self.accept_op("="):
                e = self.parse_expr()
            members.append((mname, mtext, e))
        mod.enums.append(EnumDef(n.val, n.text, members, private, line))


def parse_module(name, mtype, code):
    mod = ModuleDef(name, mtype)
    try:
        toks = tokenize(code)
    except VBACompileErr as e:
        e.module = name
        mod.errors.append(e)
        return mod
    Parser(toks, name).parse_module(mod)
    return mod


def parse_statements(code, name="Immediate"):
    toks = tokenize(code)
    p = Parser(toks, name)
    body = p.parse_block(())
    return body

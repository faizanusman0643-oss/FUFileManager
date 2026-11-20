"""Minimal stand-in for portable_excel.vba (GUI tests / screenshots only).

Understands just enough to run demo macros: Sub/End Sub, MsgBox "..",
Range("A1").Value = .., Cells(r, c).Value = .., For i = a To b / Next,
Debug.Print ...
"""
import re

from portable_excel.model import parse_a1


class VBAError(Exception):
    def __init__(self, message, module="", line=0):
        super().__init__(message)
        self.message, self.module, self.line = message, module, line


class VBAHost:
    def msgbox(self, prompt, buttons=0, title=""): return 1
    def inputbox(self, prompt, title="", default=""): return default
    def debug_print(self, text): pass
    def select(self, sheet, r1, c1, r2, c2): pass
    def get_selection(self): return None
    def activate_sheet(self, sheet): pass
    def refresh(self): pass
    def status_bar(self, text): pass
    def process_events(self): pass


class VBARuntime:
    def __init__(self, book, engine, host):
        self.book, self.engine, self.host = book, engine, host
        self.procs = {}
        self.stop = False

    def load_modules(self, modules):
        self.procs = {}
        errs = []
        for m in modules:
            cur = None
            for i, line in enumerate(m.get("code", "").splitlines(), 1):
                s = line.strip()
                mm = re.match(r"(?:Public\s+|Private\s+)?Sub\s+(\w+)", s, re.I)
                if mm:
                    cur = mm.group(1)
                    self.procs[cur] = (m["name"], [])
                elif re.match(r"End\s+Sub", s, re.I):
                    cur = None
                elif cur:
                    self.procs[cur][1].append((i, s))
        return errs

    def list_macros(self):
        return sorted(self.procs)

    def run(self, name):
        if name not in self.procs:
            raise VBAError(f"Sub or Function not defined: {name}")
        mod, lines = self.procs[name]
        self._exec(mod, lines, {})

    def _expr(self, e, env):
        e = e.strip()
        for k, v in env.items():
            e = re.sub(rf"\b{k}\b", str(v), e)
        e = e.replace("&", "+")
        try:
            return eval(e, {"__builtins__": {}}, {})
        except Exception:
            return e.strip('"')

    def _exec(self, mod, lines, env):
        i = 0
        sh = self.book.active
        while i < len(lines):
            ln, s = lines[i]
            if self.stop:
                raise VBAError("Execution interrupted", mod, ln)
            m = re.match(r"For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+)", s, re.I)
            if m:
                depth, j = 1, i + 1
                while j < len(lines):
                    if re.match(r"For\b", lines[j][1], re.I):
                        depth += 1
                    if re.match(r"Next\b", lines[j][1], re.I):
                        depth -= 1
                        if not depth:
                            break
                    j += 1
                for v in range(int(self._expr(m.group(2), env)), int(self._expr(m.group(3), env)) + 1):
                    env[m.group(1)] = v
                    self._exec(mod, lines[i + 1:j], env)
                i = j + 1
                continue
            m = re.match(r'MsgBox\s+(.+)', s, re.I)
            if m:
                self.host.msgbox(str(self._expr(m.group(1), env)), 64, "PortableExcel")
            m = re.match(r'Debug\.Print\s+(.+)', s, re.I)
            if m:
                self.host.debug_print(str(self._expr(m.group(1), env)))
            m = re.match(r'Range\("(\w+)"\)(?:\.Value)?\s*=\s*(.+)', s, re.I)
            if m:
                r, c = parse_a1(m.group(1))
                self.book.set_cell(sh, r, c, self._expr(m.group(2), env))
            m = re.match(r'Cells\((.+?),\s*(.+?)\)(?:\.Value)?\s*=\s*(.+)', s, re.I)
            if m:
                self.book.set_cell(sh, int(self._expr(m.group(1), env)) - 1,
                                   int(self._expr(m.group(2), env)) - 1,
                                   self._expr(m.group(3), env))
            if s.lower().startswith("err.raise"):
                raise VBAError("Application-defined or object-defined error", mod, ln)
            self.host.process_events()
            i += 1

    def fire_event(self, module_name, proc, args=()):
        if proc in self.procs:
            self.run(proc)

    def notify_change(self, sheet, r1, c1, r2, c2):
        pass

    def make_range(self, sheet, r1, c1, r2, c2):
        return (sheet, r1, c1, r2, c2)

    def request_stop(self):
        self.stop = True

    def evaluate_immediate(self, text):
        t = text.strip()
        if t.startswith("?"):
            return str(self._expr(t[1:], {}))
        return ""

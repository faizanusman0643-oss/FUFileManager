"""Error types used by the VBA interpreter."""
from __future__ import annotations

ERROR_MESSAGES = {
    3: "Return without GoSub",
    5: "Invalid procedure call or argument",
    6: "Overflow",
    7: "Out of memory",
    9: "Subscript out of range",
    10: "This array is fixed or temporarily locked",
    11: "Division by zero",
    13: "Type mismatch",
    14: "Out of string space",
    16: "Expression too complex",
    17: "Can't perform requested operation",
    18: "User interrupt occurred",
    20: "Resume without error",
    28: "Out of stack space",
    35: "Sub or Function not defined",
    48: "Error in loading DLL",
    51: "Internal error",
    52: "Bad file name or number",
    53: "File not found",
    54: "Bad file mode",
    55: "File already open",
    57: "Device I/O error",
    58: "File already exists",
    61: "Disk full",
    62: "Input past end of file",
    67: "Too many files",
    68: "Device unavailable",
    70: "Permission denied",
    71: "Disk not ready",
    74: "Can't rename with different drive",
    75: "Path/File access error",
    76: "Path not found",
    91: "Object variable or With block variable not set",
    92: "For loop not initialized",
    93: "Invalid pattern string",
    94: "Invalid use of Null",
    424: "Object required",
    429: "ActiveX component can't create object",
    438: "Object doesn't support this property or method",
    440: "Automation error",
    445: "Object doesn't support this action",
    446: "Object doesn't support named arguments",
    448: "Named argument not found",
    449: "Argument not optional",
    450: "Wrong number of arguments or invalid property assignment",
    451: "Property let procedure not defined and property get procedure did not return an object",
    457: "This key is already associated with an element of this collection",
    1004: "Application-defined or object-defined error",
    32811: "Element not found",
}


def error_message(number: int) -> str:
    return ERROR_MESSAGES.get(number, "Application-defined or object-defined error")


class VBAError(Exception):
    """Error reported to the host (GUI).

    message    -- full text, e.g. "Run-time error '9': Subscript out of range"
    module     -- module name where it happened (or None)
    line       -- 1-based line number in that module's code (or None)
    vba_number -- Err.Number (0 for compile/parse errors)
    """

    def __init__(self, message: str, module: str | None = None, line: int | None = None,
                 vba_number: int = 0, description: str | None = None,
                 procedure: str | None = None):
        super().__init__(message)
        self.message = message
        self.module = module
        self.line = line
        self.vba_number = vba_number
        self.description = description if description is not None else message
        self.procedure = procedure

    def __str__(self) -> str:
        loc = ""
        if self.module:
            loc = f" (in {self.module}"
            if self.line:
                loc += f", line {self.line}"
            loc += ")"
        return self.message + loc


class VBARuntimeErr(Exception):
    """Internal: a trappable VBA run-time error (Err.Raise, type mismatch...)."""

    def __init__(self, number: int, description: str | None = None, source: str = ""):
        if description is None or description == "":
            description = error_message(number)
        super().__init__(description)
        self.number = number
        self.description = description
        self.source = source
        self.module = None
        self.line = None
        self.procedure = None
        self.fatal = False   # not trappable by On Error (e.g. unsupported features)


class VBACompileErr(Exception):
    """Internal: compile/syntax error."""

    def __init__(self, message: str, line: int | None = None, module: str | None = None):
        super().__init__(message)
        self.message = message
        self.line = line
        self.module = module


class VBAControl(Exception):
    """Base for non-error control flow that unwinds the whole call stack."""


class VBAEnd(VBAControl):
    """The End statement."""


class VBAInterrupt(VBAControl):
    """User pressed Esc / Ctrl+Break."""


class VBAStop(VBAControl):
    """Stop statement / failed Debug.Assert."""

    def __init__(self, message="Stop statement"):
        super().__init__(message)
        self.message = message
        self.module = None
        self.line = None


def rt_error(number: int, description: str | None = None) -> VBARuntimeErr:
    return VBARuntimeErr(number, description)


def unsupported(what: str) -> VBARuntimeErr:
    e = VBARuntimeErr(5, f"{what} is not supported in PortableExcel")
    e.fatal = False
    return e

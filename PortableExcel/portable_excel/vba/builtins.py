"""VBA built-in functions and constants."""
from __future__ import annotations

import inspect
import math
import os
import re
import time
import glob as _glob
import shutil
import tempfile
from datetime import datetime, timedelta
from decimal import Decimal

from .errors import rt_error, VBARuntimeErr, error_message
from .values import (EMPTY, NULL, NOTHING, MISSING, VDate, VErr, VBAArray, UDTValue, VBAObject,
                     to_str, to_num, to_float, to_double, to_bool, to_date, to_integer, to_long,
                     to_longlong, to_byte, to_single, to_currency, to_value, index_int,
                     str_to_num_or_none, type_name, var_type, vba_round, fmt_double,
                     concat_str, _to_decimal, round_half_up)
from . import dates as D
from .objects import bind

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------
CONSTANTS = {
    "vbcrlf": "\r\n", "vbcr": "\r", "vblf": "\n", "vbnewline": "\r\n", "vbtab": "\t",
    "vbnullstring": "", "vbnullchar": "\0", "vbback": "\b", "vbformfeed": "\x0c",
    "vbverticaltab": "\x0b", "vbobjecterror": -2147221504,
    "vbokonly": 0, "vbokcancel": 1, "vbabortretryignore": 2, "vbyesnocancel": 3, "vbyesno": 4,
    "vbretrycancel": 5, "vbcritical": 16, "vbquestion": 32, "vbexclamation": 48,
    "vbinformation": 64, "vbdefaultbutton1": 0, "vbdefaultbutton2": 256,
    "vbdefaultbutton3": 512, "vbdefaultbutton4": 768, "vbapplicationmodal": 0,
    "vbsystemmodal": 4096, "vbmsgboxhelpbutton": 16384, "vbmsgboxsetforeground": 65536,
    "vbmsgboxright": 524288, "vbmsgboxrtlreading": 1048576,
    "vbok": 1, "vbcancel": 2, "vbabort": 3, "vbretry": 4, "vbignore": 5, "vbyes": 6, "vbno": 7,
    "vbbinarycompare": 0, "vbtextcompare": 1, "vbdatabasecompare": 2,
    "vbblack": 0, "vbred": 255, "vbgreen": 65280, "vbyellow": 65535, "vbblue": 16711680,
    "vbmagenta": 16711935, "vbcyan": 16776960, "vbwhite": 16777215,
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4, "vbdouble": 5,
    "vbcurrency": 6, "vbdate": 7, "vbstring": 8, "vbobject": 9, "vberror": 10,
    "vbboolean": 11, "vbvariant": 12, "vbdataobject": 13, "vbdecimal": 14, "vbbyte": 17,
    "vblonglong": 20, "vbuserdefinedtype": 36, "vbarray": 8192,
    "vbusesystemdayofweek": 0, "vbsunday": 1, "vbmonday": 2, "vbtuesday": 3,
    "vbwednesday": 4, "vbthursday": 5, "vbfriday": 6, "vbsaturday": 7,
    "vbusesystem": 0, "vbfirstjan1": 1, "vbfirstfourdays": 2, "vbfirstfullweek": 3,
    "vbuppercase": 1, "vblowercase": 2, "vbpropercase": 3, "vbwide": 4, "vbnarrow": 8,
    "vbunicode": 64, "vbfromunicode": 128,
    "vbmethod": 1, "vbget": 2, "vblet": 4, "vbset": 8,
    "vbtrue": -1, "vbfalse": 0, "vbusedefault": -2,
    "vbnormal": 0, "vbreadonly": 1, "vbhidden": 2, "vbsystem": 4, "vbvolume": 8,
    "vbdirectory": 16, "vbarchive": 32,
    "vbgeneraldate": 0, "vblongdate": 1, "vbshortdate": 2, "vblongtime": 3, "vbshorttime": 4,
    "vbkeyreturn": 13, "vbkeyescape": 27, "vbkeytab": 9, "vbkeyspace": 32, "vbkeyback": 8,
    "vbhide": 0, "vbnormalfocus": 1, "vbminimizedfocus": 2, "vbmaximizedfocus": 3,
    # Excel
    "xlup": -4162, "xldown": -4121, "xltoleft": -4159, "xltoright": -4161,
    "xlcenter": -4108, "xlleft": -4131, "xlright": -4152, "xlgeneral": 1, "xljustify": -4130,
    "xltop": -4160, "xlbottom": -4107, "xldistributed": -4117, "xlfill": 5,
    "xlcenteracrossselection": 7,
    "xlvalues": -4163, "xlformulas": -4123, "xlcomments": -4144, "xlpart": 2, "xlwhole": 1,
    "xlascending": 1, "xldescending": 2, "xlyes": 1, "xlno": 2, "xlguess": 0,
    "xlcalculationautomatic": -4105, "xlcalculationmanual": -4135,
    "xlcalculationsemiautomatic": 2, "xlautomatic": -4105, "xlmanual": -4135,
    "xlpastevalues": -4163, "xlpasteformats": -4122, "xlpasteall": -4104,
    "xlpasteformulas": -4123, "xlpastevaluesandnumberformats": 12,
    "xlpasteformulasandnumberformats": 11, "xlpasteallexceptborders": 7,
    "xlpastecolumnwidths": 8, "xlpastecomments": -4144, "xlpastespecialoperationnone": -4142,
    "xlpastespecialoperationadd": 2, "xlpastespecialoperationsubtract": 3,
    "xlpastespecialoperationmultiply": 4, "xlpastespecialoperationdivide": 5,
    "xlcelltypelastcell": 11, "xlcelltypeconstants": 2, "xlcelltypeformulas": -4123,
    "xlcelltypeblanks": 4, "xlcelltypevisible": 12, "xlcelltypecomments": -4144,
    "xllastcell": 11, "xltextvalues": 2, "xlnumbers": 1, "xllogical": 4, "xlerrors": 16,
    "xlsheetvisible": -1, "xlsheethidden": 0, "xlsheetveryhidden": 2,
    "xlnone": -4142, "xlcontinuous": 1, "xldash": -4115, "xldot": -4118, "xldouble": -4119,
    "xldashdot": 4, "xldashdotdot": 5, "xlslantdashdot": 13,
    "xllinestylenone": -4142, "xlhairline": 1, "xlthin": 2, "xlmedium": -4138, "xlthick": 4,
    "xledgeleft": 7, "xledgetop": 8, "xledgebottom": 9, "xledgeright": 10,
    "xlinsidevertical": 11, "xlinsidehorizontal": 12, "xldiagonaldown": 5, "xldiagonalup": 6,
    "xlshiftup": -4162, "xlshiftdown": -4121, "xlshifttoleft": -4159, "xlshifttoright": -4161,
    "xlformatfromleftorabove": 0, "xlformatfromrightorbelow": 1,
    "xlbyrows": 1, "xlbycolumns": 2, "xlnext": 1, "xlprevious": 2,
    "xlfilldefault": 0, "xlfillcopy": 1, "xlfillseries": 2, "xlfillformats": 3,
    "xlfillvalues": 4, "xlfilldays": 5, "xlfillweekdays": 6, "xlfillmonths": 7,
    "xlfillyears": 8, "xllineartrend": 9, "xlgrowthtrend": 10,
    "xla1": 1, "xlr1c1": -4150, "xlsolid": 1, "xlgray50": -4125,
    "xlunderlinestylesingle": 2, "xlunderlinestylenone": -4142, "xlunderlinestyledouble": -4119,
    "xlcolorindexnone": -4142, "xlcolorindexautomatic": -4105, "xlworksheet": -4167,
    "xlerrdiv0": 2007, "xlerrna": 2042, "xlerrname": 2029, "xlerrnull": 2000,
    "xlerrnum": 2036, "xlerrref": 2023, "xlerrvalue": 2015,
    "xltoptobottom": 1, "xllefttoright": 2, "xlsortonvalues": 0, "xlsortnormal": 0,
    "xlsorttextasnumbers": 1, "xlpinyin": 1, "xlstroke": 2, "xlsortcolumns": 1, "xlsortrows": 2,
    "xlcsv": 6, "xlopenxmlworkbook": 51, "xlopenxmlworkbookmacroenabled": 52,
    "xlworkbooknormal": -4143, "xlexcel8": 56, "xlexcel12": 50, "xlworkbookdefault": 51,
    "xlmaximized": -4137, "xlminimized": -4140, "xlnormal": -4143, "xlwait": 2,
    "xldefault": -4143, "xlnorthwestarrow": 1, "xlibeam": 3, "xlfiltervalues": 7, "xland": 1,
    "xlor": 2, "xltop10items": 3, "xllandscape": 2, "xlportrait": 1, "xldatabase": 1,
    "xlcontext": -5002, "xlltr": -5003, "xlrtl": -5004, "xlerrorhandler": 2,
    "xlinterrupt": 1, "xldisabled": 0, "xlchart": -4109, "xlthemecolordark1": 1,
    "xlthemecolorlight1": 2, "xlthemecolordark2": 3, "xlthemecolorlight2": 4,
    "xlthemecoloraccent1": 5, "xlthemecoloraccent2": 6, "xlthemecoloraccent3": 7,
    "xlthemecoloraccent4": 8, "xlthemecoloraccent5": 9, "xlthemecoloraccent6": 10,
    "xlpatternsolid": 1, "xlpatternnone": -4142, "xlpatternautomatic": -4105,
    "xlinsertdeleteformat": 0, "xlbyrow": 1, "xlnotes": -4144, "xlexclusive": 3,
    "xllocalsessionchanges": 2, "xlusercancel": 1, "xlcount": -4112, "xlsum": -4157,
    "xlaverage": -4106, "xlmax": -4136, "xlmin": -4139, "xlcelltypeallvalidation": -4174,
    "xlcelltypeallformatconditions": -4172, "xlcelltypesameformatconditions": -4173,
    "xlcelltypesamevalidation": -4175, "xldialogopen": 1, "xlhorizontal": -4128,
    "xlvertical": -4166, "xlclipboardformattext": 0, "xlvalidatelist": 3,
    "xlvalidalertstop": 1, "xlbetween": 1, "xlequal": 3, "xlgreater": 5, "xlless": 6,
    "xlcellvalue": 1, "xlexpression": 2,
    "xltypepdf": 0, "xltypexps": 1, "xlqualitystandard": 0, "xlqualityminimum": 1,
    "xldialogprintersetup": 9, "xldialogprint": 8, "xldialogsaveas": 5,
    "fmmultiselectsingle": 0, "fmmultiselectmulti": 1, "fmmultiselectextended": 2,
    "fmscrollbarsnone": 0, "fmscrollbarshorizontal": 1, "fmscrollbarsvertical": 2,
    "fmscrollbarsboth": 3, "fmspecialeffectflat": 0, "fmspecialeffectraised": 1,
    "fmspecialeffectsunken": 2, "fmbackstyletransparent": 0, "fmbackstyleopaque": 1,
    "fmtextalignleft": 1, "fmtextaligncenter": 2, "fmtextalignright": 3,
    "fmborderstylenone": 0, "fmborderstylesingle": 1, "vbmodeless": 0, "vbmodal": 1,
    "vbformcontrolmenu": 0, "vbformcode": 1, "vbappwindows": 2, "vbapptaskmanager": 3,
    "fmbackstyletransparent": 0,
    "fmbackstyleopaque": 1,
    "fmborderstylenone": 0,
    "fmborderstylesingle": 1,
    "fmspecialeffectflat": 0,
    "fmspecialeffectraised": 1,
    "fmspecialeffectsunken": 2,
    "fmspecialeffectetched": 3,
    "fmspecialeffectbump": 6,
    "fmtextalignleft": 1,
    "fmtextaligncenter": 2,
    "fmtextalignright": 3,
    "fmscrollbarsnone": 0,
    "fmscrollbarshorizontal": 1,
    "fmscrollbarsvertical": 2,
    "fmscrollbarsboth": 3,
    "fmmultiselectsingle": 0,
    "fmmultiselectmulti": 1,
    "fmmultiselectextended": 2,
    "fmliststyleplain": 0,
    "fmliststyleoption": 1,
    "fmstyledropdowncombo": 0,
    "fmstyledropdownlist": 2,
    "fmpicturesizemodeclip": 0,
    "fmpicturesizemodestretch": 1,
    "fmpicturesizemodezoom": 3,
    "fmmousepointerdefault": 0,
    "fmmousepointerarrow": 1,
    "fmmousepointercross": 2,
    "fmmousepointeribeam": 3,
    "fmmousepointersizenesw": 6,
    "fmmousepointersizens": 7,
    "fmmousepointersizenwse": 8,
    "fmmousepointersizewe": 9,
    "fmmousepointeruparrow": 10,
    "fmmousepointerhourglass": 11,
    "fmmousepointernodrop": 12,
    "fmmousepointerappstarting": 13,
    "fmmousepointerhelp": 14,
    "fmmousepointersizeall": 15,
    "fmmousepointercustom": 99,
    "fmorientationauto": -1,
    "fmorientationvertical": 0,
    "fmorientationhorizontal": 1,
    "fmkeepscrollbarsvisiblenone": 0,
    "fmkeepscrollbarsvisiblehorizontal": 1,
    "fmkeepscrollbarsvisiblevertical": 2,
    "fmkeepscrollbarsvisibleboth": 3,
    "fmcycleallforms": 0,
    "fmcyclecurrentform": 2,
    "fmtransitioneffectnone": 0,
    "fmshowlistwhennever": 0,
    "fmshowlistwhenbutton": 1,
    "fmshowlistwhenfocus": 2,
    "fmshowlistwhenalways": 3,
    "fmdropbuttonstylearrow": 1,
    "fmmatchentryfirstletter": 0,
    "fmmatchentrycomplete": 1,
    "fmmatchentrynone": 2,
    "fmenterfieldbehaviorselectall": 0,
    "fmenterfieldbehaviorrecallselection": 1,
    "fmdragbehaviordisabled": 0,
    "fmtabstyletabs": 0,
    "fmtabstylebuttons": 1,
    "fmtabstylenone": 2,
    "fmtaborientationtop": 0,
    "fmbuttoneffectflat": 0,
    "fmbuttoneffectsunken": 2,
    "fmpicturepositionabovecenter": 7,
    "fmzoomnone": 100,
    "fmmodeless": 0,
    "fmalignmentleft": 0,
    "fmalignmentright": 1,
    "fmscrollactionnochange": 0,
    "fmscrollactionlineup": 1,
    "fmscrollactionlinedown": 2,
    "fmpictureappearanceflat": 0,
    "fmverticalscrollbarsideright": 0,
    "msotrue": -1, "msofalse": 0, "msofiledialogfolderpicker": 4, "msofiledialogfilepicker": 3,
}

XL_ERR_CONST = {"xlerrdiv0", "xlerrna", "xlerrname", "xlerrnull", "xlerrnum", "xlerrref",
                "xlerrvalue"}

# --------------------------------------------------------------------------
# Registry
# --------------------------------------------------------------------------
BUILTINS = {}


class Builtin:
    __slots__ = ("name", "fn", "params", "nparams", "varargs")

    def __init__(self, name, fn):
        self.name = name
        self.fn = fn
        sig = inspect.signature(fn)
        ps = list(sig.parameters.values())[1:]
        self.varargs = any(p.kind == p.VAR_POSITIONAL for p in ps)
        self.params = [p.name.rstrip("_") for p in ps if p.kind != p.VAR_POSITIONAL]
        self.nparams = len(self.params)

    def call(self, rt, args, named):
        if named:
            args = bind(args, named, self.params)
            while args and args[-1] is MISSING:
                args.pop()
        elif not self.varargs and len(args) > self.nparams:
            raise rt_error(450)
        return self.fn(rt, *args)


def builtin(*names):
    def deco(fn):
        b = Builtin(names[0], fn)
        for n in names:
            BUILTINS[n] = b
        return fn
    return deco


def _m(v):
    return v is MISSING


def _nullish(v):
    v = to_value(v) if isinstance(v, VBAObject) else v
    return v


# --------------------------------------------------------------------------
# Interaction
# --------------------------------------------------------------------------
@builtin("msgbox")
def b_msgbox(rt, prompt, buttons=MISSING, title=MISSING, helpfile=MISSING, context=MISSING):
    p = to_str(prompt)
    b = 0 if _m(buttons) else index_int(buttons)
    t = "Microsoft Excel" if _m(title) else to_str(title)
    r = rt.host.msgbox(p, b, t)
    return 1 if r is None else int(r)


@builtin("inputbox")
def b_inputbox(rt, prompt, title=MISSING, default=MISSING, xpos=MISSING, ypos=MISSING,
               helpfile=MISSING, context=MISSING):
    r = rt.host.inputbox(to_str(prompt), "Microsoft Excel" if _m(title) else to_str(title),
                         "" if _m(default) else to_str(default))
    return "" if r is None else str(r)


@builtin("doevents")
def b_doevents(rt):
    rt.host.process_events()
    rt.check_stop()
    return 0


@builtin("beep")
def b_beep(rt):
    return EMPTY


@builtin("shell")
def b_shell(rt, pathname, windowstyle=MISSING):
    raise VBARuntimeErr(5, "Shell is not supported in PortableExcel (running external programs is disabled)")


@builtin("sendkeys")
def b_sendkeys(rt, s, wait=MISSING):
    raise VBARuntimeErr(5, "SendKeys is not supported in PortableExcel")


@builtin("appactivate")
def b_appactivate(rt, title, wait=MISSING):
    raise VBARuntimeErr(5, "AppActivate is not supported in PortableExcel")


@builtin("environ", "environ$")
def b_environ(rt, expression):
    v = to_value(expression)
    if not isinstance(v, str):
        return ""
    k = v.upper()
    if k == "USERNAME":
        import getpass
        try:
            return getpass.getuser()
        except Exception:
            return os.environ.get("USER", "")
    if k == "COMPUTERNAME":
        import platform
        return platform.node()
    if k in ("TEMP", "TMP"):
        return tempfile.gettempdir()
    if k == "USERPROFILE":
        return os.path.expanduser("~")
    return ""


@builtin("createobject")
def b_createobject(rt, cls, servername=MISSING):
    return rt.create_com(to_str(cls))


@builtin("getobject")
def b_getobject(rt, pathname=MISSING, cls=MISSING):
    raise rt_error(429)


@builtin("callbyname")
def b_callbyname(rt, obj, procname, calltype, *args):
    name = to_str(procname).lower()
    ct = index_int(calltype)
    args = list(args)
    if not isinstance(obj, VBAObject):
        raise rt_error(424)
    if ct in (4, 8):
        value = args.pop()
        obj.vba_let(name, args, {}, value, ct == 8)
        return EMPTY
    return obj.vba_get(name, args, {})


@builtin("load")
def b_load(rt, obj):
    from .userforms import FormInstance
    if not isinstance(obj, FormInstance):
        raise rt_error(424)
    return EMPTY


@builtin("unload")
def b_unload(rt, obj):
    from .userforms import FormInstance, ControlObj
    if isinstance(obj, FormInstance):
        obj.unload(nested=True)
        return EMPTY
    raise rt_error(424)


@builtin("randomize")
def b_randomize(rt, number=MISSING):
    if _m(number):
        rt.rng.seed()
    else:
        rt.rng.seed(to_float(number))
    return EMPTY


@builtin("rnd")
def b_rnd(rt, number=MISSING):
    if not _m(number):
        n = to_float(number)
        if n < 0:
            rt.rng.seed(n)
            rt.last_rnd = rt.rng.random()
            return rt.last_rnd
        if n == 0:
            return rt.last_rnd
    rt.last_rnd = rt.rng.random()
    return rt.last_rnd


@builtin("rgb")
def b_rgb(rt, red, green, blue):
    def c(x):
        x = index_int(x)
        return max(0, min(255, x))
    return c(red) + c(green) * 256 + c(blue) * 65536


QB_COLORS = [0x000000, 0x800000, 0x008000, 0x808000, 0x000080, 0x800080, 0x008080, 0xC0C0C0,
             0x808080, 0xFF0000, 0x00FF00, 0xFFFF00, 0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF]


@builtin("qbcolor")
def b_qbcolor(rt, color):
    i = index_int(color)
    if not 0 <= i <= 15:
        raise rt_error(5)
    return QB_COLORS[i]


@builtin("error", "error$")
def b_error(rt, errornumber=MISSING):
    if _m(errornumber):
        return rt.err.description
    n = index_int(errornumber)
    return "" if n == 0 else error_message(n)


@builtin("erl")
def b_erl(rt):
    return rt.err.erl


@builtin("cverr")
def b_cverr(rt, expression):
    return VErr(index_int(expression))


# --------------------------------------------------------------------------
# Type information
# --------------------------------------------------------------------------
@builtin("isnumeric")
def b_isnumeric(rt, expression):
    v = expression
    if isinstance(v, VBAObject):
        try:
            v = to_value(v)
        except VBARuntimeErr:
            return False
    t = type(v)
    if t in (int, float, bool, Decimal) or v is EMPTY:
        return True
    if t is str:
        return str_to_num_or_none(v) is not None
    return False


@builtin("isdate")
def b_isdate(rt, expression):
    v = expression
    if isinstance(v, VBAObject):
        try:
            v = to_value(v)
        except VBARuntimeErr:
            return False
    if type(v) is VDate:
        return True
    if isinstance(v, str):
        return D.parse_date_string(v) is not None
    return False


@builtin("isempty")
def b_isempty(rt, expression):
    v = expression
    if isinstance(v, VBAObject):
        try:
            v = to_value(v)
        except VBARuntimeErr:
            return False
    return v is EMPTY


@builtin("isnull")
def b_isnull(rt, expression):
    v = expression
    if isinstance(v, VBAObject):
        try:
            v = to_value(v)
        except VBARuntimeErr:
            return False
    return v is NULL


@builtin("isarray")
def b_isarray(rt, varname):
    return type(varname) is VBAArray


@builtin("isobject")
def b_isobject(rt, identifier):
    return isinstance(identifier, VBAObject) or identifier is NOTHING


@builtin("iserror")
def b_iserror(rt, expression):
    v = expression
    if isinstance(v, VBAObject):
        try:
            v = to_value(v)
        except VBARuntimeErr:
            return False
    return type(v) is VErr


@builtin("ismissing")
def b_ismissing(rt, argname):
    return argname is MISSING


@builtin("typename")
def b_typename(rt, varname):
    return type_name(varname)


@builtin("vartype")
def b_vartype(rt, varname):
    return var_type(varname)


@builtin("nz")
def b_nz(rt, value, valueifnull=MISSING):
    v = _nullish(value)
    if v is NULL or v is EMPTY:
        return "" if _m(valueifnull) else valueifnull
    return v


# --------------------------------------------------------------------------
# Conversion
# --------------------------------------------------------------------------
def _nn(f):
    def g(rt, expression):
        v = expression
        if v is NULL:
            raise rt_error(94)
        return f(v)
    return g


BUILTINS["cint"] = Builtin("cint", _nn(to_integer))
BUILTINS["clng"] = Builtin("clng", _nn(to_long))
BUILTINS["clnglng"] = Builtin("clnglng", _nn(to_longlong))
BUILTINS["clngptr"] = Builtin("clngptr", _nn(to_longlong))
BUILTINS["cbyte"] = Builtin("cbyte", _nn(to_byte))
BUILTINS["cdbl"] = Builtin("cdbl", _nn(to_double))
BUILTINS["csng"] = Builtin("csng", _nn(to_single))
BUILTINS["ccur"] = Builtin("ccur", _nn(to_currency))
BUILTINS["cdec"] = Builtin("cdec", _nn(_to_decimal))
BUILTINS["cbool"] = Builtin("cbool", _nn(to_bool))
BUILTINS["cdate"] = Builtin("cdate", _nn(to_date))
BUILTINS["cvdate"] = BUILTINS["cdate"]


@builtin("cstr")
def b_cstr(rt, expression):
    return to_str(expression)


@builtin("cvar")
def b_cvar(rt, expression):
    return to_value(expression) if isinstance(expression, VBAObject) else expression


@builtin("val")
def b_val(rt, string):
    s = re.sub(r"[ \t\r\n]", "", to_str(string))
    m = re.match(r"&[Hh]([0-9A-Fa-f]+)", s)
    if m:
        v = int(m.group(1), 16)
        if v <= 0xFFFF and len(m.group(1)) <= 4 and v >= 0x8000:
            v -= 0x10000
        return float(v)
    m = re.match(r"&[Oo]?([0-7]+)", s)
    if m and s.startswith("&"):
        return float(int(m.group(1), 8))
    m = re.match(r"[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?", s)
    if not m:
        return 0.0
    return float(m.group(0).replace("d", "e").replace("D", "e"))


@builtin("str", "str$")
def b_str(rt, number):
    v = to_value(number) if isinstance(number, VBAObject) else number
    if v is NULL:
        return NULL
    if isinstance(v, str):
        n = str_to_num_or_none(v)
        if n is None:
            raise rt_error(13)
        v = n
    s = to_str(v)
    if type(v) in (int, float, Decimal, bool) and not s.startswith("-"):
        if type(v) is bool:
            return s
        return " " + s
    return s


@builtin("hex", "hex$")
def b_hex(rt, number):
    v = _nullish(number)
    if v is NULL:
        return NULL
    n = index_int(v)
    if n < 0:
        n = n & 0xFFFF if n >= -32768 else n & 0xFFFFFFFF if n >= -2147483648 else n & 0xFFFFFFFFFFFFFFFF
    return "%X" % n


@builtin("oct", "oct$")
def b_oct(rt, number):
    v = _nullish(number)
    if v is NULL:
        return NULL
    n = index_int(v)
    if n < 0:
        n = n & 0xFFFF if n >= -32768 else n & 0xFFFFFFFF
    return "%o" % n


# --------------------------------------------------------------------------
# Math
# --------------------------------------------------------------------------
def _num1(v):
    v = _nullish(v)
    if v is NULL:
        return NULL
    return to_num(v)


@builtin("int")
def b_int(rt, number):
    v = _num1(number)
    if v is NULL:
        return NULL
    if type(v) is int:
        return v
    if type(v) is Decimal:
        return Decimal(math.floor(v))
    return float(math.floor(v))


@builtin("fix")
def b_fix(rt, number):
    v = _num1(number)
    if v is NULL:
        return NULL
    if type(v) is int:
        return v
    if type(v) is Decimal:
        return Decimal(int(v))
    return float(math.trunc(v))


@builtin("abs")
def b_abs(rt, number):
    v = _num1(number)
    if v is NULL:
        return NULL
    return abs(v)


@builtin("sgn")
def b_sgn(rt, number):
    v = _num1(number)
    if v is NULL:
        return NULL
    return (v > 0) - (v < 0)


@builtin("sqr")
def b_sqr(rt, number):
    v = to_float(number)
    if v < 0:
        raise rt_error(5)
    return math.sqrt(v)


@builtin("exp")
def b_exp(rt, number):
    try:
        return math.exp(to_float(number))
    except OverflowError:
        raise rt_error(6)


@builtin("log")
def b_log(rt, number):
    v = to_float(number)
    if v <= 0:
        raise rt_error(5)
    return math.log(v)


@builtin("sin")
def b_sin(rt, number):
    return math.sin(to_float(number))


@builtin("cos")
def b_cos(rt, number):
    return math.cos(to_float(number))


@builtin("tan")
def b_tan(rt, number):
    return math.tan(to_float(number))


@builtin("atn")
def b_atn(rt, number):
    return math.atan(to_float(number))


@builtin("round")
def b_round(rt, number, numdigitsafterdecimal=MISSING):
    v = _num1(number)
    if v is NULL:
        return NULL
    d = 0 if _m(numdigitsafterdecimal) else index_int(numdigitsafterdecimal)
    if d < 0:
        raise rt_error(5)
    if type(v) is int:
        return v
    if type(v) is Decimal:
        return v.quantize(Decimal(1).scaleb(-d))
    return vba_round(v, d)


# --------------------------------------------------------------------------
# Strings
# --------------------------------------------------------------------------
def _s(v):
    """String argument with Null propagation (returns NULL)."""
    v = _nullish(v)
    if v is NULL:
        return NULL
    return to_str(v)


@builtin("len")
def b_len(rt, expression):
    v = _nullish(expression)
    if v is NULL:
        return NULL
    if type(v) is UDTValue:
        return 0
    return len(to_str(v))


@builtin("lenb")
def b_lenb(rt, expression):
    v = _nullish(expression)
    if v is NULL:
        return NULL
    return 2 * len(to_str(v))


@builtin("left", "left$", "leftb", "leftb$")
def b_left(rt, string, length):
    s = _s(string)
    n = index_int(length)
    if n < 0:
        raise rt_error(5)
    if s is NULL:
        return NULL
    return s[:n]


@builtin("right", "right$", "rightb", "rightb$")
def b_right(rt, string, length):
    s = _s(string)
    n = index_int(length)
    if n < 0:
        raise rt_error(5)
    if s is NULL:
        return NULL
    return s[len(s) - n:] if n else ""


@builtin("mid", "mid$", "midb", "midb$")
def b_mid(rt, string, start, length=MISSING):
    s = _s(string)
    st = index_int(start)
    if st < 1:
        raise rt_error(5)
    if s is NULL:
        return NULL
    if _m(length):
        return s[st - 1:]
    n = index_int(length)
    if n < 0:
        raise rt_error(5)
    return s[st - 1:st - 1 + n]


@builtin("ucase", "ucase$")
def b_ucase(rt, string):
    s = _s(string)
    return s if s is NULL else s.upper()


@builtin("lcase", "lcase$")
def b_lcase(rt, string):
    s = _s(string)
    return s if s is NULL else s.lower()


@builtin("trim", "trim$")
def b_trim(rt, string):
    s = _s(string)
    return s if s is NULL else s.strip(" ")


@builtin("ltrim", "ltrim$")
def b_ltrim(rt, string):
    s = _s(string)
    return s if s is NULL else s.lstrip(" ")


@builtin("rtrim", "rtrim$")
def b_rtrim(rt, string):
    s = _s(string)
    return s if s is NULL else s.rstrip(" ")


def _cmp_mode(rt, compare):
    if _m(compare):
        return rt.cur_text_compare()
    return index_int(compare) == 1


@builtin("instr", "instrb")
def b_instr(rt, *args):
    args = list(args)
    if len(args) < 2:
        raise rt_error(449)
    if len(args) == 2:
        start, s1, s2, compare = 1, args[0], args[1], MISSING
    else:
        start = args[0]
        s1, s2 = args[1], args[2]
        compare = args[3] if len(args) > 3 else MISSING
        start = 1 if _m(start) else index_int(start)
    if start < 1:
        raise rt_error(5)
    a = _s(s1)
    b = _s(s2)
    if a is NULL or b is NULL:
        return NULL
    if _cmp_mode(rt, compare):
        a, b = a.lower(), b.lower()
    if start > len(a):
        return 0
    if b == "":
        return start
    return a.find(b, start - 1) + 1


@builtin("instrrev")
def b_instrrev(rt, stringcheck, stringmatch, start=MISSING, compare=MISSING):
    a = _s(stringcheck)
    b = _s(stringmatch)
    if a is NULL or b is NULL:
        return NULL
    st = -1 if _m(start) else index_int(start)
    if st == 0 or st < -1:
        raise rt_error(5)
    if _cmp_mode(rt, compare):
        a, b = a.lower(), b.lower()
    if st == -1:
        st = len(a)
    if st > len(a):
        return 0
    if b == "":
        return st
    return a.rfind(b, 0, st) + 1


@builtin("replace", "replace$")
def b_replace(rt, expression, find, replace, start=MISSING, count=MISSING, compare=MISSING):
    s = _s(expression)
    if s is NULL:
        return NULL
    f = to_str(find)
    r = to_str(replace)
    st = 1 if _m(start) else index_int(start)
    cnt = -1 if _m(count) else index_int(count)
    if st < 1:
        raise rt_error(5)
    s = s[st - 1:]
    if f == "" or cnt == 0:
        return s
    if _cmp_mode(rt, compare):
        return re.sub(re.escape(f), lambda m: r, s, count=0 if cnt < 0 else cnt, flags=re.I)
    return s.replace(f, r, cnt)


@builtin("split")
def b_split(rt, expression, delimiter=MISSING, limit=MISSING, compare=MISSING):
    s = to_str(expression)
    d = " " if _m(delimiter) else to_str(delimiter)
    lim = -1 if _m(limit) else index_int(limit)
    str_ts = rt.type_spec("string")
    if s == "":
        return VBAArray([(0, -1)], str_ts, [])
    if d == "":
        return VBAArray.from_list([s], 0, str_ts)
    if _cmp_mode(rt, compare) and not _m(compare):
        parts = re.split(re.escape(d), s, maxsplit=0 if lim < 0 else max(lim - 1, 0), flags=re.I)
    else:
        parts = s.split(d, -1 if lim < 0 else max(lim - 1, 0))
    if lim == 0:
        parts = []
    return VBAArray.from_list(parts, 0, str_ts)


@builtin("join")
def b_join(rt, sourcearray, delimiter=MISSING):
    if type(sourcearray) is not VBAArray:
        raise rt_error(13)
    d = " " if _m(delimiter) else to_str(delimiter)
    return d.join(concat_str(x) for x in sourcearray.data)


@builtin("filter")
def b_filter(rt, sourcearray, match, include=MISSING, compare=MISSING):
    if type(sourcearray) is not VBAArray:
        raise rt_error(13)
    m = to_str(match)
    inc = True if _m(include) else to_bool(include)
    text = not _m(compare) and index_int(compare) == 1
    out = []
    for x in sourcearray.data:
        s = to_str(x)
        hit = (m.lower() in s.lower()) if text else (m in s)
        if hit == inc:
            out.append(s)
    return VBAArray([(0, len(out) - 1)], rt.type_spec("string"), out)


@builtin("strcomp")
def b_strcomp(rt, string1, string2, compare=MISSING):
    a = _s(string1)
    b = _s(string2)
    if a is NULL or b is NULL:
        return NULL
    if not _m(compare) and index_int(compare) == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


@builtin("strconv")
def b_strconv(rt, string, conversion, lcid=MISSING):
    s = _s(string)
    if s is NULL:
        return NULL
    c = index_int(conversion)
    if c & 3 == 1:
        s = s.upper()
    elif c & 3 == 2:
        s = s.lower()
    elif c & 3 == 3:
        out = []
        prev = " "
        for ch in s:
            out.append(ch.upper() if prev in " \t\r\n\0" else ch.lower())
            prev = ch
        s = "".join(out)
    if c & 128:
        b = s.encode("cp1252", errors="replace")
        return VBAArray.from_list(list(b), 0, rt.type_spec("byte"))
    return s


@builtin("strreverse")
def b_strreverse(rt, expression):
    return to_str(expression)[::-1]


@builtin("space", "space$")
def b_space(rt, number):
    n = index_int(number)
    if n < 0:
        raise rt_error(5)
    return " " * n


@builtin("spc")
def b_spc(rt, n):
    return " " * max(0, index_int(n))


@builtin("tab")
def b_tab(rt, n=MISSING):
    return "\t" if _m(n) else " " * max(0, index_int(n) - 1)


@builtin("string", "string$")
def b_string(rt, number, character):
    n = index_int(number)
    if n < 0:
        raise rt_error(5)
    c = _nullish(character)
    if isinstance(c, str):
        if not c:
            raise rt_error(5)
        ch = c[0]
    else:
        ch = chr(index_int(c) % 256)
    return ch * n


def _ansi_chr(n):
    if n < 128 or n > 159:
        return chr(n)
    try:
        return bytes([n]).decode("cp1252")
    except UnicodeDecodeError:
        return chr(n)


@builtin("chr", "chr$", "chrb", "chrb$")
def b_chr(rt, charcode):
    n = index_int(charcode)
    if n < 0 or n > 255:
        if -32768 <= n <= 65535:
            return chr(n & 0xFFFF)
        raise rt_error(5)
    return _ansi_chr(n)


@builtin("chrw", "chrw$")
def b_chrw(rt, charcode):
    n = index_int(charcode)
    if n < -32768 or n > 65535:
        raise rt_error(5)
    return chr(n & 0xFFFF)


@builtin("asc", "ascb")
def b_asc(rt, string):
    s = to_str(string)
    if not s:
        raise rt_error(5)
    try:
        b = s[0].encode("cp1252")
        return b[0]
    except UnicodeEncodeError:
        return ord(s[0])


@builtin("ascw")
def b_ascw(rt, string):
    s = to_str(string)
    if not s:
        raise rt_error(5)
    return ord(s[0])


# --------------------------------------------------------------------------
# Arrays & misc
# --------------------------------------------------------------------------
@builtin("lbound")
def b_lbound(rt, arrayname, dimension=MISSING):
    if type(arrayname) is not VBAArray:
        raise rt_error(13)
    return arrayname.lbound(1 if _m(dimension) else index_int(dimension))


@builtin("ubound")
def b_ubound(rt, arrayname, dimension=MISSING):
    if type(arrayname) is not VBAArray:
        raise rt_error(13)
    return arrayname.ubound(1 if _m(dimension) else index_int(dimension))


@builtin("array")
def b_array(rt, *items):
    return VBAArray.from_list(list(items), rt.cur_option_base())


@builtin("iif")
def b_iif(rt, expression, truepart, falsepart):
    from .values import truthy
    return truepart if truthy(_nullish(expression)) else falsepart


@builtin("choose")
def b_choose(rt, index, *choices):
    i = index_int(index)
    if 1 <= i <= len(choices):
        return choices[i - 1]
    return NULL


@builtin("switch")
def b_switch(rt, *args):
    from .values import truthy
    if len(args) % 2:
        raise rt_error(5)
    for i in range(0, len(args), 2):
        if truthy(_nullish(args[i])):
            return args[i + 1]
    return NULL


# --------------------------------------------------------------------------
# Date & time
# --------------------------------------------------------------------------
def _dt(v):
    v = _nullish(v)
    if v is NULL:
        return None
    return D.serial_to_dt(to_date(v))


def _now_serial():
    return D.dt_to_serial(datetime.now())


@builtin("now")
def b_now(rt):
    return VDate(round(_now_serial() * 86400) / 86400)


@builtin("date", "date$")
def b_date(rt):
    return VDate(math.floor(_now_serial()))


@builtin("time", "time$")
def b_time(rt):
    s = _now_serial()
    return VDate(round((s - math.floor(s)) * 86400) / 86400)


@builtin("timer")
def b_timer(rt):
    n = datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second + n.microsecond / 1e6


def _part(f):
    def g(rt, date):
        d = _dt(date)
        if d is None:
            return NULL
        return f(d)
    return g


BUILTINS["year"] = Builtin("year", _part(lambda d: d.year))
BUILTINS["month"] = Builtin("month", _part(lambda d: d.month))
BUILTINS["day"] = Builtin("day", _part(lambda d: d.day))
BUILTINS["hour"] = Builtin("hour", _part(lambda d: d.hour))
BUILTINS["minute"] = Builtin("minute", _part(lambda d: d.minute))
BUILTINS["second"] = Builtin("second", _part(lambda d: d.second))


def _weekday(d, firstday=1):
    sun = (d.weekday() + 1) % 7 + 1     # Sunday = 1
    if firstday == 0:
        firstday = 1
    return (sun - firstday) % 7 + 1


@builtin("weekday")
def b_weekday(rt, date, firstdayofweek=MISSING):
    d = _dt(date)
    if d is None:
        return NULL
    return _weekday(d, 1 if _m(firstdayofweek) else index_int(firstdayofweek))


@builtin("weekdayname")
def b_weekdayname(rt, weekday, abbreviate=MISSING, firstdayofweek=MISSING):
    w = index_int(weekday)
    if not 1 <= w <= 7:
        raise rt_error(5)
    fd = 1 if _m(firstdayofweek) else index_int(firstdayofweek) or 1
    name = D.DAYS[(w - 1 + fd - 1) % 7].title()
    return name[:3] if (not _m(abbreviate) and to_bool(abbreviate)) else name


@builtin("monthname")
def b_monthname(rt, month, abbreviate=MISSING):
    m = index_int(month)
    if not 1 <= m <= 12:
        raise rt_error(5)
    name = D.MONTHS[m - 1].title()
    return name[:3] if (not _m(abbreviate) and to_bool(abbreviate)) else name


@builtin("dateserial")
def b_dateserial(rt, year, month, day):
    try:
        return VDate(D.date_serial(index_int(year), index_int(month), index_int(day)))
    except (ValueError, OverflowError):
        raise rt_error(5)


@builtin("timeserial")
def b_timeserial(rt, hour, minute, second):
    return VDate(D.time_serial(index_int(hour), index_int(minute), index_int(second)))


@builtin("datevalue")
def b_datevalue(rt, date):
    v = to_date(date)
    return VDate(math.floor(v))


@builtin("timevalue")
def b_timevalue(rt, time):
    v = to_date(time)
    return VDate(v - math.floor(v))


def _add_months(d, n):
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    import calendar
    day = min(d.day, calendar.monthrange(y, m)[1])
    return d.replace(year=y, month=m, day=day)


@builtin("dateadd")
def b_dateadd(rt, interval, number, date):
    iv = to_str(interval).lower()
    n = to_float(number)
    v = to_date(date)
    d = D.serial_to_dt(v)
    try:
        if iv == "yyyy":
            r = _add_months(d, int(n) * 12)
        elif iv == "q":
            r = _add_months(d, int(n) * 3)
        elif iv == "m":
            r = _add_months(d, int(n))
        elif iv in ("y", "d", "w"):
            r = d + timedelta(days=int(n))
        elif iv == "ww":
            r = d + timedelta(weeks=int(n))
        elif iv == "h":
            r = d + timedelta(hours=int(n))
        elif iv == "n":
            r = d + timedelta(minutes=int(n))
        elif iv == "s":
            r = d + timedelta(seconds=int(n))
        else:
            raise rt_error(5)
    except OverflowError:
        raise rt_error(6)
    return VDate(D.dt_to_serial(r))


@builtin("datediff")
def b_datediff(rt, interval, date1, date2, firstdayofweek=MISSING, firstweekofyear=MISSING):
    iv = to_str(interval).lower()
    a = to_date(date1)
    b = to_date(date2)
    da, db = D.serial_to_dt(a), D.serial_to_dt(b)
    if iv == "yyyy":
        return db.year - da.year
    if iv == "q":
        return (db.year - da.year) * 4 + ((db.month - 1) // 3 - (da.month - 1) // 3)
    if iv == "m":
        return (db.year - da.year) * 12 + db.month - da.month
    if iv in ("d", "y"):
        return math.floor(b) - math.floor(a)
    if iv == "w":
        return int((math.floor(b) - math.floor(a)) / 7)
    if iv == "ww":
        fd = 1 if _m(firstdayofweek) else index_int(firstdayofweek) or 1

        def wstart(s):
            dd = D.serial_to_dt(s)
            return math.floor(s) - (_weekday(dd, fd) - 1)
        return int((wstart(b) - wstart(a)) // 7)
    if iv == "h":
        return math.floor(round(b * 24, 6)) - math.floor(round(a * 24, 6))
    if iv == "n":
        return math.floor(round(b * 1440, 4)) - math.floor(round(a * 1440, 4))
    if iv == "s":
        return round(b * 86400) - round(a * 86400)
    raise rt_error(5)


@builtin("datepart")
def b_datepart(rt, interval, date, firstdayofweek=MISSING, firstweekofyear=MISSING):
    iv = to_str(interval).lower()
    d = _dt(date)
    if d is None:
        return NULL
    fd = 1 if _m(firstdayofweek) else index_int(firstdayofweek) or 1
    if iv == "yyyy":
        return d.year
    if iv == "q":
        return (d.month - 1) // 3 + 1
    if iv == "m":
        return d.month
    if iv == "y":
        return d.timetuple().tm_yday
    if iv == "d":
        return d.day
    if iv == "w":
        return _weekday(d, fd)
    if iv == "ww":
        jan1 = d.replace(month=1, day=1)
        off = _weekday(jan1, fd) - 1
        return (d.timetuple().tm_yday - 1 + off) // 7 + 1
    if iv == "h":
        return d.hour
    if iv == "n":
        return d.minute
    if iv == "s":
        return d.second
    raise rt_error(5)


# --------------------------------------------------------------------------
# Format
# --------------------------------------------------------------------------
NAMED_NUM = {
    "general number": None, "currency": "$#,##0.00;($#,##0.00)", "fixed": "0.00",
    "standard": "#,##0.00", "percent": "0.00%", "scientific": "0.00E+00",
}
NAMED_DATE = {
    "long date": "dddd, mmmm d, yyyy", "medium date": "dd-mmm-yy", "short date": "m/d/yyyy",
    "long time": "h:mm:ss AM/PM", "medium time": "hh:mm AM/PM", "short time": "hh:mm",
}

_DATE_TOK = re.compile(r'"[^"]*"|\\.|AM/PM|am/pm|A/P|a/p|AMPM|ampm|([dDmMyYhHnNsSwWqQtTcC])\1*|.', re.S)


def _is_date_format(fmt):
    stripped = re.sub(r'"[^"]*"|\\.', "", fmt)
    return re.search(r"[dDyYhHnNsSwWqQ]|AM/PM|am/pm|A/P|[mM]", stripped) is not None and \
        not re.search(r"[0#]", stripped)


def format_date(serial, fmt):
    d = D.serial_to_dt(serial)
    toks = []
    for m in _DATE_TOK.finditer(fmt):
        t = m.group(0)
        if t.startswith('"'):
            toks.append(("lit", t[1:-1]))
        elif t.startswith("\\") and len(t) == 2:
            toks.append(("lit", t[1]))
        elif t.upper() in ("AM/PM", "A/P", "AMPM"):
            toks.append(("ampm", t))
        elif m.group(1):
            toks.append((t[0].lower(), len(t)))
        else:
            toks.append(("lit", t))
    has_ampm = any(k == "ampm" for k, _ in toks)
    # m after h or before s means minutes
    for i, (k, n) in enumerate(toks):
        if k == "m" and n <= 2:
            prev = next((toks[j][0] for j in range(i - 1, -1, -1) if toks[j][0] != "lit"), None)
            nxt = next((toks[j][0] for j in range(i + 1, len(toks)) if toks[j][0] != "lit"), None)
            if prev == "h" or nxt == "s":
                toks[i] = ("n", n)
    out = []
    for k, n in toks:
        if k == "lit":
            out.append(n)
        elif k == "d":
            if n == 1:
                out.append(str(d.day))
            elif n == 2:
                out.append(f"{d.day:02d}")
            elif n == 3:
                out.append(D.DAYS[(d.weekday() + 1) % 7][:3].title())
            elif n == 4:
                out.append(D.DAYS[(d.weekday() + 1) % 7].title())
            elif n == 5:
                out.append(format_date(serial, "m/d/yyyy"))
            else:
                out.append(format_date(serial, "dddd, mmmm d, yyyy"))
        elif k == "m":
            if n == 1:
                out.append(str(d.month))
            elif n == 2:
                out.append(f"{d.month:02d}")
            elif n == 3:
                out.append(D.MONTHS[d.month - 1][:3].title())
            else:
                out.append(D.MONTHS[d.month - 1].title())
        elif k == "y":
            if n == 1:
                out.append(str(d.timetuple().tm_yday))
            elif n == 2:
                out.append(f"{d.year % 100:02d}")
            else:
                out.append(str(d.year))
        elif k == "h":
            h = d.hour
            if has_ampm:
                h = h % 12 or 12
            out.append(str(h) if n == 1 else f"{h:02d}")
        elif k == "n":
            out.append(str(d.minute) if n == 1 else f"{d.minute:02d}")
        elif k == "s":
            out.append(str(d.second) if n == 1 else f"{d.second:02d}")
        elif k == "w":
            if n == 1:
                out.append(str(_weekday(d)))
            else:
                jan1 = d.replace(month=1, day=1)
                out.append(str((d.timetuple().tm_yday - 1 + _weekday(jan1) - 1) // 7 + 1))
        elif k == "q":
            out.append(str((d.month - 1) // 3 + 1))
        elif k == "t":
            out.append(format_date(serial, "h:mm:ss AM/PM"))
        elif k == "c":
            from .values import fmt_date_default
            out.append(fmt_date_default(serial))
        elif k == "ampm":
            pm = d.hour >= 12
            u = n.upper()
            if u in ("AM/PM", "AMPM"):
                s = "PM" if pm else "AM"
                out.append(s if n[0].isupper() else s.lower())
            else:
                s = "P" if pm else "A"
                out.append(s if n[0].isupper() else s.lower())
    return "".join(out)


def _format_number(x, fmt):
    try:
        from ..engine import format_value
    except Exception:
        format_value = None
    if format_value is not None:
        return format_value(float(x), fmt)
    return fmt_double(float(x))


def _format_string(s, fmt):
    if fmt in (">", "<"):
        return s.upper() if fmt == ">" else s.lower()
    if "@" in fmt or "&" in fmt:
        up = fmt.startswith(">")
        lo = fmt.startswith("<")
        f = fmt.lstrip("<>!")
        n = f.count("@") + f.count("&")
        if not fmt.startswith("!") and len(s) < n:
            s2 = s.rjust(n)
        else:
            s2 = s
        out = []
        i = 0
        k = 0
        for ch in f:
            if ch in "@&":
                if k < len(s2):
                    out.append(s2[k])
                elif ch == "@":
                    out.append(" ")
                k += 1
            else:
                out.append(ch)
        if k < len(s2):
            out.append(s2[k:])
        r = "".join(out)
        return r.upper() if up else r.lower() if lo else r
    return s


def vba_format(v, fmt):
    v = _nullish(v)
    if v is NULL:
        return NULL if fmt is None else ""
    if v is EMPTY:
        if fmt and fmt.lower() not in NAMED_NUM and not _is_date_format(fmt):
            pass
        else:
            return ""
    if fmt is None or fmt == "":
        return to_str(v)
    low = fmt.lower()
    if low in ("yes/no", "true/false", "on/off"):
        b = to_bool(v)
        a, c = low.split("/")
        return (a if b else c).title()
    if low == "general date":
        return to_str(to_date(v))
    if low in NAMED_DATE:
        fmt = NAMED_DATE[low]
        low = fmt.lower()
    if low in NAMED_NUM:
        f2 = NAMED_NUM[low]
        if f2 is None:
            return to_str(to_num(v) if isinstance(v, str) else v)
        fmt = f2
    if _is_date_format(fmt):
        if isinstance(v, str):
            d = D.parse_date_string(v)
            if d is None:
                n = str_to_num_or_none(v)
                if n is None:
                    return v
                d = n
            return format_date(d, fmt)
        if type(v) is bool:
            return to_str(v)
        return format_date(float(to_num(v)) if type(v) is not VDate else float(v), fmt)
    if isinstance(v, str):
        n = str_to_num_or_none(v)
        if n is None or "@" in fmt or fmt in ("<", ">"):
            return _format_string(v, fmt)
        v = n
    if type(v) is bool:
        v = -1 if v else 0
    if v is EMPTY:
        v = 0
    if type(v) is VErr:
        return to_str(v)
    return _format_number(to_num(v), fmt)


@builtin("format", "format$")
def b_format(rt, expression, format=MISSING, firstdayofweek=MISSING, firstweekofyear=MISSING):
    return vba_format(expression, None if _m(format) else to_str(format))


def _numfmt(digits, leading, parens, group, prefix="", suffix=""):
    pat = ("#,##0" if group else "0") if leading else ("#,###" if group else "#")
    if digits > 0:
        pat += "." + "0" * digits
    pos = prefix + pat + suffix
    if parens:
        return pos + ";(" + pos + ")"
    return pos


def _tri(v, default):
    if _m(v):
        return default
    i = index_int(v)
    if i == -2:
        return default
    return i != 0


@builtin("formatnumber")
def b_formatnumber(rt, expression, numdigitsafterdecimal=MISSING, includeleadingdigit=MISSING,
                   useparensfornegativenumbers=MISSING, groupdigits=MISSING):
    d = 2 if _m(numdigitsafterdecimal) or index_int(numdigitsafterdecimal) == -1 else index_int(numdigitsafterdecimal)
    f = _numfmt(d, _tri(includeleadingdigit, True), _tri(useparensfornegativenumbers, False),
                _tri(groupdigits, True))
    return _format_number(to_double(expression), f)


@builtin("formatcurrency")
def b_formatcurrency(rt, expression, numdigitsafterdecimal=MISSING, includeleadingdigit=MISSING,
                     useparensfornegativenumbers=MISSING, groupdigits=MISSING):
    d = 2 if _m(numdigitsafterdecimal) or index_int(numdigitsafterdecimal) == -1 else index_int(numdigitsafterdecimal)
    f = _numfmt(d, _tri(includeleadingdigit, True), _tri(useparensfornegativenumbers, True),
                _tri(groupdigits, True), "$")
    return _format_number(to_double(expression), f)


@builtin("formatpercent")
def b_formatpercent(rt, expression, numdigitsafterdecimal=MISSING, includeleadingdigit=MISSING,
                    useparensfornegativenumbers=MISSING, groupdigits=MISSING):
    d = 2 if _m(numdigitsafterdecimal) or index_int(numdigitsafterdecimal) == -1 else index_int(numdigitsafterdecimal)
    f = _numfmt(d, _tri(includeleadingdigit, True), _tri(useparensfornegativenumbers, False),
                _tri(groupdigits, True), "", "%")
    return _format_number(to_double(expression), f)


@builtin("formatdatetime")
def b_formatdatetime(rt, date, namedformat=MISSING):
    v = to_date(date)
    k = 0 if _m(namedformat) else index_int(namedformat)
    if k == 0:
        return to_str(v)
    return format_date(v, {1: "dddd, mmmm d, yyyy", 2: "m/d/yyyy", 3: "h:mm:ss AM/PM",
                           4: "hh:mm"}.get(k, "m/d/yyyy"))


# --------------------------------------------------------------------------
# File system
# --------------------------------------------------------------------------
@builtin("curdir", "curdir$")
def b_curdir(rt, drive=MISSING):
    return rt.cwd()


@builtin("chdir")
def b_chdir(rt, path):
    p = rt.resolve_path(to_str(path))
    if not os.path.isdir(p):
        raise rt_error(76)
    rt.cur_dir = p
    return EMPTY


@builtin("chdrive")
def b_chdrive(rt, drive):
    return EMPTY


@builtin("mkdir")
def b_mkdir(rt, path):
    p = rt.resolve_path(to_str(path))
    if os.path.exists(p):
        raise rt_error(75)
    try:
        os.mkdir(p)
    except FileNotFoundError:
        raise rt_error(76)
    return EMPTY


@builtin("rmdir")
def b_rmdir(rt, path):
    p = rt.resolve_path(to_str(path))
    if not os.path.isdir(p):
        raise rt_error(76)
    try:
        os.rmdir(p)
    except OSError:
        raise rt_error(75)
    return EMPTY


@builtin("kill")
def b_kill(rt, pathname):
    pat = rt.resolve_path(to_str(pathname))
    files = [f for f in _glob.glob(pat) if os.path.isfile(f)]
    if not files:
        raise rt_error(53)
    for f in files:
        os.remove(f)
    return EMPTY


@builtin("filecopy")
def b_filecopy(rt, source, destination):
    s = rt.resolve_path(to_str(source))
    d = rt.resolve_path(to_str(destination))
    if not os.path.isfile(s):
        raise rt_error(53)
    shutil.copyfile(s, d)
    return EMPTY


@builtin("filelen")
def b_filelen(rt, pathname):
    p = rt.resolve_path(to_str(pathname))
    if not os.path.exists(p):
        raise rt_error(53)
    return os.path.getsize(p)


@builtin("filedatetime")
def b_filedatetime(rt, pathname):
    p = rt.resolve_path(to_str(pathname))
    if not os.path.exists(p):
        raise rt_error(53)
    return VDate(D.dt_to_serial(datetime.fromtimestamp(os.path.getmtime(p))))


@builtin("getattr")
def b_getattr(rt, pathname):
    p = rt.resolve_path(to_str(pathname))
    if not os.path.exists(p):
        raise rt_error(53)
    a = 16 if os.path.isdir(p) else 0
    if not os.access(p, os.W_OK):
        a |= 1
    if os.path.basename(p).startswith("."):
        a |= 2
    return a if a else 32 if not os.path.isdir(p) else 16


@builtin("setattr")
def b_setattr(rt, pathname, attributes):
    return EMPTY


@builtin("dir", "dir$")
def b_dir(rt, pathname=MISSING, attributes=MISSING):
    if _m(pathname):
        if rt.dir_iter is None:
            raise rt_error(5)
        try:
            return next(rt.dir_iter)
        except StopIteration:
            rt.dir_iter = None
            return ""
    attrs = 0 if _m(attributes) else index_int(attributes)
    pat = rt.resolve_path(to_str(pathname))
    if pat.endswith(os.sep) or os.path.isdir(pat) and not (attrs & 16):
        pat = os.path.join(pat, "*")
    if os.path.isdir(pat) and attrs & 16:
        names = [os.path.basename(pat.rstrip(os.sep))]
    else:
        names = []
        for f in sorted(_glob.glob(pat)):
            if os.path.isdir(f):
                if attrs & 16:
                    names.append(os.path.basename(f))
            else:
                names.append(os.path.basename(f))
        if attrs & 16 and ("*" in pat or "?" in pat):
            names = [".", ".."] + names
    rt.dir_iter = iter(names)
    try:
        return next(rt.dir_iter)
    except StopIteration:
        rt.dir_iter = None
        return ""


@builtin("freefile")
def b_freefile(rt, rangenumber=MISSING):
    n = 1
    while n in rt.files:
        n += 1
    return n


@builtin("eof")
def b_eof(rt, filenumber):
    return rt.file_handle(filenumber).eof()


@builtin("lof")
def b_lof(rt, filenumber):
    return rt.file_handle(filenumber).lof()


@builtin("loc")
def b_loc(rt, filenumber):
    return rt.file_handle(filenumber).pos


@builtin("input", "input$")
def b_input(rt, number, filenumber):
    return rt.file_handle(filenumber).read_chars(index_int(number))


@builtin("macscript")
def b_macscript(rt, script):
    raise VBARuntimeErr(5, "MacScript is not supported")


# --------------------------------------------------------------------------
# Financial (basic)
# --------------------------------------------------------------------------
@builtin("pmt")
def b_pmt(rt, rate, nper, pv, fv=MISSING, type_=MISSING):
    r, n, p = to_float(rate), to_float(nper), to_float(pv)
    f = 0.0 if _m(fv) else to_float(fv)
    t = 0 if _m(type_) else index_int(type_)
    if r == 0:
        return -(p + f) / n
    q = (1 + r) ** n
    return -(r * (p * q + f)) / ((1 + r * t) * (q - 1))


@builtin("fv")
def b_fv(rt, rate, nper, pmt, pv=MISSING, type_=MISSING):
    r, n, m = to_float(rate), to_float(nper), to_float(pmt)
    p = 0.0 if _m(pv) else to_float(pv)
    t = 0 if _m(type_) else index_int(type_)
    if r == 0:
        return -(p + m * n)
    q = (1 + r) ** n
    return -(p * q + m * (1 + r * t) * (q - 1) / r)


@builtin("pv")
def b_pv(rt, rate, nper, pmt, fv=MISSING, type_=MISSING):
    r, n, m = to_float(rate), to_float(nper), to_float(pmt)
    f = 0.0 if _m(fv) else to_float(fv)
    t = 0 if _m(type_) else index_int(type_)
    if r == 0:
        return -(f + m * n)
    q = (1 + r) ** n
    return -(f + m * (1 + r * t) * (q - 1) / r) / q

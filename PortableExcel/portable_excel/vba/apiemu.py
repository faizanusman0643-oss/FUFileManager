"""Emulation of the Windows API calls commonly Declare'd by Excel VBA code
(window styles of UserForms, mouse-wheel hooks...).  Anything else raises a
clear "not supported" error."""
from __future__ import annotations

import time

from .errors import VBARuntimeErr
from .values import (EMPTY, MISSING, VBAObject, UDTValue, to_str, to_value, index_int, NOTHING)

GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_CAPTION = 0x00C00000
WS_SYSMENU = 0x00080000
WS_THICKFRAME = 0x00040000
WS_MINIMIZEBOX = 0x00020000
WS_MAXIMIZEBOX = 0x00010000
DEFAULT_FORM_STYLE = 0x84C80080      # popup | clipsiblings | caption | sysmenu | modalframe
APP_HWND = 0x00010010
WM_MOUSEWHEEL = 0x20A
HOOK_LPARAM = 0x7E000000             # fake MSLLHOOKSTRUCT address


def _i(v):
    if v is MISSING or v is EMPTY or v is NOTHING:
        return 0
    if isinstance(v, VBAObject):
        v = to_value(v)
    try:
        return index_int(v)
    except VBARuntimeErr:
        return 0


def _s(v):
    if v is MISSING or v is EMPTY:
        return ""
    try:
        return to_str(v)
    except VBARuntimeErr:
        return ""


def _set_fields(udt, **vals):
    if type(udt) is not UDTValue:
        return
    for k, v in vals.items():
        for low, disp, ts, dims in udt.tdef.fields:
            if low == k:
                try:
                    udt.fields[low] = ts.coerce(v)
                except VBARuntimeErr:
                    udt.fields[low] = v


class ApiEmulator:
    def __init__(self, rt):
        self.rt = rt
        self.next_hwnd = 0x00A10000
        self.hwnds = {}          # hwnd -> FormInstance
        self.styles = {}         # (hwnd, index) -> value
        self.hooks = {}          # handle -> (idHook, proc)
        self.next_hook = 0x5100
        self.wheel_delta = 0
        self.cursor = None

    # -- helpers ------------------------------------------------------------
    def hwnd_for(self, inst):
        h = getattr(inst, "hwnd", None)
        if not h:
            h = self.next_hwnd
            self.next_hwnd += 0x10
            inst.hwnd = h
            self.hwnds[h] = inst
            self.styles.setdefault((h, GWL_STYLE), DEFAULT_FORM_STYLE)
        return h

    def form_rect(self, inst):
        from ..forms import points_to_px
        f = inst.form
        l = points_to_px(float(f.get("left") or 0))
        t = points_to_px(float(f.get("top") or 0))
        return l, t, l + points_to_px(float(f.get("width") or 0)), t + points_to_px(float(f.get("height") or 0))

    def apply_style(self, inst):
        h = self.hwnd_for(inst)
        st = self.styles.get((h, GWL_STYLE), DEFAULT_FORM_STYLE)
        inst.window_flags = {"resizable": bool(st & WS_THICKFRAME),
                             "minimize": bool(st & WS_MINIMIZEBOX),
                             "maximize": bool(st & WS_MAXIMIZEBOX),
                             "caption": (st & WS_CAPTION) == WS_CAPTION}
        inst.apply_window()

    # -- dispatcher -------------------------------------------------------------
    def call(self, name, args):
        """Returns (result, {arg_index: new_value})."""
        n = name.lower()
        for suffix in ("a", "w"):
            if n.endswith(suffix) and hasattr(self, "f_" + n[:-1]) and not hasattr(self, "f_" + n):
                n = n[:-1]
        if n.startswith("setwindowlongptr") or n.startswith("getwindowlongptr"):
            n = n.replace("ptr", "")
        fn = getattr(self, "f_" + n, None)
        if fn is None:
            e = VBARuntimeErr(5, f"Windows API calls are not supported ({name})")
            e.fatal = True
            raise e
        r = fn(list(args))
        if isinstance(r, tuple):
            return r
        return r, {}

    # -- user32 ---------------------------------------------------------------------
    def f_findwindow(self, a):
        cls = _s(a[0]) if a else ""
        cap = _s(a[1]) if len(a) > 1 else ""
        if cls.upper() == "XLMAIN" and not cap:
            return APP_HWND
        for inst in list(self.rt.live_forms):
            if not cap or inst.form.get("caption", "") == cap:
                if cls and cls.lower() not in ("thunderdframe", "thunderxframe", "thunderframe"):
                    continue
                return self.hwnd_for(inst)
        if cls.upper() == "XLMAIN" or (cap and cap == self.rt.app.caption):
            return APP_HWND
        return 0

    def f_findwindowex(self, a):
        return 0

    def f_getwindowlong(self, a):
        h, idx = _i(a[0]), _i(a[1])
        if h in self.hwnds:
            self.hwnd_for(self.hwnds[h])
        return self.styles.get((h, idx), DEFAULT_FORM_STYLE if idx == GWL_STYLE else 0)

    def f_setwindowlong(self, a):
        h, idx, val = _i(a[0]), _i(a[1]), _i(a[2])
        old = self.f_getwindowlong(a[:2])
        self.styles[(h, idx)] = val & 0xFFFFFFFF if val < 0 else val
        if idx == GWL_STYLE and h in self.hwnds:
            self.apply_style(self.hwnds[h])
        return old

    def f_drawmenubar(self, a):
        return 1

    def f_showwindow(self, a):
        h, cmd = _i(a[0]), _i(a[1])
        inst = self.hwnds.get(h)
        if inst is not None:
            state = {3: "maximized", 9: "normal", 1: "normal", 5: "normal", 6: "minimized",
                     2: "minimized", 7: "minimized", 11: "minimized"}.get(cmd)
            if cmd == 0:
                inst.hide()
            elif state:
                inst.set_window_state(state)
        return 1

    def f_iszoomed(self, a):
        inst = self.hwnds.get(_i(a[0]))
        return 1 if inst is not None and getattr(inst, "window_state", "normal") == "maximized" else 0

    def f_isiconic(self, a):
        inst = self.hwnds.get(_i(a[0]))
        return 1 if inst is not None and getattr(inst, "window_state", "normal") == "minimized" else 0

    def f_iswindow(self, a):
        h = _i(a[0])
        return 1 if h == APP_HWND or (h in self.hwnds and self.hwnds[h].loaded) else 0

    def f_iswindowvisible(self, a):
        inst = self.hwnds.get(_i(a[0]))
        return 1 if inst is not None and inst.visible else 0

    def f_setwindowpos(self, a):
        return 1

    def f_movewindow(self, a):
        return 1

    def f_setforegroundwindow(self, a):
        return 1

    def f_bringwindowtotop(self, a):
        return 1

    def f_enablewindow(self, a):
        return 0

    def f_setfocus(self, a):
        return 0

    def f_getparent(self, a):
        return 0

    def f_setparent(self, a):
        return 0

    def f_getforegroundwindow(self, a):
        for inst in reversed(self.rt.live_forms):
            if inst.visible:
                return self.hwnd_for(inst)
        return APP_HWND

    f_getactivewindow = f_getforegroundwindow

    def f_getdesktopwindow(self, a):
        return 0x10014

    def f_setwindowshookex(self, a):
        idhook, lpfn = _i(a[0]), _i(a[1])
        proc = self.rt.addressof_procs.get(lpfn)
        h = self.next_hook
        self.next_hook += 1
        self.hooks[h] = (idhook, proc)
        return h

    def f_unhookwindowshookex(self, a):
        self.hooks.pop(_i(a[0]), None)
        return 1

    def f_callnexthookex(self, a):
        return 0

    def f_getcursorpos(self, a):
        x, y = self._cursor()
        if a:
            _set_fields(a[0], x=x, y=y)
        return 1, {}

    def _cursor(self):
        if self.cursor is not None:
            return self.cursor
        h = getattr(self.rt.host, "get_cursor_pos", None)
        if h is not None:
            try:
                p = h()
                if p:
                    return int(p[0]), int(p[1])
            except Exception:
                pass
        return 0, 0

    def f_getwindowrect(self, a):
        h = _i(a[0])
        inst = self.hwnds.get(h)
        l, t, r, b = self.form_rect(inst) if inst is not None else (0, 0, 1920, 1080)
        if len(a) > 1:
            _set_fields(a[1], left=l, top=t, right=r, bottom=b)
        return 1

    def f_getclientrect(self, a):
        h = _i(a[0])
        inst = self.hwnds.get(h)
        l, t, r, b = self.form_rect(inst) if inst is not None else (0, 0, 1920, 1080)
        if len(a) > 1:
            _set_fields(a[1], left=0, top=0, right=r - l, bottom=b - t)
        return 1

    def f_screentoclient(self, a):
        return 1

    f_clienttoscreen = f_screentoclient

    def f_windowfrompoint(self, a):
        return self.f_getforegroundwindow(a)

    def f_getsystemmetrics(self, a):
        return {0: 1920, 1: 1080, 4: 23, 5: 1, 6: 1, 7: 3, 8: 3, 16: 1920, 17: 1040,
                32: 4, 33: 4}.get(_i(a[0]), 0)

    def f_getdc(self, a):
        return 0x1D0

    def f_releasedc(self, a):
        return 1

    def f_getdevicecaps(self, a):
        return {88: 96, 90: 96, 8: 1920, 10: 1080}.get(_i(a[1]) if len(a) > 1 else 0, 0)

    def f_getasynckeystate(self, a):
        return 0

    f_getkeystate = f_getasynckeystate

    def f_sendmessage(self, a):
        return 0

    f_postmessage = f_sendmessage

    def f_messagebox(self, a):
        text = _s(a[1]) if len(a) > 1 else ""
        cap = _s(a[2]) if len(a) > 2 else ""
        return self.rt.host.msgbox(text, _i(a[3]) if len(a) > 3 else 0, cap)

    def f_messagebeep(self, a):
        return 1

    def f_lockwindowupdate(self, a):
        return 1

    def f_extracticon(self, a):
        return 0

    def f_sendmessagea(self, a):
        return 0

    def f_seticon(self, a):
        return 0

    # -- kernel32 ---------------------------------------------------------------
    def f_rtlmovememory(self, a):
        # Only the mouse-hook pattern is meaningful: CopyMemory var, lParam + 8, 4
        if len(a) >= 2 and _i(a[1]) == HOOK_LPARAM + 8:
            d = self.wheel_delta
            return 0, {0: (d & 0xFFFF) * 65536 if d >= 0 else ((d & 0xFFFF) << 16) - (1 << 32)}
        return 0

    f_copymemory = f_rtlmovememory
    f_movememory = f_rtlmovememory

    def f_sleep(self, a):
        ms = max(0, _i(a[0]))
        end = time.time() + min(ms, 10000) / 1000.0
        while time.time() < end:
            self.rt.host.process_events()
            self.rt.check_stop()
            time.sleep(min(0.02, max(0.0, end - time.time())))
        return 0

    def f_gettickcount(self, a):
        return int(time.monotonic() * 1000) & 0x7FFFFFFF

    f_gettickcount64 = f_gettickcount
    f_timegettime = f_gettickcount

    def f_getcurrentthreadid(self, a):
        return 1

    def f_getcurrentprocessid(self, a):
        return 1

    def f_getmodulehandle(self, a):
        return 0x400000

    def f_getlasterror(self, a):
        return 0

    def f_queryperformancefrequency(self, a):
        return 1, {0: 1000000}

    def f_queryperformancecounter(self, a):
        return 1, {0: int(time.perf_counter() * 1000000)}

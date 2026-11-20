"""UserForm support: form instances, controls and the runtime side of the
runtime <-> GUI bridge described in portable_excel/forms.py."""
from __future__ import annotations

import copy

from ..forms import normalize_form, new_control, unique_control_name, PROGIDS, CONTROL_TYPES
from .errors import rt_error, VBARuntimeErr, VBAError, VBAEnd
from .values import (VBAObject, VBAArray, EMPTY, NULL, NOTHING, MISSING, to_str, to_bool,
                     to_value, index_int, to_double, to_num, str_to_num_or_none)
from .objects import bind, index_value

LIST_TYPES = ("ComboBox", "ListBox")
CLICK_ON_CHANGE = ("ComboBox", "ListBox", "OptionButton", "CheckBox", "ToggleButton")
BOOL_TYPES = ("CheckBox", "OptionButton", "ToggleButton")
NUM_TYPES = ("SpinButton", "ScrollBar", "MultiPage", "TabStrip")


def _to_hex(v):
    from .excel import vba_color_to_hex
    n = index_int(to_value(v))
    if n < 0:   # system colours (&H8000000F&): use sensible defaults
        return None
    return vba_color_to_hex(n)


def _from_hex(h, default):
    from .excel import hex_to_vba_color
    return hex_to_vba_color(h, default)


# VBA property name (lower) -> dict key
SIMPLE_PROPS = {
    "name": "name", "caption": "caption", "left": "left", "top": "top", "width": "width",
    "height": "height", "visible": "visible", "enabled": "enabled", "locked": "locked",
    "tabindex": "tab_index", "tabstop": "tab_stop", "tag": "tag",
    "controltiptext": "control_tip_text", "multiline": "multi_line",
    "passwordchar": "password_char", "maxlength": "max_length", "scrollbars": "scroll_bars",
    "groupname": "group_name", "columncount": "column_count", "columnwidths": "column_widths",
    "boundcolumn": "bound_column", "multiselect": "multi_select", "style": "style",
    "liststyle": "list_style", "min": "min", "max": "max", "smallchange": "small_change",
    "largechange": "large_change", "orientation": "orientation", "default": "default",
    "cancel": "cancel", "accelerator": "accelerator", "controlsource": "control_source",
    "wordwrap": "word_wrap", "autosize": "auto_size", "textalign": "text_align",
    "backstyle": "back_style", "picturesizemode": "picture_size_mode",
    "matchrequired": "match_required",
}
BOOL_PROPS = {"visible", "enabled", "locked", "tab_stop", "multi_line", "default", "cancel",
              "word_wrap", "auto_size", "match_required"}
INT_PROPS = {"tab_index", "max_length", "scroll_bars", "column_count", "bound_column",
             "multi_select", "style", "list_style", "orientation", "text_align", "back_style",
             "picture_size_mode"}
FLOAT_PROPS = {"left", "top", "width", "height"}
STR_PROPS = {"name", "caption", "tag", "control_tip_text", "password_char", "group_name",
             "column_widths", "accelerator"}


def _conv_prop(key, v):
    v = to_value(v) if isinstance(v, VBAObject) else v
    if key in BOOL_PROPS:
        return bool(to_bool(v))
    if key in INT_PROPS:
        return index_int(v)
    if key in FLOAT_PROPS:
        return float(to_double(v))
    if key in STR_PROPS:
        return to_str(v)
    if key in ("min", "max", "small_change", "large_change"):
        return index_int(v)
    if key == "control_source":
        s = to_str(v)
        return s or None
    return v


def _py(v):
    """VBA value -> plain python for the GUI dicts."""
    if isinstance(v, VBAObject):
        v = to_value(v)
    if v is EMPTY or v is MISSING:
        return None
    if v is NULL:
        return None
    if isinstance(v, (bool, int, float, str)) or v is None:
        return v
    return to_str(v)


def _vba(v):
    if v is None:
        return EMPTY
    return v


class FontProxy(VBAObject):
    vba_typename = "Font"
    KEYS = {"bold": "bold", "italic": "italic", "underline": "underline", "size": "font_size",
            "name": "font_name"}

    def __init__(self, owner):
        self.owner = owner      # ControlObj or FormInstance

    def vba_get(self, name, args, named):
        k = self.KEYS.get(name)
        if k is None:
            if name in ("strikethrough", "weight", "charset"):
                return False if name == "strikethrough" else 400
            raise rt_error(438)
        v = self.owner.d.get(k)
        if v is None:
            if k == "font_size":
                return self.owner.font_default("font_size")
            if k == "font_name":
                return self.owner.font_default("font_name")
            return False
        return v

    def vba_let(self, name, args, named, value, is_set=False):
        k = self.KEYS.get(name)
        if k is None:
            if name in ("strikethrough", "weight", "charset"):
                return
            raise rt_error(438)
        if k in ("bold", "italic", "underline"):
            v = bool(to_bool(value))
        elif k == "font_size":
            v = float(to_double(value))
        else:
            v = to_str(value)
        self.owner.set_key(k, v)


class ControlsColl(VBAObject):
    vba_typename = "Controls"

    def __init__(self, inst, container=None):
        self.inst = inst
        self.container = container   # name of Frame/page, None = all

    def _list(self):
        ctrls = self.inst.form["controls"]
        if self.container is None:
            return ctrls
        return [c for c in ctrls if (c.get("parent") or "").lower() == self.container.lower()]

    def g_count(self, a, n):
        return len(self._list())

    def g_item(self, args, named):
        ix = to_value(args[0]) if args else MISSING
        if isinstance(ix, str):
            c = self.inst.find_ctrl(ix)
            if c is None:
                raise VBARuntimeErr(-2147024809, "Could not find the specified object.")
            return self.inst.control(c["name"])
        i = index_int(ix)
        lst = self._list()
        if 0 <= i < len(lst):
            return self.inst.control(lst[i]["name"])
        raise VBARuntimeErr(-2147024809, "Could not find the specified object.")

    g__default = g_item

    def vba_iter(self):
        return [self.inst.control(c["name"]) for c in self._list()]

    def g_add(self, args, named):
        progid, name, visible = bind(args, named, ["bstrprogid", "name", "visible"], 1)
        pid = to_str(progid).strip().lower()
        if pid.startswith("msforms."):
            pid = "forms." + pid[8:]
        if not pid.endswith(".1") and pid.startswith("forms."):
            pid += ".1"
        ctype = PROGIDS.get(pid)
        if ctype is None:
            ctype = next((t for t in CONTROL_TYPES if t.lower() == pid), None)
        if ctype is None:
            raise VBARuntimeErr(-2147221005, f"Invalid class string: {to_str(progid)}")
        nm = to_str(name) if name is not MISSING and to_str(name) else \
            unique_control_name(self.inst.form, ctype)
        if self.inst.find_ctrl(nm) is not None:
            raise VBARuntimeErr(-2147024809, "Ambiguous name")
        d = new_control(ctype, nm)
        if self.container is not None:
            d["parent"] = self.container
        if visible is not MISSING:
            d["visible"] = bool(to_bool(visible))
        self.inst.form["controls"].append(d)
        self.inst._ctrl_objs.pop(nm.lower(), None)
        if self.inst.view is not None:
            self.inst.view.add_control(d)
        return self.inst.control(nm)

    def g_remove(self, args, named):
        ix = to_value(args[0])
        if isinstance(ix, str):
            c = self.inst.find_ctrl(ix)
        else:
            lst = self._list()
            i = index_int(ix)
            c = lst[i] if 0 <= i < len(lst) else None
        if c is None:
            raise VBARuntimeErr(-2147024809, "Could not find the specified object.")
        self.inst.form["controls"].remove(c)
        self.inst._ctrl_objs.pop(c["name"].lower(), None)
        if self.inst.view is not None:
            self.inst.view.remove_control(c["name"])
        return EMPTY

    def g_clear(self, a, n):
        for c in list(self._list()):
            self.g_remove([c["name"]], {})
        return EMPTY


class PagesColl(VBAObject):
    vba_typename = "Pages"

    def __init__(self, ctrl, key="pages"):
        self.ctrl = ctrl
        self.key = key

    def _pages(self):
        return self.ctrl.d.setdefault(self.key, [])

    def g_count(self, a, n):
        return len(self._pages())

    def g_item(self, args, named):
        ix = to_value(args[0])
        pages = self._pages()
        if isinstance(ix, str):
            for i, p in enumerate(pages):
                if p.get("name", "").lower() == ix.lower():
                    return PageObj(self, i)
            raise rt_error(5)
        i = index_int(ix)
        if 0 <= i < len(pages):
            return PageObj(self, i)
        raise rt_error(5)

    g__default = g_item

    def vba_iter(self):
        return [PageObj(self, i) for i in range(len(self._pages()))]

    def g_add(self, args, named):
        name, caption, index = bind(args, named, ["bstrname", "bstrcaption", "lindex"])
        pages = self._pages()
        nm = to_str(name) if name is not MISSING else f"Page{len(pages) + 1}"
        cap = to_str(caption) if caption is not MISSING else nm
        p = {"name": nm, "caption": cap}
        if index is MISSING:
            pages.append(p)
        else:
            pages.insert(index_int(index), p)
        self.ctrl.changed(self.key, copy.deepcopy(pages))
        return PageObj(self, pages.index(p))

    def g_remove(self, args, named):
        pages = self._pages()
        pages.pop(index_int(to_value(args[0])))
        self.ctrl.changed(self.key, copy.deepcopy(pages))
        return EMPTY


class PageObj(VBAObject):
    vba_typename = "Page"

    def __init__(self, coll, i):
        self.coll = coll
        self.i = i

    def vba_same(self, other):
        return isinstance(other, PageObj) and other.coll.ctrl is self.coll.ctrl and other.i == self.i

    def vba_get(self, name, args, named):
        p = self.coll._pages()[self.i]
        if name in ("caption", ""):
            return p.get("caption", "")
        if name == "name":
            return p.get("name", "")
        if name == "index":
            return self.i
        if name == "controls":
            c = ControlsColl(self.coll.ctrl.inst, p.get("name"))
            return c.g_item(args, named) if args else c
        if name in ("visible", "enabled"):
            return p.get(name, True)
        if name == "tag":
            return p.get("tag", "")
        raise rt_error(438)

    def vba_let(self, name, args, named, value, is_set=False):
        p = self.coll._pages()[self.i]
        if name in ("caption", ""):
            p["caption"] = to_str(value)
        elif name == "name":
            p["name"] = to_str(value)
        elif name in ("visible", "enabled"):
            p[name] = bool(to_bool(value))
        elif name == "tag":
            p["tag"] = to_str(value)
        else:
            raise rt_error(438)
        self.coll.ctrl.changed(self.coll.key, copy.deepcopy(self.coll._pages()))


# --------------------------------------------------------------------------
# Controls
# --------------------------------------------------------------------------
class ControlObj(VBAObject):
    def __init__(self, inst, d):
        self.inst = inst
        self.d = d

    @property
    def vba_typename(self):
        return self.d["type"]

    @property
    def ctype(self):
        return self.d["type"]

    @property
    def name(self):
        return self.d["name"]

    def vba_same(self, other):
        return isinstance(other, ControlObj) and other.d is self.d

    def font_default(self, key):
        return self.inst.form.get(key) or ("Tahoma" if key == "font_name" else 8.0)

    # state helpers -------------------------------------------------------
    def changed(self, key, value):
        """Notify the GUI of a property change made by VBA code."""
        v = self.inst.view
        if v is not None and not self.inst._suppress_echo:
            v.set_prop(self.d["name"], key, value)

    def set_key(self, key, value):
        self.d[key] = value
        self.changed(key, value)

    # list helpers ----------------------------------------------------------
    def rows(self):
        lst = self.d.get("list")
        if lst is None:
            lst = self.d["list"] = []
        return lst

    def ncols(self):
        n = max(1, int(self.d.get("column_count") or 1))
        for r in self.rows():
            if isinstance(r, list):
                n = max(n, len(r))
        return n

    def cell(self, i, j):
        r = self.rows()[i]
        if isinstance(r, list):
            return r[j] if j < len(r) else None
        return r if j == 0 else None

    def set_cell(self, i, j, v):
        rows = self.rows()
        r = rows[i]
        if not isinstance(r, list):
            r = [r]
        while len(r) <= j:
            r.append(None)
        r[j] = v
        rows[i] = r if len(r) > 1 else r[0]

    def bound_col(self):
        return max(1, int(self.d.get("bound_column") or 1)) - 1

    def _sync_from_index(self, fire=True):
        i = self.d.get("list_index", -1)
        rows = self.rows()
        old = self.d.get("value")
        if 0 <= i < len(rows):
            val = self.cell(i, self.bound_col())
            text = self.cell(i, 0)
            self.d["value"] = val
            self.d["text"] = "" if text is None else to_str(_vba(text))
        else:
            self.d["value"] = "" if self.ctype == "ComboBox" else None
            self.d["text"] = ""
        if self.ctype == "ListBox" and int(self.d.get("multi_select") or 0) == 0:
            sel = [False] * len(rows)
            if 0 <= i < len(rows):
                sel[i] = True
            self.d["selected"] = sel
        self.changed("list_index", i)
        self.changed("value", self.d["value"])
        self.changed("text", self.d["text"])
        if fire and old != self.d["value"]:
            self.inst.code_changed(self)

    def _list_changed(self):
        self.d["selected"] = [False] * len(self.rows())
        self.changed("list", copy.deepcopy(self.rows()))

    # value ------------------------------------------------------------------
    def get_value(self):
        t = self.ctype
        d = self.d
        if t in BOOL_TYPES:
            v = d.get("value")
            return NULL if v is None else bool(v)
        if t in NUM_TYPES:
            return int(d.get("value") or 0)
        if t == "ListBox":
            if int(d.get("multi_select") or 0) != 0:
                return NULL
            return _vba(d.get("value")) if d.get("list_index", -1) >= 0 else NULL
        if t == "ComboBox":
            v = d.get("value")
            return "" if v is None else v
        if t == "TextBox":
            v = d.get("value")
            return "" if v is None else to_str(v)
        if t == "Label":
            return d.get("caption", "")
        if t == "CommandButton":
            return False
        if t == "Frame":
            return d.get("caption", "")
        return _vba(d.get("value"))

    def set_value(self, v, fire=True):
        t = self.ctype
        d = self.d
        if isinstance(v, VBAObject):
            v = to_value(v)
        old = d.get("value")
        if t in BOOL_TYPES:
            nv = None if v is NULL else bool(to_bool(v))
            d["value"] = nv
            self.changed("value", nv)
            if t == "OptionButton" and nv:
                self.inst.clear_option_group(self)
        elif t in NUM_TYPES:
            n = index_int(v)
            if t in ("SpinButton", "ScrollBar"):
                lo, hi = int(d.get("min", 0)), int(d.get("max", 100))
                if lo <= hi:
                    n = max(lo, min(hi, n))
                else:
                    n = max(hi, min(lo, n))
            elif t in ("MultiPage", "TabStrip"):
                pages = d.get("pages" if t == "MultiPage" else "tabs") or []
                if not (0 <= n < max(1, len(pages))):
                    raise rt_error(380, "Invalid property value.")
            d["value"] = n
            self.changed("value", n)
        elif t in LIST_TYPES:
            if v is NULL or v is EMPTY:
                d["list_index"] = -1
                self._sync_from_index(fire)
                return
            s = v
            bc = self.bound_col()
            idx = -1
            for i in range(len(self.rows())):
                cv = self.cell(i, bc)
                if cv == s or (cv is not None and to_str(_vba(cv)) == to_str(s)):
                    idx = i
                    break
            if idx >= 0:
                d["list_index"] = idx
                self._sync_from_index(fire)
                return
            if t == "ListBox":
                raise rt_error(380, "Could not set the Value property. Invalid property value.")
            d["list_index"] = -1
            d["value"] = to_str(s)
            d["text"] = to_str(s)
            self.changed("list_index", -1)
            self.changed("value", d["value"])
            self.changed("text", d["text"])
        elif t == "Label":
            d["caption"] = to_str(v)
            self.changed("caption", d["caption"])
            return
        else:
            s = to_str(v) if v is not EMPTY else ""
            ml = int(d.get("max_length") or 0)
            if ml and t == "TextBox":
                s = s[:ml]
            d["value"] = s
            d["text"] = s
            self.changed("value", s)
            self.changed("text", s)
        if fire and old != d.get("value"):
            self.inst.code_changed(self)

    def get_text(self):
        t = self.ctype
        if t in ("TextBox",):
            return self.get_value()
        if t in LIST_TYPES:
            if t == "ComboBox":
                v = self.d.get("text")
                return to_str(self.d.get("value") or "") if v in (None, "") else v
            return self.d.get("text") or ""
        if t in BOOL_TYPES:
            v = self.d.get("value")
            return "" if v is None else ("True" if v else "False")
        return to_str(self.get_value())

    # VBA members -----------------------------------------------------------
    def vba_get(self, name, args, named):
        d = self.d
        if name in ("", "value"):
            v = self.get_value()
            return index_value(v, list(args)) if args else v
        if name == "text":
            return self.get_text()
        m = getattr(self, "g_" + name, None)
        if m is not None:
            return m(args, named)
        key = SIMPLE_PROPS.get(name)
        if key is not None:
            v = d.get(key)
            if v is None:
                if key in ("left", "top", "width", "height"):
                    return 0.0
                if key in STR_PROPS:
                    return ""
                if key == "control_source":
                    return ""
            return v
        if name == "backcolor":
            return _from_hex(d.get("back_color"), 0xFFFFFF if self.ctype in ("TextBox", "ListBox", "ComboBox") else 0x8000000F & 0x7FFFFFFF)
        if name == "forecolor":
            return _from_hex(d.get("fore_color"), 0)
        if name == "rowsource":
            return d.get("row_source") or ""
        raise rt_error(438)

    def vba_let(self, name, args, named, value, is_set=False):
        d = self.d
        if name in ("", "value"):
            if args:
                raise rt_error(438)
            self.set_value(value)
            return
        if name == "text":
            if self.ctype in LIST_TYPES:
                self.set_value(to_str(value))
            else:
                self.set_value(value)
            return
        m = getattr(self, "s_" + name, None)
        if m is not None:
            m(args, named, value)
            return
        key = SIMPLE_PROPS.get(name)
        if key is not None:
            if key == "name":
                nm = to_str(value)
                old = d["name"]
                self.inst._ctrl_objs.pop(old.lower(), None)
                d["name"] = nm
                self.inst._ctrl_objs[nm.lower()] = self
                v = self.inst.view
                if v is not None:
                    v.set_prop(old, "name", nm)
                return
            self.set_key(key, _conv_prop(key, value))
            if key == "control_source":
                self.inst.load_control_source(self)
            if key == "column_count":
                self.changed("list", copy.deepcopy(self.rows()))
            return
        if name == "backcolor":
            self.set_key("back_color", _to_hex(value))
            return
        if name == "forecolor":
            self.set_key("fore_color", _to_hex(value))
            return
        if name == "rowsource":
            s = to_str(value)
            d["row_source"] = s or None
            self.changed("row_source", d["row_source"])
            self.inst.load_row_source(self)
            return
        raise rt_error(438)

    # methods / complex props -------------------------------------------------
    def g_font(self, a, n):
        return FontProxy(self)

    def g_insidewidth(self, a, n):
        return max(0.0, float(self.d.get("width") or 0) - 4)

    def g_insideheight(self, a, n):
        cap = 10 if self.ctype == "Frame" and self.d.get("caption") else 0
        return max(0.0, float(self.d.get("height") or 0) - 4 - cap)

    def g_scrollheight(self, a, n):
        return float(self.d.get("scroll_height", 0))

    def s_scrollheight(self, a, n, v):
        self.set_key("scroll_height", float(to_double(v)))

    def g_scrollwidth(self, a, n):
        return float(self.d.get("scroll_width", 0))

    def s_scrollwidth(self, a, n, v):
        self.set_key("scroll_width", float(to_double(v)))

    def g_scrolltop(self, a, n):
        return float(self.d.get("scroll_top", 0))

    def s_scrolltop(self, a, n, v):
        self.set_key("scroll_top", float(to_double(v)))

    def g_scrollleft(self, a, n):
        return float(self.d.get("scroll_left", 0))

    def s_scrollleft(self, a, n, v):
        self.set_key("scroll_left", float(to_double(v)))

    def g_scrollbars(self, a, n):
        return int(self.d.get("scroll_bars", 0) or 0)

    def s_scrollbars(self, a, n, v):
        self.set_key("scroll_bars", index_int(v))

    def g_specialeffect(self, a, n):
        return int(self.d.get("special_effect", 0))

    def s_specialeffect(self, a, n, v):
        self.set_key("special_effect", index_int(v))

    def g_borderstyle(self, a, n):
        return int(self.d.get("border_style", 0))

    def s_borderstyle(self, a, n, v):
        self.set_key("border_style", index_int(v))

    def g_bordercolor(self, a, n):
        return _from_hex(self.d.get("border_color"), 0)

    def s_bordercolor(self, a, n, v):
        self.set_key("border_color", _to_hex(v))

    def g_parent(self, a, n):
        p = self.d.get("parent")
        if p:
            c = self.inst.find_ctrl(p)
            if c is not None:
                return self.inst.control(c["name"])
        return self.inst

    def g_object(self, a, n):
        return self

    def g_setfocus(self, a, n):
        self.inst.active_control = self.d["name"]
        v = self.inst.view
        if v is not None:
            v.set_focus(self.d["name"])
        return EMPTY

    def g_zorder(self, a, n):
        return EMPTY

    def g_move(self, args, named):
        l, t, w, h = bind(args, named, ["left", "top", "width", "height"])
        for k, v in (("left", l), ("top", t), ("width", w), ("height", h)):
            if v is not MISSING:
                self.set_key(k, float(to_double(v)))
        return EMPTY

    def g_controls(self, args, named):
        c = ControlsColl(self.inst, self.d["name"])
        return c.g_item(args, named) if args else c

    def g_pages(self, args, named):
        c = PagesColl(self, "pages")
        return c.g_item(args, named) if args else c

    def g_tabs(self, args, named):
        c = PagesColl(self, "tabs")
        return c.g_item(args, named) if args else c

    def g_selectedpage(self, a, n):
        return PagesColl(self).g_item([int(self.d.get("value") or 0)], {})

    # TextBox selection
    def g_selstart(self, a, n):
        return int(self.d.get("sel_start", len(to_str(self.get_text()))))

    def s_selstart(self, a, n, v):
        self.d["sel_start"] = index_int(v)
        self.d["sel_length"] = 0

    def g_sellength(self, a, n):
        return int(self.d.get("sel_length", 0))

    def s_sellength(self, a, n, v):
        self.d["sel_length"] = index_int(v)

    def g_seltext(self, a, n):
        s = to_str(self.get_text())
        st = self.g_selstart(a, n)
        return s[st:st + self.g_sellength(a, n)]

    def s_seltext(self, a, n, v):
        s = to_str(self.get_text())
        st = self.g_selstart(a, n)
        ln = self.g_sellength(a, n)
        rep = to_str(v)
        self.set_value(s[:st] + rep + s[st + ln:])
        self.d["sel_start"] = st + len(rep)
        self.d["sel_length"] = 0

    def g_linecount(self, a, n):
        return to_str(self.get_text()).count("\n") + 1

    # List controls
    def _need_list(self):
        if self.ctype not in LIST_TYPES:
            raise rt_error(438)

    def g_additem(self, args, named):
        self._need_list()
        item, index = bind(args, named, ["pvargitem", "pvargindex"])
        if self.d.get("row_source"):
            raise rt_error(70, "Permission denied")
        rows = self.rows()
        v = None if item is MISSING else _py(item)
        ncols = int(self.d.get("column_count") or 1)
        row = v if ncols <= 1 else [v] + [None] * (ncols - 1)
        if index is MISSING:
            rows.append(row)
        else:
            i = index_int(index)
            if i < 0 or i > len(rows):
                raise rt_error(-2147024809, "Invalid argument.")
            rows.insert(i, row)
            li = self.d.get("list_index", -1)
            if li >= i >= 0:
                self.d["list_index"] = li + 1
        sel = self.d.setdefault("selected", [])
        while len(sel) < len(rows):
            sel.append(False)
        self.changed("list", copy.deepcopy(rows))
        return EMPTY

    def g_removeitem(self, args, named):
        self._need_list()
        i = index_int(to_value(args[0]))
        rows = self.rows()
        if not 0 <= i < len(rows):
            raise rt_error(-2147024809, "Invalid argument.")
        rows.pop(i)
        sel = self.d.get("selected") or []
        if i < len(sel):
            sel.pop(i)
        li = self.d.get("list_index", -1)
        self.changed("list", copy.deepcopy(rows))
        if li == i:
            self.d["list_index"] = -1
            self._sync_from_index()
        elif li > i:
            self.d["list_index"] = li - 1
            self.changed("list_index", li - 1)
        return EMPTY

    def g_clear(self, args, named):
        if self.ctype not in LIST_TYPES:
            if self.ctype == "TextBox":
                self.set_value("")
                return EMPTY
            raise rt_error(438)
        if self.d.get("row_source"):
            raise rt_error(70, "Permission denied")
        self.d["list"] = []
        self.d["selected"] = []
        self.changed("list", [])
        if self.d.get("list_index", -1) != -1 or self.d.get("value") not in (None, ""):
            self.d["list_index"] = -1
            self._sync_from_index()
        return EMPTY

    def g_listcount(self, a, n):
        self._need_list()
        return len(self.rows())

    def g_listindex(self, a, n):
        self._need_list()
        return int(self.d.get("list_index", -1))

    def s_listindex(self, a, n, v):
        self._need_list()
        i = index_int(v)
        if i < -1 or i >= len(self.rows()):
            raise rt_error(380, "Could not set the ListIndex property. Invalid property value.")
        if i == self.d.get("list_index", -1) and self.ctype == "ListBox":
            return
        self.d["list_index"] = i
        self._sync_from_index()

    def g_list(self, args, named):
        self._need_list()
        rows = self.rows()
        if args:
            i = index_int(to_value(args[0]))
            j = index_int(to_value(args[1])) if len(args) > 1 and args[1] is not MISSING else 0
            if not 0 <= i < len(rows):
                raise rt_error(381, "Could not get the List property. Invalid property array index.")
            return _vba(self.cell(i, j))
        if not rows:
            return EMPTY
        nc = self.ncols()
        data = [_vba(self.cell(i, j)) if self.cell(i, j) is not None else NULL
                for i in range(len(rows)) for j in range(nc)]
        return VBAArray([(0, len(rows) - 1), (0, nc - 1)], None, data)

    def s_list(self, args, named, value):
        self._need_list()
        if args:
            i = index_int(to_value(args[0]))
            j = index_int(to_value(args[1])) if len(args) > 1 else 0
            rows = self.rows()
            if not 0 <= i < len(rows):
                raise rt_error(381, "Could not set the List property. Invalid property array index.")
            self.set_cell(i, j, _py(value))
            self.changed("list", copy.deepcopy(rows))
            return
        if isinstance(value, VBAObject):
            value = to_value(value)
        if type(value) is not VBAArray:
            if value is EMPTY or value is NULL:
                self.d["list"] = []
                self._list_changed()
                return
            raise rt_error(13)
        if value.ndims == 1:
            rows = [_py(x) for x in value.data]
        else:
            rows = []
            for r in value.to_rows():
                r = [_py(x) for x in r]
                rows.append(r if len(r) > 1 else r[0])
            if value.ndims == 2 and value.dims[1][1] - value.dims[1][0] + 1 > int(self.d.get("column_count") or 1):
                pass
        self.d["list"] = rows
        self.d["list_index"] = -1
        self._list_changed()

    def g_column(self, args, named):
        self._need_list()
        rows = self.rows()
        if not args:
            nc = self.ncols()
            data = [_vba(self.cell(i, j)) for j in range(nc) for i in range(len(rows))]
            return VBAArray([(0, nc - 1), (0, len(rows) - 1)], None, data) if rows else EMPTY
        j = index_int(to_value(args[0]))
        if len(args) > 1 and args[1] is not MISSING:
            i = index_int(to_value(args[1]))
        else:
            i = self.d.get("list_index", -1)
            if i < 0:
                return NULL
        if not 0 <= i < len(rows):
            raise rt_error(381, "Could not get the Column property. Invalid property array index.")
        v = self.cell(i, j)
        return NULL if v is None else v

    def s_column(self, args, named, value):
        self._need_list()
        j = index_int(to_value(args[0]))
        i = index_int(to_value(args[1])) if len(args) > 1 else self.d.get("list_index", -1)
        self.set_cell(i, j, _py(value))
        self.changed("list", copy.deepcopy(self.rows()))

    def g_selected(self, args, named):
        self._need_list()
        i = index_int(to_value(args[0]))
        rows = self.rows()
        if not 0 <= i < len(rows):
            raise rt_error(381, "Could not get the Selected property. Invalid property array index.")
        sel = self.d.get("selected") or []
        if int(self.d.get("multi_select") or 0) == 0:
            return self.d.get("list_index", -1) == i
        return bool(sel[i]) if i < len(sel) else False

    def s_selected(self, args, named, value):
        self._need_list()
        i = index_int(to_value(args[0]))
        rows = self.rows()
        if not 0 <= i < len(rows):
            raise rt_error(381, "Could not set the Selected property. Invalid property array index.")
        on = bool(to_bool(value))
        if int(self.d.get("multi_select") or 0) == 0:
            if on:
                self.s_listindex((), {}, i)
            elif self.d.get("list_index", -1) == i:
                self.s_listindex((), {}, -1)
            return
        sel = self.d.setdefault("selected", [])
        while len(sel) < len(rows):
            sel.append(False)
        if sel[i] != on:
            sel[i] = on
            self.d["list_index"] = i
            self.changed("selected", list(sel))
            self.inst.code_changed(self)

    def g_topindex(self, a, n):
        return int(self.d.get("top_index", 0))

    def s_topindex(self, a, n, v):
        self.set_key("top_index", index_int(v))

    def g_picture(self, a, n):
        return NOTHING if not self.d.get("picture") else self.d.get("picture")

    def s_picture(self, a, n, v):
        if v is NOTHING or v is EMPTY:
            self.set_key("picture", None)
            return
        if isinstance(v, str):
            self.set_key("picture", v)
            return
        raise VBARuntimeErr(5, "Setting Picture at run time is not supported (LoadPicture)")

    def g_dropdown(self, a, n):
        return EMPTY

    def g_copy(self, a, n):
        return EMPTY

    g_paste = g_copy
    g_cut = g_copy
    g_redraw = g_copy


# --------------------------------------------------------------------------
# Form instance
# --------------------------------------------------------------------------
class FormInstance(VBAObject):
    """A live UserForm (default instance or New UserForm1)."""

    def __init__(self, rt, module, default=False):
        self.rt = rt
        self.module = module
        self.name = module.name
        self.vba_typename = module.name
        self.form = normalize_form(module.form_def)
        if not module.form_def or "caption" not in (module.form_def or {}):
            self.form["caption"] = module.name
        self.form.setdefault("left", 0.0)
        self.form.setdefault("top", 0.0)
        self.form.setdefault("visible", False)
        self.form.setdefault("enabled", True)
        self.form.setdefault("tag", "")
        for c in self.form["controls"]:
            if c["type"] in LIST_TYPES:
                c.setdefault("selected", [False] * len(c.get("list") or []))
        self.d = self.form
        self.vars = module.new_instance_vars()
        if getattr(module, "we_vars", None):
            rt.we_holders.add(self)
        self.view = None
        self.hwnd = 0
        self.window_flags = None
        self.window_state = "normal"
        self.loaded = True
        self.visible = False
        self.is_default = default
        self.terminated = False
        self.active_control = None
        self._ctrl_objs = {}
        self._suppress_echo = False
        self._firing = {}
        rt.live_forms.append(self)
        for c in self.form["controls"]:
            if c.get("row_source"):
                try:
                    self.load_row_source(self.control(c["name"]), fire=False)
                except VBARuntimeErr:
                    pass
            if c.get("control_source"):
                try:
                    self.load_control_source(self.control(c["name"]), fire=False)
                except VBARuntimeErr:
                    pass

    def __repr__(self):
        return f"<FormInstance {self.name}>"

    def font_default(self, key):
        return "Tahoma" if key == "font_name" else 8.0

    def set_key(self, key, value):
        self.form[key] = value
        if self.view is not None:
            self.view.set_prop(None, key, value)

    # -- controls --------------------------------------------------------
    def find_ctrl(self, name):
        low = name.lower()
        for c in self.form["controls"]:
            if c["name"].lower() == low:
                return c
        return None

    def control(self, name):
        o = self._ctrl_objs.get(name.lower())
        if o is None or o.d not in self.form["controls"]:
            c = self.find_ctrl(name)
            if c is None:
                raise rt_error(438, f"Method or data member not found: {name}")
            o = ControlObj(self, c)
            self._ctrl_objs[name.lower()] = o
        return o

    def clear_option_group(self, ctrl):
        g = (ctrl.d.get("group_name") or "").lower()
        parent = (ctrl.d.get("parent") or "").lower()
        for c in self.form["controls"]:
            if c is ctrl.d or c["type"] != "OptionButton" or not c.get("value"):
                continue
            cg = (c.get("group_name") or "").lower()
            same = (cg == g) if (g or cg) else ((c.get("parent") or "").lower() == parent)
            if same:
                o = self.control(c["name"])
                c["value"] = False
                o.changed("value", False)
                self.code_changed(o, click=False)

    def code_changed(self, ctrl, click=None):
        """A value was changed by VBA code: fire Change (and Click)."""
        if click is None:
            click = ctrl.ctype in CLICK_ON_CHANGE and not (
                ctrl.ctype in BOOL_TYPES and ctrl.d.get("value") is False and ctrl.ctype == "OptionButton")
        if not self.loaded:
            return
        self.dispatch(ctrl.name, "Change", [], nested=True)
        if click:
            self.dispatch(ctrl.name, "Click", [], nested=True)

    # -- binding ------------------------------------------------------------
    def load_row_source(self, ctrl, fire=True):
        src = ctrl.d.get("row_source")
        if not src:
            return
        rng = self.rt.range_from_address(src.lstrip("="), None)
        rows = []
        r2 = rng.r2
        if rng.nrows > 100000:
            nr, _ = rng.sheet.used_bounds()
            r2 = min(r2, max(rng.r1, nr - 1))
        ncols = max(int(ctrl.d.get("column_count") or 1), 1)
        for r in range(rng.r1, r2 + 1):
            row = []
            for c in range(rng.c1, rng.c2 + 1):
                v = self.rt.cell_value(rng.sheet, r, c)
                row.append(_py(v) if v is not EMPTY else None)
            row = row[:max(ncols, 1)] if ncols < len(row) and False else row
            rows.append(row if len(row) > 1 else row[0])
        ctrl.d["list"] = rows
        ctrl.d["selected"] = [False] * len(rows)
        ctrl.changed("list", copy.deepcopy(rows))

    def load_control_source(self, ctrl, fire=False):
        src = ctrl.d.get("control_source")
        if not src:
            return
        rng = self.rt.range_from_address(src.lstrip("="), None)
        v = self.rt.cell_value(rng.sheet, rng.r1, rng.c1)
        if ctrl.ctype in BOOL_TYPES:
            v = bool(to_bool(v)) if v is not EMPTY else False
        elif ctrl.ctype in NUM_TYPES:
            v = index_int(v) if v is not EMPTY else 0
        elif v is EMPTY:
            v = ""
        try:
            ctrl.set_value(v, fire=fire)
        except VBARuntimeErr:
            pass

    def write_control_source(self, ctrl):
        src = ctrl.d.get("control_source")
        if not src:
            return
        from .excel import RangeObj
        rng = self.rt.range_from_address(src.lstrip("="), None)
        cell = RangeObj(self.rt, rng.sheet, rng.r1, rng.c1, rng.r1, rng.c1)
        cell.s_value((), {}, ctrl.get_value())

    # -- events ----------------------------------------------------------------
    def find_handler(self, ctrl, event):
        m = self.module
        name = f"{ctrl or 'userform'}_{event}".lower()
        p = m.procs.get(name)
        if p is None and ctrl is None:
            p = m.procs.get(f"form_{event}".lower())
        return p

    def dispatch(self, ctrl, event, args, nested):
        """Run an event handler. nested=True: called from running VBA code."""
        if event.lower() == "afterupdate" and ctrl is not None:
            c = self.find_ctrl(ctrl)
            if c is not None and c.get("control_source"):
                self.write_control_source(self.control(c["name"]))
        targets = []
        p = self.find_handler(ctrl, event)
        if p is not None:
            targets.append((p, self))
        if ctrl is not None:
            c = self.find_ctrl(ctrl)
            if c is not None:
                targets += self.rt.we_targets(
                    lambda v: isinstance(v, ControlObj) and v.d is c, event)
        args = list(args)
        if not targets:
            return args
        key = (ctrl or "", event.lower())
        depth = self._firing.get(key, 0)
        if depth >= 10:
            return args
        vargs = [(_vba(a) if not isinstance(a, VBAObject) else a) for a in args]
        self._firing[key] = depth + 1
        try:
            out = self.rt.run_targets(targets, vargs, nested=nested)
        finally:
            self._firing[key] = depth
        return [_py(v) if not isinstance(v, VBAObject) else v for v in out]

    # -- bridge API (GUI -> runtime) ---------------------------------------------
    def fire(self, ctrl_name, event, args=()):
        return self.dispatch(ctrl_name, event, list(args), nested=self.rt.running > 0 and False)

    def user_set(self, ctrl_name, prop, value):
        if ctrl_name is None:
            self.form[prop] = value
            if prop == "left" or prop == "top":
                pass
            return
        c = self.find_ctrl(ctrl_name)
        if c is None:
            return
        o = self.control(c["name"])
        self._suppress_echo = True
        try:
            if prop in ("value", "text"):
                if c["type"] in LIST_TYPES and prop == "value" and isinstance(value, int) \
                        and not isinstance(value, bool) and False:
                    pass
                if c["type"] == "TextBox":
                    c["value"] = "" if value is None else str(value)
                    c["text"] = c["value"]
                elif c["type"] in LIST_TYPES:
                    rows = o.rows()
                    s = "" if value is None else value
                    idx = -1
                    for i in range(len(rows)):
                        cv = o.cell(i, o.bound_col() if prop == "value" else 0)
                        if cv == s or (cv is not None and str(cv) == str(s)):
                            idx = i
                            break
                    c["list_index"] = idx
                    if idx >= 0:
                        c["value"] = o.cell(idx, o.bound_col())
                        c["text"] = "" if o.cell(idx, 0) is None else str(o.cell(idx, 0))
                    else:
                        c["value"] = s if c["type"] == "ComboBox" else None
                        c["text"] = str(s) if s is not None else ""
                elif c["type"] in BOOL_TYPES:
                    c["value"] = None if value is None else bool(value)
                    if c["type"] == "OptionButton" and value:
                        for oc in self.form["controls"]:
                            if oc is not c and oc["type"] == "OptionButton" and oc.get("value"):
                                g1 = (c.get("group_name") or "").lower()
                                g2 = (oc.get("group_name") or "").lower()
                                if (g1 == g2) if (g1 or g2) else (
                                        (oc.get("parent") or "") == (c.get("parent") or "")):
                                    oc["value"] = False
                else:
                    c["value"] = value
            elif prop == "list_index":
                c["list_index"] = int(value)
                rows = o.rows()
                i = int(value)
                if 0 <= i < len(rows):
                    c["value"] = o.cell(i, o.bound_col())
                    c["text"] = "" if o.cell(i, 0) is None else str(o.cell(i, 0))
                    if int(c.get("multi_select") or 0) == 0:
                        c["selected"] = [k == i for k in range(len(rows))]
                else:
                    c["value"] = "" if c["type"] == "ComboBox" else None
                    c["text"] = ""
            elif prop == "selected":
                if isinstance(value, (list, tuple)):
                    c["selected"] = [bool(x) for x in value]
            else:
                c[prop] = value
        finally:
            self._suppress_echo = False

    def request_close(self, close_mode=0):
        res = self.dispatch(None, "QueryClose", [False, int(close_mode)], nested=False)
        cancel = res[0] if res else False
        if cancel not in (False, 0, None):
            return False
        self._terminate_and_close(nested=False)
        return True

    # -- lifecycle -----------------------------------------------------------------
    def initialize(self):
        self.dispatch(None, "Initialize", [], nested=True)

    def show(self, modal=True):
        if not self.loaded:
            raise rt_error(91)
        rt = self.rt
        if self.view is None:
            factory = getattr(rt.host, "create_form_view", None)
            if factory is None:
                e = VBARuntimeErr(5, "UserForms need the GUI (the current host cannot display forms)")
                e.fatal = True
                raise e
            self.view = factory(self)
            self.apply_window()
        self.visible = True
        self.form["visible"] = True
        self.dispatch(None, "Activate", [], nested=True)
        if not self.loaded:
            return
        rt.modal_depth += modal
        try:
            self.view.show(bool(modal))
        finally:
            rt.modal_depth -= modal
        if rt.end_pending:
            raise VBAEnd()

    def apply_window(self):
        v = self.view
        if v is None:
            return
        if self.window_flags is not None and hasattr(v, "set_window_flags"):
            v.set_window_flags(**self.window_flags)
        if self.window_state != "normal" and hasattr(v, "set_window_state"):
            v.set_window_state(self.window_state)

    def set_window_state(self, state):
        self.window_state = state
        v = self.view
        if v is not None and hasattr(v, "set_window_state"):
            v.set_window_state(state)

    def hide(self):
        self.visible = False
        self.form["visible"] = False
        if self.view is not None:
            self.view.hide()

    def unload(self, nested=True):
        if not self.loaded:
            return
        res = self.dispatch(None, "QueryClose", [False, 1], nested=nested)
        if res and res[0] not in (False, 0, None):
            return
        self._terminate_and_close(nested)

    def _terminate_and_close(self, nested):
        if not self.loaded:
            return
        self.loaded = False
        try:
            self.dispatch(None, "Terminate", [], nested=nested)
        finally:
            self.close_view()

    def close_view(self):
        self.loaded = False
        self.visible = False
        self.terminated = True
        v = self.view
        self.view = None
        if v is not None:
            try:
                v.close()
            except Exception:
                pass
        rt = self.rt
        if self in rt.live_forms:
            rt.live_forms.remove(self)
        if self.is_default and rt.form_defaults.get(self.module.lname) is self:
            del rt.form_defaults[self.module.lname]

    # -- VBA members -----------------------------------------------------------------
    def vba_get(self, name, args, named):
        m = self.module
        p = m.prop_get.get(name) or m.procs.get(name)
        if p is not None and not p.private:
            return self.rt.invoke(p, list(args), named, self)[0]
        i = m.var_index.get(name)
        if i is not None and not m.var_private[i]:
            v = self.vars[i]
            return index_value(v, list(args)) if args else v
        if name and self.find_ctrl(name) is not None:
            o = self.control(name)
            return o.vba_get("", args, named) if args else o
        if name == "" and args:
            return self.g_controls(args, named)
        meth = getattr(self, "g_" + (name or "_default"), None)
        if meth is not None:
            return meth(args, named)
        raise rt_error(438)

    def vba_let(self, name, args, named, value, is_set=False):
        m = self.module
        p = (m.prop_set.get(name) if is_set else None) or m.prop_let.get(name)
        if p is not None and not p.private:
            self.rt.invoke(p, list(args) + [value], named, self)
            return
        i = m.var_index.get(name)
        if i is not None and not m.var_private[i]:
            if args:
                cont = self.vars[i]
                if type(cont) is VBAArray:
                    cont.set(list(args), value)
                    return
            self.vars[i] = value if is_set else m.var_ts[i].coerce(to_value(value) if isinstance(value, VBAObject) else value)
            return
        if name and self.find_ctrl(name) is not None:
            self.control(name).vba_let("", args, named, value)
            return
        meth = getattr(self, "s_" + name, None)
        if meth is not None:
            meth(args, named, value)
            return
        raise rt_error(438)

    def g__default(self, args, named):
        return self.g_controls(args, named)

    def g_show(self, args, named):
        (mode,) = bind(args, named, ["modal"])
        modal = True if mode is MISSING else index_int(to_value(mode)) != 0
        self.show(modal)
        return EMPTY

    def g_hide(self, a, n):
        self.hide()
        return EMPTY

    def g_controls(self, args, named):
        c = ControlsColl(self)
        return c.g_item(args, named) if args else c

    def g_name(self, a, n):
        return self.name

    def g_caption(self, a, n):
        return self.form.get("caption", "")

    def s_caption(self, a, n, v):
        self.set_key("caption", to_str(v))

    def _num(key):
        def g(self, a, n):
            return float(self.form.get(key) or 0.0)

        def s(self, a, n, v):
            self.set_key(key, float(to_double(v)))
        return g, s
    g_width, s_width = _num("width")
    g_height, s_height = _num("height")
    g_left, s_left = _num("left")
    g_top, s_top = _num("top")
    def g_insidewidth(self, a, n):
        v = self.form.get("inside_width")
        if v is not None:
            return float(v)
        return max(0.0, float(self.form.get("width") or 0) - 4)

    def g_insideheight(self, a, n):
        v = self.form.get("inside_height")
        if v is not None:
            return float(v)
        caption = (self.window_flags or {}).get("caption", True)
        return max(0.0, float(self.form.get("height") or 0) - (22 if caption else 4))

    def g_backcolor(self, a, n):
        return _from_hex(self.form.get("back_color"), 0x8000000F)

    def s_backcolor(self, a, n, v):
        self.set_key("back_color", _to_hex(v))

    def g_forecolor(self, a, n):
        return _from_hex(self.form.get("fore_color"), 0)

    def s_forecolor(self, a, n, v):
        self.set_key("fore_color", _to_hex(v))

    def g_font(self, a, n):
        return FontProxy(self)

    def g_visible(self, a, n):
        return self.visible

    def s_visible(self, a, n, v):
        if to_bool(v):
            self.show(False)
        else:
            self.hide()

    def g_enabled(self, a, n):
        return bool(self.form.get("enabled", True))

    def s_enabled(self, a, n, v):
        self.set_key("enabled", bool(to_bool(v)))

    def g_tag(self, a, n):
        return self.form.get("tag", "")

    def s_tag(self, a, n, v):
        self.form["tag"] = to_str(v)

    def g_startupposition(self, a, n):
        return int(self.form.get("start_up_position", 1))

    def s_startupposition(self, a, n, v):
        self.set_key("start_up_position", index_int(v))

    def g_activecontrol(self, a, n):
        if self.active_control and self.find_ctrl(self.active_control) is not None:
            return self.control(self.active_control)
        for c in sorted(self.form["controls"], key=lambda c: c.get("tab_index", 0)):
            if c.get("visible", True) and c.get("enabled", True) and c.get("tab_stop", True) \
                    and c["type"] not in ("Label", "Frame", "Image"):
                return self.control(c["name"])
        return NOTHING

    def g_repaint(self, a, n):
        if self.view is not None:
            self.view.repaint()
        return EMPTY

    def g_setfocus(self, a, n):
        return EMPTY

    def g_printform(self, a, n):
        raise VBARuntimeErr(5, "PrintForm is not supported")

    def g_move(self, args, named):
        l, t, w, h = bind(args, named, ["left", "top", "width", "height"])
        for k, v in (("left", l), ("top", t), ("width", w), ("height", h)):
            if v is not MISSING:
                self.set_key(k, float(to_double(v)))
        return EMPTY

    def g_whatsthisbutton(self, a, n):
        return False

    def g_scrollbars(self, a, n):
        return 0

    def s_scrollbars(self, a, n, v):
        pass

    def g_picture(self, a, n):
        return NOTHING

    def g_activate(self, a, n):
        return EMPTY

    def vba_iter(self):
        return ControlsColl(self).vba_iter()

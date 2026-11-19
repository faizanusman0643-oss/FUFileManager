"""UserForm definitions shared by fileio (parsing .xlsm designers), the VBA
runtime (form object model + events) and the GUI (rendering + designer).

A form module in ``Workbook.vba_modules`` looks like::

    {"name": "UserForm1", "type": "form", "code": "...VBA...", "form": FORM_DICT}

FORM_DICT (all sizes/positions are in POINTS, like VBA; GUI converts with
points_to_px)::

    {
      "caption": "UserForm1",
      "width": 240.0, "height": 180.0,          # outer size, like UserForm.Width/Height
      "back_color": "#F0F0F0" | None,
      "fore_color": "#000000" | None,
      "font_name": "Tahoma", "font_size": 8.0,
      "start_up_position": 1,                   # 0 manual, 1 center owner, 2 center screen
      "controls": [CONTROL_DICT, ...],          # in tab / z order
    }

CONTROL_DICT keys (only "type" and "name" are required; missing keys take the
defaults from CONTROL_DEFAULTS / COMMON_DEFAULTS)::

    type        one of CONTROL_TYPES
    name        "TextBox1"
    parent      name of the containing Frame / MultiPage page, or None for the form
    left, top, width, height   (points, relative to the parent's client area)
    caption, value, text
    visible, enabled, locked, tab_index, tab_stop, tag, control_tip_text
    fore_color, back_color ("#RRGGBB" or None), back_style (0 transparent, 1 opaque)
    font_name, font_size, bold, italic, underline
    text_align (1 left, 2 center, 3 right), word_wrap, auto_size
    multi_line, password_char, max_length, scroll_bars                  (TextBox)
    default, cancel, accelerator                                       (CommandButton)
    group_name                                                         (OptionButton)
    list (list of str), list_index, column_count, column_widths,
    bound_column, row_source, multi_select (0 single, 1 multi, 2 extended),
    style (ComboBox: 0 dropdown combo, 2 dropdown list), list_style     (ComboBox/ListBox)
    min, max, small_change, large_change, orientation                  (SpinButton/ScrollBar)
    pages: [{"name": "Page1", "caption": "Page1"}], value = page index (MultiPage)
    tabs:  [{"name": "Tab1", "caption": "Tab1"}]                       (TabStrip)
    picture: base64 PNG/BMP data or None, picture_size_mode           (Image)
    control_source: "Sheet1!A1" or None
"""
from __future__ import annotations

import copy

CONTROL_TYPES = (
    "Label", "TextBox", "CommandButton", "CheckBox", "OptionButton",
    "ToggleButton", "ComboBox", "ListBox", "Frame", "MultiPage", "TabStrip",
    "SpinButton", "ScrollBar", "Image",
)

# ProgIDs used by Controls.Add("Forms.TextBox.1") and the designer toolbox
PROGIDS = {f"forms.{t.lower()}.1": t for t in CONTROL_TYPES}

COMMON_DEFAULTS = {
    "parent": None,
    "left": 0.0, "top": 0.0, "width": 72.0, "height": 18.0,
    "caption": "", "value": None, "text": "",
    "visible": True, "enabled": True, "locked": False,
    "tab_index": 0, "tab_stop": True, "tag": "", "control_tip_text": "",
    "fore_color": None, "back_color": None, "back_style": 1,
    "font_name": None, "font_size": None, "bold": False, "italic": False,
    "underline": False, "text_align": 1, "word_wrap": False, "auto_size": False,
    "control_source": None,
}

CONTROL_DEFAULTS = {
    "Label": {"width": 72.0, "height": 18.0, "back_style": 0, "word_wrap": True},
    "TextBox": {"width": 72.0, "height": 18.0, "value": "", "multi_line": False,
                "password_char": "", "max_length": 0, "scroll_bars": 0},
    "CommandButton": {"width": 72.0, "height": 24.0, "default": False,
                      "cancel": False, "accelerator": "", "text_align": 2},
    "CheckBox": {"width": 108.0, "height": 18.0, "value": False, "back_style": 0},
    "OptionButton": {"width": 108.0, "height": 18.0, "value": False,
                     "group_name": "", "back_style": 0},
    "ToggleButton": {"width": 72.0, "height": 24.0, "value": False, "text_align": 2},
    "ComboBox": {"width": 72.0, "height": 18.0, "value": "", "list": [],
                 "list_index": -1, "column_count": 1, "column_widths": "",
                 "bound_column": 1, "row_source": None, "style": 0,
                 "list_style": 0, "match_required": False},
    "ListBox": {"width": 72.0, "height": 72.0, "value": None, "list": [],
                "list_index": -1, "column_count": 1, "column_widths": "",
                "bound_column": 1, "row_source": None, "multi_select": 0,
                "list_style": 0},
    "Frame": {"width": 144.0, "height": 72.0, "caption": "Frame"},
    "MultiPage": {"width": 180.0, "height": 120.0, "value": 0,
                  "pages": [{"name": "Page1", "caption": "Page1"},
                            {"name": "Page2", "caption": "Page2"}]},
    "TabStrip": {"width": 180.0, "height": 120.0, "value": 0,
                 "tabs": [{"name": "Tab1", "caption": "Tab1"},
                          {"name": "Tab2", "caption": "Tab2"}]},
    "SpinButton": {"width": 13.0, "height": 26.0, "value": 0, "min": 0, "max": 100,
                   "small_change": 1, "orientation": -1},
    "ScrollBar": {"width": 13.0, "height": 72.0, "value": 0, "min": 0, "max": 32767,
                  "small_change": 1, "large_change": 1, "orientation": -1},
    "Image": {"width": 72.0, "height": 72.0, "picture": None,
              "picture_size_mode": 0, "back_style": 1},
}

FORM_DEFAULTS = {
    "caption": "UserForm1", "width": 240.0, "height": 180.0,
    "back_color": None, "fore_color": None,
    "font_name": "Tahoma", "font_size": 8.0, "start_up_position": 1,
    "controls": [],
}

# Events the runtime dispatches, by control type (form events: see FORM_EVENTS)
FORM_EVENTS = ("Initialize", "Activate", "Deactivate", "QueryClose", "Terminate",
               "Click", "DblClick", "KeyDown", "KeyUp", "KeyPress",
               "MouseDown", "MouseUp", "MouseMove", "Resize")
CONTROL_EVENTS = {
    "_all": ("Enter", "Exit", "KeyDown", "KeyUp", "KeyPress",
             "MouseDown", "MouseUp", "MouseMove"),
    "Label": ("Click", "DblClick"),
    "TextBox": ("Change", "BeforeUpdate", "AfterUpdate", "DblClick"),
    "CommandButton": ("Click",),
    "CheckBox": ("Click", "Change", "AfterUpdate"),
    "OptionButton": ("Click", "Change", "AfterUpdate"),
    "ToggleButton": ("Click", "Change", "AfterUpdate"),
    "ComboBox": ("Change", "Click", "DropButtonClick", "AfterUpdate", "BeforeUpdate"),
    "ListBox": ("Change", "Click", "DblClick", "AfterUpdate", "BeforeUpdate"),
    "Frame": ("Click", "DblClick"),
    "MultiPage": ("Change", "Click"),
    "TabStrip": ("Change", "Click"),
    "SpinButton": ("Change", "SpinUp", "SpinDown"),
    "ScrollBar": ("Change", "Scroll"),
    "Image": ("Click", "DblClick"),
}


def points_to_px(points: float, dpi: float = 96.0) -> int:
    return int(round(points * dpi / 72.0))


def px_to_points(px: float, dpi: float = 96.0) -> float:
    return round(px * 72.0 / dpi, 2)


def new_form(caption: str = "UserForm1") -> dict:
    f = copy.deepcopy(FORM_DEFAULTS)
    f["caption"] = caption
    return f


def new_control(ctype: str, name: str, **props) -> dict:
    if ctype not in CONTROL_TYPES:
        raise ValueError(f"Unknown control type {ctype!r}")
    d = {"type": ctype, "name": name}
    d.update(copy.deepcopy(COMMON_DEFAULTS))
    d.update(copy.deepcopy(CONTROL_DEFAULTS.get(ctype, {})))
    if ctype in ("CommandButton", "Label", "CheckBox", "OptionButton",
                 "ToggleButton", "Frame"):
        d["caption"] = name
    d.update(props)
    return d


def normalize_form(form: dict | None) -> dict:
    """Fill defaults so consumers can index keys directly."""
    out = copy.deepcopy(FORM_DEFAULTS)
    out.update(copy.deepcopy(form or {}))
    ctrls = []
    for c in out.get("controls") or []:
        ctype = c.get("type", "Label")
        if ctype not in CONTROL_TYPES:
            ctype = "Label"
        full = new_control(ctype, c.get("name") or f"{ctype}{len(ctrls) + 1}")
        full.update({k: v for k, v in c.items() if k != "type"})
        ctrls.append(full)
    out["controls"] = ctrls
    return out


def unique_control_name(form: dict, ctype: str) -> str:
    names = {c["name"].lower() for c in form.get("controls", [])}
    i = 1
    while f"{ctype}{i}".lower() in names:
        i += 1
    return f"{ctype}{i}"


# --------------------------------------------------------------------------
# Runtime <-> GUI bridge
# --------------------------------------------------------------------------
class FormView:
    """What the GUI implements for one live form instance (one QDialog).

    The VBA runtime creates a form instance, then calls
    ``host.create_form_view(instance) -> FormView`` (host = the VBAHost).
    ``instance`` exposes:
        instance.name           -> "UserForm1"
        instance.form           -> normalized FORM_DICT holding the *current*
                                   property values (runtime keeps it updated)
        instance.fire(ctrl_name_or_None, event, args=()) -> list
            GUI -> runtime: run the event handler (e.g. "CommandButton1",
            "Click") re-entrantly; returns the final ByRef arg values
            (e.g. Cancel for Exit / QueryClose). ctrl None = form event.
        instance.user_set(ctrl_name, prop, value)
            GUI -> runtime: the user changed a value in the UI (typing,
            ticking, selecting). The runtime stores it WITHOUT echoing back
            to the view, then the GUI fires Change/Click as appropriate.
        instance.request_close(close_mode=0)
            GUI -> runtime: user clicked the window's X. Runtime fires
            QueryClose(Cancel, CloseMode); if not cancelled it unloads.
    """

    def show(self, modal: bool) -> None:
        """Show the window. If modal, BLOCK (nested Qt event loop) until close()
        or hide() is called; event handlers run re-entrantly meanwhile."""

    def hide(self) -> None:
        """Hide (Me.Hide). Ends a modal show() loop."""

    def close(self) -> None:
        """Destroy the window (Unload). Ends a modal show() loop."""

    def set_prop(self, ctrl_name: str | None, prop: str, value) -> None:
        """Runtime -> GUI: a property changed from VBA code (ctrl None = form
        property such as caption/width/back_color)."""

    def add_control(self, ctrl: dict) -> None:
        """Runtime -> GUI: Controls.Add created a control at run time."""

    def remove_control(self, ctrl_name: str) -> None:
        """Runtime -> GUI: Controls.Remove."""

    def set_focus(self, ctrl_name: str) -> None:
        """Runtime -> GUI: Control.SetFocus."""

    def repaint(self) -> None:
        """Me.Repaint / DoEvents."""

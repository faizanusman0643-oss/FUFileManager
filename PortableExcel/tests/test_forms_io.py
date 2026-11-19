"""UserForm designer parsing ([MS-OFORMS]) and persistence in fileio.

Run:  cd PortableExcel && python3 -m pytest tests/test_forms_io.py -q
"""
from __future__ import annotations

import copy
import io
import json
import logging
import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "fixtures"))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from make_vba_project import build_form_designer, build_vba_project, module_header  # noqa: E402

from portable_excel import fileio  # noqa: E402
from portable_excel.fileio import extract_vba_modules, load_workbook, save_workbook  # noqa: E402
from portable_excel.forms import new_control, new_form, normalize_form  # noqa: E402
from portable_excel.model import Workbook  # noqa: E402

FORM = {
    "caption": "Order entry", "width": 300, "height": 200, "back_color": "#FFFFCC",
    "controls": [
        {"type": "Label", "name": "lblName", "left": 6, "top": 6, "width": 60, "height": 12,
         "caption": "Name:", "fore_color": "#FF0000", "text_align": 3},
        {"type": "TextBox", "name": "txtName", "left": 72, "top": 6, "width": 120, "height": 18,
         "value": "Grüße €", "multi_line": True, "max_length": 40, "password_char": "*",
         "control_tip_text": "Your name", "tag": "t1"},
        {"type": "CommandButton", "name": "btnOK", "left": 200, "top": 150, "width": 72,
         "height": 24, "caption": "OK", "default": True},
        {"type": "CommandButton", "name": "btnCancel", "left": 120, "top": 150.75, "width": 72,
         "height": 24, "caption": "Cancel", "cancel": True},
        {"type": "CheckBox", "name": "chkGift", "left": 6, "top": 40, "width": 100, "height": 18,
         "caption": "Gift wrap", "value": True},
        {"type": "OptionButton", "name": "optA", "left": 6, "top": 60, "width": 100,
         "height": 18, "caption": "Option A", "value": False, "group_name": "grp"},
        {"type": "ComboBox", "name": "cboSize", "left": 120, "top": 40, "width": 80,
         "height": 18, "style": 2, "list_style": 1, "value": "M", "column_count": 2},
        {"type": "ListBox", "name": "lstItems", "left": 120, "top": 70, "width": 80,
         "height": 60, "multi_select": 1},
        {"type": "Frame", "name": "fraOpts", "left": 6, "top": 90, "width": 100, "height": 60,
         "caption": "Options"},
        {"type": "CheckBox", "name": "chkFast", "parent": "fraOpts", "left": 6, "top": 6,
         "width": 80, "height": 18, "caption": "Fast", "value": False},
        {"type": "Label", "name": "lblHidden", "parent": "fraOpts", "left": 6, "top": 30,
         "width": 80, "height": 12, "caption": "hidden", "visible": False},
        {"type": "SpinButton", "name": "spn", "left": 210, "top": 40, "width": 13, "height": 26,
         "min": 1, "max": 10, "value": 3},
        {"type": "ScrollBar", "name": "scr", "left": 230, "top": 40, "width": 13, "height": 60,
         "min": 0, "max": 50, "value": 7, "large_change": 5},
        {"type": "TabStrip", "name": "ts", "left": 250, "top": 70, "width": 40, "height": 60,
         "value": 1, "tabs": [{"name": "t1", "caption": "One"}, {"name": "t2", "caption": "Two"}]},
        {"type": "Image", "name": "img", "left": 250, "top": 140, "width": 30, "height": 30},
    ],
}
FORM_CODE = (module_header("UserForm1", "form")
             + "Private Sub btnOK_Click()\r\n    Me.Hide\r\nEnd Sub\r\n")


def _project(form=FORM, designer=None):
    mods = [("ThisWorkbook", "document", module_header("ThisWorkbook", "document", True)),
            ("Module1", "standard", module_header("Module1", "standard") + "Sub Go()\r\n    UserForm1.Show\r\nEnd Sub\r\n"),
            ("UserForm1", "form", FORM_CODE)]
    return build_vba_project(mods, forms={"UserForm1": designer if designer is not None else form})


@pytest.fixture(scope="module")
def form_bin():
    return _project()


def _by_name(form):
    return {c["name"]: c for c in form["controls"]}


def test_form_module_detected_with_designer(form_bin):
    mods = {m["name"]: m for m in extract_vba_modules(form_bin)}
    assert mods["UserForm1"]["type"] == "form"
    assert mods["UserForm1"]["code"] == FORM_CODE.replace("\r\n", "\n")
    assert "form" not in mods["Module1"] and "form" not in mods["ThisWorkbook"]
    form = mods["UserForm1"]["form"]
    assert form == normalize_form(form)
    assert form["caption"] == "Order entry"
    assert form["width"] == 300 and form["height"] == 200
    assert form["back_color"] == "#FFFFCC"
    assert form["fore_color"] == "#000000"
    assert form["font_name"] == "Tahoma" and form["font_size"] == 8.25
    assert form["start_up_position"] == 1


def test_control_types_positions_and_order(form_bin):
    form = extract_vba_modules(form_bin)[2]["form"]
    exp = {c["name"]: c for c in FORM["controls"]}
    got = _by_name(form)
    assert list(got) == ["lblName", "txtName", "btnOK", "btnCancel", "chkGift", "optA",
                         "cboSize", "lstItems", "fraOpts", "chkFast", "lblHidden", "spn",
                         "scr", "ts", "img"]
    for name, c in exp.items():
        g = got[name]
        assert g["type"] == c["type"], name
        for key in ("left", "top", "width", "height"):
            assert g[key] == pytest.approx(c[key], abs=0.01), (name, key)
        assert g["parent"] == c.get("parent"), name
    # tab index restarts inside the frame
    assert got["chkFast"]["tab_index"] == 0 and got["lblHidden"]["tab_index"] == 1
    assert got["btnOK"]["tab_index"] == 2
    # frame precedes its children
    names = [c["name"] for c in form["controls"]]
    assert names.index("fraOpts") < names.index("chkFast")


def test_control_properties(form_bin):
    got = _by_name(extract_vba_modules(form_bin)[2]["form"])
    assert got["lblName"]["caption"] == "Name:"
    assert got["lblName"]["fore_color"] == "#FF0000"
    assert got["lblName"]["text_align"] == 3
    assert got["lblName"]["tab_stop"] is False
    t = got["txtName"]
    assert t["value"] == "Grüße €" and t["text"] == "Grüße €"
    assert t["multi_line"] is True and t["max_length"] == 40 and t["password_char"] == "*"
    assert t["control_tip_text"] == "Your name" and t["tag"] == "t1"
    assert t["enabled"] is True and t["visible"] is True and t["tab_stop"] is True
    assert got["btnOK"]["caption"] == "OK" and got["btnOK"]["default"] is True
    assert got["btnOK"]["cancel"] is False and got["btnCancel"]["cancel"] is True
    assert got["btnOK"]["text_align"] == 2
    assert got["chkGift"]["value"] is True and got["chkGift"]["caption"] == "Gift wrap"
    assert got["optA"]["value"] is False and got["optA"]["group_name"] == "grp"
    assert got["optA"]["caption"] == "Option A"
    cbo = got["cboSize"]
    assert cbo["style"] == 2 and cbo["list_style"] == 1 and cbo["value"] == "M"
    assert cbo["column_count"] == 2 and cbo["list"] == []
    lst = got["lstItems"]
    assert lst["multi_select"] == 1 and lst["list_style"] == 0 and lst["value"] is None
    assert got["fraOpts"]["caption"] == "Options"
    assert got["chkFast"]["caption"] == "Fast" and got["chkFast"]["value"] is False
    assert got["lblHidden"]["visible"] is False
    assert (got["spn"]["min"], got["spn"]["max"], got["spn"]["value"]) == (1, 10, 3)
    assert (got["scr"]["max"], got["scr"]["value"], got["scr"]["large_change"]) == (50, 7, 5)
    assert got["ts"]["tabs"] == [{"name": "t1", "caption": "One"}, {"name": "t2", "caption": "Two"}]
    assert got["ts"]["value"] == 1
    assert got["img"]["picture"] is None


def test_ole_colors():
    c = fileio.ole_color_to_hex
    assert c(0x000000FF) == "#FF0000"
    assert c(0x00FF0000) == "#0000FF"
    assert c(0x8000000F) == "#F0F0F0"   # button face
    assert c(0x80000012) == "#000000"   # button text
    assert c(0x80000005) == "#FFFFFF"   # window
    assert c(0x80000008) == "#000000"   # window text
    assert c(None) is None


def test_malformed_designer_never_crashes(caplog):
    designer = build_form_designer("UserForm1", FORM)
    designer["f"] = designer["f"][:40] + b"\xff" * 30
    data = _project(designer=designer)
    with caplog.at_level(logging.WARNING, logger="portable_excel.fileio"):
        mods = {m["name"]: m for m in extract_vba_modules(data)}
    form = mods["UserForm1"]["form"]
    assert form["caption"] == "Order entry"
    assert isinstance(form["controls"], list)
    assert any("UserForm" in r.getMessage() for r in caplog.records)


def test_garbage_object_stream_keeps_sites(caplog):
    designer = build_form_designer("UserForm1", FORM)
    designer["o"] = os.urandom(len(designer["o"]))
    with caplog.at_level(logging.WARNING, logger="portable_excel.fileio"):
        form = {m["name"]: m for m in extract_vba_modules(_project(designer=designer))}["UserForm1"]["form"]
    got = _by_name(form)
    assert got["btnOK"]["type"] == "CommandButton"
    assert got["btnOK"]["left"] == 200 and got["btnOK"]["default"] is True


def test_missing_designer_gives_empty_form():
    mods = [("UserForm1", "form", FORM_CODE)]
    data = build_vba_project(mods)   # BaseClass= in PROJECT but no storage
    m = extract_vba_modules(data)[0]
    assert m["type"] == "form"
    assert m["form"] == normalize_form(new_form("UserForm1"))


def test_unknown_control_type_skipped():
    designer = build_form_designer("UserForm1", {"controls": [
        {"type": "Label", "name": "A", "caption": "a"},
        {"type": "CommandButton", "name": "B", "caption": "b"}]})
    # patch the first site's ClsidCacheIndex (21 = Label) to an unknown 99
    # (the index is the last DataBlock field, right before the name "A" + padding)
    f = designer["f"]
    assert f.count(b"\x15\x00A\x00\x00\x00") == 1
    designer["f"] = f.replace(b"\x15\x00A\x00\x00\x00", b"\x63\x00A\x00\x00\x00")
    form = extract_vba_modules(_project(designer=designer))[2]["form"]
    assert [c["name"] for c in form["controls"]] == ["B"]
    assert form["controls"][0]["caption"] == "b"


# --------------------------------------------------------------------------
# persistence through .xlsm
# --------------------------------------------------------------------------
def _book(vbin):
    b = Workbook()
    b.active.set_raw(0, 0, "1")
    b.vba_project_bin = vbin
    b.vba_modules = extract_vba_modules(vbin)
    return b


def test_xlsm_load_has_forms_and_bin_kept(tmp_path, form_bin):
    p = str(tmp_path / "f.xlsm")
    save_workbook(_book(form_bin), p)
    zf = zipfile.ZipFile(p)
    assert zf.read("xl/vbaProject.bin") == form_bin
    assert fileio.VBA_JSON_PART not in zf.namelist()   # unchanged -> bin only
    b = load_workbook(p)
    form = {m["name"]: m for m in b.vba_modules}["UserForm1"]["form"]
    assert _by_name(form)["chkFast"]["parent"] == "fraOpts"
    assert b.vba_modified is False


def test_edited_form_round_trips_via_json(tmp_path, form_bin):
    b = _book(form_bin)
    mod = {m["name"]: m for m in b.vba_modules}["UserForm1"]
    btn = _by_name(mod["form"])["btnOK"]
    btn["left"] = 33.5
    btn["caption"] = "Go!"
    mod["form"]["controls"].append(new_control("TextBox", "txtNew", left=1, top=2,
                                               parent="fraOpts"))
    mod["form"]["caption"] = "Edited"
    b.vba_modified = True
    p = str(tmp_path / "edited.xlsm")
    save_workbook(b, p)
    zf = zipfile.ZipFile(p)
    payload = json.loads(zf.read(fileio.VBA_JSON_PART))
    stored = {m["name"]: m for m in payload["modules"]}["UserForm1"]
    assert stored["form"]["caption"] == "Edited"
    assert zf.read("xl/vbaProject.bin") == form_bin     # original designer kept for Excel
    b2 = load_workbook(p)
    form = {m["name"]: m for m in b2.vba_modules}["UserForm1"]["form"]
    got = _by_name(form)
    assert form["caption"] == "Edited"
    assert got["btnOK"]["left"] == 33.5 and got["btnOK"]["caption"] == "Go!"
    assert got["txtNew"]["parent"] == "fraOpts"
    # re-save without changes keeps the JSON edits (vba_modified False after load)
    p2 = str(tmp_path / "edited2.xlsm")
    save_workbook(b2, p2)
    form3 = {m["name"]: m for m in load_workbook(p2).vba_modules}["UserForm1"]["form"]
    assert form3 == form


def test_form_only_change_detected_without_flag(tmp_path, form_bin):
    b = _book(form_bin)
    mod = {m["name"]: m for m in b.vba_modules}["UserForm1"]
    mod["form"]["width"] = 999.0
    p = str(tmp_path / "w.xlsm")
    save_workbook(b, p)   # vba_modified not set, but the form differs from the bin
    form = {m["name"]: m for m in load_workbook(p).vba_modules}["UserForm1"]["form"]
    assert form["width"] == 999.0


def test_form_created_in_app_round_trips(tmp_path):
    b = Workbook()
    form = new_form("Dialog")
    form["controls"] = [new_control("CommandButton", "btn1", left=10, top=20),
                        new_control("Frame", "fr"),
                        new_control("OptionButton", "o1", parent="fr", value=True)]
    b.vba_modules = [{"name": "UserForm1", "type": "form", "code": "Private Sub btn1_Click()\nEnd Sub\n",
                      "form": copy.deepcopy(form)}]
    p = str(tmp_path / "new.xlsm")
    save_workbook(b, p)
    b2 = load_workbook(p)
    m = b2.vba_modules[0]
    assert m["type"] == "form" and m["code"] == b.vba_modules[0]["code"]
    assert m["form"] == normalize_form(form)


def test_json_form_missing_or_bad_is_normalized(tmp_path):
    b = Workbook()
    b.vba_modules = [{"name": "UF", "type": "form", "code": ""}]
    p = str(tmp_path / "nf.xlsm")
    save_workbook(b, p)
    m = load_workbook(p).vba_modules[0]
    assert m["form"] == normalize_form(new_form("UF"))


# --------------------------------------------------------------------------
# optional: real VBE-written sample (not stored in the repo)
# --------------------------------------------------------------------------
SAMPLE = os.environ.get("PORTABLEEXCEL_OLEFORM_SAMPLE")


@pytest.mark.skipif(not SAMPLE or not os.path.exists(SAMPLE or ""),
                    reason="set PORTABLEEXCEL_OLEFORM_SAMPLE to oletools' oleform-PR314.docm")
def test_real_world_sample():
    mods = {m["name"]: m for m in extract_vba_modules(SAMPLE)}
    form = mods["UserFormTEST1"]["form"]
    got = _by_name(form)
    assert form["caption"] == "UserForm1_capt"
    assert got["Label1"]["caption"] == "Label1-test" and got["Label1"]["left"] == 54
    assert got["TextBox1"]["value"] == "heyhey"
    assert got["ComboBox1"]["value"] == "none dd"
    assert got["CheckBox1"]["caption"] == "mouahaha" and got["CheckBox1"]["value"] is True
    assert got["TextBox2"]["parent"] == "Frame1" and got["TextBox2"]["value"] == "abcd"
    assert got["MultiPage1"]["pages"] == [{"name": "Page1", "caption": "Page1"},
                                          {"name": "Page2", "caption": "Page2"}]
    assert got["TextBox3"]["parent"] == "Page1"
    assert got["TabStrip1"]["tabs"][1]["caption"] == "Tab2"
    assert got["ListBox1"]["list_style"] == 1

"""Tests for portable_excel.fileio.

Run:  cd PortableExcel && python3 -m pytest tests/test_fileio.py -q
"""
from __future__ import annotations

import datetime as dt
import io
import json
import os
import sys
import zipfile

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "fixtures"))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from make_vba_project import build_vba_project, compress, module_header, write_cfb  # noqa: E402

from portable_excel import fileio  # noqa: E402
from portable_excel.fileio import (FileFormatError, PasswordRequired,  # noqa: E402
                                   extract_vba_modules, load_workbook, save_workbook)
from portable_excel.model import Workbook, literal_value  # noqa: E402

openpyxl = pytest.importorskip("openpyxl")


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def _crlf(s: str) -> str:
    return s.replace("\n", "\r\n")


MODULE1_BODY = 'Option Explicit\n\nSub Hello()\n    Range("A1").Value = "Hi"\n    MsgBox "Grüße"\nEnd Sub\n'
CLASS1_BODY = "Public Total As Double\n\nPublic Sub Add(x As Double)\n    Total = Total + x\nEnd Sub\n"
WB_BODY = "Private Sub Workbook_Open()\n    Module1.Hello\nEnd Sub\n"
SHEET1_BODY = "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n"


def _vba_modules_spec():
    return [
        ("ThisWorkbook", "document", module_header("ThisWorkbook", "document", True) + _crlf(WB_BODY)),
        ("Sheet1", "document", module_header("Sheet1", "document") + _crlf(SHEET1_BODY)),
        ("Module1", "standard", module_header("Module1", "standard") + _crlf(MODULE1_BODY)),
        ("Class1", "class", module_header("Class1", "class") + _crlf(CLASS1_BODY)),
    ]


@pytest.fixture(scope="module")
def vba_bin() -> bytes:
    return build_vba_project(_vba_modules_spec())


def _expected_modules():
    return {name: (kind, code.replace("\r\n", "\n")) for name, kind, code in _vba_modules_spec()}


def _sample_book() -> Workbook:
    b = Workbook()
    s = b.active
    s.set_raw(0, 0, "12")
    s.set_raw(0, 1, "=A1*2")
    s.set_raw(0, 2, "=SUM(A1:B1)")
    s.set_raw(1, 0, "3.25")
    s.set_raw(1, 1, "-0.1")
    s.set_raw(1, 2, "1e-07")
    s.set_raw(2, 0, "hello world")
    s.set_raw(2, 1, "TRUE")
    s.set_raw(2, 2, "FALSE")
    s.set_raw(3, 0, "'00123")
    s.set_raw(3, 1, "'=not a formula")
    s.set_raw(3, 2, "#N/A")
    s.set_raw(4, 0, "123456789012")
    s.set_raw(4, 1, "0.1234567890123456789")
    s.set_raw(4, 2, "'TRUE")
    s.set_raw(5, 0, "'it''s")
    return b


# --------------------------------------------------------------------------
# xlsx round trips
# --------------------------------------------------------------------------
def test_values_and_formulas_round_trip(tmp_path):
    b = _sample_book()
    p = str(tmp_path / "v.xlsx")
    assert save_workbook(b, p) == []
    b2 = load_workbook(p)
    s, s2 = b.active, b2.active
    assert b2.path == p
    for key, raw in s.cells.items():
        got = s2.cells[key]
        if raw.startswith("="):
            assert got == raw
        else:
            assert literal_value(got) == literal_value(raw), (key, raw, got)
    assert s2.cells[(0, 0)] == "12"
    assert s2.cells[(1, 0)] == "3.25"
    assert s2.cells[(3, 0)] == "'00123"
    assert s2.cells[(3, 1)] == "'=not a formula"
    assert s2.cells[(2, 1)] == "TRUE"
    assert s2.cells[(4, 2)] == "'TRUE"
    assert s2.cells[(3, 2)] == "#N/A"
    # stored as real types in the xlsx
    ws = openpyxl.load_workbook(p).active
    assert ws["A1"].value == 12 and isinstance(ws["A1"].value, int)
    assert ws["B1"].value == "=A1*2"
    assert ws["B3"].value is True
    assert ws["A4"].value == "00123" and ws["A4"].data_type == "s"
    assert ws["B4"].value == "=not a formula" and ws["B4"].data_type == "s"
    assert ws["C4"].data_type == "e"


def test_float_precision_exact(tmp_path):
    b = Workbook()
    vals = [0.1, 1 / 3, 0.10070953948708206, 1e-300, 123456.789e10, -2.5e-8, 2 ** 0.5]
    for i, v in enumerate(vals):
        b.active.set_raw(i, 0, repr(v))
    p = str(tmp_path / "f.xlsx")
    save_workbook(b, p)
    b2 = load_workbook(p)
    for i, v in enumerate(vals):
        assert literal_value(b2.active.cells[(i, 0)]) == v


def test_numeric_looking_text_from_excel_file(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    texts = ["00123", "1e5", "=notformula", "TRUE", " 42 ", "50%", "'quoted", "plain", "#N/A", "1,234"]
    for i, t in enumerate(texts, start=1):
        c = ws.cell(row=i, column=1, value=t)
        c.data_type = "s"
    p = str(tmp_path / "t.xlsx")
    wb.save(p)
    s = load_workbook(p).active
    for i, t in enumerate(texts):
        raw = s.cells[(i, 0)]
        assert literal_value(raw) == t, (t, raw)
    assert s.cells[(7, 0)] == "plain"
    assert s.cells[(0, 0)] == "'00123"


def test_dates_become_serials(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = dt.datetime(2024, 1, 15)
    ws["A1"].number_format = "yyyy-mm-dd"
    ws["A2"] = dt.datetime(2024, 1, 15, 18, 0)
    ws["A2"].number_format = "yyyy-mm-dd hh:mm"
    ws["A3"] = dt.date(1900, 3, 1)
    ws["A4"] = dt.time(6, 0)
    ws["A4"].number_format = "hh:mm"
    p = str(tmp_path / "d.xlsx")
    wb.save(p)
    s = load_workbook(p).active
    assert s.cells[(0, 0)] == "45306"
    assert s.styles[(0, 0)]["number_format"] == "yyyy-mm-dd"
    assert s.cells[(1, 0)] == "45306.75"
    assert s.cells[(2, 0)] == "61"
    assert s.cells[(2, 0)] and "number_format" in s.styles[(2, 0)]
    assert s.cells[(3, 0)] == "0.25"
    assert s.styles[(3, 0)]["number_format"] == "hh:mm"


def test_styles_round_trip(tmp_path):
    b = Workbook()
    s = b.active
    full = {"bold": True, "italic": True, "underline": True, "color": "#FF0000",
            "fill": "#00FF00", "align": "center", "font_size": 14.0, "font_name": "Arial",
            "number_format": "#,##0.00", "wrap": True}
    s.set_raw(0, 0, "1234.5")
    s.styles[(0, 0)] = dict(full)
    s.set_raw(1, 0, "x")
    s.styles[(1, 0)] = {"bold": True}
    s.set_raw(2, 0, "y")
    s.styles[(2, 0)] = {"align": "right"}
    s.styles[(3, 3)] = {"fill": "#123456"}          # styled empty cell
    for r in range(10, 200):                       # many cells sharing a style
        s.set_raw(r, 1, str(r))
        s.styles[(r, 1)] = {"italic": True, "color": "#0000FF"}
    p = str(tmp_path / "s.xlsx")
    save_workbook(b, p)
    s2 = load_workbook(p).active
    assert s2.styles[(0, 0)] == full
    assert s2.styles[(1, 0)] == {"bold": True}
    assert s2.styles[(2, 0)] == {"align": "right"}
    assert s2.styles[(3, 3)] == {"fill": "#123456"}
    assert (3, 3) not in s2.cells
    for r in range(10, 200):
        assert s2.styles[(r, 1)] == {"italic": True, "color": "#0000FF"}
    assert (5, 5) not in s2.styles
    # style dicts must be independent objects
    s2.styles[(10, 1)]["bold"] = True
    assert "bold" not in s2.styles[(11, 1)]


def test_percent_typed_value_gets_percent_format(tmp_path):
    b = Workbook()
    b.active.set_raw(0, 0, "50%")
    b.active.set_raw(1, 0, "12.5%")
    p = str(tmp_path / "pct.xlsx")
    save_workbook(b, p)
    s = load_workbook(p).active
    assert literal_value(s.cells[(0, 0)]) == 0.5
    assert s.styles[(0, 0)]["number_format"] == "0%"
    assert s.styles[(1, 0)]["number_format"] == "0.00%"


def test_theme_colours_ignored_and_default_font_not_reported(tmp_path):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.styles.colors import Color
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "t"
    ws["A1"].font = Font(name="Calibri", sz=11, color=Color(theme=4))
    ws["A1"].fill = PatternFill("solid", fgColor=Color(theme=5))
    ws["A2"] = "b"
    ws["A2"].font = Font(bold=True, name="Calibri", sz=11)
    p = str(tmp_path / "th.xlsx")
    wb.save(p)
    s = load_workbook(p).active
    assert s.styles.get((0, 0), {}) == {}
    assert s.styles[(1, 0)] == {"bold": True}


def test_column_widths_and_row_heights(tmp_path):
    b = Workbook()
    s = b.active
    s.set_raw(0, 0, "a")
    s.col_widths = {0: 100, 2: 64, 5: 200}
    s.row_heights = {0: 40, 3: 20}
    p = str(tmp_path / "w.xlsx")
    save_workbook(b, p)
    s2 = load_workbook(p).active
    assert s2.col_widths == {0: 100, 2: 64, 5: 200}
    assert s2.row_heights == {0: 40, 3: 20}


def test_column_width_groups_from_excel(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws.column_dimensions.group("B", "D", hidden=False)
    ws.column_dimensions["B"].width = 20
    ws.row_dimensions[2].height = 30
    p = str(tmp_path / "g.xlsx")
    wb.save(p)
    s = load_workbook(p).active
    assert s.col_widths[1] == s.col_widths[2] == s.col_widths[3] == int(20 * 7 + 5)
    assert 0 not in s.col_widths and 4 not in s.col_widths
    assert s.row_heights[1] == 40


def test_multiple_sheets_visibility_active_and_codenames(tmp_path):
    b = Workbook()
    b.active.set_raw(0, 0, "first")
    s2 = b.add_sheet("Data")
    s2.set_raw(4, 3, "=Sheet1!A1")
    s3 = b.add_sheet("Hidden one")
    s3.visible = False
    s3.set_raw(0, 0, "secret")
    b.active_index = 1
    p = str(tmp_path / "m.xlsx")
    save_workbook(b, p)
    b2 = load_workbook(p)
    assert [s.name for s in b2.sheets] == ["Sheet1", "Data", "Hidden one"]
    assert [s.code_name for s in b2.sheets] == ["Sheet1", "Sheet2", "Sheet3"]
    assert [s.visible for s in b2.sheets] == [True, True, False]
    assert b2.active_index == 1
    assert b2.sheets[1].cells == {(4, 3): "=Sheet1!A1"}
    assert b2.sheets[2].cells == {(0, 0): "secret"}


def test_codename_default_when_missing(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "Alpha"
    wb.create_sheet("Beta")
    p = str(tmp_path / "c.xlsx")
    wb.save(p)
    b = load_workbook(p)
    assert [s.code_name for s in b.sheets] == ["Sheet1", "Sheet2"]


def test_defined_names_round_trip(tmp_path):
    b = Workbook()
    b.add_sheet("Data")
    b.defined_names = {"MyRange": "Sheet1!$A$1:$B$2", "Rate": "Data!$C$3"}
    p = str(tmp_path / "n.xlsx")
    save_workbook(b, p)
    b2 = load_workbook(p)
    assert b2.defined_names == {"MyRange": "Sheet1!$A$1:$B$2", "Rate": "Data!$C$3"}


def test_defined_names_from_openpyxl_file(tmp_path):
    from openpyxl.workbook.defined_name import DefinedName
    wb = openpyxl.Workbook()
    wb.defined_names["Total"] = DefinedName("Total", attr_text="Sheet!$A$1:$A$10")
    p = str(tmp_path / "n2.xlsx")
    wb.save(p)
    assert load_workbook(p).defined_names == {"Total": "Sheet!$A$1:$A$10"}


def test_array_formula_loaded_as_text(tmp_path):
    from openpyxl.worksheet.formula import ArrayFormula
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["A2"] = 2
    ws["B1"] = ArrayFormula("B1", "=SUM(A1:A2*2)")
    p = str(tmp_path / "af.xlsx")
    wb.save(p)
    assert load_workbook(p).active.cells[(0, 1)] == "=SUM(A1:A2*2)"


def test_illegal_chars_and_sheet_name_rules(tmp_path):
    b = Workbook()
    b.active.name = "A very long sheet name that exceeds thirty one chars"
    b.active.set_raw(0, 0, "bad\x01char\x07s")
    b.add_sheet("x[1]:y")
    p = str(tmp_path / "i.xlsx")
    warnings = save_workbook(b, p)
    assert any("sheet names" in w for w in warnings)
    b2 = load_workbook(p)
    assert len(b2.sheets[0].name) == 31
    assert b2.sheets[0].cells[(0, 0)] == "badchars"
    assert b2.sheets[1].name == "x_1__y"


def test_template_formats(tmp_path):
    b = _sample_book()
    p = str(tmp_path / "t.xltx")
    save_workbook(b, p)
    ct = zipfile.ZipFile(p).read("[Content_Types].xml").decode()
    assert "spreadsheetml.template.main+xml" in ct
    assert load_workbook(p).active.cells[(0, 1)] == "=A1*2"


def test_empty_workbook(tmp_path):
    p = str(tmp_path / "e.xlsx")
    save_workbook(Workbook(), p)
    b = load_workbook(p)
    assert len(b.sheets) == 1 and b.active.cells == {}


def test_large_workbook_fast(tmp_path):
    import time
    b = Workbook()
    s = b.active
    for r in range(5000):
        for c in range(10):
            s.cells[(r, c)] = f"=A{r + 1}+{c}" if c == 9 else (str(r * 10 + c) if c % 2 else f"t{r}")
            if c == 1:
                s.styles[(r, c)] = {"bold": True, "number_format": "0.00"}
    p = str(tmp_path / "big.xlsx")
    t = time.time()
    save_workbook(b, p)
    b2 = load_workbook(p)
    elapsed = time.time() - t
    assert b2.active.cells == s.cells
    assert b2.active.styles == s.styles
    assert elapsed < 15


# --------------------------------------------------------------------------
# atomic save
# --------------------------------------------------------------------------
def test_atomic_save_keeps_original_on_failure(tmp_path, monkeypatch):
    p = str(tmp_path / "keep.xlsx")
    save_workbook(_sample_book(), p)
    before = open(p, "rb").read()

    def boom(*a, **k):
        raise RuntimeError("crash while writing")

    monkeypatch.setattr(fileio, "_postprocess_zip", boom)
    b = _sample_book()
    with pytest.raises(RuntimeError):
        save_workbook(b, str(tmp_path / "keep.xlsm"))
    monkeypatch.setattr(os, "replace", boom)
    with pytest.raises(RuntimeError):
        save_workbook(b, p)
    assert open(p, "rb").read() == before
    assert sorted(os.listdir(tmp_path)) == ["keep.xlsx"]


@pytest.mark.skipif(sys.platform == "win32", reason="POSIX permission bits do not exist on Windows")
def test_save_preserves_permissions(tmp_path):
    p = str(tmp_path / "perm.xlsx")
    save_workbook(_sample_book(), p)
    os.chmod(p, 0o640)
    save_workbook(_sample_book(), p)
    assert os.stat(p).st_mode & 0o777 == 0o640


# --------------------------------------------------------------------------
# CSV
# --------------------------------------------------------------------------
def test_csv_semicolon_cp1252(tmp_path):
    p = tmp_path / "Umsätze.csv"
    p.write_bytes("Name;Betrag;Stadt\r\nMüller;12,50;Köln\r\n\"A;B\";3;€uro\r\n".encode("cp1252"))
    b = load_workbook(str(p))
    s = b.active
    assert s.name == "Umsätze"
    assert len(b.sheets) == 1
    assert s.cells[(1, 0)] == "Müller"
    assert s.cells[(1, 1)] == "12,50"
    assert s.cells[(1, 2)] == "Köln"
    assert s.cells[(2, 0)] == "A;B"
    assert s.cells[(2, 2)] == "€uro"


def test_csv_utf8_bom_comma_and_tab(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes("﻿a,b,c\n1,2.5,=A2+B2\n,,x\n".encode("utf-8"))
    s = load_workbook(str(p)).active
    assert s.cells == {(0, 0): "a", (0, 1): "b", (0, 2): "c", (1, 0): "1", (1, 1): "2.5",
                       (1, 2): "=A2+B2", (2, 2): "x"}
    p2 = tmp_path / "t.csv"
    p2.write_text("x\ty, z\n1\t2, 3\n", encoding="utf-8")
    s2 = load_workbook(str(p2)).active
    assert s2.cells[(0, 1)] == "y, z"
    assert s2.cells[(1, 0)] == "1"


def test_csv_latin1_fallback(tmp_path):
    p = tmp_path / "l.csv"
    p.write_bytes(b"a,\x81\x8d\n")  # bytes undefined in cp1252 -> latin-1
    s = load_workbook(str(p)).active
    assert s.cells[(0, 1)] == "\x81\x8d"


def test_csv_save(tmp_path):
    b = Workbook()
    s = b.active
    s.set_raw(0, 0, "Name")
    s.set_raw(0, 1, "Wert")
    s.set_raw(1, 0, "Grüße, Welt")
    s.set_raw(1, 1, "=1+1")
    s.set_raw(2, 0, "'00123")
    s.set_raw(3, 2, "end")
    b.add_sheet("Other")
    p = str(tmp_path / "out.csv")
    warnings = save_workbook(b, p)
    assert any("only the active sheet" in w for w in warnings)
    data = open(p, "rb").read()
    assert data.startswith(b"\xef\xbb\xbf")
    text = data.decode("utf-8-sig")
    assert text == 'Name,Wert\r\n"Grüße, Welt",=1+1\r\n00123\r\n,,end\r\n'

    class Engine:
        def display_text(self, sheet, r, c):
            return "2" if (r, c) == (1, 1) else sheet.get_raw(r, c).lstrip("'")

    save_workbook(b, p, engine=Engine())
    assert open(p, encoding="utf-8-sig").read().splitlines()[1] == '"Grüße, Welt",2'
    # csv round trip
    b2 = load_workbook(p)
    assert b2.active.cells[(1, 0)] == "Grüße, Welt"


def test_csv_password_rejected(tmp_path):
    with pytest.raises(FileFormatError):
        save_workbook(Workbook(), str(tmp_path / "x.csv"), password="secret")


# --------------------------------------------------------------------------
# VBA fixture builder self-test
# --------------------------------------------------------------------------
def test_fixture_compression_round_trip():
    from oletools.olevba import decompress_stream
    samples = [b"", b"a", b"abcabcabcabcabc", b"Sub A()\r\nEnd Sub\r\n" * 500,
               bytes(range(256)) * 20, os.urandom(9000)]
    for s in samples:
        assert decompress_stream(bytearray(compress(s))) == s


def test_fixture_vba_project_is_valid_ole(vba_bin):
    import olefile
    ole = olefile.OleFileIO(io.BytesIO(vba_bin))
    streams = {"/".join(e) for e in ole.listdir()}
    assert {"PROJECT", "VBA/dir", "VBA/_VBA_PROJECT", "VBA/Module1", "VBA/ThisWorkbook"} <= streams
    assert ole.openstream("VBA/_VBA_PROJECT").read()[:4] == b"\xcc\x61\xff\xff"


# --------------------------------------------------------------------------
# VBA extraction / preservation
# --------------------------------------------------------------------------
def test_extract_vba_modules_from_bin(vba_bin):
    mods = extract_vba_modules(vba_bin)
    exp = _expected_modules()
    assert [m["name"] for m in mods] == list(exp)
    for m in mods:
        kind, code = exp[m["name"]]
        assert m["type"] == kind
        assert m["code"] == code
        assert "\r" not in m["code"]
        assert m["code"].startswith('Attribute VB_Name = "%s"' % m["name"])


def test_extract_large_module_and_form_kind():
    big = module_header("Big", "standard") + "".join(
        f"Sub P{i}()\r\n    Debug.Print \"line {i} {os.urandom(4).hex()}\"\r\nEnd Sub\r\n" for i in range(400))
    form = ('Attribute VB_Name = "UserForm1"\r\n'
            'Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}{00000000-0000-0000-0000-000000000000}"\r\n'
            "Attribute VB_PredeclaredId = True\r\nPrivate Sub UserForm_Click()\r\nEnd Sub\r\n")
    data = build_vba_project([("Big", "standard", big), ("UserForm1", "form", form)])
    mods = {m["name"]: m for m in extract_vba_modules(data)}
    assert mods["Big"]["code"] == big.replace("\r\n", "\n")
    assert len(mods["Big"]["code"]) > 4096
    assert mods["UserForm1"]["type"] == "form"


def test_kind_heuristics_without_project_stream():
    g = fileio._guess_kind
    assert g("ThisWorkbook", "ThisWorkbook.cls", "", set(), {}) == "document"
    assert g("Sheet7", "Sheet7.cls", 'Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"', set(), {}) == "document"
    assert g("Module9", "Module9.bas", "Sub X()\nEnd Sub", set(), {}) == "standard"
    assert g("UF", "UF.frm", "", set(), {}) == "form"
    assert g("UF2", "UF2.cls", "Attribute VB_PredeclaredId = True", {"uf2"}, {}) == "form"
    assert g("C", "C.cls", "VERSION 1.0 CLASS\nAttribute VB_Exposed = False", set(), {}) == "class"


def test_extract_vba_modules_none():
    assert extract_vba_modules(b"not a file") == []


def _xlsm_book(vba_bin) -> Workbook:
    b = _sample_book()
    b.add_sheet("Data")
    b.vba_project_bin = vba_bin
    b.vba_modules = extract_vba_modules(vba_bin)
    return b


def test_xlsm_preserves_vba_project(tmp_path, vba_bin):
    p = str(tmp_path / "macro.xlsm")
    warnings = save_workbook(_xlsm_book(vba_bin), p)
    assert warnings == []
    zf = zipfile.ZipFile(p)
    assert zf.read("xl/vbaProject.bin") == vba_bin
    ct = zf.read("[Content_Types].xml").decode()
    assert "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in ct
    assert 'PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"' in ct
    assert "spreadsheetml.sheet.main+xml" not in ct
    rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    assert "relationships/vbaProject" in rels and 'Target="vbaProject.bin"' in rels
    assert fileio.VBA_JSON_PART not in zf.namelist()
    assert 'codeName="ThisWorkbook"' in zf.read("xl/workbook.xml").decode()
    zf.close()

    # openpyxl can open it and keeps the VBA part
    wb = openpyxl.load_workbook(p, keep_vba=True)
    assert wb.vba_archive is not None
    assert wb.worksheets[0]["B1"].value == "=A1*2"
    wb.vba_archive.close()

    # oletools extracts the original macros from the saved file
    exp = _expected_modules()
    got = {m["name"]: (m["type"], m["code"]) for m in extract_vba_modules(p)}
    assert got == exp

    b2 = load_workbook(p)
    assert b2.vba_project_bin == vba_bin
    assert {m["name"]: (m["type"], m["code"]) for m in b2.vba_modules} == exp
    assert b2.vba_modified is False
    assert b2.active.cells[(0, 1)] == "=A1*2"

    # save again (second generation) - still byte-identical VBA part
    p2 = str(tmp_path / "macro2.xlsm")
    save_workbook(b2, p2)
    assert zipfile.ZipFile(p2).read("xl/vbaProject.bin") == vba_bin
    assert {m["name"]: (m["type"], m["code"]) for m in extract_vba_modules(p2)} == exp


def test_xlsm_edited_macros_stored_as_json(tmp_path, vba_bin):
    b = _xlsm_book(vba_bin)
    new_code = 'Attribute VB_Name = "Module1"\nSub Hello()\n    MsgBox "edited ✓"\nEnd Sub\n'
    for m in b.vba_modules:
        if m["name"] == "Module1":
            m["code"] = new_code
    b.vba_modules.append({"name": "Module2", "type": "standard", "code": "Sub New()\nEnd Sub\n"})
    b.vba_modified = True
    p = str(tmp_path / "edited.xlsm")
    warnings = save_workbook(b, p)
    assert any("original macros" in w for w in warnings)
    zf = zipfile.ZipFile(p)
    assert zf.read("xl/vbaProject.bin") == vba_bin
    payload = json.loads(zf.read(fileio.VBA_JSON_PART).decode("utf-8"))
    assert {m["name"] for m in payload["modules"]} >= {"Module1", "Module2"}
    ct = zf.read("[Content_Types].xml").decode()
    assert '<Default Extension="json" ContentType="application/json"/>' in ct
    zf.close()
    openpyxl.load_workbook(p, keep_vba=True).vba_archive.close()

    b2 = load_workbook(p)
    assert b2.vba_modified is False
    mods = {m["name"]: m for m in b2.vba_modules}
    assert mods["Module1"]["code"] == new_code
    assert mods["Module2"]["type"] == "standard"
    assert mods["ThisWorkbook"]["type"] == "document"
    assert b2.vba_project_bin == vba_bin
    # the original project is untouched
    assert {m["name"] for m in extract_vba_modules(p)} == set(_expected_modules())

    # re-saving without further edits (vba_modified is False after load) keeps the edits
    p2 = str(tmp_path / "edited2.xlsm")
    save_workbook(b2, p2)
    assert {m["name"]: m["code"] for m in load_workbook(p2).vba_modules}["Module1"] == new_code
    assert zipfile.ZipFile(p2).read("xl/vbaProject.bin") == vba_bin


def test_xlsm_macros_from_scratch(tmp_path):
    b = Workbook()
    b.active.set_raw(0, 0, "1")
    b.vba_modules = [{"name": "Module1", "type": "standard", "code": "Sub A()\n  Range(\"A1\").Value = 2\nEnd Sub\n"}]
    p = str(tmp_path / "scratch.xlsm")
    warnings = save_workbook(b, p)
    assert warnings and "PortableExcel" in warnings[0]
    zf = zipfile.ZipFile(p)
    assert fileio.VBA_JSON_PART in zf.namelist()
    assert "xl/vbaProject.bin" not in zf.namelist()
    assert "macroEnabled.main+xml" in zf.read("[Content_Types].xml").decode()
    zf.close()
    openpyxl.load_workbook(p)
    b2 = load_workbook(p)
    assert b2.vba_modules == b.vba_modules
    assert b2.vba_project_bin is None
    assert extract_vba_modules(p) == []


def test_xlsx_drops_macros_with_warning(tmp_path, vba_bin):
    p = str(tmp_path / "nomacro.xlsx")
    warnings = save_workbook(_xlsm_book(vba_bin), p)
    assert any("Macros are not saved in .xlsx" in w for w in warnings)
    zf = zipfile.ZipFile(p)
    assert "xl/vbaProject.bin" not in zf.namelist()
    assert "macroEnabled" not in zf.read("[Content_Types].xml").decode()
    b = load_workbook(p)
    assert b.vba_modules == [] and b.vba_project_bin is None


def test_xltm_template(tmp_path, vba_bin):
    p = str(tmp_path / "t.xltm")
    save_workbook(_xlsm_book(vba_bin), p)
    ct = zipfile.ZipFile(p).read("[Content_Types].xml").decode()
    assert "application/vnd.ms-excel.template.macroEnabled.main+xml" in ct
    b = load_workbook(p)
    assert {m["name"] for m in b.vba_modules} == set(_expected_modules())


# --------------------------------------------------------------------------
# encryption
# --------------------------------------------------------------------------
def test_encrypted_round_trip(tmp_path):
    pytest.importorskip("msoffcrypto")
    b = _sample_book()
    b.active.styles[(0, 0)] = {"bold": True}
    p = str(tmp_path / "enc.xlsx")
    save_workbook(b, p, password="secret")
    raw = open(p, "rb").read()
    assert raw[:8] == bytes.fromhex("D0CF11E0A1B11AE1")
    assert not zipfile.is_zipfile(p)
    with pytest.raises(PasswordRequired):
        load_workbook(p)
    with pytest.raises(PasswordRequired):
        load_workbook(p, password="wrong")
    b2 = load_workbook(p, password="secret")
    assert b2.active.cells == load_workbook_plain_cells(b, tmp_path)
    assert b2.active.styles[(0, 0)] == {"bold": True}


def load_workbook_plain_cells(b, tmp_path):
    p = str(tmp_path / "plain.xlsx")
    save_workbook(b, p)
    return load_workbook(p).active.cells


def test_encrypted_xlsm_keeps_vba(tmp_path, vba_bin):
    pytest.importorskip("msoffcrypto")
    p = str(tmp_path / "enc.xlsm")
    save_workbook(_xlsm_book(vba_bin), p, password="secret")
    b = load_workbook(p, password="secret")
    assert b.vba_project_bin == vba_bin
    assert {m["name"] for m in b.vba_modules} == set(_expected_modules())


# --------------------------------------------------------------------------
# errors
# --------------------------------------------------------------------------
def test_legacy_xls_rejected(tmp_path):
    p = tmp_path / "old.xls"
    p.write_bytes(write_cfb({"Workbook": b"\x09\x08" + b"\x00" * 600}))
    with pytest.raises(FileFormatError, match="Old .xls format is not supported"):
        load_workbook(str(p))
    p2 = tmp_path / "renamed.xlsx"          # xls content with an xlsx name
    p2.write_bytes(p.read_bytes())
    with pytest.raises(FileFormatError, match="Old .xls"):
        load_workbook(str(p2))


def test_garbage_and_missing_files(tmp_path):
    p = tmp_path / "bad.xlsx"
    p.write_bytes(b"this is not a zip")
    with pytest.raises(FileFormatError):
        load_workbook(str(p))
    p2 = tmp_path / "trunc.xlsx"
    good = tmp_path / "good.xlsx"
    save_workbook(_sample_book(), str(good))
    p2.write_bytes(good.read_bytes()[:200])
    with pytest.raises(FileFormatError):
        load_workbook(str(p2))
    with pytest.raises(FileFormatError):
        load_workbook(str(tmp_path / "missing.xlsx"))
    p3 = tmp_path / "doc.xlsx"
    with zipfile.ZipFile(p3, "w") as z:
        z.writestr("word/document.xml", "<x/>")
    with pytest.raises(FileFormatError):
        load_workbook(str(p3))
    with pytest.raises(FileFormatError):
        save_workbook(Workbook(), str(tmp_path / "x.ods"))
    with pytest.raises(FileFormatError):
        save_workbook(Workbook(), str(tmp_path / "x.xls"))


def test_public_constants():
    assert "*.xlsm" in fileio.SUPPORTED_OPEN and "*.csv" in fileio.SUPPORTED_OPEN
    assert fileio.SUPPORTED_SAVE.startswith("Excel Macro-Enabled Workbook (*.xlsm)")
    assert issubclass(PasswordRequired, Exception) and issubclass(FileFormatError, Exception)


# --------------------------------------------------------------------------
# pictures (barcodes) in .xlsx / .xlsm
# --------------------------------------------------------------------------
def _png(color, size=(40, 20), fmt="PNG"):
    from PIL import Image as PILImage
    b = io.BytesIO()
    PILImage.new("RGB", size, color).save(b, fmt)
    return b.getvalue()


BARCODE_META = {"symbology": "code128", "source": "A2", "text": "", "show_text": True,
                "module_px": 2, "bar_height_px": 40}


def _image_book():
    b = Workbook()
    s = b.active
    s.set_raw(1, 0, "12345")
    s.images = [
        {"id": "bc1", "name": "Barcode 1", "row": 1, "col": 2, "dx": 3, "dy": 4,
         "width": 120, "height": 50, "data": _png("black", (120, 50)), "fmt": "png",
         "barcode": dict(BARCODE_META)},
        {"id": "logo", "name": "Logo", "row": 5, "col": 0, "dx": 0, "dy": 0,
         "width": 40, "height": 20, "data": _png("blue"), "fmt": "png", "barcode": None},
    ]
    return b


def test_images_round_trip(tmp_path):
    from PIL import Image as PILImage
    b = _image_book()
    p = str(tmp_path / "pics.xlsx")
    assert save_workbook(b, p) == []
    zf = zipfile.ZipFile(p)
    meta = json.loads(zf.read(fileio.META_JSON_PART))
    assert [e["name"] for e in meta["images"]["Sheet1"]] == ["Barcode 1", "Logo"]
    assert all("data" not in e for e in meta["images"]["Sheet1"])
    assert meta["images"]["Sheet1"][0]["barcode"] == BARCODE_META
    assert '<Default Extension="json" ContentType="application/json"/>' in \
        zf.read("[Content_Types].xml").decode()
    zf.close()

    # openpyxl (and so Excel) sees two real pictures at the right anchors
    wb = openpyxl.load_workbook(p)
    imgs = wb.active._images
    assert len(imgs) == 2
    a0, a1 = imgs[0].anchor, imgs[1].anchor
    assert type(a0).__name__ == "OneCellAnchor"
    assert (a0._from.row, a0._from.col, a0._from.colOff, a0._from.rowOff) == (1, 2, 3 * 9525, 4 * 9525)
    assert (a0.ext.width, a0.ext.height) == (120 * 9525, 50 * 9525)
    assert a0.pic.nvPicPr.cNvPr.name == "Barcode 1"
    assert (a1._from.row, a1._from.col) == (5, 0)

    s = load_workbook(p).active
    assert len(s.images) == 2
    bc, logo = s.images
    for got, exp in zip(s.images, b.active.images):
        for key in ("id", "name", "row", "col", "dx", "dy", "width", "height", "fmt", "barcode"):
            assert got[key] == exp[key], key
    with PILImage.open(io.BytesIO(bc["data"])) as im:
        assert im.format == "PNG" and im.size == (120, 50)
    assert logo["barcode"] is None

    # a second generation keeps everything
    p2 = str(tmp_path / "pics2.xlsx")
    save_workbook(load_workbook(p), p2)
    assert [i["barcode"] for i in load_workbook(p2).active.images] == [BARCODE_META, None]


def test_images_without_meta_and_two_cell_anchor(tmp_path):
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.column_dimensions["B"].width = 20          # 145 px
    ws.row_dimensions[3].height = 30              # 40 px
    img = XLImage(io.BytesIO(_png("red", (50, 50))))
    img.anchor = TwoCellAnchor(_from=AnchorMarker(col=1, colOff=9525 * 5, row=2, rowOff=0),
                               to=AnchorMarker(col=3, colOff=9525 * 10, row=4, rowOff=9525 * 2))
    ws.add_image(img)
    jpg = XLImage(io.BytesIO(_png("green", (30, 10), "JPEG")))
    jpg.anchor = "E7"
    ws.add_image(jpg)
    p = str(tmp_path / "two.xlsx")
    wb.save(p)
    s = load_workbook(p).active
    assert len(s.images) == 2
    a, b = sorted(s.images, key=lambda i: i["row"])   # anchor types may be reordered
    assert (a["row"], a["col"], a["dx"], a["dy"]) == (2, 1, 5, 0)
    # cols B (145) + C (64) - 5 + 10 ; rows 3 (40) + 4 (20) + 2
    assert (a["width"], a["height"]) == (145 + 64 - 5 + 10, 40 + 20 + 2)
    assert a["barcode"] is None and a["fmt"] == "png"
    assert a["data"][:8] == b"\x89PNG\r\n\x1a\n"
    assert (b["row"], b["col"]) == (6, 4)
    assert b["data"][:8] == b"\x89PNG\r\n\x1a\n"     # JPEG converted to PNG
    assert len({a["id"], b["id"]}) == 2
    # user pictures survive a load/save round trip
    p2 = str(tmp_path / "two2.xlsx")
    save_workbook(load_workbook(p), p2)
    assert len(openpyxl.load_workbook(p2).active._images) == 2


def test_images_with_macros_and_multiple_sheets(tmp_path, vba_bin):
    b = _xlsm_book(vba_bin)
    b.sheets[1].images = _image_book().active.images
    p = str(tmp_path / "both.xlsm")
    save_workbook(b, p)
    zf = zipfile.ZipFile(p)
    assert zf.read("xl/vbaProject.bin") == vba_bin
    assert list(json.loads(zf.read(fileio.META_JSON_PART))["images"]) == ["Data"]
    b2 = load_workbook(p)
    assert b2.sheets[0].images == []
    assert [i["barcode"] for i in b2.sheets[1].images] == [BARCODE_META, None]
    assert {m["name"] for m in b2.vba_modules} == set(_expected_modules())


def test_bad_image_data_warns(tmp_path):
    b = Workbook()
    b.active.images = [{"id": "x", "name": "Broken", "row": 0, "col": 0, "dx": 0, "dy": 0,
                        "width": 10, "height": 10, "data": b"not an image", "fmt": "png",
                        "barcode": None}]
    p = str(tmp_path / "bad.xlsx")
    warnings = save_workbook(b, p)
    assert any("Broken" in w for w in warnings)
    assert load_workbook(p).active.images == []


def test_csv_warns_about_pictures(tmp_path):
    warnings = save_workbook(_image_book(), str(tmp_path / "p.csv"))
    assert any("Pictures" in w and "not saved" in w for w in warnings)

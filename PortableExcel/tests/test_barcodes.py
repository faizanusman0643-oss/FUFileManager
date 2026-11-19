"""Barcode generation: patterns verified by decoding the rendered PNGs."""
import io
import os
import sys

import pytest
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from portable_excel import barcodes as B  # noqa: E402
from portable_excel.model import Workbook  # noqa: E402
from portable_excel.engine import FormulaEngine, insert_delete  # noqa: E402


def png_row_modules(png, module_px, y=None):
    """Decode one pixel row of a 1-D barcode PNG into a '1'/'0' module string
    (quiet zones stripped)."""
    img = Image.open(io.BytesIO(png)).convert("L")
    w, h = img.size
    y = module_px * 2 + 5 if y is None else y
    px = [img.getpixel((x, y)) for x in range(w)]
    assert set(px) <= {0, 255}, "bars must be crisp (no anti-aliasing)"
    bits = "".join("1" if p == 0 else "0" for p in px)
    first, last = bits.index("1"), bits.rindex("1")
    core = bits[first:last + 1]
    out = []
    i = 0
    while i < len(core):
        j = i
        while j < len(core) and core[j] == core[i]:
            j += 1
        n = j - i
        assert n % module_px == 0, "module widths must be integer multiples"
        out.append(core[i] * (n // module_px))
        i = j
    return "".join(out), first, w - 1 - last


# --------------------------------------------------------------------------
# EAN-13
# --------------------------------------------------------------------------
L_CODES = ["0001101", "0011001", "0010011", "0111101", "0100011",
           "0110001", "0101111", "0111011", "0110111", "0001011"]
G_CODES = [c.translate(str.maketrans("01", "10"))[::-1] for c in L_CODES]
R_CODES = [c.translate(str.maketrans("01", "10")) for c in L_CODES]
PARITY = ["LLLLLL", "LLGLGG", "LLGGLG", "LLGGGL", "LGLLGG",
          "LGGLLG", "LGGGLL", "LGLGLG", "LGLGGL", "LGGLGL"]


def decode_ean13(mods):
    assert len(mods) == 95
    assert mods[:3] == "101" and mods[45:50] == "01010" and mods[92:] == "101"
    left = [mods[3 + 7 * i:10 + 7 * i] for i in range(6)]
    right = [mods[50 + 7 * i:57 + 7 * i] for i in range(6)]
    digits, parity = [], ""
    for p in left:
        if p in L_CODES:
            digits.append(L_CODES.index(p))
            parity += "L"
        else:
            digits.append(G_CODES.index(p))
            parity += "G"
    first = PARITY.index(parity)
    digits += [R_CODES.index(p) for p in right]
    return str(first) + "".join(map(str, digits))


@pytest.mark.parametrize("mp", [1, 2, 3])
def test_ean13_decode_from_png(mp):
    png, w, h = B.render_png("ean13", "590123412345", module_px=mp, show_text=True)
    mods, ql, qr = png_row_modules(png, mp)
    code = decode_ean13(mods)
    assert code == "5901234123457"
    assert B.gtin_check_digit(code[:12]) == code[12]
    assert ql == qr == 10 * mp       # quiet zones
    assert w == (95 + 20) * mp


def test_ean8_and_upca_guards():
    mods, _, _ = png_row_modules(B.render_png("ean8", "9638507")[0], 2)
    assert len(mods) == 67 and mods[:3] == "101" and mods[31:36] == "01010" and mods[-3:] == "101"
    mods, _, _ = png_row_modules(B.render_png("upca", "03600029145")[0], 2)
    assert len(mods) == 95
    # UPC-A is EAN-13 with a leading 0
    assert decode_ean13(mods) == "0" + "036000291452"


# --------------------------------------------------------------------------
# Code 128
# --------------------------------------------------------------------------
C128 = ["212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312",
        "132212", "221213", "221312", "231212", "112232", "122132", "122231", "113222",
        "123122", "123221", "223211", "221132", "221231", "213212", "223112", "312131",
        "311222", "321122", "321221", "312212", "322112", "322211", "212123", "212321",
        "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
        "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121",
        "313121", "211331", "231131", "213113", "213311", "213131", "311123", "311321",
        "331121", "312113", "312311", "332111", "314111", "221411", "431111", "111224",
        "111422", "121124", "121421", "141122", "141221", "112214", "112412", "122114",
        "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
        "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112",
        "421211", "212141", "214121", "412121", "111143", "111341", "131141", "114113",
        "114311", "411113", "411311", "113141", "114131", "311141", "411131", "211412",
        "211214", "211232"]
STOP = "2331112"


def widths(mods):
    out = []
    i = 0
    while i < len(mods):
        j = i
        while j < len(mods) and mods[j] == mods[i]:
            j += 1
        out.append(str(j - i))
        i = j
    return out


def decode_code128(mods):
    ws = widths(mods)
    assert "".join(ws[-7:]) == STOP, "stop pattern"
    vals = []
    for k in range(0, len(ws) - 7, 6):
        vals.append(C128.index("".join(ws[k:k + 6])))
    start, data, check = vals[0], vals[1:-1], vals[-1]
    assert start in (103, 104, 105)
    assert (start + sum((i + 1) * v for i, v in enumerate(data))) % 103 == check, "checksum"
    text = ""
    code = {103: "A", 104: "B", 105: "C"}[start]
    for v in data:
        if v == 99:
            code = "C"
        elif v == 100:
            code = "B"
        elif v == 101:
            code = "A"
        elif code == "C":
            text += f"{v:02d}"
        elif code == "B":
            text += chr(v + 32)
        else:
            text += chr(v + 32) if v < 64 else chr(v - 64)
    return text


@pytest.mark.parametrize("text", ["PE-Test", "Hello world!", "ABC123456789"])
def test_code128_decode_from_png(text):
    png, w, h = B.render_png("code128", text, module_px=2)
    mods, ql, qr = png_row_modules(png, 2)
    assert decode_code128(mods) == text
    assert ql == qr == 20


def test_code39_and_itf_render():
    mods, _, _ = png_row_modules(B.render_png("code39", "abc-1")[0], 2)
    assert mods.startswith("1") and len(mods) > 50
    png, w, h = B.render_png("itf", "123", module_px=1)
    assert w > 0 and B.normalize("itf", "123") == "0123"


# --------------------------------------------------------------------------
# QR
# --------------------------------------------------------------------------
def test_qr_matrix_and_quiet_zone():
    import segno
    q = segno.make("https://example.com", micro=False)
    n = 17 + 4 * q.version
    assert q.symbol_size(border=0) == (n, n)
    png, w, h = B.render_png("qr", "https://example.com", module_px=3)
    assert w == h == (n + 8) * 3
    img = Image.open(io.BytesIO(png)).convert("L")
    border = 4 * 3
    assert all(img.getpixel((x, y)) == 255 for x in range(w) for y in range(border))
    assert all(img.getpixel((x, y)) == 255 for y in range(h) for x in range(border))
    # top-left finder pattern starts right after the quiet zone
    assert img.getpixel((border, border)) == 0
    assert img.getpixel((border + 3 * 7 - 1, border)) == 0
    assert img.getpixel((border + 3 * 7, border)) == 255


# --------------------------------------------------------------------------
# validation / normalisation
# --------------------------------------------------------------------------
def test_validate_messages():
    assert B.validate("ean13", "123") == "EAN-13 needs 12 or 13 digits"
    assert B.validate("ean13", "5901234123458") == "EAN-13 check digit is wrong (expected 7)"
    assert B.validate("ean13", "5901234123457") is None
    assert B.validate("ean13", "59012341234a") == "EAN-13 can only contain digits"
    assert B.validate("ean8", "123") == "EAN-8 needs 7 or 8 digits"
    assert B.validate("upca", "036000291452") is None
    assert B.validate("code39", "abc_") .startswith("Code 39 cannot encode")
    assert B.validate("code128", "héllo") == "Code 128 can only encode ASCII characters"
    assert B.validate("itf", "12a") == "ITF can only contain digits"
    assert B.validate("qr", "") == "No data to encode"
    assert B.validate("nope", "1").startswith("Unknown barcode type")
    assert "datamatrix" not in B.SYMBOLOGIES


def test_normalize():
    assert B.normalize("ean13", "590123412345") == "5901234123457"
    assert B.normalize("ean8", "9638507") == "96385074"
    assert B.normalize("upca", "03600029145") == "036000291452"
    assert B.normalize("code39", "abc") == "ABC"
    assert B.normalize("code128", "abc") == "abc"
    with pytest.raises(ValueError):
        B.render_png("ean13", "12")


# --------------------------------------------------------------------------
# sheet integration
# --------------------------------------------------------------------------
def make_book():
    book = Workbook()
    eng = FormulaEngine(book)
    return book, eng, book.sheets[0]


def test_make_barcode_image_from_numeric_cell():
    book, eng, sh = make_book()
    book.set_cell(sh, 1, 0, "1234567890128")
    img = B.make_barcode_image(sh, "A2", None, "ean13", engine=eng, row=1, col=1)
    assert img["barcode"]["last_value"] == "1234567890128"
    assert img["barcode"]["error"] is None
    assert img["name"] == "Barcode 1" and img["fmt"] == "png" and len(img["id"]) == 32
    mods, _, _ = png_row_modules(img["data"], 2)
    assert decode_ean13(mods) == "1234567890128"
    sh.images.append(img)
    img2 = B.make_barcode_image(sh, None, "fixed", "qr", engine=eng, row=0, col=0)
    assert img2["name"] == "Barcode 2"
    assert (img2["width"], img2["height"]) == Image.open(io.BytesIO(img2["data"])).size


def test_refresh_after_cell_change_and_invalid_placeholder():
    book, eng, sh = make_book()
    book.set_cell(sh, 0, 0, "PE-1")
    img = B.make_barcode_image(sh, "A1", None, "code128", engine=eng, row=0, col=2)
    sh.images.append(img)
    assert B.refresh_barcodes(book, eng) == []
    book.set_cell(sh, 0, 0, "PE-2")
    eng.invalidate()
    assert B.refresh_barcodes(book, eng, changed_cells=[(5, 5)]) == []
    upd = B.refresh_barcodes(book, eng, sh, changed_cells=[(sh, 0, 0)])
    assert upd == [img]
    assert decode_code128(png_row_modules(img["data"], 2)[0]) == "PE-2"
    # invalid data -> red placeholder, no exception
    img3 = B.make_barcode_image(sh, "A1", None, "ean13", engine=eng, row=3, col=3)
    assert img3["barcode"]["error"] == "EAN-13 can only contain digits"
    pic = Image.open(io.BytesIO(img3["data"])).convert("RGB")
    assert (204, 0, 0) in {pic.getpixel((x, y)) for x in range(pic.size[0]) for y in range(pic.size[1])}
    sh.images.append(img3)
    book.set_cell(sh, 0, 0, "590123412345")
    eng.invalidate()
    B.refresh_barcodes(book, eng)
    assert img3["barcode"]["error"] is None
    assert decode_ean13(png_row_modules(img3["data"], 2)[0]) == "5901234123457"


# --------------------------------------------------------------------------
# insert/delete keeps anchors and sources in step
# --------------------------------------------------------------------------
def test_insert_delete_shifts_images_and_sources():
    book, eng, sh = make_book()
    book.set_cell(sh, 5, 0, "X1")
    img = B.make_barcode_image(sh, "A6", None, "code128", engine=eng, row=5, col=2)
    sh.images.append(img)
    insert_delete(book, sh, "row", 0, 2)
    assert img["row"] == 7 and img["barcode"]["source"] == "A8"
    insert_delete(book, sh, "col", 0, 1)
    assert img["col"] == 3 and img["barcode"]["source"] == "B8"
    insert_delete(book, sh, "row", 10, 3)          # after the image: no change
    assert img["row"] == 7
    insert_delete(book, sh, "row", 6, -3)          # deletes rows 7..9 incl. anchor and source
    assert img["row"] == 6
    assert img["barcode"]["source"] == "#REF!"
    eng.invalidate()
    B.refresh_barcodes(book, eng)
    assert img["barcode"]["error"] == "Source cell was deleted"
    insert_delete(book, sh, "col", 3, -1)
    assert img["col"] == 3


# --------------------------------------------------------------------------
# VBA bridge
# --------------------------------------------------------------------------
def test_vba_pe_addbarcode_and_shapes():
    from portable_excel.vba import VBARuntime, VBAHost

    class H(VBAHost):
        def __init__(self):
            super().__init__()
            self.prints = []

        def debug_print(self, t):
            self.prints.append(t)
    book, eng, sh = make_book()
    host = H()
    rt = VBARuntime(book, eng, host)
    rt.load_modules([{"name": "Module1", "type": "standard", "code": '''
Sub Go()
    Range("A2").Value = "ITEM-42"
    Dim nm As String
    nm = Application.Run("PE_AddBarcode", "A2", "code128", "B2")
    Debug.Print nm, ActiveSheet.Shapes.Count, ActiveSheet.Shapes(1).Name
    Debug.Print ActiveSheet.Shapes(nm).Width > 0, ActiveSheet.Shapes(1).TopLeftCell.Address
    Range("A2").Value = "ITEM-43"
    ActiveSheet.Shapes(1).Top = 0
    ActiveSheet.Shapes("Barcode 1").Delete
    Debug.Print ActiveSheet.Shapes.Count
End Sub
Sub Keep()
    Range("A2").Value = "ITEM-1"
    Application.Run "PE_AddBarcode", "A2", "code128", "C3"
    Range("A2").Value = "ITEM-2"
End Sub'''}])
    rt.run("Go")
    assert host.prints[0].split() == ["Barcode", "1", "1", "Barcode", "1"]
    assert host.prints[1].split() == ["True", "$B$2"]
    assert host.prints[2].strip() == "0"
    rt.run("Keep")
    img = sh.images[0]
    assert (img["row"], img["col"]) == (2, 2)
    assert decode_code128(png_row_modules(img["data"], 2)[0]) == "ITEM-2"


# --------------------------------------------------------------------------
# ISBT 128 (blood bank labels) — Code 128 with exact data
# --------------------------------------------------------------------------
ISBT_SAMPLES = ["=%6200", "=<E0212V00", "=PK0262601234400", "&>0262352359"]


def text_of_png(png, mp):
    """Human-readable band present below the bars (non-white pixels)."""
    img = Image.open(io.BytesIO(png)).convert("L")
    w, h = img.size
    return any(img.getpixel((x, y)) < 128 for y in range(mp * 2 + 61, h) for x in range(w))


@pytest.mark.parametrize("data", ISBT_SAMPLES)
@pytest.mark.parametrize("sym", ["code128", "isbt128"])
def test_isbt_strings_encode_exactly(sym, data):
    assert B.validate(sym, data) is None
    assert B.normalize(sym, data) == data
    assert B.human_text(sym, data) == data
    png, w, h = B.render_png(sym, data, module_px=2)
    mods, ql, qr = png_row_modules(png, 2)
    assert decode_code128(mods) == data
    assert text_of_png(png, 2)


def test_isbt_uses_code_set_c_optimally():
    vals = B.code128_values("=PK0262601234400")
    # start B, "=PK0" in B, switch to C, then six digit pairs, checksum
    assert vals[0] == 104
    assert vals[1:5] == [ord(c) - 32 for c in "=PK0"]
    assert vals[5] == 99
    assert vals[6:12] == [26, 26, 1, 23, 44, 0]
    assert len(vals) == 13
    assert B.code128_values("&>0262352359")[3] == 99
    # an all-digit even string starts directly in code set C
    assert B.code128_values("123456")[0] == 105 and len(B.code128_values("123456")) == 5


def test_isbt_validation_and_warnings():
    assert B.SYMBOLOGIES["isbt128"] == "ISBT 128 (Code 128)"
    assert B.validate("isbt128", "6200").startswith("ISBT 128 data must start")
    assert B.validate("isbt128", "=").startswith("ISBT 128 data must start")
    for d in ISBT_SAMPLES:
        assert B.warnings("isbt128", d) == []
    assert B.isbt_identifier("=%6200")[1] == "ABO/Rh blood group"
    assert B.isbt_identifier("=PK0262601234400")[1] == "Donation identification number"
    w = B.warnings("isbt128", "=@123")
    assert w == ["Unknown ISBT 128 data identifier '=@'"]
    assert B.validate("isbt128", "=@123") is None          # warning only, still renders
    mods, _, _ = png_row_modules(B.render_png("isbt128", "=@123")[0], 2)
    assert decode_code128(mods) == "=@123"
    assert B.warnings("isbt128", "=%62") == ["ISBT 128 abo/rh blood group should have 6 characters (has 4)"]


def test_din_check_character_iso7064():
    alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ*"
    body = "PK02626012344"
    c = B.din_check_character("=PK0262601234400")
    assert c == B.iso7064_mod37_2(body)
    # ISO 7064 MOD 37-2 property: sum(v_i * 2^(n-i+1)) + check == 1 (mod 37)
    n = len(body)
    total = sum(alphabet.index(ch) * pow(2, n - i, 37) for i, ch in enumerate(body))
    assert (total + alphabet.index(c)) % 37 == 1
    assert B.din_check_character("=%6200") is None
    png_plain = B.render_png("isbt128", "=PK0262601234400")[0]
    png_box = B.render_png("isbt128", "=PK0262601234400", din_check=True)[0]
    assert png_plain != png_box
    # bars are identical with or without the check box
    assert png_row_modules(png_plain, 2)[0] == png_row_modules(png_box, 2)[0]


def test_isbt_from_xlookup_formula_cells():
    book, eng, sh = make_book()
    book.set_cell(sh, 1, 6, "A+")
    book.set_cell(sh, 1, 0, '=IFERROR(_xlfn.XLOOKUP(G2,{"A+","B+"},{"=%6200","=%7300"}),"")')
    img = B.make_barcode_image(sh, "A2", None, "isbt128", engine=eng, row=1, col=1)
    assert img["barcode"]["error"] is None and img["barcode"]["warning"] is None
    assert decode_code128(png_row_modules(img["data"], 2)[0]) == "=%6200"
    sh.images.append(img)
    book.set_cell(sh, 1, 6, "B+")
    eng.invalidate()
    assert B.refresh_barcodes(book, eng) == [img]
    assert decode_code128(png_row_modules(img["data"], 2)[0]) == "=%7300"
    book.set_cell(sh, 1, 6, "?")          # XLOOKUP misses -> "" -> no data
    eng.invalidate()
    B.refresh_barcodes(book, eng)
    assert img["barcode"]["error"] == "No data to encode"

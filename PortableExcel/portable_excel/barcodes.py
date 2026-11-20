"""Barcode generation for worksheet images (no Qt).

Bars are rendered from the module patterns of python-barcode (1-D codes) and
segno (QR) with Pillow: integer module widths, no anti-aliasing, white
background.  Data Matrix is not offered: no pure-Python encoder is installed.

Code 128 (and "isbt128", which is Code 128 with ISBT 128 data-identifier
checks) is encoded by our own encoder: the data is kept exactly as given
(e.g. "=%6200", "&>0262352359") and code sets A/B/C are chosen by dynamic
programming for the fewest symbols (code set C for digit runs).  The text
under the bars is exactly the encoded data.

Sheet.images barcode dicts get these runtime keys in addition to the
model.py schema (fileio does not need to save them):
    last_value  the data last rendered (str), or ("#error", reason) when the
                source could not be read; refresh_barcodes() compares the
                source cell's current value with it to skip unchanged codes.
    error       None, or the reason the data could not be encoded; the image
                then shows a red "Invalid barcode: <reason>" placeholder.
    warning     None, or non-blocking remarks (e.g. unknown ISBT 128 data
                identifier); the barcode is still rendered.
    din_check   (isbt128 only) bool: draw the ISO 7064 MOD 37-2 DIN check
                character in a box after the human-readable DIN. Off by default.
"""
from __future__ import annotations

import io
import re
import uuid

from PIL import Image, ImageDraw, ImageFont

SYMBOLOGIES = {
    "code128": "Code 128",
    "isbt128": "ISBT 128 (Code 128)",
    "code39": "Code 39",
    "ean13": "EAN-13",
    "ean8": "EAN-8",
    "upca": "UPC-A",
    "itf": "ITF (Interleaved 2 of 5)",
    "qr": "QR Code",
}

CODE39_CHARS = set("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%")
_GTIN = {"ean13": (12, "EAN-13"), "ean8": (7, "EAN-8"), "upca": (11, "UPC-A")}


# --------------------------------------------------------------------------
# validation / normalisation
# --------------------------------------------------------------------------
def gtin_check_digit(digits: str) -> str:
    """Check digit for EAN-13/EAN-8/UPC-A data (without the check digit)."""
    total = 0
    for i, ch in enumerate(reversed(digits)):
        total += int(ch) * (3 if i % 2 == 0 else 1)
    return str((10 - total % 10) % 10)


def _sym(symbology: str) -> str:
    s = (symbology or "").strip().lower().replace("-", "").replace(" ", "").replace("_", "")
    s = {"upc": "upca", "qrcode": "qr", "interleaved2of5": "itf", "i2of5": "itf"}.get(s, s)
    if s not in SYMBOLOGIES:
        raise ValueError(f"Unknown barcode type: {symbology}")
    return s


# --------------------------------------------------------------------------
# Code 128 encoder (optimal code-set selection, exact data)
# --------------------------------------------------------------------------
C128_WIDTHS = [
    "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312",
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
C128_STOP = "2331112"
_START = {"A": 103, "B": 104, "C": 105}
_SWITCH = {"A": 101, "B": 100, "C": 99}


def _val_a(ch):
    o = ord(ch)
    if 32 <= o <= 95:
        return o - 32
    if 0 <= o < 32:
        return o + 64
    return None


def _val_b(ch):
    o = ord(ch)
    return o - 32 if 32 <= o <= 127 else None


def code128_values(text: str):
    """Symbol values (start .. data .. checksum, without stop) using the
    fewest symbols: dynamic programming over code sets A/B/C."""
    n = len(text)
    if n == 0:
        raise ValueError("No data to encode")
    INF = 10 ** 9
    # best[i][set] = (cost, path) for encoding text[i:] being in `set`
    sets = ("A", "B", "C")
    best = [{k: (INF, None) for k in sets} for _ in range(n + 1)]
    for k in sets:
        best[n][k] = (0, [])
    for i in range(n - 1, -1, -1):
        for cur in sets:
            options = []
            for nxt in sets:
                sw = 0 if nxt == cur else 1
                if nxt == "C":
                    if i + 1 < n and text[i].isdigit() and text[i + 1].isdigit() \
                            and text[i].isascii() and text[i + 1].isascii():
                        rest = best[i + 2][nxt]
                        if rest[0] < INF:
                            vals = ([_SWITCH[nxt]] if sw else []) + [int(text[i:i + 2])]
                            options.append((sw + 1 + rest[0], vals + rest[1]))
                else:
                    v = _val_a(text[i]) if nxt == "A" else _val_b(text[i])
                    if v is not None:
                        rest = best[i + 1][nxt]
                        if rest[0] < INF:
                            vals = ([_SWITCH[nxt]] if sw else []) + [v]
                            options.append((sw + 1 + rest[0], vals + rest[1]))
            if options:
                best[i][cur] = min(options, key=lambda t: t[0])
    # choose start set: path from best[0][s] must not begin with a switch
    starts = []
    for k in sets:
        c, path = best[0][k]
        if c < INF and not (path and path[0] in (99, 100, 101) and False):
            starts.append((c, k, path))
    if not starts:
        raise ValueError("Code 128 can only encode ASCII characters")
    c, k, path = min(starts, key=lambda t: (t[0], "CBA".index(t[1])))
    vals = [_START[k]] + path
    chk = (vals[0] + sum((i + 1) * v for i, v in enumerate(vals[1:]))) % 103
    return vals + [chk]


def code128_modules(text: str) -> str:
    out = []
    for v in code128_values(text):
        w = C128_WIDTHS[v]
        out.append("".join(("1" if j % 2 == 0 else "0") * int(d) for j, d in enumerate(w)))
    out.append("".join(("1" if j % 2 == 0 else "0") * int(d) for j, d in enumerate(C128_STOP)))
    return "".join(out)


# --------------------------------------------------------------------------
# ISBT 128
# --------------------------------------------------------------------------
ISBT_IDENTIFIERS = {
    "=%": ("ABO/Rh blood group", 6), "=<": ("Product code", 10),
    "=>": ("Expiration date", 8), "&>": ("Expiration date and time", 12),
    "=*": ("Collection date", 8), "&*": ("Collection date and time", 12),
    "=}": ("Production date", 8), "&}": ("Production date and time", 12),
    "&(": ("Special testing: general", None), "={": ("Special testing: red blood cell antigens", None),
    "&{": ("Special testing: red blood cell antigens", None),
    "=\\": ("Special testing: platelet HLA and antigens", None),
    "&\\": ("Special testing: HLA-A and -B alleles", None),
    "=[": ("Special testing: HLA-DRB1 alleles", None),
    "=)": ("Container manufacturer and catalog number", None),
    "&)": ("Container lot number", None), "=;": ("Donor identification number", None),
    "='": ("Staff member identification number", None), "=-": ("Manufacturer and catalog number", None),
    "&-": ("Lot number", None), "=#": ("Infectious markers", None),
    "&#": ("Infectious markers", None), "=$": ("Product consignment", None),
    "&$": ("Dimensions", None), "&\"": ("Red cell antigens with test history", None),
    "=+": ("Flexible date and time", None), "=,": ("Product divisions", None),
    "=/": ("Processing facility information code", None), "=.": ("Processor product id", None),
    "&,": ("MPHO lot number", None), "&+": ("MPHO supplemental identification", None),
    "&!": ("Compound message", None), "=!": ("Compound message", None),
    "&%": ("Patient date of birth", None), "&&": ("Patient identification number", None),
    "&;": ("Expiration month and year", None), "&:": ("Transfusion transmitted infection marker", None),
    "&'": ("Global registration identifier for donors", None), "=&": ("Single European Code", None),
}
_DIN_FIRST = set("ABCDEFGHIJKLMNPQRSTUVWXYZ123456789")


def isbt_identifier(text: str):
    """(identifier, description) for ISBT 128 data, or (None, None)."""
    if len(text) < 2 or text[0] not in "=&":
        return None, None
    if text[0] == "=" and text[1] in _DIN_FIRST:
        return "=α", "Donation identification number"
    di = text[:2]
    if di in ISBT_IDENTIFIERS:
        return di, ISBT_IDENTIFIERS[di][0]
    return di, None


def iso7064_mod37_2(data: str) -> str:
    """ISO/IEC 7064 MOD 37-2 check character (as used for the ISBT 128 DIN
    keyboard-entry check character): 0-9, A-Z, '*'."""
    alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ*"
    p = 0
    for ch in data.upper():
        v = alphabet.index(ch)
        p = ((p + v) * 2) % 37
    return alphabet[(38 - p) % 37]


def din_check_character(text: str):
    """Check character for an ISBT 128 DIN "=αppppyynnnnnnff" (computed over the
    13 characters after "=", flags excluded), or None if text is not a DIN."""
    if isbt_identifier(text)[0] != "=α" or len(text) < 14:
        return None
    body = text[1:14]
    if not body.isalnum() or not body.isascii():
        return None
    return iso7064_mod37_2(body)


def warnings(symbology: str, text: str) -> list:
    """Non-blocking warnings (e.g. unknown ISBT 128 data identifier)."""
    try:
        s = _sym(symbology)
    except ValueError:
        return []
    t = "" if text is None else str(text)
    out = []
    if s == "isbt128" and len(t) >= 2 and t[0] in "=&":
        di, desc = isbt_identifier(t)
        if desc is None:
            out.append(f"Unknown ISBT 128 data identifier {t[:2]!r}")
        elif di == "=α":
            if len(t) != 16:
                out.append(f"ISBT 128 donation identification number should have 16 characters (has {len(t)})")
        else:
            n = ISBT_IDENTIFIERS[di][1]
            if n is not None and len(t) != n:
                out.append(f"ISBT 128 {desc.lower()} should have {n} characters (has {len(t)})")
    return out


def validate(symbology: str, text: str) -> str | None:
    try:
        s = _sym(symbology)
    except ValueError as e:
        return str(e)
    t = "" if text is None else str(text)
    if t == "":
        return "No data to encode"
    if s in _GTIN:
        n, label = _GTIN[s]
        if not t.isdigit() or not t.isascii():
            return f"{label} can only contain digits"
        if len(t) not in (n, n + 1):
            return f"{label} needs {n} or {n + 1} digits"
        if len(t) == n + 1:
            want = gtin_check_digit(t[:-1])
            if t[-1] != want:
                return f"{label} check digit is wrong (expected {want})"
        return None
    if s == "itf":
        if not t.isdigit() or not t.isascii():
            return "ITF can only contain digits"
        return None
    if s == "code39":
        bad = sorted({ch for ch in t.upper() if ch not in CODE39_CHARS})
        if bad:
            return "Code 39 cannot encode: " + " ".join(repr(b) for b in bad)
        return None
    if s == "isbt128":
        if len(t) < 2 or t[0] not in "=&":
            return "ISBT 128 data must start with a data identifier ('=' or '&' plus one character)"
    if s in ("code128", "isbt128"):
        bad = sorted({ch for ch in t if ord(ch) > 127})
        if bad:
            return "Code 128 can only encode ASCII characters"
        if len(t) > 80:
            return "Code 128 text is too long (max 80 characters)"
        return None
    if s == "qr":
        if len(t.encode("utf-8")) > 2900:
            return "Too much data for a QR code"
        return None
    return None


def normalize(symbology: str, text: str) -> str:
    s = _sym(symbology)
    t = "" if text is None else str(text)
    if s in _GTIN:
        n, _ = _GTIN[s]
        if len(t) == n and t.isdigit():
            t += gtin_check_digit(t)
        return t
    if s == "code39":
        return t.upper()
    if s == "itf":
        return t if len(t) % 2 == 0 else "0" + t
    return t


# --------------------------------------------------------------------------
# module patterns
# --------------------------------------------------------------------------
def modules(symbology: str, text: str):
    """1-D: string of '1'/'0' modules (no quiet zone). QR: list of rows of 0/1."""
    s = _sym(symbology)
    err = validate(s, text)
    if err:
        raise ValueError(err)
    t = normalize(s, text)
    if s in ("code128", "isbt128"):
        return code128_modules(t)
    if s == "qr":
        import segno
        q = segno.make(t, micro=False)
        return [[1 if v else 0 for v in row] for row in q.matrix]
    import barcode
    cls = barcode.get_barcode_class(s)
    if s == "code39":
        bc = cls(t, add_checksum=False)
    elif s in _GTIN:
        bc = cls(t[:-1]) if s != "upca" else cls(t[:-1])
    else:
        bc = cls(t)
    return bc.build()[0]


def human_text(symbology: str, text: str) -> str:
    s = _sym(symbology)
    t = normalize(s, text)
    if s == "code39":
        return "*" + t + "*"
    return t


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------
def _font(size: int):
    try:
        return ImageFont.load_default(size=size)
    except TypeError:               # very old Pillow
        return ImageFont.load_default()


def _png(img: Image.Image, dpi: int) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(dpi, dpi))
    return buf.getvalue()


def render_png(symbology: str, text: str, *, module_px: int = 2, bar_height_px: int = 60,
               show_text: bool = True, quiet_zone: bool = True, dpi: int = 96,
               din_check: bool = False):
    """Render a barcode; returns (png_bytes, width, height). Raises ValueError
    with a readable message for invalid data."""
    s = _sym(symbology)
    mp = max(1, int(module_px))
    pat = modules(s, text)
    if s == "qr":
        n = len(pat)
        q = 4 if quiet_zone else 0
        size = (n + 2 * q) * mp
        img = Image.new("L", (size, size), 255)
        d = ImageDraw.Draw(img)
        for y, row in enumerate(pat):
            for x, v in enumerate(row):
                if v:
                    x0 = (x + q) * mp
                    y0 = (y + q) * mp
                    d.rectangle([x0, y0, x0 + mp - 1, y0 + mp - 1], fill=0)
        return _png(img.convert("RGB"), dpi), size, size
    q = 10 if quiet_zone else 0
    width = (len(pat) + 2 * q) * mp
    bh = max(1, int(bar_height_px))
    top = mp * 2 if quiet_zone else 0
    caption = human_text(s, text) if show_text else ""
    font = None
    th = 0
    if caption:
        fsize = max(10, min(bh // 4, 9 * mp + 4))
        font = _font(fsize)
        box = font.getbbox(caption)
        th = box[3] - min(box[1], 0) + 4
    height = top + bh + th + (mp * 2 if quiet_zone else 0)
    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.fontmode = "1"
    x = q * mp
    i = 0
    while i < len(pat):
        if pat[i] == "1":
            j = i
            while j < len(pat) and pat[j] == "1":
                j += 1
            d.rectangle([x + i * mp, top, x + j * mp - 1, top + bh - 1], fill=(0, 0, 0))
            i = j
        else:
            i += 1
    if caption:
        chk = din_check_character(caption) if (din_check and s == "isbt128") else None
        tw = d.textlength(caption, font=font)
        box_w = 0
        if chk:
            cw = d.textlength(chk, font=font)
            box_w = int(cw) + 8
        x0 = (width - tw - box_w) / 2
        d.text((x0, top + bh + 2), caption, fill=(0, 0, 0), font=font)
        if chk:
            bx = int(x0 + tw + 4)
            by = top + bh + 1
            d.rectangle([bx, by, bx + box_w - 2, by + th - 2], outline=(0, 0, 0))
            d.text((bx + 3, top + bh + 2), chk, fill=(0, 0, 0), font=font)
    return _png(img, dpi), width, height


def render_error_png(message: str, dpi: int = 96):
    msg = f"Invalid barcode: {message}"
    font = _font(12)
    tmp = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    tw = int(tmp.textlength(msg, font=font)) + 12
    w, h = max(tw, 80), 24
    img = Image.new("RGB", (w, h), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w - 1, h - 1], outline=(204, 0, 0))
    d.text((6, 5), msg, fill=(204, 0, 0), font=font)
    return _png(img, dpi), w, h


# --------------------------------------------------------------------------
# sheet integration
# --------------------------------------------------------------------------
_REF = re.compile(r"^(?:(?:'([^']+)'|([^!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$")


def _parse_source(sheet, source):
    from .model import name_to_col
    m = _REF.match(source.strip())
    if not m:
        return None
    return int(m.group(4)) - 1, name_to_col(m.group(3))


def source_text(sheet, source, engine=None):
    """Text for the barcode from a source cell, as displayed (no 1.23E+12)."""
    pos = _parse_source(sheet, source)
    if pos is None:
        raise ValueError("Source cell was deleted" if "#REF" in source.upper()
                         else f"Bad source cell: {source}")
    r, c = pos
    from .model import literal_value, XLError
    if engine is not None:
        v = engine.value(sheet, r, c)
    else:
        raw = sheet.get_raw(r, c)
        v = literal_value(raw) if not raw.startswith("=") else raw
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, XLError):
        raise ValueError(f"Source cell shows {v.code}")
    if isinstance(v, (int, float)):
        if float(v).is_integer() and abs(v) < 1e16:
            return str(int(v))
        if engine is not None:
            t = engine.display_text(sheet, r, c)
            if "E+" not in t.upper() and "#" not in t:
                return t
        return repr(float(v))
    return str(v)


def _render_spec(bc, text):
    """Render according to a barcode dict; returns (png, w, h, error)."""
    try:
        if text is None:
            raise ValueError("No data to encode")
        err = validate(bc["symbology"], text)
        if err:
            raise ValueError(err)
        png, w, h = render_png(bc["symbology"], text, module_px=bc.get("module_px", 2),
                               bar_height_px=bc.get("bar_height_px", 60),
                               show_text=bc.get("show_text", True),
                               din_check=bc.get("din_check", False))
        bc["warning"] = "; ".join(warnings(bc["symbology"], text)) or None
        return png, w, h, None
    except ValueError as e:
        png, w, h = render_error_png(str(e))
        return png, w, h, str(e)


def _next_name(sheet):
    n = 0
    for img in getattr(sheet, "images", []):
        m = re.match(r"Barcode (\d+)$", img.get("name", ""))
        if m:
            n = max(n, int(m.group(1)))
    return f"Barcode {n + 1}"


def make_barcode_image(sheet, source, text, symbology, *, engine=None, row, col, **opts):
    sym = _sym(symbology)
    module_px = int(opts.get("module_px", 4 if sym == "qr" else 2))
    bc = {"symbology": sym, "source": source or None, "text": "" if text is None else str(text),
          "show_text": bool(opts.get("show_text", True)), "module_px": module_px,
          "bar_height_px": int(opts.get("bar_height_px", 60)), "last_value": None, "error": None,
          "warning": None}
    if sym == "isbt128":
        bc["din_check"] = bool(opts.get("din_check", False))
    if source:
        try:
            value = source_text(sheet, source, engine)
        except ValueError as e:
            value = None
            png, w, h = render_error_png(str(e))
            bc["error"] = str(e)
        if value is not None:
            png, w, h, bc["error"] = _render_spec(bc, value)
        bc["last_value"] = value
    else:
        png, w, h, bc["error"] = _render_spec(bc, bc["text"])
        bc["last_value"] = bc["text"]
    return {"id": uuid.uuid4().hex, "name": opts.get("name") or _next_name(sheet),
            "row": int(row), "col": int(col), "dx": int(opts.get("dx", 0)),
            "dy": int(opts.get("dy", 0)), "width": w, "height": h, "data": png, "fmt": "png",
            "barcode": bc}


def rerender(img, sheet, engine=None, force=False):
    """Re-render one barcode image if its data changed. Returns True if updated."""
    bc = img.get("barcode")
    if not bc:
        return False
    if bc.get("source"):
        try:
            value = source_text(sheet, bc["source"], engine)
            err = None
        except ValueError as e:
            value, err = None, str(e)
    else:
        value, err = bc.get("text", ""), None
    key = value if err is None else ("#error", err)
    if not force and key == bc.get("last_value"):
        return False
    if err is not None:
        png, w, h = render_error_png(err)
    else:
        png, w, h, err = _render_spec(bc, value)
    bc["last_value"] = key
    bc["error"] = err
    img["data"], img["width"], img["height"] = png, w, h
    return True


def refresh_barcodes(book, engine, sheet=None, changed_cells=None):
    """Re-render barcodes whose source cell value changed.

    changed_cells: optional iterable of (row, col) or (sheet, row, col); when
    given, only barcodes whose source is among them are considered."""
    sheets = [sheet] if sheet is not None else list(book.sheets)
    changed = None
    if changed_cells is not None:
        changed = set()
        for t in changed_cells:
            if len(t) == 3:
                changed.add((id(t[0]), t[1], t[2]))
            else:
                changed.add((None, t[0], t[1]))
    out = []
    for sh in sheets:
        for img in getattr(sh, "images", []):
            bc = img.get("barcode")
            if not bc or not bc.get("source"):
                continue
            if changed is not None:
                pos = _parse_source(sh, bc["source"])
                if pos is not None and (id(sh), pos[0], pos[1]) not in changed and \
                        (None, pos[0], pos[1]) not in changed:
                    continue
            if rerender(img, sh, engine):
                out.append(img)
    return out

"""Barcode generation for worksheet images (no Qt).

Bars are rendered from the module patterns of python-barcode (1-D codes) and
segno (QR) with Pillow: integer module widths, no anti-aliasing, white
background.  Data Matrix is not offered: no pure-Python encoder is installed.
"""
from __future__ import annotations

import io
import re
import uuid

from PIL import Image, ImageDraw, ImageFont

SYMBOLOGIES = {
    "code128": "Code 128",
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
    if s == "code128":
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
               show_text: bool = True, quiet_zone: bool = True, dpi: int = 96):
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
        tw = d.textlength(caption, font=font)
        d.text(((width - tw) / 2, top + bh + 2), caption, fill=(0, 0, 0), font=font)
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
                               show_text=bc.get("show_text", True))
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
          "bar_height_px": int(opts.get("bar_height_px", 60)), "last_value": None, "error": None}
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

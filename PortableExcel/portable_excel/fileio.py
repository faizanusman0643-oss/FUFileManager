"""File input/output for PortableExcel.

Public API::

    load_workbook(path, password=None) -> Workbook
    save_workbook(book, path, engine=None, password=None) -> list[str]   # warnings
    extract_vba_modules(path_or_bytes) -> list[dict]                     # {"name","type","code"}

Supported: .xlsx/.xlsm/.xltx/.xltm (openpyxl), .csv, password-protected
(ECMA-376 encrypted) OOXML files via msoffcrypto-tool.  VBA source is read
with oletools.olevba; the original xl/vbaProject.bin is preserved on .xlsm
save, and edited / newly written macros are stored as JSON in the extra zip
part ``customXml/portableexcel_vba.json``.

Third-party packages are imported lazily so that a missing package produces
a clear FileFormatError instead of an ImportError at application start.
"""
from __future__ import annotations

import csv
import datetime as _dt
import io
import json
import logging
import os
import re
import struct
import tempfile
import zipfile
from typing import Dict, List, Optional, Tuple

from .model import Sheet, Workbook, literal_value


class PasswordRequired(Exception):
    """The file is encrypted and no password, or a wrong one, was given."""


class FileFormatError(Exception):
    """Unreadable or unsupported file; str(exc) is a message for the user."""


SUPPORTED_OPEN = "Excel files (*.xlsx *.xlsm *.xltx *.xltm);;CSV (*.csv);;All files (*)"
SUPPORTED_SAVE = "Excel Macro-Enabled Workbook (*.xlsm);;Excel Workbook (*.xlsx);;CSV UTF-8 (*.csv)"

VBA_JSON_PART = "customXml/portableexcel_vba.json"
META_JSON_PART = "customXml/portableexcel_meta.json"   # picture / barcode metadata
_VBA_BIN_PART = "xl/vbaProject.bin"
_CT_VBA = "application/vnd.ms-office.vbaProject"
_CT_MAIN = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "xltx": "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "xlsm": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "xltm": "application/vnd.ms-excel.template.macroEnabled.main+xml",
}
_REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
_MACRO_EXTS = ("xlsm", "xltm")
_OOXML_EXTS = ("xlsx", "xlsm", "xltx", "xltm")
_CSV_EXTS = ("csv", "txt", "tsv")
_OLE_MAGIC = bytes.fromhex("D0CF11E0A1B11AE1")
_XLS_MSG = "Old .xls format is not supported; open it in Excel/LibreOffice and save as .xlsx"
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def _require(module: str, package: str):
    try:
        return __import__(module, fromlist=["_"])
    except ImportError as exc:  # pragma: no cover - depends on environment
        raise FileFormatError(
            f"The Python package '{package}' is required for this operation but is not "
            f"installed ({exc}). Install it with: pip install {package}") from exc


def _ext(path: str) -> str:
    return os.path.splitext(path)[1].lower().lstrip(".")


# ==========================================================================
# Loading
# ==========================================================================
def load_workbook(path: str, password: Optional[str] = None) -> Workbook:
    """Load an Excel (.xlsx/.xlsm/.xltx/.xltm) or CSV file into a Workbook."""
    try:
        with open(path, "rb") as f:
            data = f.read()
    except OSError as exc:
        raise FileFormatError(f"Cannot read '{path}': {exc.strerror or exc}") from exc

    ext = _ext(path)
    if data[:8] == _OLE_MAGIC:
        data = _decrypt_ole(data, password)          # raises for .xls / wrong password
    if data[:4] == b"PK\x03\x04" or (data[:2] == b"PK" and zipfile.is_zipfile(io.BytesIO(data))):
        book = _load_ooxml(data, ext)
    elif ext in _OOXML_EXTS:
        raise FileFormatError(f"'{os.path.basename(path)}' is not a valid Excel workbook "
                              "(the file is damaged or has the wrong extension).")
    elif ext == "xls":
        raise FileFormatError(_XLS_MSG)
    else:
        book = _load_csv(data, path)
    book.path = path
    return book


def _decrypt_ole(data: bytes, password: Optional[str]) -> bytes:
    """Return the decrypted OOXML package of an encrypted file, or raise."""
    olefile = _require("olefile", "olefile")
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
        streams = {"/".join(p) for p in ole.listdir(streams=True, storages=False)}
        ole.close()
    except Exception as exc:
        raise FileFormatError(f"The file is damaged or not a spreadsheet ({exc}).") from exc
    if "EncryptionInfo" not in streams or "EncryptedPackage" not in streams:
        if "Workbook" in streams or "Book" in streams:
            raise FileFormatError(_XLS_MSG)
        raise FileFormatError("This file is an OLE compound document, not an Excel workbook.")
    msoffcrypto = _require("msoffcrypto", "msoffcrypto-tool")
    try:
        office = msoffcrypto.OfficeFile(io.BytesIO(data))
        if not office.is_encrypted():  # pragma: no cover - defensive
            raise FileFormatError("Unsupported OLE container.")
    except FileFormatError:
        raise
    except Exception as exc:
        raise FileFormatError(f"Unsupported encrypted file ({exc}).") from exc
    if password is None:
        raise PasswordRequired("This workbook is password protected.")
    out = io.BytesIO()
    try:
        office.load_key(password=password)
        office.decrypt(out)
    except Exception as exc:
        name = type(exc).__name__
        if "InvalidKey" in name or "Decrypt" in name or isinstance(exc, (ValueError, KeyError)):
            raise PasswordRequired("The password is incorrect.") from exc
        raise FileFormatError(f"Cannot decrypt the workbook ({exc}).") from exc
    dec = out.getvalue()
    if not zipfile.is_zipfile(io.BytesIO(dec)):
        raise PasswordRequired("The password is incorrect.")
    return dec


# -- cell value conversion -------------------------------------------------
def _number_text(v) -> str:
    if isinstance(v, int):
        return str(v)
    if v != v or v in (float("inf"), float("-inf")):
        return "#NUM!"
    if v.is_integer() and abs(v) < 1e16:
        return str(int(v))
    return repr(v)


def _text_raw(s: str) -> str:
    """Raw text for a string cell value so that it stays text in the model."""
    if s.startswith(("=", "'")):
        return "'" + s
    if not isinstance(literal_value(s), str):
        return "'" + s
    return s


def _cell_raw(cell, to_excel) -> Optional[str]:
    v = cell.value
    if v is None:
        return None
    t = type(v)
    if t is str:
        if cell.data_type == "f":
            return v if v.startswith("=") else "=" + v
        if cell.data_type == "e":
            return v
        if v == "":
            return None
        return _text_raw(v)
    if t is bool:
        return "TRUE" if v else "FALSE"
    if t is int or t is float:
        return _number_text(v)
    if isinstance(v, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        try:
            return _number_text(to_excel(v))
        except Exception:
            return _text_raw(str(v))
    cls = t.__name__
    if cls == "ArrayFormula":
        text = v.text or ""
        return text if text.startswith("=") else "=" + text
    if cls == "DataTableFormula":
        return None
    if cls == "CellRichText":  # only when rich_text=True
        return _text_raw(str(v))
    if isinstance(v, (int, float)):  # numpy-ish / Decimal subclasses
        return _number_text(float(v) if not isinstance(v, int) else v)
    return _text_raw(str(v))


# -- styles ----------------------------------------------------------------
def _color_hex(color) -> Optional[str]:
    """openpyxl Color -> '#RRGGBB' (rgb and indexed colours; theme colours ignored)."""
    if color is None:
        return None
    try:
        ctype = color.type
        if ctype == "rgb":
            rgb = color.rgb
            if isinstance(rgb, str) and len(rgb) >= 6 and re.fullmatch(r"[0-9A-Fa-f]+", rgb):
                return "#" + rgb[-6:].upper()
        elif ctype == "indexed":
            from openpyxl.styles.colors import COLOR_INDEX
            idx = color.indexed
            if isinstance(idx, int) and 0 <= idx < 64:
                return "#" + COLOR_INDEX[idx][-6:].upper()
    except Exception:
        pass
    return None


def _style_from_cell(cell) -> dict:
    st: dict = {}
    font = cell.font
    if font is not None:
        if font.b:
            st["bold"] = True
        if font.i:
            st["italic"] = True
        if font.u and font.u != "none":
            st["underline"] = True
        col = _color_hex(font.color)
        if col and col != "#000000":
            st["color"] = col
        sz = font.sz
        if sz is not None and float(sz) != 11.0:
            st["font_size"] = float(sz)
        name = font.name
        if name and name != "Calibri":
            st["font_name"] = name
    fill = cell.fill
    if fill is not None and getattr(fill, "patternType", None) == "solid":
        col = _color_hex(fill.fgColor)
        if col:
            st["fill"] = col
    al = cell.alignment
    if al is not None:
        h = al.horizontal
        if h in ("left", "center", "right"):
            st["align"] = h
        elif h in ("centerContinuous",):
            st["align"] = "center"
        if al.wrap_text:
            st["wrap"] = True
    nf = cell.number_format
    if nf and nf != "General":
        st["number_format"] = nf
    return st


# -- ooxml -----------------------------------------------------------------
def _load_ooxml(data: bytes, ext: str) -> Workbook:
    openpyxl = _require("openpyxl", "openpyxl")
    from openpyxl.utils.datetime import to_excel

    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
        names = set(zf.namelist())
    except zipfile.BadZipFile as exc:
        raise FileFormatError(f"The workbook is damaged ({exc}).") from exc
    if "xl/workbook.xml" not in names and not any(n.endswith("workbook.xml") for n in names):
        zf.close()
        raise FileFormatError("This zip file is not an Excel workbook "
                              "(Word/PowerPoint files are not supported).")
    has_vba = _VBA_BIN_PART in names
    try:
        # keep_vba is not needed: the workbook is rebuilt from the model on save and
        # xl/vbaProject.bin is read (and re-injected) by this module directly.
        import warnings as _warnings
        with _warnings.catch_warnings(record=True) as caught:
            _warnings.simplefilter("always")
            wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=False,
                                        data_only=False, keep_links=False)
        load_warnings = [str(w.message) for w in caught
                         if "image" in str(w.message).lower() or "drawing" in str(w.message).lower()]
    except Exception as exc:
        zf.close()
        raise FileFormatError(f"Cannot read the workbook: {exc}") from exc

    book = Workbook()
    book.sheets = []
    used_codes = set()
    image_counter = [0]
    style_cache: Dict[tuple, dict] = {}
    for i, ws in enumerate(wb.worksheets, start=1):
        code = None
        try:
            code = ws.sheet_properties.codeName
        except Exception:
            pass
        if not code or code.lower() in used_codes:
            code = f"Sheet{i}"
            n = i
            while code.lower() in used_codes:
                n += 1
                code = f"Sheet{n}"
        used_codes.add(code.lower())
        sh = Sheet(ws.title, code_name=code)
        sh.visible = ws.sheet_state == "visible"
        _read_cells(ws, sh, to_excel, style_cache)
        _read_dimensions(ws, sh)
        sh.images = _read_images(ws, sh, image_counter, load_warnings)
        book.sheets.append(sh)
    if not book.sheets:
        book.sheets = [Sheet("Sheet1")]
    try:
        active = wb.active
        book.active_index = next((k for k, ws in enumerate(wb.worksheets) if ws is active), 0)
    except Exception:
        book.active_index = 0

    book.defined_names = _read_defined_names(wb)
    if META_JSON_PART in names:
        try:
            _apply_image_meta(book, json.loads(zf.read(META_JSON_PART).decode("utf-8")))
        except Exception as exc:
            _log.warning("Cannot read picture metadata (%s)", exc)
    for msg in load_warnings:
        _log.warning("%s", msg)
    # Not a model field: warnings for the GUI to show after opening (optional).
    book.load_warnings = load_warnings

    # VBA
    if has_vba:
        book.vba_project_bin = zf.read(_VBA_BIN_PART)
        try:
            book.vba_modules = extract_vba_modules(book.vba_project_bin)
        except Exception:
            book.vba_modules = []
    if VBA_JSON_PART in names:
        try:
            mods = json.loads(zf.read(VBA_JSON_PART).decode("utf-8"))
            if isinstance(mods, dict):
                mods = mods.get("modules", [])
            book.vba_modules = [_module_from_json(m) for m in mods
                                if isinstance(m, dict) and "name" in m]
        except Exception:
            pass
    book.vba_modified = False
    zf.close()
    try:
        wb.close()
    except Exception:
        pass
    return book


# -- pictures --------------------------------------------------------------
_EMU_PER_PX = 9525
_DEFAULT_COL_PX = 64
_DEFAULT_ROW_PX = 20


def _col_px(sh: Sheet, c: int) -> int:
    return sh.col_widths.get(c, _DEFAULT_COL_PX)


def _row_px(sh: Sheet, r: int) -> int:
    return sh.row_heights.get(r, _DEFAULT_ROW_PX)


def _span_px(size_fn, sh: Sheet, a: int, b: int) -> int:
    """Pixels from the start of index a to the start of index b (b >= a)."""
    return sum(size_fn(sh, i) for i in range(a, max(a, min(b, a + 20000))))


def _abs_to_cell(size_fn, sh: Sheet, px: int):
    i = 0
    while px >= size_fn(sh, i) and i < 1_048_575:
        px -= size_fn(sh, i)
        i += 1
    return i, px


def _to_png(data: bytes) -> Tuple[Optional[bytes], str]:
    """-> (PNG bytes or None, source format)."""
    from PIL import Image as PILImage
    with PILImage.open(io.BytesIO(data)) as im:
        fmt = (im.format or "").lower()
        if fmt == "png":
            return data, fmt
        if fmt in ("wmf", "emf"):
            return None, fmt
        im.load()
        if im.mode not in ("RGB", "RGBA", "L", "LA", "P", "1"):
            im = im.convert("RGBA")
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue(), fmt


def _read_images(ws, sh: Sheet, counter: list, load_warnings: list) -> list:
    images = []
    for idx, img in enumerate(getattr(ws, "_images", []) or []):
        try:
            ref = img.ref
            if hasattr(ref, "seek"):
                ref.seek(0)
                raw = ref.read()
            else:
                with open(ref, "rb") as f:
                    raw = f.read()
            png, fmt = _to_png(raw)
        except Exception as exc:
            load_warnings.append(f"Picture {idx + 1} on '{sh.name}' could not be read ({exc}) "
                                 "and was dropped.")
            continue
        if png is None:
            load_warnings.append(f"Picture {idx + 1} on '{sh.name}' is in {fmt.upper()} format, "
                                 "which is not supported; it was dropped.")
            continue
        name, descr = None, None
        row = col = dx = dy = 0
        width, height = int(img.width or 0), int(img.height or 0)
        anchor = img.anchor
        try:
            child = getattr(anchor, "pic", None)
            if child is not None and child.nvPicPr is not None:
                name = child.nvPicPr.cNvPr.name
                descr = child.nvPicPr.cNvPr.descr
            kind = type(anchor).__name__
            if kind in ("OneCellAnchor", "TwoCellAnchor"):
                fr = anchor._from
                row, col = int(fr.row), int(fr.col)
                dx, dy = int(round(fr.colOff / _EMU_PER_PX)), int(round(fr.rowOff / _EMU_PER_PX))
                if kind == "OneCellAnchor":
                    ext = anchor.ext
                    if ext is not None and ext.width and ext.height:
                        width = int(round(ext.width / _EMU_PER_PX))
                        height = int(round(ext.height / _EMU_PER_PX))
                else:
                    to = anchor.to
                    width = (_span_px(_col_px, sh, col, int(to.col)) - dx
                             + int(round(to.colOff / _EMU_PER_PX)))
                    height = (_span_px(_row_px, sh, row, int(to.row)) - dy
                              + int(round(to.rowOff / _EMU_PER_PX)))
            elif kind == "AbsoluteAnchor":
                col, dx = _abs_to_cell(_col_px, sh, int(round(anchor.pos.x / _EMU_PER_PX)))
                row, dy = _abs_to_cell(_row_px, sh, int(round(anchor.pos.y / _EMU_PER_PX)))
                width = int(round(anchor.ext.width / _EMU_PER_PX))
                height = int(round(anchor.ext.height / _EMU_PER_PX))
            elif isinstance(anchor, str):
                from openpyxl.utils.cell import coordinate_to_tuple
                r1, c1 = coordinate_to_tuple(anchor)
                row, col = r1 - 1, c1 - 1
        except Exception as exc:
            _log.warning("Picture %d on %r: cannot read its anchor (%s)", idx + 1, sh.name, exc)
        counter[0] += 1
        images.append({
            "id": f"img{counter[0]}", "name": name or f"Picture {idx + 1}",
            "descr": descr or "",
            "row": max(row, 0), "col": max(col, 0), "dx": max(dx, 0), "dy": max(dy, 0),
            "width": max(width, 1), "height": max(height, 1),
            "data": png, "fmt": "png", "barcode": None,
        })
    return images


def _apply_image_meta(book: Workbook, meta: dict) -> None:
    sheets = (meta or {}).get("images", {})
    for sh in book.sheets:
        entries = sheets.get(sh.name)
        if not entries:
            continue
        by_name = {}
        for img in sh.images:
            by_name.setdefault(img["name"], img)
        for e in entries:
            img = by_name.get(e.get("name"))
            if img is None:
                i = e.get("index")
                if not isinstance(i, int) or not 0 <= i < len(sh.images):
                    continue
                img = sh.images[i]
            if e.get("id"):
                img["id"] = str(e["id"])
            bc = e.get("barcode")
            img["barcode"] = dict(bc) if isinstance(bc, dict) else None


def _write_images(ws, sh: Sheet, warnings: List[str]) -> list:
    """Add sh.images to the openpyxl sheet; return the metadata entries."""
    images = getattr(sh, "images", None) or []
    if not images:
        return []
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import (AnchorMarker, OneCellAnchor,
                                                      SpreadsheetDrawing)
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    meta = []
    used = set()
    helper = SpreadsheetDrawing()
    for i, im in enumerate(images):
        try:
            data = im.get("data")
            if not data:
                raise ValueError("no picture data")
            png, _fmt = _to_png(bytes(data))
            if png is None:
                raise ValueError("unsupported picture format")
            xl = XLImage(io.BytesIO(png))
        except Exception as exc:
            warnings.append(f"Picture '{im.get('name') or i + 1}' on '{sh.name}' "
                            f"could not be saved ({exc}).")
            continue
        name = str(im.get("name") or f"Picture {i + 1}")
        base, n = name, 2
        while name in used:
            name = f"{base} ({n})"
            n += 1
        used.add(name)
        w = int(im.get("width") or xl.width)
        h = int(im.get("height") or xl.height)
        xl.width, xl.height = w, h
        marker = AnchorMarker(col=int(im.get("col", 0)), colOff=int(im.get("dx", 0)) * _EMU_PER_PX,
                              row=int(im.get("row", 0)), rowOff=int(im.get("dy", 0)) * _EMU_PER_PX)
        anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(w * _EMU_PER_PX, h * _EMU_PER_PX))
        pic = helper._picture_frame(len(ws._images) + 1)
        pic.nvPicPr.cNvPr.name = name
        pic.nvPicPr.cNvPr.descr = str(im.get("descr") or name)
        anchor.pic = pic
        xl.anchor = anchor
        ws.add_image(xl)
        entry = {"index": len(ws._images) - 1, "name": name, "id": str(im.get("id") or "")}
        if im.get("barcode"):
            entry["barcode"] = {k: v for k, v in dict(im["barcode"]).items()
                                if isinstance(v, (str, int, float, bool, type(None)))}
        meta.append(entry)
    return meta


def _module_from_json(m: dict) -> dict:
    mod = {"name": str(m["name"]), "type": str(m.get("type", "standard")),
           "code": str(m.get("code", "")).replace("\r\n", "\n").replace("\r", "\n")}
    if mod["type"] == "form" or isinstance(m.get("form"), dict):
        from .forms import new_form, normalize_form
        form = m.get("form")
        try:
            mod["form"] = normalize_form(form if isinstance(form, dict) else new_form(mod["name"]))
        except Exception as exc:
            _log.warning("UserForm %s: invalid stored form (%s)", mod["name"], exc)
            mod["form"] = normalize_form(new_form(mod["name"]))
    return mod


def _read_cells(ws, sh: Sheet, to_excel, style_cache) -> None:
    cells = sh.cells
    styles = sh.styles
    store = getattr(ws, "_cells", None)
    if store is None:  # pragma: no cover - fallback for other openpyxl versions
        store = {(c.row, c.column): c for row in ws.iter_rows() for c in row}
    for (row, col), cell in store.items():
        key = (row - 1, col - 1)
        raw = _cell_raw(cell, to_excel)
        if raw is not None and raw != "":
            cells[key] = raw
        sa = getattr(cell, "_style", None)
        if sa is not None:
            if not any(sa):
                continue
            sk = tuple(sa)
            st = style_cache.get(sk)
            if st is None:
                st = _style_from_cell(cell)
                style_cache[sk] = st
        else:  # pragma: no cover
            st = _style_from_cell(cell) if cell.has_style else {}
        if st:
            styles[key] = dict(st)


def _read_dimensions(ws, sh: Sheet) -> None:
    from openpyxl.utils import column_index_from_string
    for letter, dim in ws.column_dimensions.items():
        width = getattr(dim, "width", None)
        if not width or width <= 0:
            continue
        try:
            lo = dim.min or column_index_from_string(letter)
            hi = dim.max or lo
        except Exception:
            continue
        hi = min(hi, lo + 16384)
        px = int(width * 7 + 5 + 1e-6)
        for c in range(lo, hi + 1):
            sh.col_widths[c - 1] = px
    for r, dim in ws.row_dimensions.items():
        ht = getattr(dim, "ht", None)
        if ht:
            sh.row_heights[r - 1] = int(round(ht * 4 / 3))


def _read_defined_names(wb) -> Dict[str, str]:
    out: Dict[str, str] = {}

    def add(dn):
        name = getattr(dn, "name", None)
        text = getattr(dn, "attr_text", None) or getattr(dn, "value", None)
        if not name or not text or name.startswith("_xlnm.") or name.startswith("_"):
            return
        if getattr(dn, "hidden", False):
            return
        out.setdefault(name, text)

    dns = wb.defined_names
    items = dns.values() if hasattr(dns, "values") else getattr(dns, "definedName", [])
    for dn in list(items):
        add(dn)
    for ws in wb.worksheets:  # sheet-scoped names (openpyxl >= 3.1)
        local = getattr(ws, "defined_names", None)
        if local and hasattr(local, "values"):
            for dn in list(local.values()):
                add(dn)
    return out


# -- csv -------------------------------------------------------------------
def _decode_text(data: bytes) -> str:
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")  # pragma: no cover


def _sniff_delimiter(text: str) -> str:
    lines = text.splitlines()[:50]
    if not lines:
        return ","
    best, best_score = ",", (-1, -1)
    for d in ("\t", ";", ","):  # on a tie prefer the rarer-in-text delimiters
        try:
            rows = [r for r in csv.reader(lines, delimiter=d) if r]
        except csv.Error:
            continue
        counts = [len(r) for r in rows]
        if not counts or max(counts) <= 1:
            score = (0, 0)
        else:
            mode = max(set(counts), key=counts.count)
            score = (counts.count(mode) if mode > 1 else 0, mode)
        if score > best_score:
            best, best_score = d, score
    return best


def _safe_sheet_name(name: str, taken=()) -> str:
    name = _INVALID_SHEET_CHARS.sub("_", name or "").strip("'") or "Sheet"
    name = name[:31]
    low = {t.lower() for t in taken}
    base, n = name, 2
    while name.lower() in low:
        suffix = f" ({n})"
        name = base[:31 - len(suffix)] + suffix
        n += 1
    return name


def _load_csv(data: bytes, path: str) -> Workbook:
    text = _decode_text(data)
    if "\x00" in text[:4096]:
        raise FileFormatError(f"'{os.path.basename(path)}' is not a text/CSV file.")
    delim = "\t" if _ext(path) == "tsv" else _sniff_delimiter(text)
    book = Workbook()
    sh = Sheet(_safe_sheet_name(os.path.splitext(os.path.basename(path))[0]), code_name="Sheet1")
    try:
        reader = csv.reader(io.StringIO(text, newline=""), delimiter=delim)
        for r, row in enumerate(reader):
            for c, val in enumerate(row):
                if val != "":
                    sh.cells[(r, c)] = val
    except csv.Error as exc:
        raise FileFormatError(f"Cannot parse the CSV file: {exc}") from exc
    book.sheets = [sh]
    book.active_index = 0
    return book


# ==========================================================================
# VBA extraction
# ==========================================================================
def _vba_bin_from(path_or_bytes) -> Optional[bytes]:
    if isinstance(path_or_bytes, (bytes, bytearray, memoryview)):
        data = bytes(path_or_bytes)
    else:
        with open(os.fspath(path_or_bytes), "rb") as f:
            data = f.read()
    if data[:8] == _OLE_MAGIC:
        return data
    if data[:2] == b"PK":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = zf.namelist()
                cand = [n for n in names if n.lower().endswith("vbaproject.bin")]
                if not cand:
                    return None
                return zf.read(_VBA_BIN_PART if _VBA_BIN_PART in cand else cand[0])
        except zipfile.BadZipFile:
            return None
    return None


def _project_kinds(ole, root: str) -> Dict[str, str]:
    """Parse the PROJECT stream: module name (lower) -> type."""
    kinds: Dict[str, str] = {}
    path = (root + "/" if root else "") + "PROJECT"
    try:
        if not ole.exists(path):
            return kinds
        raw = ole.openstream(path).read()
    except Exception:
        return kinds
    for line in raw.decode("cp1252", errors="replace").splitlines():
        line = line.strip()
        if line.startswith("["):
            break
        if "=" not in line:
            continue
        key, val = line.split("=", 1)
        val = val.strip().strip('"')
        if key == "Document":
            kinds[val.split("/", 1)[0].lower()] = "document"
        elif key == "Module":
            kinds[val.lower()] = "standard"
        elif key == "Class":
            kinds[val.lower()] = "class"
        elif key == "BaseClass":
            kinds[val.lower()] = "form"
    return kinds


_VB_BASE_RE = re.compile(r'^Attribute\s+VB_Base\s*=\s*"([^"]*)"', re.I | re.M)
_DOC_GUIDS = ("{00020819-", "{00020820-")  # Excel workbook / worksheet
_FORM_GUIDS = ("{c62a69f0-",)              # MSForms UserForm


def _guess_kind(name: str, filename: str, code: str, ole_storages: set,
                kinds: Dict[str, str]) -> str:
    k = kinds.get(name.lower())
    if k:
        return k
    low = code.lower()
    m = _VB_BASE_RE.search(code)
    base = m.group(1).lower() if m else ""
    if name.lower() == "thisworkbook" or any(g.lower() in base for g in _DOC_GUIDS):
        return "document"
    if (filename.lower().endswith(".frm") or "begin {" in low
            or any(g in base for g in _FORM_GUIDS) or name.lower() in ole_storages):
        return "form"
    if filename.lower().endswith(".bas"):
        return "standard"
    if "version 1.0 class" in low or "vb_exposed" in low or "vb_predeclaredid" in low \
            or filename.lower().endswith(".cls"):
        return "class"
    return "standard"


# ==========================================================================
# UserForm designers ([MS-OFORMS])
# ==========================================================================
# A form module "UserForm1" has a designer storage "UserForm1" next to the
# VBA storage, holding:
#   \x03VBFrame  text: Caption, ClientWidth/ClientHeight (twips), StartUpPosition
#   f            FormControl: form properties + one OleSiteConcrete per control
#   o            the controls' own property records, back to back; each site's
#                ObjectStreamSize says how many bytes belong to it
#   i<ID>        nested storage (own f / o) of a Frame or MultiPage with that ID
#
# Every record is "version(2) cb(2) PropMask DataBlock ExtraDataBlock": a mask
# bit says the property is stored (else it has its default).  DataBlock holds
# fixed-size values in mask-bit order, each aligned to its own size; strings
# and arrays store a length there and their bytes in ExtraDataBlock, which
# also holds the 8-byte size/position pairs, each item padded to 4 bytes.
_log = logging.getLogger("portable_excel.fileio")

_HM_PT = 72.0 / 2540.0      # HIMETRIC -> points

# kinds: u8 u16 i16 u32 i32 -> data; str / arr -> 4-byte length in data + bytes
# in extra; pair -> 8 bytes in extra; pic / font -> 2-byte marker in data
# (object in the stream data); flag -> no data (presence toggles a boolean)
_FORM_LAYOUT = [
    (1, "u32", "back_color"), (2, "u32", "fore_color"), (3, "u32", "next_id"),
    (6, "u32", "bool"), (7, "u8", "border_style"), (8, "u8", "mouse_pointer"),
    (9, "u8", "scroll_bars"), (10, "pair", "size"), (11, "pair", "logical_size"),
    (12, "pair", "scroll_pos"), (13, "u32", "group_cnt"), (15, "pic", "mouse_icon"),
    (16, "u8", "cycle"), (17, "u8", "special_effect"), (18, "u32", "border_color"),
    (19, "str", "caption"), (20, "font", "font"), (21, "pic", "picture"),
    (22, "u32", "zoom"), (23, "u8", "picture_alignment"), (24, "flag", "picture_tiling"),
    (25, "u8", "picture_size_mode"), (26, "u32", "shape_cookie"), (27, "u32", "draw_buffer"),
]
_SITE_LAYOUT = [
    (0, "str", "name"), (1, "str", "tag"), (2, "u32", "id"), (3, "u32", "help_context"),
    (4, "u32", "bit_flags"), (5, "u32", "object_size"), (6, "i16", "tab_index"),
    (7, "u16", "clsid_index"), (8, "pair", "position"), (9, "u16", "group_id"),
    (11, "str", "control_tip_text"), (12, "str", "license_key"),
    (13, "str", "control_source"), (14, "str", "row_source"),
]
_MORPH_LAYOUT = [
    (0, "u32", "flags"), (1, "u32", "back_color"), (2, "u32", "fore_color"),
    (3, "u32", "max_length"), (4, "u8", "border_style"), (5, "u8", "scroll_bars"),
    (6, "u8", "display_style"), (7, "u8", "mouse_pointer"), (8, "pair", "size"),
    (9, "u16", "password_char"), (10, "u32", "list_width"), (11, "u16", "bound_column"),
    (12, "i16", "text_column"), (13, "i16", "column_count"), (14, "i16", "list_rows"),
    (15, "u16", "column_info_count"), (16, "u8", "match_entry"), (17, "u8", "list_style"),
    (18, "u8", "show_drop_button_when"), (20, "u8", "drop_button_style"),
    (21, "u8", "multi_select"), (22, "str", "value"), (23, "str", "caption"),
    (24, "u32", "picture_position"), (25, "u32", "border_color"),
    (26, "u32", "special_effect"), (27, "pic", "mouse_icon"), (28, "pic", "picture"),
    (29, "u16", "accelerator"), (32, "str", "group_name"),
]
_LABEL_LAYOUT = [
    (0, "u32", "fore_color"), (1, "u32", "back_color"), (2, "u32", "flags"),
    (3, "str", "caption"), (4, "u32", "picture_position"), (5, "pair", "size"),
    (6, "u8", "mouse_pointer"), (7, "u32", "border_color"), (8, "u16", "border_style"),
    (9, "u16", "special_effect"), (10, "pic", "picture"), (11, "u16", "accelerator"),
    (12, "pic", "mouse_icon"),
]
_BUTTON_LAYOUT = [
    (0, "u32", "fore_color"), (1, "u32", "back_color"), (2, "u32", "flags"),
    (3, "str", "caption"), (4, "u32", "picture_position"), (5, "pair", "size"),
    (6, "u8", "mouse_pointer"), (7, "pic", "picture"), (8, "u16", "accelerator"),
    (9, "flag", "no_take_focus"), (10, "pic", "mouse_icon"),
]
_IMAGE_LAYOUT = [
    (2, "flag", "auto_size"), (3, "u32", "border_color"), (4, "u32", "back_color"),
    (5, "u8", "border_style"), (6, "u8", "mouse_pointer"), (7, "u8", "picture_size_mode"),
    (8, "u8", "special_effect"), (9, "pair", "size"), (10, "pic", "picture"),
    (11, "u8", "picture_alignment"), (12, "flag", "picture_tiling"), (13, "u32", "flags"),
    (14, "pic", "mouse_icon"),
]
_SPIN_LAYOUT = [
    (0, "u32", "fore_color"), (1, "u32", "back_color"), (2, "u32", "flags"),
    (3, "pair", "size"), (5, "i32", "min"), (6, "i32", "max"), (7, "i32", "position"),
    (8, "u32", "prev_enabled"), (9, "u32", "next_enabled"), (10, "i32", "small_change"),
    (11, "i32", "orientation"), (12, "i32", "delay"), (13, "pic", "mouse_icon"),
    (14, "u8", "mouse_pointer"),
]
_SCROLL_LAYOUT = [
    (0, "u32", "fore_color"), (1, "u32", "back_color"), (2, "u32", "flags"),
    (3, "pair", "size"), (4, "u8", "mouse_pointer"), (5, "i32", "min"), (6, "i32", "max"),
    (7, "i32", "position"), (9, "u32", "prev_enabled"), (10, "u32", "next_enabled"),
    (11, "i32", "small_change"), (12, "i32", "large_change"), (13, "i32", "orientation"),
    (14, "i16", "proportional_thumb"), (15, "i32", "delay"), (16, "pic", "mouse_icon"),
]
_TABSTRIP_LAYOUT = [
    (0, "i32", "list_index"), (1, "u32", "back_color"), (2, "u32", "fore_color"), (4, "pair", "size"),
    (5, "arr", "items"), (6, "u8", "mouse_pointer"), (8, "u32", "tab_orientation"),
    (9, "u32", "tab_style"), (10, "flag", "multi_row"), (11, "u32", "tab_fixed_width"),
    (12, "u32", "tab_fixed_height"), (13, "flag", "tooltips"), (15, "arr", "tip_strings"),
    (17, "arr", "names"), (18, "u32", "flags"), (19, "flag", "new_version"),
    (20, "u32", "tabs_allocated"), (21, "arr", "tags"), (22, "u32", "tab_data"),
    (23, "arr", "accelerators"), (24, "pic", "mouse_icon"),
]
_TEXTPROPS_LAYOUT = [
    (0, "str", "font_name"), (1, "u32", "font_effects"), (2, "u32", "font_height"),
    (4, "u8", "charset"), (5, "u8", "pitch_family"), (6, "u8", "paragraph_align"),
    (7, "u16", "font_weight"),
]
_FIXED = {"u8": ("<B", 1), "u16": ("<H", 2), "i16": ("<h", 2), "u32": ("<I", 4),
          "i32": ("<i", 4), "str": ("<I", 4), "arr": ("<I", 4), "pic": ("<H", 2),
          "font": ("<H", 2)}

# ClsidCacheIndex -> control type ([MS-OFORMS] 2.4.5); 15 = generic MorphData
_CLSID_TYPES = {7: "Page", 12: "Image", 14: "Frame", 15: "MorphData", 16: "SpinButton",
                17: "CommandButton", 18: "TabStrip", 21: "Label", 23: "TextBox",
                24: "ListBox", 25: "ComboBox", 26: "CheckBox", 27: "OptionButton",
                28: "ToggleButton", 47: "ScrollBar", 57: "MultiPage"}
_DISPLAY_STYLE_TYPES = {1: "TextBox", 2: "ListBox", 3: "ComboBox", 4: "CheckBox",
                        5: "OptionButton", 6: "ToggleButton", 7: "ComboBox"}
_MORPH_DEFAULT_STYLE = {"TextBox": 1, "ListBox": 2, "ComboBox": 3, "CheckBox": 4,
                        "OptionButton": 5, "ToggleButton": 6}

# Windows 10 default system colours, by COLOR_* index (OLE_COLOR 0x800000xx)
_SYS_COLORS = {
    0: "#C8C8C8", 1: "#000000", 2: "#99B4D1", 3: "#BFCDDB", 4: "#F0F0F0", 5: "#FFFFFF",
    6: "#646464", 7: "#000000", 8: "#000000", 9: "#000000", 10: "#B4B4B4", 11: "#F4F7FC",
    12: "#ABABAB", 13: "#0078D7", 14: "#FFFFFF", 15: "#F0F0F0", 16: "#A0A0A0",
    17: "#6D6D6D", 18: "#000000", 19: "#000000", 20: "#FFFFFF", 21: "#696969",
    22: "#E3E3E3", 23: "#000000", 24: "#FFFFE1", 26: "#0066CC", 27: "#B9D1EA",
    28: "#D7E4F2", 29: "#3399FF", 30: "#F0F0F0",
}
_STDFONT_GUID = bytes.fromhex("0352e30b918fce119de300aa004bb851")    # {0BE35203-8F91-11CE-9DE3-00AA004BB851}
_TEXTPROPS_GUID = bytes.fromhex("2009c2af4edace11b94300aa006887b4")  # {AFC20920-DA4E-11CE-B943-00AA006887B4}
_STDPIC_GUID = bytes.fromhex("0452e30b918fce119de300aa004bb851")     # {0BE35204-...}

# VariousPropertyBits (shared meaning across controls)
_VPB_ENABLED = 0x00000002
_VPB_LOCKED = 0x00000004
_VPB_OPAQUE = 0x00000008
_VPB_MATCH_REQUIRED = 0x00001000
_VPB_WORD_WRAP = 0x00800000
_VPB_AUTO_SIZE = 0x10000000
_VPB_MULTI_LINE = 0x80000000
_DEFAULT_FLAGS = {"morph": 0x2C80081B, "label": 0x0080001B, "button": 0x0000001B,
                  "image": 0x0000001B, "spin": 0x0000001B, "scroll": 0x0000001B,
                  "tabstrip": 0x0000001B}
# site BitFlags
_SITE_TABSTOP, _SITE_VISIBLE, _SITE_DEFAULT, _SITE_CANCEL = 0x1, 0x2, 0x4, 0x8
_SITE_DEFAULT_FLAGS = 0x33


class _FormParseError(Exception):
    pass


def ole_color_to_hex(v) -> Optional[str]:
    """OLE_COLOR (0x00BBGGRR, or 0x800000nn system colour) -> '#RRGGBB'."""
    if v is None:
        return None
    v = int(v) & 0xFFFFFFFF
    if v & 0x80000000:
        return _SYS_COLORS.get(v & 0xFF, "#000000")
    r, g, b = v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def _pt(himetric) -> float:
    """HIMETRIC -> points, preferring the shortest value that maps back to the
    same HIMETRIC integer (VBE stores round(points * 2540 / 72))."""
    exact = himetric * _HM_PT
    for nd in (0, 1, 2):
        cand = round(exact, nd)
        back = cand / _HM_PT
        if round(back) == himetric or int(back + 1e-9) == himetric:
            return float(cand)
    return round(exact, 2)


def _u(fmt: str, buf: bytes, pos: int):
    size = struct.calcsize(fmt)
    if pos < 0 or pos + size > len(buf):
        raise _FormParseError(f"truncated record at {pos}")
    return struct.unpack_from(fmt, buf, pos)[0]


def _align(pos: int, base: int, n: int) -> int:
    off = (pos - base) % n
    return pos + (n - off if off else 0)


def _read_ms_string(buf: bytes, pos: int, count: int, compressed: bool) -> str:
    raw = buf[pos:pos + count]
    if len(raw) < count:
        raise _FormParseError("truncated string")
    if compressed:
        return raw.decode("latin-1")
    return raw.decode("utf-16-le", errors="replace")


def _parse_props(buf: bytes, start: int, end: int, mask: int, layout) -> dict:
    """Decode one DataBlock + ExtraDataBlock (start = just after PropMask,
    end = end of the record as given by its cb field)."""
    vals: dict = {}
    extras = []
    pos = start
    for bit, kind, key in layout:
        if not (mask >> bit) & 1:
            continue
        if kind == "flag":
            vals[key] = True
            continue
        if kind == "pair":
            extras.append((key, "pair", None))
            continue
        fmt, size = _FIXED[kind]
        pos = _align(pos, start, size)
        v = _u(fmt, buf, pos)
        pos += size
        if kind == "str":
            extras.append((key, "str", (v & 0x7FFFFFFF, bool(v & 0x80000000))))
        elif kind == "arr":
            extras.append((key, "arr", v))
        elif kind in ("pic", "font"):
            vals[key] = True
        else:
            vals[key] = v
    data_end = _align(pos, start, 4)
    extra_size = 0
    for _key, kind, aux in extras:
        if kind == "pair":
            extra_size += 8
        elif kind == "str":
            extra_size += _align(aux[0], 0, 4)
        else:
            extra_size += _align(aux, 0, 4)
    # The extra block ends the record; locating it from the end tolerates
    # DataBlock fields we do not know about.
    epos = end - extra_size if end - extra_size >= data_end else data_end
    for key, kind, aux in extras:
        if kind == "pair":
            vals[key] = (_u("<i", buf, epos), _u("<i", buf, epos + 4))
            epos += 8
        elif kind == "str":
            n, comp = aux
            vals[key] = _read_ms_string(buf, epos, n, comp)
            epos += _align(n, 0, 4)
        else:
            items, p, stop = [], epos, epos + aux
            while p + 4 <= stop:
                v = _u("<I", buf, p)
                n = v & 0x7FFFFFFF
                p += 4
                items.append(_read_ms_string(buf, p, n, bool(v & 0x80000000)))
                p += _align(n, 0, 4)
            vals[key] = items
            epos += _align(aux, 0, 4)
    return vals


def _record(buf: bytes, pos: int, mask_fmt: str = "<I"):
    """-> (mask, data_start, record_end) of a versioned property record."""
    cb = _u("<H", buf, pos + 2)
    mask = _u(mask_fmt, buf, pos + 4)
    end = pos + 4 + cb
    if end > len(buf):
        raise _FormParseError("record exceeds stream")
    return mask, pos + 4 + struct.calcsize(mask_fmt), end


def _skip_picture(buf: bytes, pos: int, keep: bool = False):
    """GuidAndPicture -> (new_pos, picture bytes or None)."""
    pos += 16  # GUID (StdPicture)
    pre = _u("<I", buf, pos)
    size = _u("<I", buf, pos + 4)
    if pre != 0x0000746C:
        raise _FormParseError("bad StdPicture preamble")
    data = buf[pos + 8:pos + 8 + size] if keep else None
    return pos + 8 + size, data


def _text_props(buf: bytes, pos: int):
    """TextProps -> (font dict, new_pos)."""
    mask, start, end = _record(buf, pos)
    v = _parse_props(buf, start, end, mask, _TEXTPROPS_LAYOUT)
    font: dict = {}
    if v.get("font_name"):
        font["font_name"] = v["font_name"]
    if "font_height" in v:
        font["font_size"] = round(v["font_height"] / 20.0, 2)
    eff = v.get("font_effects", 0)
    font["bold"] = bool(eff & 1) or v.get("font_weight", 400) >= 700
    font["italic"] = bool(eff & 2)
    font["underline"] = bool(eff & 4)
    if "paragraph_align" in v:  # 1 left, 2 right, 3 center -> fmTextAlign
        font["text_align"] = {1: 1, 2: 3, 3: 2}.get(v["paragraph_align"], 1)
    return font, end


def _guid_and_font(buf: bytes, pos: int):
    guid = buf[pos:pos + 16]
    pos += 16
    if guid == _STDFONT_GUID:
        # StdFont: version(1) charset(2) flags(1) weight(2) height(4, pt*10000) len(1) face
        flags = _u("<B", buf, pos + 3)
        weight = _u("<H", buf, pos + 4)
        height = _u("<I", buf, pos + 6)
        n = _u("<B", buf, pos + 10)
        face = buf[pos + 11:pos + 11 + n].decode("latin-1")
        return ({"font_name": face, "font_size": round(height / 10000.0, 2),
                 "bold": weight >= 700, "italic": bool(flags & 2),
                 "underline": bool(flags & 4)}, pos + 11 + n)
    if guid == _TEXTPROPS_GUID:
        return _text_props(buf, pos)
    raise _FormParseError("unknown font GUID")


def _parse_form_control(f: bytes) -> dict:
    """FormControl ("f" stream) -> {"props": {...}, "font": {...}, "sites": [...]}."""
    if len(f) < 8 or f[1] != 4:
        raise _FormParseError("not a FormControl")
    mask, start, end = _record(f, 0)
    props = _parse_props(f, start, end, mask, _FORM_LAYOUT)
    pos = end
    font: dict = {}
    if props.get("mouse_icon"):
        pos, _ = _skip_picture(f, pos)
    if props.get("font"):
        font, pos = _guid_and_font(f, pos)
    if props.get("picture"):
        pos, _ = _skip_picture(f, pos)
    if not props.get("bool", 0x4) & 0x8000:  # FORM_FLAG_DONTSAVECLASSTABLE clear
        n = _u("<H", f, pos)
        pos += 2
        for _ in range(n):  # SiteClassInfo: version(2) cb(2) data
            pos += 4 + _u("<H", f, pos + 2)
    sites = []
    if pos + 8 <= len(f):
        count = _u("<I", f, pos)
        pos += 8  # CountOfSites, CountOfBytes
        base = pos
        remaining = count
        while remaining > 0:
            tc = _u("<B", f, pos + 1)
            pos += 2
            if tc & 0x80:
                n = tc & 0x7F
                pos += 1
            else:
                n = 1
            if n == 0:
                break
            remaining -= n
        pos = _align(pos, base, 4)
        for _ in range(count):
            mask_s, sstart, send = _record(f, pos)
            site = _parse_props(f, sstart, send, mask_s, _SITE_LAYOUT)
            site.setdefault("bit_flags", _SITE_DEFAULT_FLAGS)
            site.setdefault("object_size", 0)
            site.setdefault("clsid_index", 0x7FFF)
            site.setdefault("id", 0)
            site.setdefault("tab_index", -1)
            sites.append(site)
            pos = send
    return {"props": props, "font": font, "sites": sites}


def _common_visual(v: dict, kind: str) -> dict:
    out: dict = {}
    if "fore_color" in v:
        out["fore_color"] = ole_color_to_hex(v["fore_color"])
    if "back_color" in v:
        out["back_color"] = ole_color_to_hex(v["back_color"])
    flags = v.get("flags", _DEFAULT_FLAGS[kind])
    out["enabled"] = bool(flags & _VPB_ENABLED)
    if kind in ("morph", "label", "button", "image"):
        out["back_style"] = 1 if flags & _VPB_OPAQUE else 0
    if kind in ("morph", "label", "button"):
        out["word_wrap"] = bool(flags & _VPB_WORD_WRAP)
        out["auto_size"] = bool(flags & _VPB_AUTO_SIZE)
    if kind == "morph":
        out["locked"] = bool(flags & _VPB_LOCKED)
    if "size" in v:
        out["width"], out["height"] = _pt(v["size"][0]), _pt(v["size"][1])
    if v.get("accelerator"):
        out["accelerator"] = chr(v["accelerator"])
    return out


def _stream_tail(buf: bytes, pos: int, v: dict, order, text: bool, keep_picture=False):
    """Consume pictures (in `order`) and TextProps after a control record."""
    font, picture = {}, None
    for key in order:
        if v.get(key):
            pos, data = _skip_picture(buf, pos, keep=keep_picture and key == "picture")
            if data is not None:
                picture = data
    if text and pos + 4 <= len(buf):
        font, pos = _text_props(buf, pos)
    return font, picture, pos


def _picture_b64(data: Optional[bytes]) -> Optional[str]:
    if not data:
        return None
    if data[:2] == b"BM" or data[:8] == b"\x89PNG\r\n\x1a\n" or data[:3] in (b"GIF", b"\xff\xd8\xff"):
        import base64
        return base64.b64encode(data).decode("ascii")
    return None


def _parse_control_object(ctype: str, o: bytes) -> dict:
    """Decode one control's record from the "o" stream -> CONTROL_DICT fields."""
    out: dict = {}
    if ctype in _MORPH_DEFAULT_STYLE or ctype == "MorphData":
        mask, start, end = _record(o, 0, "<Q")
        v = _parse_props(o, start, end, mask, _MORPH_LAYOUT)
        style = v.get("display_style", _MORPH_DEFAULT_STYLE.get(ctype, 1))
        if ctype == "MorphData":
            ctype = _DISPLAY_STYLE_TYPES.get(style, "TextBox")
        out.update(_common_visual(v, "morph"))
        flags = v.get("flags", _DEFAULT_FLAGS["morph"])
        value = v.get("value", "")
        if ctype == "TextBox":
            out.update(value=value, text=value, multi_line=bool(flags & _VPB_MULTI_LINE),
                       max_length=v.get("max_length", 0), scroll_bars=v.get("scroll_bars", 0),
                       password_char=chr(v["password_char"]) if v.get("password_char") else "")
        elif ctype in ("CheckBox", "OptionButton", "ToggleButton"):
            out["caption"] = v.get("caption", "")
            out["value"] = value.strip() not in ("", "0")
            if ctype == "OptionButton":
                out["group_name"] = v.get("group_name", "")
        else:  # ComboBox / ListBox
            out.update(column_count=v.get("column_count", 1),
                       bound_column=v.get("bound_column", 1),
                       list_style=v.get("list_style", 0))
            if ctype == "ComboBox":
                out.update(value=value, text=value, style=2 if style == 7 else 0,
                           match_required=bool(flags & _VPB_MATCH_REQUIRED),
                           max_length=v.get("max_length", 0))
            else:
                out.update(value=value if value != "" else None,
                           multi_select=v.get("multi_select", 0))
        font, picture, _ = _stream_tail(o, end, v, ("mouse_icon", "picture"), True)
    elif ctype == "Label":
        mask, start, end = _record(o, 0)
        v = _parse_props(o, start, end, mask, _LABEL_LAYOUT)
        out.update(_common_visual(v, "label"))
        out["caption"] = v.get("caption", "")
        font, _, _ = _stream_tail(o, end, v, ("picture", "mouse_icon"), True)
    elif ctype == "CommandButton":
        mask, start, end = _record(o, 0)
        v = _parse_props(o, start, end, mask, _BUTTON_LAYOUT)
        out.update(_common_visual(v, "button"))
        out["caption"] = v.get("caption", "")
        out["take_focus_on_click"] = not v.get("no_take_focus", False)
        font, _, _ = _stream_tail(o, end, v, ("picture", "mouse_icon"), True)
    elif ctype == "Image":
        mask, start, end = _record(o, 0)
        v = _parse_props(o, start, end, mask, _IMAGE_LAYOUT)
        out.update(_common_visual(v, "image"))
        out["picture_size_mode"] = v.get("picture_size_mode", 0)
        out["auto_size"] = bool(v.get("auto_size"))
        font, picture, _ = _stream_tail(o, end, v, ("picture", "mouse_icon"), False, True)
        out["picture"] = _picture_b64(picture)
    elif ctype in ("SpinButton", "ScrollBar"):
        layout, kind = (_SPIN_LAYOUT, "spin") if ctype == "SpinButton" else (_SCROLL_LAYOUT, "scroll")
        mask, start, end = _record(o, 0)
        v = _parse_props(o, start, end, mask, layout)
        out.update(_common_visual(v, kind))
        out.update(min=v.get("min", 0), max=v.get("max", 100 if ctype == "SpinButton" else 32767),
                   value=v.get("position", 0), small_change=v.get("small_change", 1),
                   orientation=v.get("orientation", -1))
        if ctype == "ScrollBar":
            out["large_change"] = v.get("large_change", 1)
        font = {}
    elif ctype == "TabStrip":
        mask, start, end = _record(o, 0)
        v = _parse_props(o, start, end, mask, _TABSTRIP_LAYOUT)
        out.update(_common_visual(v, "tabstrip"))
        items = v.get("items", [])
        names = v.get("names", [])
        out["tabs"] = [{"name": names[i] if i < len(names) and names[i] else f"Tab{i + 1}",
                        "caption": cap} for i, cap in enumerate(items)]
        out["value"] = max(v.get("list_index", 0), 0)
        font, _, _ = _stream_tail(o, end, v, ("mouse_icon",), True)
    else:
        return {}
    out.update({k: val for k, val in font.items() if k in
                ("font_name", "font_size", "bold", "italic", "underline", "text_align")})
    out["type"] = ctype
    return out


def _child_storage(ole, path: List[str], site_id: int) -> Optional[List[str]]:
    for entry in ole.listdir(streams=False, storages=True):
        if len(entry) == len(path) + 1 and [e.lower() for e in entry[:-1]] == [p.lower() for p in path]:
            nm = entry[-1]
            if nm[:1].lower() == "i" and nm[1:].isdigit() and int(nm[1:]) == site_id:
                return entry
    return None


def _read_stream(ole, path: List[str], name: str) -> bytes:
    p = path + [name]
    return ole.openstream(p).read() if ole.exists("/".join(p)) else b""


def _parse_container(ole, path: List[str], parent: Optional[str], controls: list,
                     depth: int = 0) -> dict:
    """Parse the f/o streams of `path` (form, Frame or MultiPage page); append
    the controls (recursively) to `controls`. Returns the FormControl info."""
    fc = _parse_form_control(_read_stream(ole, path, "f"))
    o = _read_stream(ole, path, "o")
    pos = 0
    for site in fc["sites"]:
        size = site["object_size"]
        chunk = o[pos:pos + size]
        pos += size
        ctype = _CLSID_TYPES.get(site["clsid_index"])
        name = site.get("name") or ""
        if ctype is None:
            _log.warning("UserForm control %r: unsupported control type (class index %s) skipped",
                         name, site["clsid_index"])
            continue
        flags = site["bit_flags"]
        ctrl = {"type": ctype, "name": name, "parent": parent,
                "tab_index": max(site["tab_index"], 0),
                "tab_stop": bool(flags & _SITE_TABSTOP),
                "visible": bool(flags & _SITE_VISIBLE)}
        if "position" in site:
            ctrl["left"], ctrl["top"] = _pt(site["position"][0]), _pt(site["position"][1])
        else:
            ctrl["left"] = ctrl["top"] = 0.0
        for key in ("tag", "control_tip_text"):
            if site.get(key):
                ctrl[key] = site[key]
        for key in ("control_source", "row_source"):
            if site.get(key):
                ctrl[key] = site[key]
        if ctype == "CommandButton":
            ctrl["default"] = bool(flags & _SITE_DEFAULT)
            ctrl["cancel"] = bool(flags & _SITE_CANCEL)
        if chunk:
            try:
                ctrl.update(_parse_control_object(ctype, chunk))
            except Exception as exc:  # keep the control with site data only
                _log.warning("UserForm control %r: cannot decode properties (%s)", name, exc)
                if ctype == "MorphData":
                    ctrl["type"] = "TextBox"
        elif ctype == "MorphData":
            ctrl["type"] = "TextBox"
        if ctype in ("Frame", "MultiPage", "Page") and depth < 16:
            sub = _child_storage(ole, path, site["id"])
            if sub is not None:
                try:
                    if ctype == "MultiPage":
                        _parse_multipage(ole, list(sub), ctrl, controls, depth + 1)
                        continue
                    insert_at = len(controls)
                    info = _parse_container(ole, list(sub), name, controls, depth + 1)
                    p = info["props"]
                    if "size" in p:
                        ctrl["width"], ctrl["height"] = _pt(p["size"][0]), _pt(p["size"][1])
                    if "caption" in p:
                        ctrl["caption"] = p["caption"]
                    _apply_container_look(ctrl, info)
                    controls.insert(insert_at, ctrl)
                    continue
                except Exception as exc:
                    _log.warning("UserForm %s %r: cannot decode contents (%s)", ctype, name, exc)
        if ctype == "Page":  # a page outside a MultiPage: keep as a Frame
            ctrl["type"] = "Frame"
        if not any(c is ctrl for c in controls):
            controls.append(ctrl)
    return fc


def _apply_container_look(ctrl: dict, info: dict) -> None:
    p = info["props"]
    if "back_color" in p:
        ctrl["back_color"] = ole_color_to_hex(p["back_color"])
    if "fore_color" in p:
        ctrl["fore_color"] = ole_color_to_hex(p["fore_color"])
    for k in ("font_name", "font_size", "bold", "italic", "underline"):
        if k in info["font"]:
            ctrl[k] = info["font"][k]


def _parse_multipage(ole, path: List[str], ctrl: dict, controls: list, depth: int) -> None:
    """MultiPage storage: its f holds a TabStrip site plus one site per page,
    its o the TabStrip (page captions = tab items), each page an i<ID> storage."""
    fc = _parse_form_control(_read_stream(ole, path, "f"))
    o = _read_stream(ole, path, "o")
    p = fc["props"]
    if "size" in p:
        ctrl["width"], ctrl["height"] = _pt(p["size"][0]), _pt(p["size"][1])
    _apply_container_look(ctrl, fc)
    captions: List[str] = []
    pos = 0
    pages = []
    for site in fc["sites"]:
        size = site["object_size"]
        chunk = o[pos:pos + size]
        pos += size
        if site["clsid_index"] == 18 and chunk:
            try:
                captions = [t["caption"] for t in _parse_control_object("TabStrip", chunk)["tabs"]]
            except Exception as exc:
                _log.warning("MultiPage %r: cannot decode tabs (%s)", ctrl["name"], exc)
        elif site["clsid_index"] == 7:
            pages.append(site)
    pages.sort(key=lambda s: s["tab_index"])
    ctrl["pages"] = []
    controls.append(ctrl)
    for i, site in enumerate(pages):
        pname = site.get("name") or f"Page{i + 1}"
        cap = captions[i] if i < len(captions) else pname
        ctrl["pages"].append({"name": pname, "caption": cap})
        sub = _child_storage(ole, path, site["id"])
        if sub is not None:
            try:
                _parse_container(ole, list(sub), pname, controls, depth + 1)
            except Exception as exc:
                _log.warning("MultiPage page %r: cannot decode (%s)", pname, exc)
    ctrl.setdefault("value", 0)


_VBFRAME_RE = re.compile(r'^\s*(\w+)\s*=\s*(?:"((?:[^"]|"")*)"|(\S+))', re.M)


def _parse_vbframe(text: str) -> dict:
    out: dict = {}
    for m in _VBFRAME_RE.finditer(text):
        key = m.group(1)
        out[key] = m.group(2).replace('""', '"') if m.group(2) is not None else m.group(3)
    return out


# Border/title-bar allowance between UserForm.Width/Height and the client
# area (VBE defaults: ClientWidth 4560 twips for Width 240, ClientHeight 3015
# for Height 180).
_FORM_EXTRA_W = 12.0
_FORM_EXTRA_H = 29.25


def parse_form_designer(ole, storage, name: str = "UserForm1") -> dict:
    """Parse a UserForm designer storage of an open olefile.OleFileIO into a
    normalized FORM_DICT (see forms.py). Never raises: on malformed data it
    logs a warning and returns what it could decode (at least the caption)."""
    from .forms import new_form, normalize_form
    path = [storage] if isinstance(storage, str) else list(storage)
    form = new_form(name)
    form["back_color"] = "#F0F0F0"
    form["fore_color"] = "#000000"
    client = {}
    caption_from_vbframe = False
    try:
        raw = _read_stream(ole, path, "\x03VBFrame")
        if raw:
            vb = _parse_vbframe(raw.decode("cp1252", errors="replace"))
            if "Caption" in vb:
                form["caption"] = vb["Caption"]
                caption_from_vbframe = True
            for key in ("ClientWidth", "ClientHeight", "Width", "Height", "StartUpPosition"):
                if key in vb:
                    try:
                        client[key] = float(vb[key])
                    except ValueError:
                        pass
    except Exception as exc:
        _log.warning("UserForm %s: cannot read VBFrame (%s)", name, exc)
    if "ClientWidth" in client:
        form["width"] = round(client["ClientWidth"] / 20.0 + _FORM_EXTRA_W, 2)
    elif "Width" in client:
        form["width"] = round(client["Width"] / 20.0, 2)
    if "ClientHeight" in client:
        form["height"] = round(client["ClientHeight"] / 20.0 + _FORM_EXTRA_H, 2)
    elif "Height" in client:
        form["height"] = round(client["Height"] / 20.0, 2)
    if "StartUpPosition" in client:
        form["start_up_position"] = int(client["StartUpPosition"])
    controls: list = []
    try:
        info = _parse_container(ole, path, None, controls)
        p = info["props"]
        if "size" in p and "ClientWidth" not in client:
            form["width"] = round(_pt(p["size"][0]) + _FORM_EXTRA_W, 2)
            form["height"] = round(_pt(p["size"][1]) + _FORM_EXTRA_H, 2)
        if "back_color" in p:
            form["back_color"] = ole_color_to_hex(p["back_color"])
        if "fore_color" in p:
            form["fore_color"] = ole_color_to_hex(p["fore_color"])
        if "caption" in p and not caption_from_vbframe:
            form["caption"] = p["caption"]
        font = info["font"]
        if font.get("font_name"):
            form["font_name"] = font["font_name"]
        if font.get("font_size"):
            form["font_size"] = font["font_size"]
    except Exception as exc:
        _log.warning("UserForm %s: cannot decode the designer (%s); controls dropped", name, exc)
    form["controls"] = controls
    return normalize_form(form)


def extract_vba_modules(path_or_bytes) -> List[dict]:
    """Extract VBA module source from an .xlsm path/bytes or vbaProject.bin bytes.

    Returns [{"name", "type", "code"}] with type in standard|document|class|form.
    Code keeps the 'Attribute VB_...' lines; line endings are normalised to '\\n'.
    """
    vbin = _vba_bin_from(path_or_bytes)
    if not vbin:
        return []
    olevba = _require("oletools.olevba", "oletools")
    olefile = _require("olefile", "olefile")
    import logging
    logging.getLogger("olevba").setLevel(logging.CRITICAL)

    kinds: Dict[str, str] = {}
    storages: set = set()
    storage_names: Dict[str, str] = {}
    ole = None
    try:
        ole = olefile.OleFileIO(io.BytesIO(vbin))
        for entry in ole.listdir(streams=False, storages=True):
            if len(entry) == 1:
                storages.add(entry[0].lower())
                storage_names[entry[0].lower()] = entry[0]
        root = ""
        if not ole.exists("PROJECT"):
            for entry in ole.listdir():
                if entry[-1] == "PROJECT":
                    root = "/".join(entry[:-1])
                    break
        kinds = _project_kinds(ole, root)
    except Exception:
        pass
    storages.discard("vba")

    try:
        parser = olevba.VBA_Parser("vbaProject.bin", data=vbin)
    except Exception as exc:
        raise FileFormatError(f"Cannot read the VBA project: {exc}") from exc
    out: List[dict] = []
    seen = set()
    try:
        if not parser.detect_vba_macros():
            return []
        for _fname, stream_path, vba_filename, code in parser.extract_macros():
            if code is None:
                continue
            if isinstance(code, bytes):
                code = code.decode("cp1252", errors="replace")
            name = os.path.splitext(vba_filename or "")[0] or str(stream_path).split("/")[-1]
            if name.lower() in seen:
                continue
            seen.add(name.lower())
            code = code.replace("\r\n", "\n").replace("\r", "\n")
            kind = _guess_kind(name, vba_filename or "", code, storages, kinds)
            mod = {"name": name, "type": kind, "code": code}
            if kind == "form":
                mod["form"] = _form_for_module(ole, storage_names.get(name.lower()), name)
            out.append(mod)
    finally:
        try:
            parser.close()
        except Exception:
            pass
        if ole is not None:
            try:
                ole.close()
            except Exception:
                pass
    return out


def _form_for_module(ole, storage: Optional[str], name: str) -> dict:
    from .forms import new_form, normalize_form
    if ole is None or storage is None:
        _log.warning("UserForm %s: designer storage not found; using an empty form", name)
        return normalize_form(new_form(name))
    try:
        return parse_form_designer(ole, storage, name)
    except Exception as exc:  # parse_form_designer is defensive; belt and braces
        _log.warning("UserForm %s: cannot read the designer (%s)", name, exc)
        return normalize_form(new_form(name))


# ==========================================================================
# Saving
# ==========================================================================
def save_workbook(book: Workbook, path: str, engine=None,
                  password: Optional[str] = None) -> List[str]:
    """Save `book` to `path` (format from the extension). Returns warnings."""
    ext = _ext(path)
    warnings: List[str] = []
    if ext in _CSV_EXTS:
        if password is not None:
            raise FileFormatError("CSV files cannot be password protected; "
                                  "save as .xlsx or .xlsm to use a password.")
        data = _csv_bytes(book, engine, warnings, delimiter="\t" if ext == "tsv" else ",")
    elif ext in _OOXML_EXTS:
        data = _ooxml_bytes(book, ext, warnings)
        if password is not None:
            data = _encrypt(data, password)
    elif ext == "xls":
        raise FileFormatError("Saving in the old .xls format is not supported; use .xlsx or .xlsm.")
    else:
        raise FileFormatError(f"Unsupported file type '.{ext}'. Use .xlsm, .xlsx or .csv.")
    _atomic_write(path, data)
    book.path = path
    return warnings


def _atomic_write(path: str, data: bytes) -> None:
    path = os.path.abspath(path)
    d = os.path.dirname(path) or "."
    try:
        mode = os.stat(path).st_mode & 0o7777
    except OSError:
        um = os.umask(0)
        os.umask(um)
        mode = 0o666 & ~um
    fd, tmp = tempfile.mkstemp(prefix=".~pe_", suffix=".tmp", dir=d)
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        try:
            os.chmod(tmp, mode)
        except OSError:  # pragma: no cover
            pass
        os.replace(tmp, path)
    except BaseException as exc:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        if isinstance(exc, OSError):
            raise FileFormatError(f"Cannot write '{path}': {exc.strerror or exc}") from exc
        raise


def _encrypt(data: bytes, password: str) -> bytes:
    _require("msoffcrypto", "msoffcrypto-tool")
    try:
        from msoffcrypto.format.ooxml import OOXMLFile
    except ImportError as exc:  # pragma: no cover
        raise FileFormatError("This version of msoffcrypto-tool cannot encrypt files; "
                              "upgrade it (pip install -U msoffcrypto-tool).") from exc
    if not hasattr(OOXMLFile, "encrypt"):  # pragma: no cover
        raise FileFormatError("This version of msoffcrypto-tool cannot encrypt files; "
                              "upgrade it (pip install -U msoffcrypto-tool). "
                              "The workbook was NOT saved.")
    out = io.BytesIO()
    try:
        OOXMLFile(io.BytesIO(data)).encrypt(password, out)
    except Exception as exc:
        raise FileFormatError(f"Encryption failed ({exc}); the workbook was NOT saved.") from exc
    enc = out.getvalue()
    if enc[:8] != _OLE_MAGIC:  # pragma: no cover
        raise FileFormatError("Encryption failed; the workbook was NOT saved.")
    return enc


# -- csv -------------------------------------------------------------------
def _display(engine, sheet: Sheet, r: int, c: int, raw: str) -> str:
    if engine is not None and hasattr(engine, "display_text"):
        try:
            v = engine.display_text(sheet, r, c)
            return "" if v is None else str(v)
        except Exception:
            pass
    if raw.startswith("'"):
        return raw[1:]
    return raw


def _csv_bytes(book: Workbook, engine, warnings: List[str], delimiter: str = ",") -> bytes:
    sheet = book.active
    if len(book.sheets) > 1:
        warnings.append(f"CSV saves only the active sheet ('{sheet.name}').")
    if sheet.styles:
        warnings.append("Formatting is not saved in CSV.")
    if any(getattr(s, "images", None) for s in book.sheets):
        warnings.append("Pictures (e.g. barcodes) are not saved in CSV.")
    if book.vba_modules:
        warnings.append("Macros are not saved in CSV.")
    n_rows, n_cols = sheet.used_bounds()
    buf = io.StringIO(newline="")
    w = csv.writer(buf, delimiter=delimiter, lineterminator="\r\n")
    for r in range(n_rows):
        row = []
        for c in range(n_cols):
            raw = sheet.cells.get((r, c), "")
            row.append(_display(engine, sheet, r, c, raw) if raw else "")
        while row and row[-1] == "":
            row.pop()
        w.writerow(row)
    return buf.getvalue().encode("utf-8-sig")


# -- ooxml -----------------------------------------------------------------
_PCT_RE = re.compile(r"^\s*[-+]?[\d.,]+(?:[eE][-+]?\d+)?%\s*$")


def _model_to_cell_value(raw: str, illegal):
    """-> (value, force_text, implied_number_format)"""
    if raw.startswith("=") and len(raw) > 1:
        return illegal.sub("", raw), False, None
    v = literal_value(raw)
    if isinstance(v, bool):
        return v, False, None
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 1e15:
            v = int(v)
        nf = None
        if _PCT_RE.match(raw):
            nf = "0.00%" if "." in raw else "0%"
        return v, False, nf
    if isinstance(v, str):
        return illegal.sub("", v), True, None
    if v is None:
        return None, False, None
    return str(v), False, None  # XLError -> "#N/A" etc. (openpyxl marks it as error)


def _build_style(st: dict):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.styles.colors import Color
    parts = {}
    if any(k in st for k in ("bold", "italic", "underline", "color", "font_size", "font_name")):
        kw = {"name": st.get("font_name") or "Calibri",
              "sz": float(st.get("font_size") or 11)}
        if st.get("bold"):
            kw["b"] = True
        if st.get("italic"):
            kw["i"] = True
        if st.get("underline"):
            kw["u"] = "single"
        col = _hex_to_argb(st.get("color"))
        if col:
            kw["color"] = Color(rgb=col)
        parts["font"] = Font(**kw)
    fill = _hex_to_argb(st.get("fill"))
    if fill:
        parts["fill"] = PatternFill(fill_type="solid", fgColor=Color(rgb=fill),
                                    bgColor=Color(rgb=fill))
    if st.get("align") in ("left", "center", "right") or st.get("wrap"):
        parts["alignment"] = Alignment(
            horizontal=st.get("align") if st.get("align") in ("left", "center", "right") else None,
            wrap_text=True if st.get("wrap") else None)
    nf = st.get("number_format")
    if nf and nf != "General":
        parts["number_format"] = str(nf)
    return parts


def _hex_to_argb(h) -> Optional[str]:
    if not h or not isinstance(h, str):
        return None
    s = h.strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6 or not re.fullmatch(r"[0-9A-Fa-f]{6}", s):
        return None
    return "FF" + s.upper()


def _ooxml_bytes(book: Workbook, ext: str, warnings: List[str]) -> bytes:
    openpyxl = _require("openpyxl", "openpyxl")
    from copy import copy
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
    from openpyxl.utils import get_column_letter

    macro = ext in _MACRO_EXTS
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if ext in ("xltx", "xltm"):
        wb.template = True

    names: List[str] = []
    renamed: Dict[str, str] = {}
    image_meta: Dict[str, list] = {}
    style_cache: Dict[tuple, object] = {}   # style key -> StyleArray
    illegal = ILLEGAL_CHARACTERS_RE
    for sh in book.sheets:
        title = _safe_sheet_name(sh.name, names)
        if title != sh.name:
            renamed[sh.name] = title
        names.append(title)
        ws = wb.create_sheet(title=title)
        if sh.code_name:
            try:
                ws.sheet_properties.codeName = sh.code_name
            except Exception:
                pass
        if not sh.visible:
            ws.sheet_state = "hidden"

        store = ws._cells
        keys = set(sh.cells) | set(sh.styles)
        for key in sorted(keys):
            r, c = key
            raw = sh.cells.get(key)
            st = sh.styles.get(key)
            cell = Cell(ws, row=r + 1, column=c + 1)
            nf_implied = None
            if raw:
                value, force_text, nf_implied = _model_to_cell_value(raw, illegal)
                if value is not None:
                    cell.value = value
                    if force_text:
                        cell.data_type = "s"
            if nf_implied and not (st and st.get("number_format")):
                st = dict(st or {})
                st["number_format"] = nf_implied
            if st:
                skey = tuple(sorted((k, v) for k, v in st.items()
                                    if k in ("bold", "italic", "underline", "color", "fill",
                                             "align", "font_size", "font_name",
                                             "number_format", "wrap") and v not in (None, False)))
                if skey:
                    sa = style_cache.get(skey)
                    if sa is None:
                        for attr, obj in _build_style(st).items():
                            setattr(cell, attr, obj)
                        style_cache[skey] = copy(cell._style)
                    else:
                        cell._style = copy(sa)
            store[(r + 1, c + 1)] = cell

        for c, px in sorted(sh.col_widths.items()):
            ws.column_dimensions[get_column_letter(c + 1)].width = max(0.0, (px - 5) / 7)
        for r, px in sorted(sh.row_heights.items()):
            ws.row_dimensions[r + 1].height = round(px * 0.75, 2)
        meta_entries = _write_images(ws, sh, warnings)
        if meta_entries:
            image_meta[title] = meta_entries

    if not wb.worksheets:  # pragma: no cover
        wb.create_sheet("Sheet1")
    if renamed:
        warnings.append("Some sheet names were changed to be valid in Excel: " +
                        ", ".join(f"'{a}' -> '{b}'" for a, b in renamed.items()))
    # visible / active sheet
    idx = min(max(book.active_index, 0), len(wb.worksheets) - 1)
    if wb.worksheets[idx].sheet_state != "visible":
        wb.worksheets[idx].sheet_state = "visible"
    wb.active = idx
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws is wb.worksheets[idx]

    _write_defined_names(wb, book.defined_names, warnings)

    write_json = False
    vbin = None
    if book.vba_modules or book.vba_project_bin:
        if macro:
            wb.code_name = "ThisWorkbook"
            vbin = book.vba_project_bin
            write_json = bool(book.vba_modules) and (
                book.vba_modified or not vbin or not _same_modules(book.vba_modules, vbin))
            if write_json and vbin:
                warnings.append("Edited macros are stored for PortableExcel; Microsoft Excel "
                                "will still run the original macros of this file.")
            elif write_json:
                warnings.append("Macros written in PortableExcel are stored for PortableExcel "
                                "only; Microsoft Excel will not see them.")
        else:
            warnings.append(f"Macros are not saved in .{ext}; save as .xlsm to keep them.")

    buf = io.BytesIO()
    try:
        with _exact_floats():
            wb.save(buf)
    except Exception as exc:
        raise FileFormatError(f"Cannot save the workbook: {exc}") from exc
    data = buf.getvalue()
    meta_bytes = None
    if image_meta:
        meta_bytes = json.dumps({"format": "portableexcel-meta", "version": 1,
                                 "images": image_meta},
                                ensure_ascii=False, indent=1).encode("utf-8")
    json_bytes = None
    if write_json:
        json_bytes = json.dumps(
            {"format": "portableexcel-vba", "version": 1,
             "modules": [_module_json(m) for m in book.vba_modules]},
            ensure_ascii=False, indent=1).encode("utf-8")
    return _postprocess_zip(data, ext, vbin if macro else None, json_bytes, meta_bytes)


def _same_modules(modules: List[dict], vbin: bytes) -> bool:
    """True if `modules` equal the source stored in vbaProject.bin `vbin`.

    Needed because a file whose macros were edited earlier loads its modules
    from the JSON part with vba_modified False; re-saving it must keep them.
    """
    try:
        orig = extract_vba_modules(vbin)
    except Exception:
        return False
    norm = lambda ms: [json.dumps(_module_json(m), sort_keys=True, default=str)  # noqa: E731
                       for m in ms]
    return norm(orig) == norm(modules)


def _module_json(m: dict) -> dict:
    d = {"name": m.get("name", ""), "type": m.get("type", "standard"), "code": m.get("code", "")}
    if m.get("form") is not None:
        d["form"] = m["form"]
    return d


class _exact_floats:
    """openpyxl writes numbers with '%.16g', which loses precision for some
    doubles; temporarily use repr() (shortest exact round-trip) instead."""

    def __enter__(self):
        try:
            from openpyxl.cell import _writer
        except ImportError:  # pragma: no cover
            self._mod = None
            return self
        self._mod = _writer
        self._orig = _writer.safe_string
        orig = self._orig

        def safe_string(value):
            if type(value) is float and value == value and value not in (_INF, -_INF):
                return repr(value)
            return orig(value)

        _writer.safe_string = safe_string
        return self

    def __exit__(self, *exc):
        if self._mod is not None:
            self._mod.safe_string = self._orig
        return False


_INF = float("inf")


def _write_defined_names(wb, defined: Dict[str, str], warnings: List[str]) -> None:
    if not defined:
        return
    from openpyxl.workbook.defined_name import DefinedName
    bad = []
    for name, ref in defined.items():
        try:
            dn = DefinedName(name, attr_text=str(ref).lstrip("="))
            if hasattr(wb.defined_names, "add"):
                wb.defined_names.add(dn)
            elif hasattr(wb.defined_names, "append"):  # pragma: no cover - openpyxl < 3.1
                wb.defined_names.append(dn)
            else:  # pragma: no cover
                wb.defined_names[name] = dn
        except Exception:
            bad.append(name)
    if bad:
        warnings.append("Some defined names could not be saved: " + ", ".join(bad))


def _postprocess_zip(data: bytes, ext: str, vbin: Optional[bytes],
                     json_bytes: Optional[bytes], meta_bytes: Optional[bytes] = None) -> bytes:
    """Fix content types for macro/template files and inject VBA / JSON parts."""
    if ext == "xlsx" and vbin is None and json_bytes is None and meta_bytes is None:
        return data
    src = zipfile.ZipFile(io.BytesIO(data))
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            name = info.filename
            if name in (_VBA_BIN_PART, VBA_JSON_PART, META_JSON_PART):
                continue
            content = src.read(name)
            if name == "[Content_Types].xml":
                content = _fix_content_types(content.decode("utf-8"), ext,
                                             vbin is not None,
                                             json_bytes is not None or meta_bytes is not None
                                             ).encode("utf-8")
            elif name == "xl/_rels/workbook.xml.rels" and vbin is not None:
                content = _add_vba_rel(content.decode("utf-8")).encode("utf-8")
            dst.writestr(info, content, compress_type=zipfile.ZIP_DEFLATED)
        if vbin is not None:
            dst.writestr(_VBA_BIN_PART, vbin, compress_type=zipfile.ZIP_DEFLATED)
        if json_bytes is not None:
            dst.writestr(VBA_JSON_PART, json_bytes, compress_type=zipfile.ZIP_DEFLATED)
        if meta_bytes is not None:
            dst.writestr(META_JSON_PART, meta_bytes, compress_type=zipfile.ZIP_DEFLATED)
    src.close()
    return out_buf.getvalue()


def _fix_content_types(xml: str, ext: str, has_vba: bool, has_json: bool) -> str:
    main = _CT_MAIN[ext]
    xml = re.sub(r'(PartName="/xl/workbook\.xml"\s+ContentType=")[^"]*(")',
                 lambda m: m.group(1) + main + m.group(2), xml)
    xml = re.sub(r'(ContentType=")[^"]*(\"\s+PartName="/xl/workbook\.xml")',
                 lambda m: m.group(1) + main + m.group(2), xml)
    xml = re.sub(r'<Override[^>]*PartName="/xl/vbaProject\.bin"[^>]*/>', "", xml)
    inserts = []
    if has_vba:
        inserts.append(f'<Override PartName="/xl/vbaProject.bin" ContentType="{_CT_VBA}"/>')
        if not re.search(r'<Default[^>]*Extension="bin"', xml):
            inserts.insert(0, f'<Default Extension="bin" ContentType="{_CT_VBA}"/>')
    if has_json and not re.search(r'<Default[^>]*Extension="json"', xml):
        inserts.insert(0, '<Default Extension="json" ContentType="application/json"/>')
    if inserts:
        # Defaults must come before Overrides; put new Defaults right after <Types ...>
        defaults = "".join(s for s in inserts if s.startswith("<Default"))
        overrides = "".join(s for s in inserts if s.startswith("<Override"))
        xml = re.sub(r"(<Types[^>]*>)", lambda m: m.group(1) + defaults, xml, count=1)
        xml = xml.replace("</Types>", overrides + "</Types>")
    return xml


def _add_vba_rel(xml: str) -> str:
    if _REL_VBA in xml:
        return xml
    ids = set(re.findall(r'Id="([^"]+)"', xml))
    n = 1
    while f"rId{n}" in ids:
        n += 1
    rel = f'<Relationship Id="rId{n}" Type="{_REL_VBA}" Target="vbaProject.bin"/>'
    return xml.replace("</Relationships>", rel + "</Relationships>")

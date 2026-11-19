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
import os
import re
import tempfile
import zipfile
from typing import Dict, List, Optional

from .model import Sheet, Workbook, literal_value


class PasswordRequired(Exception):
    """The file is encrypted and no password, or a wrong one, was given."""


class FileFormatError(Exception):
    """Unreadable or unsupported file; str(exc) is a message for the user."""


SUPPORTED_OPEN = "Excel files (*.xlsx *.xlsm *.xltx *.xltm);;CSV (*.csv);;All files (*)"
SUPPORTED_SAVE = "Excel Macro-Enabled Workbook (*.xlsm);;Excel Workbook (*.xlsx);;CSV UTF-8 (*.csv)"

VBA_JSON_PART = "customXml/portableexcel_vba.json"
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
        wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=False,
                                    data_only=False, keep_links=False)
    except Exception as exc:
        zf.close()
        raise FileFormatError(f"Cannot read the workbook: {exc}") from exc

    book = Workbook()
    book.sheets = []
    used_codes = set()
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
        book.sheets.append(sh)
    if not book.sheets:
        book.sheets = [Sheet("Sheet1")]
    try:
        active = wb.active
        book.active_index = next((k for k, ws in enumerate(wb.worksheets) if ws is active), 0)
    except Exception:
        book.active_index = 0

    book.defined_names = _read_defined_names(wb)

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
            book.vba_modules = [
                {"name": str(m["name"]), "type": str(m.get("type", "standard")),
                 "code": str(m.get("code", "")).replace("\r\n", "\n").replace("\r", "\n")}
                for m in mods if isinstance(m, dict) and "name" in m]
        except Exception:
            pass
    book.vba_modified = False
    zf.close()
    try:
        wb.close()
    except Exception:
        pass
    return book


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
    try:
        ole = olefile.OleFileIO(io.BytesIO(vbin))
        for entry in ole.listdir(streams=False, storages=True):
            if len(entry) == 1:
                storages.add(entry[0].lower())
        root = ""
        if not ole.exists("PROJECT"):
            for entry in ole.listdir():
                if entry[-1] == "PROJECT":
                    root = "/".join(entry[:-1])
                    break
        kinds = _project_kinds(ole, root)
        ole.close()
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
            out.append({"name": name, "type": kind, "code": code})
    finally:
        try:
            parser.close()
        except Exception:
            pass
    return out


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
    json_bytes = None
    if write_json:
        json_bytes = json.dumps(
            {"format": "portableexcel-vba", "version": 1,
             "modules": [{"name": m.get("name", ""), "type": m.get("type", "standard"),
                          "code": m.get("code", "")} for m in book.vba_modules]},
            ensure_ascii=False, indent=1).encode("utf-8")
    return _postprocess_zip(data, ext, vbin if macro else None, json_bytes)


def _same_modules(modules: List[dict], vbin: bytes) -> bool:
    """True if `modules` equal the source stored in vbaProject.bin `vbin`.

    Needed because a file whose macros were edited earlier loads its modules
    from the JSON part with vba_modified False; re-saving it must keep them.
    """
    try:
        orig = extract_vba_modules(vbin)
    except Exception:
        return False
    norm = lambda ms: [(m.get("name"), m.get("type"), m.get("code")) for m in ms]  # noqa: E731
    return norm(orig) == norm(modules)


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
                     json_bytes: Optional[bytes]) -> bytes:
    """Fix content types for macro/template files and inject VBA parts."""
    if ext == "xlsx" and vbin is None and json_bytes is None:
        return data
    src = zipfile.ZipFile(io.BytesIO(data))
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            name = info.filename
            if name in (_VBA_BIN_PART, VBA_JSON_PART):
                continue
            content = src.read(name)
            if name == "[Content_Types].xml":
                content = _fix_content_types(content.decode("utf-8"), ext,
                                             vbin is not None, json_bytes is not None).encode("utf-8")
            elif name == "xl/_rels/workbook.xml.rels" and vbin is not None:
                content = _add_vba_rel(content.decode("utf-8")).encode("utf-8")
            dst.writestr(info, content, compress_type=zipfile.ZIP_DEFLATED)
        if vbin is not None:
            dst.writestr(_VBA_BIN_PART, vbin, compress_type=zipfile.ZIP_DEFLATED)
        if json_bytes is not None:
            dst.writestr(VBA_JSON_PART, json_bytes, compress_type=zipfile.ZIP_DEFLATED)
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

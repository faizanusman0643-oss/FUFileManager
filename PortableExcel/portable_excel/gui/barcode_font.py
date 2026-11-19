"""Barcode-font support for cells.

Workbooks often format cells with a Code 128 barcode font ("Code 128",
"IDAutomationSC128XS DEMO", "Libre Barcode 128 Text", ...) and fill them with
font-encoded text produced by VBA. When such a font is not installed we
emulate it: decode the characters to Code 128 symbol values and paint the bars.

The bundled OFL Libre Barcode fonts (gui/fonts) are registered at start-up.
"""
from __future__ import annotations

import logging
import os
import re
from typing import List, Optional, Tuple

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QPainter

log = logging.getLogger("PortableExcel")

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
_BARCODE_FONT_RE = re.compile(r"code\s*128|idautomation.*128|libre\s*barcode\s*128", re.I)
_TEXT_VARIANT_RE = re.compile(r"\btext\b|libre\s*barcode\s*128\s*text", re.I)
# offsets some generators add to start/stop/check characters (IDAutomation
# ReturnType 6-9 variants)
_HIGH_OFFSETS = (11500, 11300, 10700, 10500)

START_A, START_B, START_C, STOP = 103, 104, 105, 106

_registered = [False]
_families_cache: Optional[set] = None


def register_bundled_fonts() -> List[str]:
    """Register the fonts shipped in gui/fonts with Qt (once)."""
    global _families_cache
    if _registered[0]:
        return []
    _registered[0] = True
    fams = []
    if os.path.isdir(FONT_DIR):
        for fn in sorted(os.listdir(FONT_DIR)):
            if fn.lower().endswith((".ttf", ".otf")):
                fid = QFontDatabase.addApplicationFont(os.path.join(FONT_DIR, fn))
                if fid >= 0:
                    fams.extend(QFontDatabase.applicationFontFamilies(fid))
    _families_cache = None
    return fams


def font_installed(name: str) -> bool:
    global _families_cache
    if _families_cache is None:
        _families_cache = {f.lower() for f in QFontDatabase.families()}
    return (name or "").strip().lower() in _families_cache


def is_barcode_font(name: Optional[str]) -> bool:
    return bool(name) and bool(_BARCODE_FONT_RE.search(name))


def is_text_variant(name: Optional[str]) -> bool:
    return bool(name) and "text" in name.lower()


def needs_emulation(name: Optional[str]) -> bool:
    return is_barcode_font(name) and not font_installed(name)


def char_value(ch: str) -> Optional[int]:
    """Code 128 symbol value of one font character (None = not a symbol)."""
    c = ord(ch)
    if c > 255:
        for off in _HIGH_OFFSETS:
            if 32 <= c - off <= 211:
                c -= off
                break
        else:
            return None
    if c == 32 or c == 194:
        return 0
    if 33 <= c <= 126:
        return c - 32
    if 195 <= c <= 206:
        return c - 100
    if 207 <= c <= 211:  # some fonts put values 95..106 at 200..211
        return c - 105
    return None


def decode(text: str) -> List[int]:
    vals = []
    for ch in text or "":
        v = char_value(ch)
        if v is not None:
            vals.append(v)
    return vals


def is_valid_symbol(vals: List[int]) -> bool:
    """Start code, data, correct check value and stop code."""
    if len(vals) < 3 or vals[0] not in (START_A, START_B, START_C) or vals[-1] != STOP:
        return False
    total = vals[0] + sum(v * i for i, v in enumerate(vals[1:-2], 1))
    return total % 103 == vals[-2]


def human_text(vals: List[int]) -> str:
    """Readable data of a decoded symbol (without start/check/stop)."""
    if not vals:
        return ""
    data = vals
    cset = "B"
    if vals[0] in (START_A, START_B, START_C):
        cset = {START_A: "A", START_B: "B", START_C: "C"}[vals[0]]
        data = vals[1:]
        if data and data[-1] == STOP:
            data = data[:-2] if len(data) >= 2 else []
    out = []
    shift = False
    for v in data:
        cur = cset
        if shift:
            cur = "A" if cset == "B" else "B"
            shift = False
        if cur == "C":
            if v < 100:
                out.append(f"{v:02d}")
            elif v == 100:
                cset = "B"
            elif v == 101:
                cset = "A"
            continue
        if v == 99:
            cset = "C"
        elif v == 100 and cur == "A":
            cset = "B"
        elif v == 101 and cur == "B":
            cset = "A"
        elif v == 98:
            shift = True
        elif v >= 96:
            continue  # FNC codes
        elif cur == "A":
            out.append(chr(v + 32) if v < 64 else chr(v - 64))
        else:
            out.append(chr(v + 32))
    return "".join(ch if ch.isprintable() else " " for ch in out)


def _widths_table():
    try:
        from .. import barcodes
        return list(barcodes.C128_WIDTHS), barcodes.C128_STOP
    except Exception:
        return None, None


def module_runs(vals: List[int]) -> List[int]:
    """Alternating bar/space widths (in modules) for the symbol values."""
    table, stop = _widths_table()
    if table is None:
        return []
    runs: List[int] = []
    for v in vals:
        pat = stop if v == STOP else (table[v] if 0 <= v < len(table) else "")
        runs.extend(int(x) for x in pat)
    return runs


def paint(p: QPainter, rect, text: str, font_name: str, color=None,
          show_text: Optional[bool] = None, text_font: Optional[QFont] = None) -> bool:
    """Draw font-encoded Code 128 text as bars into rect. False if nothing drawn."""
    vals = decode(text)
    runs = module_runs(vals)
    if not runs:
        return False
    if show_text is None:
        show_text = is_text_variant(font_name)
    r = QRectF(rect).adjusted(2, 2, -2, -2)
    if r.width() <= 4 or r.height() <= 4:
        return False
    th = 0.0
    if show_text:
        th = min(r.height() * 0.35, max(9.0, r.height() * 0.22))
    total = sum(runs)
    quiet = 6 if total + 12 < r.width() else 0
    unit = (r.width() - 2 * quiet) / max(1, total)
    if unit <= 0:
        return False
    # centre a reasonable-width barcode (not wider than 3 px per module)
    unit = min(unit, 3.0)
    width = unit * total
    x = r.left() + (r.width() - width) / 2
    bar_h = r.height() - th
    p.save()
    p.setRenderHint(QPainter.Antialiasing, False)
    col = QColor(color) if color else QColor("#000000")
    bar = True
    for w in runs:
        if bar:
            p.fillRect(QRectF(x, r.top(), w * unit, bar_h), col)
        x += w * unit
        bar = not bar
    if show_text:
        f = QFont(text_font) if text_font is not None else QFont()
        f.setPixelSize(max(6, int(th * 0.85)))
        p.setFont(f)
        p.setPen(col)
        p.drawText(QRectF(r.left(), r.top() + bar_h, r.width(), th), Qt.AlignCenter,
                   human_text(vals))
    p.restore()
    return True


def encode_b(data: str) -> str:
    """Font-encode data as Code 128 set B (for tests / demo data)."""
    vals = [START_B] + [ord(c) - 32 for c in data]
    check = (vals[0] + sum(v * i for i, v in enumerate(vals[1:], 1))) % 103
    vals += [check, STOP]

    def ch(v):
        if v == 0:
            return chr(194)
        return chr(v + 32) if v < 95 else chr(v + 100)
    return "".join(ch(v) for v in vals)

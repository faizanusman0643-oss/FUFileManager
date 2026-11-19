"""Throwaway stand-in for portable_excel.barcodes (GUI tests / screenshots only).

Draws a fake bar pattern derived from the text; validation follows the real
rules closely enough for the GUI (digits / check digits for EAN/UPC).
"""
import uuid

from portable_excel.model import parse_a1, value_to_text

SYMBOLOGIES = {
    "code128": "Code 128", "code39": "Code 39", "ean13": "EAN-13", "ean8": "EAN-8",
    "upca": "UPC-A", "itf": "ITF (Interleaved 2 of 5)", "qr": "QR Code",
    "datamatrix": "Data Matrix",
}
_DIGITS = {"ean13": 12, "ean8": 7, "upca": 11}


def _check(digits):
    s = sum(int(d) * (3 if i % 2 == 0 else 1) for i, d in enumerate(reversed(digits)))
    return str((10 - s % 10) % 10)


def validate(sym, text):
    text = "" if text is None else str(text)
    if sym not in SYMBOLOGIES:
        return f"Unknown symbology {sym!r}"
    if not text:
        return "No data to encode"
    if sym in _DIGITS:
        n = _DIGITS[sym]
        if not text.isdigit():
            return "Only digits are allowed"
        if len(text) not in (n, n + 1):
            return f"Needs {n} digits (or {n + 1} with check digit)"
        if len(text) == n + 1 and _check(text[:-1]) != text[-1]:
            return f"Wrong check digit (expected {_check(text[:-1])})"
    if sym == "itf" and (not text.isdigit() or len(text) % 2):
        return "ITF needs an even number of digits"
    if sym == "code39" and any(ch not in "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
                               for ch in text.upper()):
        return "Invalid character for Code 39"
    return None


def normalize(sym, text):
    text = "" if text is None else str(text)
    if sym in _DIGITS and text.isdigit() and len(text) == _DIGITS[sym]:
        return text + _check(text)
    if sym == "code39":
        return text.upper()
    return text


def render_png(sym, text, module_px=2, bar_height_px=60, show_text=True):
    from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QRect, Qt
    from PySide6.QtGui import QColor, QFont, QImage, QPainter
    err = validate(sym, text)
    if err:
        raise ValueError(err)
    text = normalize(sym, text)
    m = max(1, int(module_px))
    if sym in ("qr", "datamatrix"):
        n = 21 if sym == "qr" else 18
        w = h = (n + 8) * m
        img = QImage(w, h, QImage.Format_RGB32)
        img.fill(Qt.white)
        p = QPainter(img)
        seed = sum(map(ord, text)) or 1
        for y in range(n):
            for x in range(n):
                seed = (seed * 1103515245 + 12345) & 0x7FFFFFFF
                corner = (x < 7 and y < 7) or (x >= n - 7 and y < 7) or (x < 7 and y >= n - 7)
                on = ((x in (0, 6) or y in (0, 6) or (2 <= x <= 4 and 2 <= y <= 4))
                      if corner else (seed >> 16) & 1)
                if corner and sym == "qr":
                    for ox, oy in ((0, 0), (n - 7, 0), (0, n - 7)):
                        if ox <= x < ox + 7 and oy <= y < oy + 7:
                            lx, ly = x - ox, y - oy
                            on = lx in (0, 6) or ly in (0, 6) or (2 <= lx <= 4 and 2 <= ly <= 4)
                if on:
                    p.fillRect((x + 4) * m, (y + 4) * m, m, m, QColor("black"))
        p.end()
    else:
        modules = []
        seed = sum(map(ord, text)) * 7 + 3
        for ch in "0" + text + "0":
            seed = (seed * 31 + ord(ch)) & 0xFFFF
            modules.extend([1, 0, (seed >> 3) & 1, 1, (seed >> 7) & 1, 0, 1, (seed >> 11) & 1])
        modules = [1, 0, 1] + modules + [1, 0, 1]
        text_h = 16 if show_text else 0
        w = (len(modules) + 20) * m
        h = int(bar_height_px) + text_h + 6
        img = QImage(w, h, QImage.Format_RGB32)
        img.fill(Qt.white)
        p = QPainter(img)
        for i, on in enumerate(modules):
            if on:
                p.fillRect((i + 10) * m, 3, m, int(bar_height_px), QColor("black"))
        if show_text:
            f = QFont("DejaVu Sans Mono")
            f.setPixelSize(12)
            p.setFont(f)
            p.drawText(QRect(0, int(bar_height_px) + 3, w, text_h), Qt.AlignCenter, text)
        p.end()
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    return bytes(ba.data()), img.width(), img.height()


def _source_text(sheet, source, engine):
    rc = parse_a1(source) if source else None
    if rc is None:
        return ""
    if engine is not None:
        try:
            return engine.display_text(sheet, rc[0], rc[1])
        except Exception:
            pass
    return value_to_text(sheet.get_raw(*rc))


def make_barcode_image(sheet, source, text, symbology, engine=None, row=0, col=0,
                       show_text=True, module_px=2, bar_height_px=60, **opts):
    data_text = _source_text(sheet, source, engine) if source else (text or "")
    png, w, h = render_png(symbology, data_text, module_px, bar_height_px, show_text)
    return {"id": uuid.uuid4().hex, "name": "Barcode", "row": row, "col": col, "dx": 2,
            "dy": 2, "width": w, "height": h, "data": png, "fmt": "png",
            "barcode": {"symbology": symbology, "source": source, "text": text or "",
                        "show_text": show_text, "module_px": module_px,
                        "bar_height_px": bar_height_px}}


def refresh_barcodes(book, engine, sheet=None, changed_cells=None):
    updated = []
    for sh in ([sheet] if sheet is not None else book.sheets):
        for img in sh.images:
            bc = img.get("barcode")
            if not bc or not bc.get("source"):
                continue
            rc = parse_a1(bc["source"])
            if changed_cells is not None and rc not in set(map(tuple, changed_cells)):
                continue
            text = _source_text(sh, bc["source"], engine)
            if validate(bc["symbology"], text):
                continue
            png, w, h = render_png(bc["symbology"], text, bc.get("module_px", 2),
                                   bc.get("bar_height_px", 60), bc.get("show_text", True))
            if png != img.get("data"):
                img["data"] = png
                updated.append(img)
    return updated

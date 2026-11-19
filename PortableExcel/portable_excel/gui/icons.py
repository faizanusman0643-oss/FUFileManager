"""Generic icons drawn with QPainter (no image files, no third-party artwork)."""
from __future__ import annotations

from typing import Callable, Dict, Optional

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QBrush, QColor, QFont, QIcon, QPainter, QPainterPath,
                           QPen, QPixmap, QPolygonF)

GREEN = QColor("#217346")
DARK = QColor("#444444")
BLUE = QColor("#2b579a")
RED = QColor("#c43e1c")

_cache: Dict[str, QIcon] = {}


def _pix(size: int) -> QPixmap:
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    return pm


def app_pixmap(size: int = 64) -> QPixmap:
    """Green square, white grid, bold letter P."""
    pm = _pix(size)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    s = size / 64.0
    p.setPen(Qt.NoPen)
    p.setBrush(GREEN)
    p.drawRoundedRect(QRectF(2 * s, 2 * s, 60 * s, 60 * s), 8 * s, 8 * s)
    # white sheet with grid on the right half
    p.setBrush(QColor("#ffffff"))
    sheet = QRectF(30 * s, 12 * s, 26 * s, 40 * s)
    p.drawRect(sheet)
    p.setPen(QPen(GREEN, max(1.0, 1.6 * s)))
    for i in range(1, 4):
        y = sheet.top() + i * sheet.height() / 4
        p.drawLine(QPointF(sheet.left(), y), QPointF(sheet.right(), y))
    x = sheet.left() + sheet.width() / 2
    p.drawLine(QPointF(x, sheet.top()), QPointF(x, sheet.bottom()))
    # darker plate with the letter
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#185c37"))
    p.drawRoundedRect(QRectF(8 * s, 16 * s, 30 * s, 32 * s), 3 * s, 3 * s)
    f = QFont("Segoe UI")
    f.setBold(True)
    f.setPixelSize(max(6, int(26 * s)))
    p.setFont(f)
    p.setPen(QColor("#ffffff"))
    p.drawText(QRectF(8 * s, 16 * s, 30 * s, 32 * s), Qt.AlignCenter, "P")
    p.end()
    return pm


def app_icon() -> QIcon:
    ic = _cache.get("__app__")
    if ic is None:
        ic = QIcon()
        for sz in (16, 24, 32, 48, 64, 128, 256):
            ic.addPixmap(app_pixmap(sz))
        _cache["__app__"] = ic
    return ic


# --------------------------------------------------------------------------
# drawing helpers (all on a 32x32 logical canvas)
# --------------------------------------------------------------------------
def _text(p: QPainter, txt: str, color=DARK, size=20, bold=True, italic=False,
          underline=False, rect=QRectF(0, 0, 32, 32), family="Segoe UI"):
    f = QFont(family)
    f.setPixelSize(size)
    f.setBold(bold)
    f.setItalic(italic)
    f.setUnderline(underline)
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(rect, Qt.AlignCenter, txt)


def _sheet(p, x=6, y=3, w=20, h=26, fold=6, fill="#ffffff", pen=DARK):
    path = QPainterPath()
    path.moveTo(x, y)
    path.lineTo(x + w - fold, y)
    path.lineTo(x + w, y + fold)
    path.lineTo(x + w, y + h)
    path.lineTo(x, y + h)
    path.closeSubpath()
    p.setPen(QPen(QColor(pen), 1.6))
    p.setBrush(QColor(fill))
    p.drawPath(path)


def _grid(p, x, y, w, h, rows=3, cols=3, color=GREEN):
    p.setPen(QPen(color, 1.4))
    p.setBrush(QColor("#ffffff"))
    p.drawRect(QRectF(x, y, w, h))
    for i in range(1, rows):
        yy = y + i * h / rows
        p.drawLine(QPointF(x, yy), QPointF(x + w, yy))
    for j in range(1, cols):
        xx = x + j * w / cols
        p.drawLine(QPointF(xx, y), QPointF(xx, y + h))


def _lines(p, aligns, color=DARK):
    p.setPen(QPen(QColor(color), 2.2))
    y = 7
    for (x1, x2) in aligns:
        p.drawLine(QPointF(x1, y), QPointF(x2, y))
        y += 6


def _arrow(p, pts, color=DARK, width=2.4):
    p.setPen(QPen(QColor(color), width, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    p.setBrush(Qt.NoBrush)
    path = QPainterPath(QPointF(*pts[0]))
    for pt in pts[1:]:
        path.lineTo(QPointF(*pt))
    p.drawPath(path)


def _undo(p, mirror=False):
    p.save()
    if mirror:
        p.translate(32, 0)
        p.scale(-1, 1)
    p.setPen(QPen(QColor("#ffffff"), 2.6, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
    p.setBrush(Qt.NoBrush)
    path = QPainterPath(QPointF(9, 13))
    path.cubicTo(QPointF(16, 6), QPointF(28, 9), QPointF(26, 20))
    p.drawPath(path)
    p.setBrush(QColor("#ffffff"))
    p.setPen(Qt.NoPen)
    p.drawPolygon(QPolygonF([QPointF(4, 8), QPointF(13, 18), QPointF(4, 19)]))
    p.restore()


def _draw_save(p, color="#ffffff"):
    c = QColor(color)
    p.setPen(QPen(c, 2))
    p.setBrush(Qt.NoBrush)
    p.drawRoundedRect(QRectF(5, 5, 22, 22), 2, 2)
    p.setBrush(c)
    p.drawRect(QRectF(10, 5, 12, 7))
    p.setBrush(Qt.NoBrush)
    p.drawRect(QRectF(9, 17, 14, 10))


def _clipboard(p):
    p.setPen(QPen(DARK, 1.6))
    p.setBrush(QColor("#c8a165"))
    p.drawRoundedRect(QRectF(5, 5, 20, 24), 2, 2)
    p.setBrush(QColor("#ffffff"))
    p.drawRect(QRectF(9, 9, 12, 16))
    p.setBrush(DARK)
    p.drawRoundedRect(QRectF(11, 3, 8, 5), 1.5, 1.5)


def _scissors(p):
    p.setPen(QPen(DARK, 2))
    p.setBrush(Qt.NoBrush)
    p.drawEllipse(QRectF(5, 19, 8, 8))
    p.drawEllipse(QRectF(19, 19, 8, 8))
    p.drawLine(QPointF(11, 20), QPointF(21, 4))
    p.drawLine(QPointF(21, 20), QPointF(11, 4))


def _copy(p):
    _sheet(p, 4, 3, 16, 20, 4)
    _sheet(p, 12, 9, 16, 20, 4, fill="#e8f3ec")


def _paint_a(p, color):
    _text(p, "A", DARK, 20, True, rect=QRectF(0, -3, 32, 28))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(color))
    p.drawRect(QRectF(4, 25, 24, 5))


def _bucket(p, color):
    p.setPen(QPen(DARK, 1.6))
    p.setBrush(QColor("#ffffff"))
    poly = QPolygonF([QPointF(6, 12), QPointF(15, 3), QPointF(24, 12), QPointF(15, 21)])
    p.drawPolygon(poly)
    p.setBrush(DARK)
    p.drawEllipse(QRectF(23, 13, 4, 6))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(color))
    p.drawRect(QRectF(4, 25, 24, 5))


def _sigma(p):
    _text(p, "Σ", DARK, 26, True)


def _sort(p, desc=False):
    a, b = ("Z", "A") if desc else ("A", "Z")
    _text(p, a, BLUE, 12, True, rect=QRectF(2, 1, 16, 14))
    _text(p, b, BLUE, 12, True, rect=QRectF(2, 16, 16, 14))
    _arrow(p, [(23, 4), (23, 27)])
    _arrow(p, [(18, 21), (23, 27), (28, 21)])


def _find(p):
    p.setPen(QPen(DARK, 2.6))
    p.setBrush(QColor("#dcebf7"))
    p.drawEllipse(QRectF(4, 4, 16, 16))
    p.setPen(QPen(DARK, 4, Qt.SolidLine, Qt.RoundCap))
    p.drawLine(QPointF(18, 18), QPointF(27, 27))


def _rows_insert(p, delete=False, cols=False):
    p.save()
    if cols:
        p.translate(32, 0)
        p.rotate(90)
    _grid(p, 4, 6, 18, 20, 3, 2, DARK)
    p.setPen(Qt.NoPen)
    p.setBrush(RED if delete else GREEN)
    p.drawRect(QRectF(4, 12.6, 18, 6.6))
    p.restore()
    p.setPen(QPen(RED if delete else GREEN, 2.6))
    if delete:
        p.drawLine(QPointF(22, 20), QPointF(30, 28))
        p.drawLine(QPointF(30, 20), QPointF(22, 28))
    else:
        p.drawLine(QPointF(26, 19), QPointF(26, 29))
        p.drawLine(QPointF(21, 24), QPointF(31, 24))


def _fx(p):
    _text(p, "fx", DARK, 18, False, True, family="Times New Roman")


def _play(p, color=GREEN):
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(color))
    p.drawPolygon(QPolygonF([QPointF(9, 5), QPointF(27, 16), QPointF(9, 27)]))


def _stop(p):
    p.setPen(Qt.NoPen)
    p.setBrush(RED)
    p.drawRect(QRectF(8, 8, 16, 16))


def _code(p):
    p.setPen(QPen(BLUE, 1.6))
    p.setBrush(QColor("#ffffff"))
    p.drawRect(QRectF(3, 5, 26, 22))
    p.setBrush(BLUE)
    p.drawRect(QRectF(3, 5, 26, 5))
    _arrow(p, [(12, 13), (7, 18), (12, 23)], GREEN, 2)
    _arrow(p, [(20, 13), (25, 18), (20, 23)], GREEN, 2)


def _macros(p):
    _grid(p, 3, 4, 18, 18, 3, 3, DARK)
    p.setPen(QPen(QColor("#ffffff"), 1))
    p.setBrush(GREEN)
    p.drawEllipse(QRectF(14, 14, 16, 16))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#ffffff"))
    p.drawPolygon(QPolygonF([QPointF(19.5, 18), QPointF(26, 22), QPointF(19.5, 26)]))


def _shield(p):
    path = QPainterPath(QPointF(16, 3))
    path.lineTo(27, 7)
    path.cubicTo(QPointF(27, 18), QPointF(22, 25), QPointF(16, 29))
    path.cubicTo(QPointF(10, 25), QPointF(5, 18), QPointF(5, 7))
    path.closeSubpath()
    p.setPen(QPen(QColor("#b58b00"), 1.6))
    p.setBrush(QColor("#ffd966"))
    p.drawPath(path)
    _text(p, "!", DARK, 16, True)


def _wrap(p):
    _arrow(p, [(4, 8), (24, 8)])
    _arrow(p, [(4, 15), (24, 15), (27, 18), (24, 22), (14, 22)])
    _arrow(p, [(18, 18), (14, 22), (18, 26)])


def _new(p):
    _sheet(p)
    _grid(p, 9, 12, 14, 12, 3, 2)


def _open(p):
    p.setPen(QPen(QColor("#b58b00"), 1.6))
    p.setBrush(QColor("#ffd966"))
    path = QPainterPath(QPointF(3, 8))
    path.lineTo(12, 8)
    path.lineTo(14, 11)
    path.lineTo(26, 11)
    path.lineTo(26, 26)
    path.lineTo(3, 26)
    path.closeSubpath()
    p.drawPath(path)
    p.setBrush(QColor("#ffe699"))
    p.drawPolygon(QPolygonF([QPointF(7, 15), QPointF(30, 15), QPointF(26, 26), QPointF(3, 26)]))


def _key(p):
    p.setPen(QPen(QColor("#b58b00"), 2.4))
    p.setBrush(QColor("#ffd966"))
    p.drawEllipse(QRectF(3, 9, 13, 13))
    p.drawLine(QPointF(16, 15.5), QPointF(29, 15.5))
    p.drawLine(QPointF(25, 15.5), QPointF(25, 21))
    p.drawLine(QPointF(21, 15.5), QPointF(21, 20))


def _lock(p):
    p.setPen(QPen(DARK, 2.6))
    p.setBrush(Qt.NoBrush)
    p.drawArc(QRectF(9, 3, 14, 18), 0, 180 * 16)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#b58b00"))
    p.drawRoundedRect(QRectF(6, 13, 20, 16), 2, 2)


def _zoom(p, sign):
    _find(p)
    p.setPen(QPen(DARK, 2))
    p.drawLine(QPointF(8, 12), QPointF(16, 12))
    if sign == "+":
        p.drawLine(QPointF(12, 8), QPointF(12, 16))


def _freeze(p, which):
    _grid(p, 4, 4, 24, 24, 4, 4, QColor("#888888"))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(33, 115, 70, 120))
    if which in ("row", "panes"):
        p.drawRect(QRectF(4, 4, 24, 6))
    if which in ("col", "panes"):
        p.drawRect(QRectF(4, 4, 6, 24))
    p.setPen(QPen(GREEN, 2.4))
    if which in ("row", "panes"):
        p.drawLine(QPointF(4, 10), QPointF(28, 10))
    if which in ("col", "panes"):
        p.drawLine(QPointF(10, 4), QPointF(10, 28))


def _gridlines(p):
    _grid(p, 4, 4, 24, 24, 4, 4, QColor("#999999"))


def _table(p):
    _grid(p, 3, 5, 26, 22, 4, 3, BLUE)
    p.setPen(Qt.NoPen)
    p.setBrush(BLUE)
    p.drawRect(QRectF(3, 5, 26, 5.5))


def _calc(p):
    p.setPen(QPen(DARK, 1.6))
    p.setBrush(QColor("#f2f2f2"))
    p.drawRoundedRect(QRectF(6, 3, 20, 26), 2, 2)
    p.setBrush(QColor("#c6e0b4"))
    p.drawRect(QRectF(9, 6, 14, 6))
    p.setBrush(DARK)
    for i in range(3):
        for j in range(3):
            p.drawRect(QRectF(9 + j * 5, 15 + i * 4.5, 3, 2.5))


def _names(p):
    _text(p, "A1", GREEN, 13, True, rect=QRectF(0, 0, 32, 16))
    p.setPen(QPen(DARK, 1.4))
    p.setBrush(QColor("#ffffff"))
    p.drawRect(QRectF(4, 17, 24, 11))
    p.drawLine(QPointF(7, 22.5), QPointF(22, 22.5))


def _duplicates(p):
    _grid(p, 3, 4, 16, 18, 3, 1, DARK)
    p.setPen(QPen(RED, 2.6))
    p.drawLine(QPointF(17, 17), QPointF(28, 28))
    p.drawLine(QPointF(28, 17), QPointF(17, 28))


def _clear(p):
    p.setPen(QPen(DARK, 1.4))
    p.setBrush(QColor("#f4b183"))
    p.drawPolygon(QPolygonF([QPointF(4, 20), QPointF(16, 8), QPointF(26, 18), QPointF(16, 28),
                             QPointF(12, 28)]))
    p.setBrush(QColor("#ffffff"))
    p.drawPolygon(QPolygonF([QPointF(4, 20), QPointF(10, 14), QPointF(20, 24),
                             QPointF(16, 28), QPointF(12, 28)]))


def _fill_down(p):
    _grid(p, 5, 3, 14, 26, 4, 1, DARK)
    _arrow(p, [(25, 6), (25, 26)], BLUE)
    _arrow(p, [(21, 21), (25, 26), (29, 21)], BLUE)


def _percent(p):
    _text(p, "%", DARK, 22, True)


def _comma(p):
    _text(p, "000", DARK, 12, True, rect=QRectF(0, 4, 32, 20))
    _text(p, ",", DARK, 16, True, rect=QRectF(0, 12, 14, 20))


def _decimals(p, more=True):
    _text(p, ".0" if not more else ".00", DARK, 11, True, rect=QRectF(0, 2, 32, 14))
    _text(p, ".00" if not more else ".000", DARK, 11, True, rect=QRectF(0, 15, 32, 14))
    _arrow(p, [(4, 22), (9, 22)] if more else [(9, 22), (4, 22)], BLUE, 2)


def _export(p, imp=False):
    _sheet(p, 4, 3, 18, 24, 5)
    _text(p, "bas", BLUE, 9, True, rect=QRectF(4, 12, 18, 10))
    if imp:
        _arrow(p, [(30, 18), (20, 18)], GREEN)
        _arrow(p, [(24, 14), (20, 18), (24, 22)], GREEN)
    else:
        _arrow(p, [(18, 18), (30, 18)], GREEN)
        _arrow(p, [(26, 14), (30, 18), (26, 22)], GREEN)


def _module(p, cls=False):
    _sheet(p, 6, 3, 20, 26, 5, fill="#ffffff", pen=BLUE)
    if cls:
        _text(p, "C", BLUE, 14, True, rect=QRectF(6, 9, 20, 18))
    else:
        p.setPen(QPen(BLUE, 1.6))
        for y in (12, 16, 20, 24):
            p.drawLine(QPointF(10, y), QPointF(22, y))


def _remove(p):
    _module(p)
    p.setPen(QPen(RED, 3))
    p.drawLine(QPointF(17, 17), QPointF(29, 29))
    p.drawLine(QPointF(29, 17), QPointF(17, 29))


def _align(kind):
    if kind == "left":
        spec = [(4, 28), (4, 20), (4, 28), (4, 18)]
    elif kind == "center":
        spec = [(4, 28), (8, 24), (4, 28), (9, 23)]
    else:
        spec = [(4, 28), (12, 28), (4, 28), (14, 28)]
    return lambda p: _lines(p, spec)


def _workbook_obj(p):
    _grid(p, 4, 6, 24, 20, 3, 3, GREEN)


def _form(p):
    p.setPen(QPen(DARK, 1.4))
    p.setBrush(QColor("#f3f3f3"))
    p.drawRect(QRectF(3, 5, 26, 22))
    p.setBrush(BLUE)
    p.drawRect(QRectF(3, 5, 26, 5))
    p.setBrush(QColor("#ffffff"))
    p.drawRect(QRectF(7, 14, 18, 4))
    p.drawRect(QRectF(16, 21, 9, 4))


def _folder(p):
    p.setPen(QPen(QColor("#b58b00"), 1.4))
    p.setBrush(QColor("#ffd966"))
    p.drawRect(QRectF(3, 9, 26, 18))
    p.drawRect(QRectF(3, 6, 10, 4))


def _barcode(p, x=4, y=5, w=24, h=18, color=DARK):
    p.setPen(Qt.NoPen)
    p.setBrush(QColor(color))
    pattern = [2, 1, 1, 2, 3, 1, 1, 1, 2, 1, 3, 1, 1, 2, 1]
    xx = x
    on = True
    unit = w / sum(pattern)
    for n in pattern:
        if on:
            p.drawRect(QRectF(xx, y, n * unit, h))
        xx += n * unit
        on = not on


def _barcode_icon(p):
    p.setPen(QPen(DARK, 1))
    p.setBrush(QColor("#FFFFFF"))
    p.drawRect(QRectF(2, 3, 28, 26))
    _barcode(p, 5, 6, 22, 15)
    _text(p, "123", DARK, 7, False, rect=QRectF(2, 21, 28, 8))


def _barcodes_icon(p):
    for i, y in enumerate((2, 12, 22)):
        p.setPen(QPen(QColor("#888888"), 1))
        p.setBrush(QColor("#FFFFFF"))
        p.drawRect(QRectF(2, y, 9, 8))
        _barcode(p, 14, y + 1, 16, 7)


def _picture(p):
    p.setPen(QPen(DARK, 1.2))
    p.setBrush(QColor("#DDEBF7"))
    p.drawRect(QRectF(3, 5, 26, 22))
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#70AD47"))
    p.drawPolygon(QPolygonF([QPointF(4, 26), QPointF(12, 14), QPointF(19, 26)]))
    p.setBrush(QColor("#548235"))
    p.drawPolygon(QPolygonF([QPointF(13, 26), QPointF(20, 17), QPointF(28, 26)]))
    p.setBrush(QColor("#FFC000"))
    p.drawEllipse(QRectF(20, 8, 6, 6))


def _printer(p):
    p.setPen(QPen(DARK, 1.4))
    p.setBrush(QColor("#FFFFFF"))
    p.drawRect(QRectF(9, 3, 14, 9))
    p.setBrush(QColor("#BFBFBF"))
    p.drawRoundedRect(QRectF(3, 11, 26, 11), 2, 2)
    p.setBrush(QColor("#FFFFFF"))
    p.drawRect(QRectF(9, 18, 14, 10))
    p.drawLine(QPointF(11, 22), QPointF(21, 22))
    p.drawLine(QPointF(11, 25), QPointF(19, 25))


def _pdf(p):
    _sheet(p, 6, 3, 20, 26, 6)
    p.setPen(Qt.NoPen)
    p.setBrush(QColor("#C00000"))
    p.drawRect(QRectF(3, 15, 22, 10))
    _text(p, "PDF", "#FFFFFF", 8, True, rect=QRectF(3, 15, 22, 10))


_DRAW: Dict[str, Callable[[QPainter], None]] = {
    "print": _printer,
    "pdf": _pdf,
    "barcode": _barcode_icon,
    "barcodes": _barcodes_icon,
    "picture": _picture,
    "save": lambda p: _draw_save(p, "#444444"),
    "save_white": lambda p: _draw_save(p, "#ffffff"),
    "undo_white": lambda p: _undo(p),
    "redo_white": lambda p: _undo(p, True),
    "undo": lambda p: (_arrow(p, [(8, 14), (20, 14), (26, 20), (20, 26), (12, 26)]),
                       _arrow(p, [(13, 9), (8, 14), (13, 19)])),
    "redo": lambda p: (_arrow(p, [(24, 14), (12, 14), (6, 20), (12, 26), (20, 26)]),
                       _arrow(p, [(19, 9), (24, 14), (19, 19)])),
    "new": _new,
    "open": _open,
    "paste": _clipboard,
    "paste_values": lambda p: (_clipboard(p), _text(p, "123", BLUE, 9, True,
                                                    rect=QRectF(8, 12, 14, 12))),
    "cut": _scissors,
    "copy": _copy,
    "bold": lambda p: _text(p, "B", DARK, 20, True),
    "italic": lambda p: _text(p, "I", DARK, 20, False, True, family="Times New Roman"),
    "underline": lambda p: _text(p, "U", DARK, 19, False, False, True),
    "font_color": lambda p: _paint_a(p, "#ff0000"),
    "fill_color": lambda p: _bucket(p, "#ffff00"),
    "align_left": _align("left"),
    "align_center": _align("center"),
    "align_right": _align("right"),
    "wrap": _wrap,
    "autosum": _sigma,
    "sort_asc": _sort,
    "sort_desc": lambda p: _sort(p, True),
    "find": _find,
    "insert_row": lambda p: _rows_insert(p),
    "delete_row": lambda p: _rows_insert(p, True),
    "insert_col": lambda p: _rows_insert(p, False, True),
    "delete_col": lambda p: _rows_insert(p, True, True),
    "function": _fx,
    "run": _play,
    "stop": _stop,
    "vba": _code,
    "macros": _macros,
    "security": _shield,
    "key": _key,
    "lock": _lock,
    "zoom_in": lambda p: _zoom(p, "+"),
    "zoom_out": lambda p: _zoom(p, "-"),
    "zoom_100": lambda p: _text(p, "100", DARK, 11, True),
    "freeze_row": lambda p: _freeze(p, "row"),
    "freeze_col": lambda p: _freeze(p, "col"),
    "freeze_panes": lambda p: _freeze(p, "panes"),
    "gridlines": _gridlines,
    "table": _table,
    "calc": _calc,
    "names": _names,
    "duplicates": _duplicates,
    "clear": _clear,
    "fill_down": _fill_down,
    "percent": _percent,
    "comma": _comma,
    "dec_more": lambda p: _decimals(p, True),
    "dec_less": lambda p: _decimals(p, False),
    "export": _export,
    "import": lambda p: _export(p, True),
    "module": _module,
    "class_module": lambda p: _module(p, True),
    "remove": _remove,
    "sheet_obj": _workbook_obj,
    "workbook_obj": lambda p: (_grid(p, 4, 6, 24, 20, 3, 3, GREEN),
                               _grid(p, 8, 2, 20, 6, 1, 3, GREEN)),
    "form": _form,
    "folder": _folder,
    "sheet": lambda p: (_sheet(p), _grid(p, 9, 12, 14, 12, 3, 2)),
    "app": lambda p: p.drawPixmap(QRectF(0, 0, 32, 32), app_pixmap(64), QRectF(0, 0, 64, 64)),
    "exit": lambda p: (_arrow(p, [(18, 6), (6, 6), (6, 26), (18, 26)]),
                       _arrow(p, [(12, 16), (28, 16)], RED),
                       _arrow(p, [(23, 11), (28, 16), (23, 21)], RED)),
    "info": lambda p: (p.setPen(Qt.NoPen), p.setBrush(BLUE), p.drawEllipse(QRectF(3, 3, 26, 26)),
                       _text(p, "i", "#ffffff", 18, True)),
    "back": lambda p: (_arrow(p, [(26, 16), (6, 16)], "#ffffff", 2.6),
                       _arrow(p, [(13, 9), (6, 16), (13, 23)], "#ffffff", 2.6)),
}


def names():
    return sorted(_DRAW)


def pixmap(name: str, size: int = 32) -> QPixmap:
    pm = _pix(size)
    fn = _DRAW.get(name)
    if fn is None:
        return pm
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    p.scale(size / 32.0, size / 32.0)
    fn(p)
    p.end()
    return pm


def icon(name: str) -> QIcon:
    ic = _cache.get(name)
    if ic is None:
        ic = QIcon()
        for sz in (16, 20, 24, 32, 48, 64):
            ic.addPixmap(pixmap(name, sz))
        _cache[name] = ic
    return ic


def color_icon(base: str, color: str) -> QIcon:
    """Font colour / fill colour icon with the chosen colour bar."""
    key = f"{base}:{color}"
    ic = _cache.get(key)
    if ic is None:
        ic = QIcon()
        for sz in (16, 24, 32):
            pm = _pix(sz)
            p = QPainter(pm)
            p.setRenderHint(QPainter.Antialiasing)
            p.scale(sz / 32.0, sz / 32.0)
            if base == "font_color":
                _paint_a(p, color)
            else:
                _bucket(p, color)
            p.end()
            ic.addPixmap(pm)
        _cache[key] = ic
    return ic

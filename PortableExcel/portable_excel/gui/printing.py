"""Printing / PDF export of a cell range (paginated, fit to page width)."""
from __future__ import annotations

import logging
import re
from typing import List, Optional, Tuple

from PySide6.QtCore import QMarginsF, QRectF, QSizeF, Qt
from PySide6.QtGui import (QColor, QFont, QFontMetricsF, QPageLayout, QPageSize, QPainter,
                           QPdfWriter, QPen, QPixmap)

from ..model import XLError, col_to_name
from . import barcode_font
from .images import image_pixmap

log = logging.getLogger("PortableExcel")

DEFAULT_COL_W = 64
DEFAULT_ROW_H = 20
HEADER_H = 18     # page header line (sheet name / page numbers), in 96-dpi px


class RangeRenderer:
    """Lays out sheet[r1..r2, c1..c2] on pages; units are 96-dpi px."""

    def __init__(self, sheet, engine, r1, c1, r2, c2, fit_width=True, gridlines=True,
                 title: Optional[str] = None):
        self.sheet, self.engine = sheet, engine
        rows, cols = sheet.used_bounds()
        for im in getattr(sheet, "images", []):
            rows = max(rows, int(im.get("row", 0)) + 1)
            cols = max(cols, int(im.get("col", 0)) + 1)
        # clamp whole-column / whole-sheet selections to the used area
        self.r1, self.c1 = r1, c1
        self.r2 = max(r1, min(r2, rows - 1 if rows else r1))
        self.c2 = max(c1, min(c2, cols - 1 if cols else c1))
        self.fit_width = fit_width
        self.gridlines = gridlines
        self.title = title or sheet.name
        self.col_w = [self._cw(c) for c in range(self.c1, self.c2 + 1)]
        self.row_h = [self._rh(r) for r in range(self.r1, self.r2 + 1)]

    def _cw(self, c):
        w = self.sheet.col_widths.get(c, DEFAULT_COL_W)
        return max(0, w)

    def _rh(self, r):
        h = self.sheet.row_heights.get(r, DEFAULT_ROW_H)
        return max(0, h)

    def total_width(self) -> float:
        return float(sum(self.col_w))

    def paginate(self, page_w: float, page_h: float) -> Tuple[float, List[Tuple[int, int]],
                                                              List[Tuple[int, int]]]:
        """Returns (scale, row_slices, col_slices) with index ranges into the range."""
        scale = 1.0
        if self.fit_width and self.total_width() > 0:
            scale = min(1.0, page_w / self.total_width())
        avail_h = (page_h - HEADER_H) / scale
        avail_w = page_w / scale
        rows = self._slices(self.row_h, avail_h)
        cols = self._slices(self.col_w, avail_w) if not self.fit_width else \
            [(0, len(self.col_w))]
        return scale, rows, cols

    @staticmethod
    def _slices(sizes, avail) -> List[Tuple[int, int]]:
        out = []
        start = 0
        acc = 0.0
        for i, s in enumerate(sizes):
            if acc + s > avail and i > start:
                out.append((start, i))
                start, acc = i, 0.0
            acc += s
        out.append((start, len(sizes)))
        return out

    def page_count(self, page_w, page_h) -> int:
        _s, rows, cols = self.paginate(page_w, page_h)
        return len(rows) * len(cols)

    # -- painting ---------------------------------------------------------------
    def render(self, painter: QPainter, device, page_rect_px: QRectF, new_page) -> int:
        """Paint all pages. new_page() is called between pages. Returns page count."""
        pw, ph = page_rect_px.width(), page_rect_px.height()
        scale, rows, cols = self.paginate(pw, ph)
        n = 0
        total = len(rows) * len(cols)
        for (ra, rb) in rows:
            for (ca, cb) in cols:
                if n:
                    new_page()
                n += 1
                painter.save()
                painter.translate(page_rect_px.left(), page_rect_px.top())
                self._header(painter, pw, n, total)
                painter.translate(0, HEADER_H)
                painter.scale(scale, scale)
                self._page(painter, ra, rb, ca, cb)
                painter.restore()
        return n

    def _header(self, p, pw, n, total):
        f = QFont()
        f.setPixelSize(10)
        p.setFont(f)
        p.setPen(QColor("#666666"))
        p.drawText(QRectF(0, 0, pw, HEADER_H - 4), Qt.AlignLeft | Qt.AlignVCenter, self.title)
        p.drawText(QRectF(0, 0, pw, HEADER_H - 4), Qt.AlignRight | Qt.AlignVCenter,
                   f"Page {n} of {total}")

    def _page(self, p: QPainter, ra, rb, ca, cb):
        xs = [0.0]
        for i in range(ca, cb):
            xs.append(xs[-1] + self.col_w[i])
        ys = [0.0]
        for j in range(ra, rb):
            ys.append(ys[-1] + self.row_h[j])
        sheet = self.sheet
        # cells
        for j in range(ra, rb):
            r = self.r1 + j
            if self.row_h[j] <= 0:
                continue
            for i in range(ca, cb):
                c = self.c1 + i
                if self.col_w[i] <= 0:
                    continue
                rect = QRectF(xs[i - ca], ys[j - ra], self.col_w[i], self.row_h[j])
                st = sheet.styles.get((r, c)) or {}
                if st.get("fill"):
                    p.fillRect(rect, QColor(st["fill"]))
                if self.gridlines:
                    p.setPen(QPen(QColor("#D9D9D9"), 0))
                    p.drawRect(rect)
                raw = sheet.cells.get((r, c))
                if raw is None:
                    continue
                try:
                    text = self.engine.display_text(sheet, r, c)
                except Exception:
                    text = "#ERROR!"
                if text:
                    self._cell_text(p, rect, r, c, text, st)
        # pictures anchored on this page
        top_row, left_col = self.r1 + ra, self.c1 + ca
        for im in getattr(sheet, "images", []):
            row, col = int(im.get("row", 0)), int(im.get("col", 0))
            if not (self.r1 + ra <= row < self.r1 + rb and self.c1 + ca <= col < self.c1 + cb):
                continue
            x = xs[col - left_col] + float(im.get("dx", 0))
            y = ys[row - top_row] + float(im.get("dy", 0))
            pm = image_pixmap(im)
            if not pm.isNull():
                p.drawPixmap(QRectF(x, y, float(im.get("width", pm.width())),
                                    float(im.get("height", pm.height()))).toRect(), pm)

    def _cell_text(self, p, rect, r, c, text, st):
        fname = st.get("font_name")
        if fname and barcode_font.needs_emulation(fname):
            if barcode_font.paint(p, rect, text, fname, st.get("color")):
                return
        f = QFont(fname or "Calibri")
        try:
            size_pt = float(st.get("font_size") or 11)
        except (TypeError, ValueError):
            size_pt = 11.0
        # pixel size keeps the layout identical on any device resolution
        f.setPixelSize(max(1, round(size_pt * 96 / 72)))
        f.setBold(bool(st.get("bold")))
        f.setItalic(bool(st.get("italic")))
        f.setUnderline(bool(st.get("underline")))
        p.setFont(f)
        p.setPen(QColor(st.get("color") or "#000000"))
        align = st.get("align")
        try:
            val = self.engine.value(self.sheet, r, c)
        except Exception:
            val = None
        if align == "center":
            h = Qt.AlignHCenter
        elif align == "right" or (not align and isinstance(val, (int, float))
                                  and not isinstance(val, bool)):
            h = Qt.AlignRight
        elif not align and isinstance(val, (bool, XLError)):
            h = Qt.AlignHCenter
        else:
            h = Qt.AlignLeft
        flags = h | Qt.AlignVCenter
        if st.get("wrap"):
            flags |= Qt.TextWordWrap
        p.save()
        p.setClipRect(rect)
        p.drawText(rect.adjusted(3, 0, -3, 0), flags, text)
        p.restore()


def _page_rect_px(device, layout_rect_pts=None) -> QRectF:
    # we paint in 96-dpi px: painter is scaled by device dpi / 96
    w = device.width() * 96.0 / device.logicalDpiX()
    h = device.height() * 96.0 / device.logicalDpiY()
    return QRectF(0, 0, w, h)


def render_to_device(renderer: RangeRenderer, device) -> int:
    p = QPainter(device)
    if not p.isActive():
        raise RuntimeError("Cannot paint on the output device")
    try:
        s = device.logicalDpiX() / 96.0
        p.scale(s, s)
        rect = _page_rect_px(device)
        return renderer.render(p, device, rect, device.newPage)
    finally:
        p.end()


def export_pdf(renderer: RangeRenderer, path: str, landscape: bool = False) -> int:
    w = QPdfWriter(path)
    w.setResolution(300)
    w.setPageSize(QPageSize(QPageSize.A4))
    w.setPageMargins(QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter)
    w.setPageOrientation(QPageLayout.Landscape if landscape else QPageLayout.Portrait)
    w.setTitle(renderer.title)
    w.setCreator("PortableExcel")
    return render_to_device(renderer, w)


def make_printer():
    from PySide6.QtPrintSupport import QPrinter
    pr = QPrinter(QPrinter.HighResolution)
    pr.setPageSize(QPageSize(QPageSize.A4))
    pr.setPageMargins(QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter)
    return pr


def excel_filter_to_qt(flt: Optional[str]) -> str:
    """'PDF Files (*.pdf), *.pdf, Text (*.txt), *.txt' -> 'PDF Files (*.pdf);;Text (*.txt)'."""
    if not flt:
        return "All files (*)"
    parts = [x.strip() for x in str(flt).split(",")]
    out = []
    for i in range(0, len(parts) - 1, 2):
        desc, pats = parts[i], parts[i + 1]
        pats = " ".join(x.strip() for x in pats.split(";") if x.strip())
        desc_clean = re.sub(r"\s*\([^)]*\)\s*$", "", desc) or desc
        out.append(f"{desc_clean} ({pats})")
    if len(parts) % 2 == 1 and parts[-1]:
        if "(" in parts[-1]:
            out.append(parts[-1])
        else:
            out.append(f"{parts[-1]} ({parts[-1]})")
    return ";;".join(out) or "All files (*)"

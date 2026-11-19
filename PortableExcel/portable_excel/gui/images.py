"""Floating pictures and barcodes over the grid (Sheet.images).

ImageController (one per main window) paints the images on the grid views,
handles selection / move / resize / context menu, and owns the undo commands.
BarcodeDialog and BatchBarcodeDialog create barcodes through
portable_excel.barcodes.
"""
from __future__ import annotations

import copy
import importlib
import logging
import os
import uuid
from typing import Dict, List, Optional, Tuple

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPoint, QRect, QSize, Qt, QTimer
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPen, QPixmap, QUndoCommand
from PySide6.QtWidgets import (QApplication, QButtonGroup, QCheckBox, QComboBox, QDialog,
                               QDialogButtonBox, QFileDialog, QFormLayout, QGroupBox,
                               QHBoxLayout, QLabel, QLineEdit, QMenu, QMessageBox,
                               QProgressDialog, QRadioButton, QScrollArea, QSpinBox,
                               QVBoxLayout, QWidget)

from ..model import a1, col_to_name, name_to_col, parse_a1, parse_range_a1

log = logging.getLogger("PortableExcel")

HANDLE = 7
PICTURE_FILTER = "Pictures (*.png *.jpg *.jpeg *.bmp *.gif);;All files (*)"
EAN_LIKE = ("ean13", "ean8", "upca")


def barcodes_module():
    """portable_excel.barcodes, or None when it is missing/broken."""
    try:
        mod = importlib.import_module("portable_excel.barcodes")
        return mod if hasattr(mod, "make_barcode_image") else None
    except Exception:
        log.exception("barcodes module could not be imported")
        return None


def symbologies() -> List[Tuple[str, str]]:
    mod = barcodes_module()
    if mod is None:
        return []
    out = []
    for key, val in getattr(mod, "SYMBOLOGIES", {}).items():
        if isinstance(val, dict):
            label = val.get("name") or val.get("label") or key
        elif isinstance(val, str):
            label = val
        else:
            label = getattr(val, "name", None) or key
        out.append((key, str(label)))
    return out


def png_bytes(img: QImage) -> bytes:
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    buf.close()
    return bytes(ba.data())


_pix_cache: Dict[Tuple[str, int], QPixmap] = {}


def image_pixmap(img: dict) -> QPixmap:
    data = img.get("data") or b""
    key = (img.get("id", ""), hash(data))
    pm = _pix_cache.get(key)
    if pm is None:
        pm = QPixmap()
        if data:
            pm.loadFromData(bytes(data))
        if len(_pix_cache) > 400:
            _pix_cache.clear()
        _pix_cache[key] = pm
    return pm


def unique_image_name(book, base: str) -> str:
    used = {str(im.get("name", "")).lower() for sh in book.sheets for im in sh.images}
    i = 1
    while f"{base} {i}".lower() in used:
        i += 1
    return f"{base} {i}"


def new_image_id(book) -> str:
    used = {im.get("id") for sh in book.sheets for im in sh.images}
    while True:
        i = uuid.uuid4().hex[:12]
        if i not in used:
            return i


class ImagesCommand(QUndoCommand):
    """Undo step for image-only changes on one sheet (list snapshots)."""

    def __init__(self, ctl, sheet, before, after, text, select=None):
        super().__init__(text)
        self.ctl, self.sheet, self.before, self.after = ctl, sheet, before, after
        self.select = select
        self._first = True

    def redo(self):
        if self._first:
            self._first = False
            return
        self.sheet.images[:] = copy.deepcopy(self.after)
        self.ctl.after_undo(self.sheet, self.select)

    def undo(self):
        self.sheet.images[:] = copy.deepcopy(self.before)
        self.ctl.after_undo(self.sheet, None)


# --------------------------------------------------------------------------
# Controller
# --------------------------------------------------------------------------
class ImageController:
    def __init__(self, win):
        self.win = win
        self.selected: Optional[str] = None   # image id
        self.drag = None
        self.clip: Optional[dict] = None

    # helpers -----------------------------------------------------------------
    @property
    def sheet(self):
        return self.win.book.active

    def image(self, image_id) -> Optional[dict]:
        for im in self.sheet.images:
            if im.get("id") == image_id:
                return im
        return None

    def selected_image(self) -> Optional[dict]:
        return self.image(self.selected) if self.selected else None

    def rect(self, view, img: dict, geom=None) -> QRect:
        z = self.win.zoom
        g = geom or img
        row = int(g.get("row", 0))
        col = int(g.get("col", 0))
        model = view.model()
        if row >= model.rowCount():
            model.ensure_rows(row + 100)
        if col >= model.columnCount():
            model.ensure_cols(col + 26)
        x = view.columnViewportPosition(col) + round(float(g.get("dx", 0)) * z)
        y = view.rowViewportPosition(row) + round(float(g.get("dy", 0)) * z)
        return QRect(x, y, max(1, round(float(g.get("width", 1)) * z)),
                     max(1, round(float(g.get("height", 1)) * z)))

    def handles(self, r: QRect) -> Dict[str, QRect]:
        s = HANDLE
        xs = {"l": r.left() - s // 2, "c": r.center().x() - s // 2, "r": r.right() - s // 2}
        ys = {"t": r.top() - s // 2, "m": r.center().y() - s // 2, "b": r.bottom() - s // 2}
        return {v + h: QRect(xs[h], ys[v], s, s) for v in ys for h in xs if v + h != "mc"}

    # painting ------------------------------------------------------------------
    def paint(self, view, p: QPainter):
        if not self.sheet.images or view.model().sheet is not self.sheet:
            return
        vr = view.viewport().rect()
        p.save()
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        for img in self.sheet.images:
            geom = None
            if self.drag and self.drag.get("id") == img.get("id") and self.drag.get("geom"):
                geom = self.drag["geom"]
            r = self.rect(view, img, geom)
            if not r.intersects(vr):
                continue
            pm = image_pixmap(img)
            if pm.isNull():
                p.fillRect(r, QColor("#F2F2F2"))
                p.setPen(QColor("#A0A0A0"))
                p.drawRect(r.adjusted(0, 0, -1, -1))
                p.drawText(r, Qt.AlignCenter, img.get("name", "Picture"))
            else:
                p.drawPixmap(r, pm)
            if img.get("id") == self.selected:
                p.setPen(QPen(QColor("#5B5B5B"), 1))
                p.setBrush(Qt.NoBrush)
                p.drawRect(r.adjusted(0, 0, -1, -1))
                for h in self.handles(r).values():
                    p.setBrush(QColor("#FFFFFF"))
                    p.setPen(QPen(QColor("#217346"), 1))
                    p.drawEllipse(h)
        p.restore()

    # hit testing / mouse --------------------------------------------------------
    def hit(self, view, pos) -> Tuple[Optional[dict], Optional[str]]:
        sel = self.selected_image()
        if sel is not None:
            r = self.rect(view, sel)
            for k, h in self.handles(r).items():
                if h.adjusted(-3, -3, 3, 3).contains(pos):
                    return sel, k
        for img in reversed(self.sheet.images):
            if self.rect(view, img).contains(pos):
                return img, None
        return None, None

    def select(self, image_id: Optional[str]):
        if image_id == self.selected:
            return
        self.selected = image_id
        self.win.view.refresh()
        self.win.on_image_selection()

    def mouse_press(self, view, pos, e) -> bool:
        img, handle = self.hit(view, pos)
        if img is None:
            if self.selected:
                self.select(None)
            return False
        self.select(img["id"])
        if e.button() == Qt.RightButton:
            self.context_menu(view.viewport().mapToGlobal(pos))
            return True
        if e.button() != Qt.LeftButton:
            return True
        r = self.rect(view, img)
        self.drag = {"id": img["id"], "start": QPoint(pos), "handle": handle, "rect": r,
                     "geom": None}
        return True

    def mouse_move(self, view, pos, e) -> bool:
        d = self.drag
        if d is None:
            if not e.buttons():
                img, handle = self.hit(view, pos)
                if img is not None:
                    cur = Qt.SizeAllCursor
                    if handle:
                        cur = {"tl": Qt.SizeFDiagCursor, "br": Qt.SizeFDiagCursor,
                               "tr": Qt.SizeBDiagCursor, "bl": Qt.SizeBDiagCursor,
                               "tc": Qt.SizeVerCursor, "bc": Qt.SizeVerCursor,
                               "ml": Qt.SizeHorCursor, "mr": Qt.SizeHorCursor}[handle]
                    view.viewport().setCursor(cur)
                    return True
            return False
        img = self.image(d["id"])
        if img is None:
            self.drag = None
            return False
        r0: QRect = d["rect"]
        dx, dy = pos.x() - d["start"].x(), pos.y() - d["start"].y()
        if d["handle"] is None:
            r = r0.translated(dx, dy)
        else:
            h = d["handle"]
            left, top, right, bottom = r0.left(), r0.top(), r0.right(), r0.bottom()
            if "l" in h[1]:
                left = min(left + dx, right - 4)
            if "r" in h[1]:
                right = max(right + dx, left + 4)
            if h[0] == "t":
                top = min(top + dy, bottom - 4)
            if h[0] == "b":
                bottom = max(bottom + dy, top + 4)
            r = QRect(QPoint(left, top), QPoint(right, bottom))
            if e.modifiers() & Qt.ShiftModifier and r0.height() > 0:
                ratio = r0.width() / r0.height()
                if h in ("tc", "bc"):
                    r.setWidth(max(4, round(r.height() * ratio)))
                else:
                    nh = max(4, round(r.width() / ratio))
                    if h[0] == "t":
                        r.setTop(r.bottom() - nh + 1)
                    else:
                        r.setHeight(nh)
        d["geom"] = self.geom_from_rect(view, r)
        view.refresh()
        return True

    def mouse_release(self, view, pos, e) -> bool:
        d, self.drag = self.drag, None
        if d is None:
            return False
        geom = d.get("geom")
        if geom:
            self.update_image(d["id"], geom,
                              "Move Picture" if d["handle"] is None else "Size Picture")
        view.refresh()
        return True

    def mouse_double(self, view, pos) -> bool:
        img, _h = self.hit(view, pos)
        if img is None:
            return False
        self.select(img["id"])
        if img.get("barcode"):
            self.win.barcode_dialog(edit=img)
        return True

    def geom_from_rect(self, view, r: QRect) -> dict:
        """Anchor a view rectangle to the cell under its top-left corner."""
        z = self.win.zoom
        hh, vh = view.horizontalHeader(), view.verticalHeader()
        ax = max(0, r.left() + hh.offset())   # absolute content coordinates
        ay = max(0, r.top() + vh.offset())
        col = hh.logicalIndexAt(ax - hh.offset()) if ax - hh.offset() >= 0 else -1
        row = vh.logicalIndexAt(ay - vh.offset()) if ay - vh.offset() >= 0 else -1
        if col < 0:
            col = 0
            while col + 1 < view.model().columnCount() and hh.sectionPosition(col + 1) <= ax:
                col += 1
        if row < 0:
            row = 0
            while row + 1 < view.model().rowCount() and vh.sectionPosition(row + 1) <= ay:
                row += 1
        dx = (ax - hh.sectionPosition(col)) / z
        dy = (ay - vh.sectionPosition(row)) / z
        return {"row": row, "col": col, "dx": round(dx), "dy": round(dy),
                "width": max(4, round(r.width() / z)), "height": max(4, round(r.height() / z))}

    # undoable changes ---------------------------------------------------------------
    def change(self, text: str, fn, select=None):
        sheet = self.sheet
        before = copy.deepcopy(sheet.images)
        fn(sheet)
        after = copy.deepcopy(sheet.images)
        if before == after:
            return False
        self.win.undo_stack.push(ImagesCommand(self, sheet, before, after, text, select))
        self.after_undo(sheet, select)
        return True

    def after_undo(self, sheet, select):
        if sheet is not self.win.book.active:
            self.win.activate_sheet(sheet)
        if select is not None:
            self.selected = select
        if self.selected and not self.image(self.selected):
            self.selected = None
        self.win.mark_dirty()
        self.win.view.refresh()
        self.win.on_image_selection()

    def update_image(self, image_id, geom: dict, text="Move Picture"):
        def fn(sheet):
            for im in sheet.images:
                if im.get("id") == image_id:
                    im.update(geom)
        return self.change(text, fn, select=image_id)

    def add_image(self, img: dict, text="Insert Picture") -> dict:
        book = self.win.book
        img = dict(img)
        if not img.get("id") or any(im.get("id") == img["id"] for sh in book.sheets
                                    for im in sh.images):
            img["id"] = new_image_id(book)
        base = "Barcode" if img.get("barcode") else "Picture"
        name = str(img.get("name") or "")
        if not name or name in (base,) or any(im.get("name") == name for sh in book.sheets
                                              for im in sh.images):
            img["name"] = unique_image_name(book, base)
        img.setdefault("fmt", "png")
        img.setdefault("dx", 0)
        img.setdefault("dy", 0)
        self.change(text, lambda sheet: sheet.images.append(img), select=img["id"])
        return img

    def delete_selected(self) -> bool:
        if not self.selected:
            return False
        iid = self.selected
        self.selected = None
        self.change("Delete Picture",
                    lambda sheet: sheet.images.__setitem__(
                        slice(None), [im for im in sheet.images if im.get("id") != iid]))
        return True

    def bring(self, front=True):
        iid = self.selected

        def fn(sheet):
            ims = [im for im in sheet.images if im.get("id") != iid]
            me = [im for im in sheet.images if im.get("id") == iid]
            sheet.images[:] = ims + me if front else me + ims
        self.change("Bring to Front" if front else "Send to Back", fn, select=iid)

    def reset_size(self):
        img = self.selected_image()
        if img is None:
            return
        w = h = None
        bc = img.get("barcode")
        mod = barcodes_module()
        if bc and mod is not None:
            try:
                text = self.win.barcode_text(self.sheet, bc)
                _png, w, h = mod.render_png(bc["symbology"], text, bc.get("module_px", 2),
                                            bc.get("bar_height_px", 60), bc.get("show_text", True))
            except Exception:
                w = h = None
        if w is None:
            pm = image_pixmap(img)
            if pm.isNull():
                return
            w, h = pm.width(), pm.height()
        self.update_image(img["id"], {"width": w, "height": h}, "Reset Picture Size")

    # clipboard ---------------------------------------------------------------------
    def copy_selected(self) -> bool:
        img = self.selected_image()
        if img is None:
            return False
        self.clip = copy.deepcopy(img)
        self.win.last_copy_kind = "image"
        qi = QImage()
        qi.loadFromData(bytes(img.get("data") or b""))
        if not qi.isNull():
            QGuiApplication.clipboard().setImage(qi)
        return True

    def paste(self) -> bool:
        if self.clip is None or self.win.last_copy_kind != "image":
            md = QGuiApplication.clipboard().mimeData()
            if md is not None and md.hasImage() and not md.hasText():
                qi = QGuiApplication.clipboard().image()
                if not qi.isNull():
                    r, c = self.win.view.current_cell()
                    self.add_image({"row": r, "col": c, "dx": 0, "dy": 0,
                                    "width": qi.width(), "height": qi.height(),
                                    "data": png_bytes(qi), "fmt": "png", "barcode": None,
                                    "name": ""}, "Paste Picture")
                    return True
            return False
        img = copy.deepcopy(self.clip)
        r, c = self.win.view.current_cell()
        img.update(row=r, col=c, dx=2, dy=2, id="", name="")
        self.add_image(img, "Paste Picture")
        return True

    # menu / files --------------------------------------------------------------
    def context_menu(self, global_pos):
        img = self.selected_image()
        if img is None:
            return None
        m = QMenu(self.win)
        if img.get("barcode"):
            m.addAction("Edit Barcode...", lambda: self.win.barcode_dialog(edit=img))
            m.addSeparator()
        m.addAction("Copy", self.copy_selected)
        m.addAction("Delete", self.delete_selected)
        m.addSeparator()
        m.addAction("Bring to Front", lambda: self.bring(True))
        m.addAction("Send to Back", lambda: self.bring(False))
        m.addAction("Reset Size", self.reset_size)
        m.addSeparator()
        m.addAction("Save Picture As PNG...", self.save_picture)
        self.win._last_menu = m
        m.popup(global_pos)
        return m

    def save_picture(self, path: Optional[str] = None) -> Optional[str]:
        img = self.selected_image()
        if img is None:
            return None
        if not path:
            path, _ = QFileDialog.getSaveFileName(self.win, "Save Picture As",
                                                  f"{img.get('name', 'Picture')}.png",
                                                  "PNG image (*.png)")
            if not path:
                return None
        if not path.lower().endswith(".png"):
            path += ".png"
        data = bytes(img.get("data") or b"")
        if img.get("fmt", "png") != "png" or not data.startswith(b"\x89PNG"):
            qi = QImage()
            qi.loadFromData(data)
            data = png_bytes(qi)
        with open(path, "wb") as fh:
            fh.write(data)
        self.win.set_status_message(f"Saved {os.path.basename(path)}")
        return path

    def insert_picture_file(self, path: Optional[str] = None) -> Optional[dict]:
        if not path:
            path, _ = QFileDialog.getOpenFileName(self.win, "Insert Picture", "", PICTURE_FILTER)
            if not path:
                return None
        qi = QImage(path)
        if qi.isNull():
            QMessageBox.warning(self.win, "PortableExcel", f"Could not read the picture:\n{path}")
            return None
        w, h = qi.width(), qi.height()
        if w > 640:
            h = round(h * 640 / w)
            w = 640
        r, c = self.win.view.current_cell()
        return self.add_image({"row": r, "col": c, "dx": 0, "dy": 0, "width": w, "height": h,
                               "data": png_bytes(qi), "fmt": "png", "barcode": None,
                               "name": ""}, "Insert Picture")


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------
class BarcodeDialog(QDialog):
    """Insert or edit one barcode, with a live preview."""

    def __init__(self, win, edit: Optional[dict] = None):
        super().__init__(win)
        self.win = win
        self.edit_img = edit
        self.setWindowTitle("Edit Barcode" if edit else "Insert Barcode")
        mod = barcodes_module()
        self.mod = mod
        lay = QVBoxLayout(self)
        form = QFormLayout()
        self.sym = QComboBox()
        for key, label in symbologies():
            self.sym.addItem(label, key)
        form.addRow("Symbology:", self.sym)
        lay.addLayout(form)
        src = QGroupBox("Data")
        g = QFormLayout(src)
        self.linked = QRadioButton("Linked to cell:")
        self.fixed = QRadioButton("Fixed text:")
        grp = QButtonGroup(self)
        grp.addButton(self.linked)
        grp.addButton(self.fixed)
        r, c = win.view.current_cell()
        self.source = QLineEdit(a1(r, c))
        self.source.setPlaceholderText("e.g. A2")
        self.text = QLineEdit()
        g.addRow(self.linked, self.source)
        g.addRow(self.fixed, self.text)
        self.value_label = QLabel("")
        self.value_label.setStyleSheet("color:#555;")
        g.addRow("Value:", self.value_label)
        lay.addWidget(src)
        opts = QGroupBox("Options")
        o = QFormLayout(opts)
        self.show_text = QCheckBox("Show human-readable text")
        self.show_text.setChecked(True)
        self.module_px = QSpinBox()
        self.module_px.setRange(1, 6)
        self.module_px.setValue(2)
        self.module_px.setSuffix(" px")
        self.bar_height = QSpinBox()
        self.bar_height.setRange(10, 400)
        self.bar_height.setValue(60)
        self.bar_height.setSuffix(" px")
        self.target = QLineEdit(a1(r, c + 1))
        o.addRow("", self.show_text)
        o.addRow("Bar width (module):", self.module_px)
        o.addRow("Bar height:", self.bar_height)
        o.addRow("Place at cell:", self.target)
        lay.addWidget(opts)
        pv = QGroupBox("Preview")
        pvl = QVBoxLayout(pv)
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumSize(320, 130)
        self.preview.setStyleSheet("background:white; border:1px solid #D4D4D4;")
        scroll = QScrollArea()
        scroll.setWidget(self.preview)
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(150)
        pvl.addWidget(scroll)
        self.error = QLabel("")
        self.error.setStyleSheet("color:#C00000;")
        self.error.setWordWrap(True)
        pvl.addWidget(self.error)
        lay.addWidget(pv)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        # defaults / edit mode
        self.linked.setChecked(True)
        default_sym = "code128"
        if edit and edit.get("barcode"):
            bc = edit["barcode"]
            default_sym = bc.get("symbology", default_sym)
            if bc.get("source"):
                self.source.setText(bc["source"])
            else:
                self.fixed.setChecked(True)
                self.text.setText(bc.get("text") or "")
            self.show_text.setChecked(bool(bc.get("show_text", True)))
            self.module_px.setValue(int(bc.get("module_px", 2)))
            self.bar_height.setValue(int(bc.get("bar_height_px", 60)))
            self.target.setText(a1(int(edit.get("row", 0)), int(edit.get("col", 0))))
        i = self.sym.findData(default_sym)
        self.sym.setCurrentIndex(max(0, i))
        for sig in (self.sym.currentIndexChanged, self.linked.toggled, self.source.textChanged,
                    self.text.textChanged, self.show_text.toggled, self.module_px.valueChanged,
                    self.bar_height.valueChanged):
            sig.connect(self.update_preview)
        self.source.textEdited.connect(self._source_edited)
        self._target_touched = bool(edit)
        self.target.textEdited.connect(lambda *_: setattr(self, "_target_touched", True))
        self.update_preview()
        self.resize(460, 560)

    def _source_edited(self, text):
        if self._target_touched:
            return
        rc = parse_a1(text)
        if rc:
            self.target.setText(a1(rc[0], rc[1] + 1))

    def symbology(self) -> str:
        return self.sym.currentData() or "code128"

    def data_text(self) -> Optional[str]:
        if self.linked.isChecked():
            rc = parse_a1(self.source.text())
            if rc is None:
                return None
            return self.win.engine.display_text(self.win.book.active, *rc)
        return self.text.text()

    def update_preview(self, *_):
        self.source.setEnabled(self.linked.isChecked())
        self.text.setEnabled(self.fixed.isChecked())
        ok_btn = self.buttons.button(QDialogButtonBox.Ok)
        err = None
        text = self.data_text()
        sym = self.symbology()
        if self.mod is None:
            err = "Barcode support is not available."
        elif text is None:
            err = "Enter a valid cell reference, e.g. A2."
        else:
            try:
                err = self.mod.validate(sym, text)
            except Exception as e:
                err = str(e)
        shown = text or ""
        if not err and sym in EAN_LIKE:
            try:
                norm = self.mod.normalize(sym, text)
                shown = f"{text}  →  {norm} (with check digit)" if norm != text else norm
            except Exception:
                pass
        self.value_label.setText(shown)
        if not err:
            try:
                png, w, h = self.mod.render_png(sym, text, self.module_px.value(),
                                                self.bar_height.value(),
                                                self.show_text.isChecked())
                pm = QPixmap()
                pm.loadFromData(png)
                self.preview.setPixmap(pm)
            except Exception as e:
                err = str(e)
        if err:
            self.preview.setPixmap(QPixmap())
            self.preview.setText("")
        self.error.setText(err or "")
        if parse_a1(self.target.text()) is None and not err:
            err = "Enter the cell where the barcode is placed."
            self.error.setText(err)
        ok_btn.setEnabled(not err)
        return err

    def values(self) -> dict:
        rc = parse_a1(self.target.text()) or (0, 0)
        return {"symbology": self.symbology(),
                "source": self.source.text().strip().upper().replace("$", "")
                if self.linked.isChecked() else None,
                "text": "" if self.linked.isChecked() else self.text.text(),
                "show_text": self.show_text.isChecked(), "module_px": self.module_px.value(),
                "bar_height_px": self.bar_height.value(), "row": rc[0], "col": rc[1]}


class BatchBarcodeDialog(QDialog):
    def __init__(self, win, rng: Tuple[int, int, int, int]):
        super().__init__(win)
        self.win = win
        self.setWindowTitle("Generate Barcodes")
        r1, c1, r2, c2 = rng
        lay = QVBoxLayout(self)
        intro = QLabel("Creates one barcode for every non-empty cell in the range. "
                       "Barcodes stay linked to their cell and update when it changes.")
        intro.setWordWrap(True)
        lay.addWidget(intro)
        form = QFormLayout()
        self.range_edit = QLineEdit(f"{a1(r1, c1)}:{a1(r2, c2)}")
        self.sym = QComboBox()
        for key, label in symbologies():
            self.sym.addItem(label, key)
        self.out_col = QLineEdit(col_to_name(self.default_output_col(rng)))
        self.out_col.setMaxLength(3)
        self.show_text = QCheckBox("Show human-readable text")
        self.show_text.setChecked(True)
        self.module_px = QSpinBox()
        self.module_px.setRange(1, 6)
        self.module_px.setValue(1)
        self.module_px.setSuffix(" px")
        self.bar_height = QSpinBox()
        self.bar_height.setRange(10, 400)
        self.bar_height.setValue(40)
        self.bar_height.setSuffix(" px")
        self.fit = QCheckBox("Fit column width and row heights to the barcodes")
        self.fit.setChecked(True)
        self.linked = QCheckBox("Link barcodes to their cells (auto-update)")
        self.linked.setChecked(True)
        form.addRow("Data range:", self.range_edit)
        form.addRow("Symbology:", self.sym)
        form.addRow("Output column:", self.out_col)
        form.addRow("Bar width (module):", self.module_px)
        form.addRow("Bar height:", self.bar_height)
        form.addRow("", self.show_text)
        form.addRow("", self.fit)
        form.addRow("", self.linked)
        lay.addLayout(form)
        self.info = QLabel("")
        self.info.setWordWrap(True)
        lay.addWidget(self.info)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.button(QDialogButtonBox.Ok).setText("Generate")
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        lay.addWidget(self.buttons)
        for sig in (self.range_edit.textChanged, self.sym.currentIndexChanged,
                    self.out_col.textChanged):
            sig.connect(self.check)
        self.check()
        self.resize(440, 380)

    def default_output_col(self, rng) -> int:
        r1, c1, r2, c2 = rng
        cells = self.win.book.active.cells
        c = c2 + 1
        while any((r, c) in cells for r in range(r1, min(r2, r1 + 5000) + 1)):
            c += 1
        return c

    def check(self, *_):
        ok = True
        rng = parse_range_a1(self.range_edit.text())
        out = self.out_col.text().strip().upper()
        if rng is None:
            self.info.setText("<span style='color:#C00000'>Invalid range.</span>")
            ok = False
        elif not out.isalpha():
            self.info.setText("<span style='color:#C00000'>Invalid output column.</span>")
            ok = False
        else:
            r1, c1, r2, c2 = self.win._clamp(*rng)
            sheet = self.win.book.active
            vals = [(r, c) for r, c, _raw in sheet.iter_cells_in(r1, c1, r2, c2)]
            mod = barcodes_module()
            bad = 0
            if mod is not None:
                for (r, c) in vals[:2000]:
                    try:
                        if mod.validate(self.sym.currentData(),
                                        self.win.engine.display_text(sheet, r, c)):
                            bad += 1
                    except Exception:
                        bad += 1
            msg = f"{len(vals)} non-empty cell(s) → {len(vals) - bad} barcode(s)"
            if bad:
                msg += f"; <span style='color:#C00000'>{bad} invalid value(s) will be skipped</span>"
            self.info.setText(msg)
            ok = len(vals) - bad > 0
        self.buttons.button(QDialogButtonBox.Ok).setEnabled(ok)

    def values(self) -> dict:
        return {"range": parse_range_a1(self.range_edit.text()),
                "symbology": self.sym.currentData(),
                "out_col": name_to_col(self.out_col.text().strip().upper()),
                "show_text": self.show_text.isChecked(), "module_px": self.module_px.value(),
                "bar_height_px": self.bar_height.value(), "fit": self.fit.isChecked(),
                "linked": self.linked.isChecked()}

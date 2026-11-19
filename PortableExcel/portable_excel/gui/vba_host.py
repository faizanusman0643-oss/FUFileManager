"""GUI implementation of the VBA host interface (MsgBox, InputBox, selection...)."""
from __future__ import annotations

import logging
from typing import Dict, Tuple

from PySide6.QtWidgets import QApplication, QInputDialog, QMessageBox

try:  # the interpreter may still be under construction
    from .. import vba as _vba
    _HostBase = getattr(_vba, "VBAHost", object)
except Exception:  # pragma: no cover
    _vba = None
    _HostBase = object

log = logging.getLogger("PortableExcel")

# VBA constants
VB_OK, VB_CANCEL, VB_ABORT, VB_RETRY, VB_IGNORE, VB_YES, VB_NO = 1, 2, 3, 4, 5, 6, 7

_BUTTON_SETS = {
    0: [(QMessageBox.Ok, VB_OK)],
    1: [(QMessageBox.Ok, VB_OK), (QMessageBox.Cancel, VB_CANCEL)],
    2: [(QMessageBox.Abort, VB_ABORT), (QMessageBox.Retry, VB_RETRY),
        (QMessageBox.Ignore, VB_IGNORE)],
    3: [(QMessageBox.Yes, VB_YES), (QMessageBox.No, VB_NO), (QMessageBox.Cancel, VB_CANCEL)],
    4: [(QMessageBox.Yes, VB_YES), (QMessageBox.No, VB_NO)],
    5: [(QMessageBox.Retry, VB_RETRY), (QMessageBox.Cancel, VB_CANCEL)],
}
_ICONS = {16: QMessageBox.Critical, 32: QMessageBox.Question, 48: QMessageBox.Warning,
          64: QMessageBox.Information}


def build_msgbox(prompt, buttons=0, title="", parent=None) -> Tuple[QMessageBox, Dict]:
    """Create (not show) a QMessageBox for VBA MsgBox; returns (box, {button: code})."""
    try:
        buttons = int(buttons or 0)
    except (TypeError, ValueError):
        buttons = 0
    box = QMessageBox(parent)
    box.setWindowTitle(str(title) if title else "PortableExcel")
    box.setText(str(prompt).replace("\r\n", "\n").replace("\r", "\n"))
    box.setIcon(_ICONS.get(buttons & 0x70, QMessageBox.NoIcon))
    spec = _BUTTON_SETS.get(buttons & 0x0F, _BUTTON_SETS[0])
    mapping = {}
    for std, code in spec:
        b = box.addButton(std)
        mapping[b] = code
    default_index = (buttons & 0xF00) >> 8
    btns = list(mapping)
    if 0 <= default_index < len(btns):
        box.setDefaultButton(btns[default_index])
    cancel_code = VB_CANCEL if any(c == VB_CANCEL for _, c in spec) else (
        VB_OK if len(spec) == 1 else None)
    if cancel_code is not None:
        for b, code in mapping.items():
            if code == cancel_code:
                box.setEscapeButton(b)
    return box, mapping


class GuiVBAHost(_HostBase):
    def __init__(self, win):
        try:
            super().__init__()
        except TypeError:
            pass
        self.win = win

    # dialogs -------------------------------------------------------------
    def msgbox(self, prompt, buttons=0, title=""):
        box, mapping = build_msgbox(prompt, buttons, title, self.win)
        box.exec()
        clicked = box.clickedButton()
        if clicked in mapping:
            return mapping[clicked]
        codes = list(mapping.values())
        return VB_CANCEL if VB_CANCEL in codes else codes[0]

    def inputbox(self, prompt, title="", default=""):
        dlg = QInputDialog(self.win)
        dlg.setWindowTitle(str(title) if title else "PortableExcel")
        dlg.setLabelText(str(prompt))
        dlg.setTextValue("" if default is None else str(default))
        if dlg.exec():
            return dlg.textValue()
        return None

    # output ----------------------------------------------------------------
    def debug_print(self, text):
        self.win.immediate_print(str(text))

    def status_bar(self, text):
        self.win.set_status_message(text)

    # selection / sheets ------------------------------------------------------
    def _sheet(self, sheet):
        book = self.win.book
        if sheet is None:
            return book.active
        if isinstance(sheet, int):
            return book.sheets[sheet]
        if isinstance(sheet, str):
            return book.sheet_by_name(sheet) or book.active
        return sheet

    def select(self, sheet, r1, c1, r2, c2):
        sh = self._sheet(sheet)
        if sh is not self.win.book.active:
            self.win.activate_sheet(sh)
        self.win.view.select_range(r1, c1, r2, c2)

    def get_selection(self):
        r1, c1, r2, c2 = self.win.view.active_range()
        return (self.win.book.active, r1, c1, r2, c2)

    def get_active_cell(self):
        r, c = self.win.view.current_cell()
        return (self.win.book.active, r, c)

    def activate_sheet(self, sheet):
        self.win.activate_sheet(self._sheet(sheet))

    def refresh(self):
        self.win.refresh_from_macro()

    def process_events(self):
        QApplication.processEvents()

    def save_workbook(self):
        return self.win.save()

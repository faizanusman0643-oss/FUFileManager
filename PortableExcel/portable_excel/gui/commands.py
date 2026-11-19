"""Undo/redo commands (QUndoStack, limited to 100 steps)."""
from __future__ import annotations

from PySide6.QtGui import QUndoCommand

from .ops import restore_book

UNDO_LIMIT = 100


class CellsCommand(QUndoCommand):
    """Cell contents and/or styles of one sheet.

    before/after: {(r, c): (raw or None, style dict or None)}
    """

    def __init__(self, win, sheet, before, after, text, select=None):
        super().__init__(text)
        self.win = win
        self.sheet = sheet
        self.before = before
        self.after = after
        self.select = select

    def _apply(self, state):
        sh = self.sheet
        for (r, c), (raw, st) in state.items():
            sh.set_raw(r, c, raw)
            if st:
                sh.styles[(r, c)] = dict(st)
            else:
                sh.styles.pop((r, c), None)
        self.win.cells_changed(sh, list(state.keys()), select=self.select)

    def redo(self):
        self._apply(self.after)

    def undo(self):
        self._apply(self.before)


class SnapshotCommand(QUndoCommand):
    """Whole-book snapshot, used for structural changes (rows/cols/sheets)."""

    def __init__(self, win, before, after, text):
        super().__init__(text)
        self.win = win
        self.before = before
        self.after = after
        self._first = True

    def redo(self):
        if self._first:  # the change was already applied when pushed
            self._first = False
            return
        restore_book(self.win.book, self.after)
        self.win.structure_changed()

    def undo(self):
        restore_book(self.win.book, self.before)
        self.win.structure_changed()

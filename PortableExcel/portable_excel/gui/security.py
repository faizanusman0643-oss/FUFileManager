"""Application password lock.

The password itself is never stored: only a random salt and the
PBKDF2-HMAC-SHA256 hash (200 000 iterations).  The built-in default is kept as
hard-coded hex constants; a password changed by the user is written to
PortableExcel.ini (see config.py).

This lock is a deterrent for casual access to the application; it does not
encrypt workbooks (use "Save with password..." for that).
"""
from __future__ import annotations

import hashlib
import hmac
import os
from typing import Optional, Tuple

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont, QPainter, QPen
from PySide6.QtWidgets import (QCheckBox, QDialog, QDialogButtonBox, QFormLayout,
                               QHBoxLayout, QLabel, QLineEdit, QMessageBox,
                               QPushButton, QVBoxLayout, QWidget)

from . import config
from . import icons

ITERATIONS = 200_000
# salt + hash of the default password, computed once (the password is not stored)
DEFAULT_SALT_HEX = "6e2d1e7d6cfa929c0180d939d0eee2ee"
DEFAULT_HASH_HEX = "47ff97e4df1efa277541ef2b505aef8da3e9105cccd0f5d60f79dd25042febf2"

MAX_FREE_ATTEMPTS = 5
LOCKOUT_SECONDS = 30
MIN_PASSWORD_LEN = 4


def hash_password(password: str, salt: bytes, iterations: int = ITERATIONS) -> bytes:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)


def stored_credentials() -> Tuple[bytes, bytes, int]:
    """(salt, hash, iterations) - from the ini if a password was set, else the default."""
    sec = config.get("app_password")
    if isinstance(sec, dict):
        try:
            salt = bytes.fromhex(sec["salt"])
            digest = bytes.fromhex(sec["hash"])
            it = int(sec.get("iterations", ITERATIONS))
            if salt and len(digest) == 32 and it >= 1000:
                return salt, digest, it
        except (KeyError, ValueError, TypeError):
            pass
    return bytes.fromhex(DEFAULT_SALT_HEX), bytes.fromhex(DEFAULT_HASH_HEX), ITERATIONS


def check_password(password: str) -> bool:
    salt, digest, it = stored_credentials()
    return hmac.compare_digest(hash_password(password or "", salt, it), digest)


def set_password(new_password: str) -> None:
    """Store salt + hash of a new app password in PortableExcel.ini."""
    if len(new_password) < MIN_PASSWORD_LEN:
        raise ValueError(f"The password must have at least {MIN_PASSWORD_LEN} characters.")
    salt = os.urandom(16)
    digest = hash_password(new_password, salt)
    config.put("app_password", {"salt": salt.hex(), "hash": digest.hex(),
                                "iterations": ITERATIONS, "algorithm": "pbkdf2_sha256"},
               persist=False)
    config.save()


def change_password(old: str, new: str, confirm: str) -> Optional[str]:
    """Returns an error message, or None on success."""
    if not check_password(old):
        return "The current password is incorrect."
    if new != confirm:
        return "The new passwords do not match."
    if len(new) < MIN_PASSWORD_LEN:
        return f"The new password must have at least {MIN_PASSWORD_LEN} characters."
    try:
        set_password(new)
    except OSError as e:
        return f"Could not write the settings file:\n{e}"
    return None


def lock_disabled() -> bool:
    return os.environ.get("PORTABLEEXCEL_NO_LOCK", "") == "1"


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------
class _Banner(QWidget):
    """Green header with a padlock drawn by QPainter."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(92)

    def paintEvent(self, _e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#217346"))
        p.drawPixmap(22, 14, icons.app_pixmap(64))
        # padlock
        p.setPen(Qt.NoPen)
        x, y = self.width() - 70, 22
        p.setBrush(QColor("#ffffff"))
        p.drawRoundedRect(x, y + 20, 40, 30, 4, 4)
        p.setPen(QPen(QColor("#ffffff"), 5))
        p.setBrush(Qt.NoBrush)
        p.drawArc(x + 7, y, 26, 40, 0, 180 * 16)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#217346"))
        p.drawEllipse(x + 16, y + 29, 8, 8)
        p.setPen(QColor("#ffffff"))
        f = QFont(self.font())
        f.setPointSizeF(16)
        f.setBold(True)
        p.setFont(f)
        p.drawText(100, 20, self.width() - 190, 34, Qt.AlignLeft | Qt.AlignVCenter, "PortableExcel")
        f.setPointSizeF(10)
        f.setBold(False)
        p.setFont(f)
        p.drawText(100, 50, self.width() - 190, 24, Qt.AlignLeft | Qt.AlignVCenter,
                   "This application is password protected")
        p.end()


class LockDialog(QDialog):
    """'PortableExcel is locked - enter password'."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PortableExcel is locked — enter password")
        self.setWindowIcon(icons.app_icon())
        self.setModal(True)
        self.failed = 0
        self._remaining = 0
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 12)
        lay.addWidget(_Banner(self))
        body = QVBoxLayout()
        body.setContentsMargins(24, 12, 24, 0)
        self.info = QLabel("PortableExcel is locked — enter password")
        f = self.info.font()
        f.setPointSizeF(f.pointSizeF() + 1.5)
        self.info.setFont(f)
        body.addWidget(self.info)
        row = QHBoxLayout()
        self.edit = QLineEdit()
        self.edit.setEchoMode(QLineEdit.Password)
        self.edit.setPlaceholderText("Password")
        self.edit.setMinimumWidth(280)
        self.edit.returnPressed.connect(self.try_unlock)
        row.addWidget(self.edit)
        self.show_box = QCheckBox("Show")
        self.show_box.toggled.connect(
            lambda on: self.edit.setEchoMode(QLineEdit.Normal if on else QLineEdit.Password))
        row.addWidget(self.show_box)
        body.addLayout(row)
        self.msg = QLabel("")
        self.msg.setStyleSheet("color:#c00000;")
        body.addWidget(self.msg)
        btns = QHBoxLayout()
        btns.addStretch(1)
        self.ok_btn = QPushButton("Unlock")
        self.ok_btn.setDefault(True)
        self.ok_btn.setObjectName("primary")
        self.ok_btn.clicked.connect(self.try_unlock)
        self.exit_btn = QPushButton("Exit")
        self.exit_btn.clicked.connect(self.reject)
        btns.addWidget(self.ok_btn)
        btns.addWidget(self.exit_btn)
        body.addLayout(btns)
        lay.addLayout(body)
        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self._tick)
        self.resize(460, 250)

    def try_unlock(self):
        if self._remaining > 0:
            return
        if check_password(self.edit.text()):
            self.edit.clear()
            self.accept()
            return
        self.failed += 1
        self.edit.clear()
        if self.failed >= MAX_FREE_ATTEMPTS:
            self._start_lockout()
        else:
            left = MAX_FREE_ATTEMPTS - self.failed
            self.msg.setText(f"Incorrect password. {left} attempt(s) left before a "
                             f"{LOCKOUT_SECONDS} s delay.")

    def _start_lockout(self):
        self._remaining = LOCKOUT_SECONDS
        self.edit.setEnabled(False)
        self.ok_btn.setEnabled(False)
        self._update_lock_msg()
        self._timer.start()

    def _update_lock_msg(self):
        self.msg.setText(f"Too many wrong attempts. Try again in {self._remaining} s.")

    def _tick(self):
        self._remaining -= 1
        if self._remaining <= 0:
            self._timer.stop()
            self._remaining = 0
            self.edit.setEnabled(True)
            self.ok_btn.setEnabled(True)
            self.msg.setText("Incorrect password.")
            self.edit.setFocus()
        else:
            self._update_lock_msg()


class ChangePasswordDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Change app password")
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Change the password required to start PortableExcel."))
        form = QFormLayout()
        self.old = QLineEdit()
        self.new = QLineEdit()
        self.confirm = QLineEdit()
        for w in (self.old, self.new, self.confirm):
            w.setEchoMode(QLineEdit.Password)
            w.setMinimumWidth(240)
        form.addRow("Current password:", self.old)
        form.addRow("New password:", self.new)
        form.addRow("Confirm new password:", self.confirm)
        lay.addLayout(form)
        self.msg = QLabel("")
        self.msg.setStyleSheet("color:#c00000;")
        lay.addWidget(self.msg)
        note = QLabel(f"Stored as a salted PBKDF2-SHA256 hash in {config.INI_NAME}.")
        note.setStyleSheet("color:#666;")
        lay.addWidget(note)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _ok(self):
        err = change_password(self.old.text(), self.new.text(), self.confirm.text())
        if err:
            self.msg.setText(err)
            return
        QMessageBox.information(self, "PortableExcel", "The app password was changed.")
        self.accept()


def unlock(parent=None) -> bool:
    """Show the lock dialog unless disabled for tests. True when unlocked."""
    if lock_disabled():
        return True
    dlg = LockDialog(parent)
    return dlg.exec() == QDialog.Accepted

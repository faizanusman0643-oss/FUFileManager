"""Portable settings file (PortableExcel.ini, JSON content).

The file lives next to the executable when frozen (PyInstaller), otherwise in
the project root.  If that folder is not writable we fall back to the user's
configuration directory.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from typing import Any, Optional

INI_NAME = "PortableExcel.ini"
LOG_NAME = "PortableExcel.log"

_override_path: Optional[str] = None
_cache: Optional[dict] = None


def project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _writable(folder: str) -> bool:
    try:
        os.makedirs(folder, exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix=".pe_write_test", dir=folder)
        os.close(fd)
        os.remove(tmp)
        return True
    except OSError:
        return False


def user_config_dir() -> str:
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "PortableExcel")


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return project_root()


def settings_dir() -> str:
    if _override_path:
        return os.path.dirname(_override_path)
    d = app_dir()
    if _writable(d):
        return d
    d = user_config_dir()
    os.makedirs(d, exist_ok=True)
    return d


def ini_path() -> str:
    if _override_path:
        return _override_path
    return os.path.join(settings_dir(), INI_NAME)


def log_path() -> str:
    return os.path.join(settings_dir(), LOG_NAME)


def set_ini_path(path: Optional[str]) -> None:
    """Use another ini file (tests)."""
    global _override_path, _cache
    _override_path = path
    _cache = None


def load() -> dict:
    global _cache
    if _cache is None:
        try:
            with open(ini_path(), "r", encoding="utf-8") as fh:
                data = json.load(fh)
            _cache = data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            _cache = {}
    return _cache


def save() -> None:
    data = load()
    path = ini_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)
    os.replace(tmp, path)


def get(key: str, default: Any = None) -> Any:
    return load().get(key, default)


def put(key: str, value: Any, persist: bool = True) -> None:
    load()[key] = value
    if persist:
        try:
            save()
        except OSError:
            pass


# recent files ---------------------------------------------------------------
MAX_RECENT = 10


def recent_files() -> list:
    return [p for p in get("recent_files", []) if isinstance(p, str)]


def add_recent(path: str) -> None:
    path = os.path.abspath(path)
    lst = [p for p in recent_files() if os.path.normcase(p) != os.path.normcase(path)]
    lst.insert(0, path)
    put("recent_files", lst[:MAX_RECENT])

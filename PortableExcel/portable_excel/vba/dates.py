"""Date helpers: OLE-automation / Excel serial dates <-> datetime, parsing."""
from __future__ import annotations

import math
import re
from datetime import datetime, timedelta, date

EPOCH = datetime(1899, 12, 30)

MONTHS = ["january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"]
MONTH_ABBR = [m[:3] for m in MONTHS]
DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]


def serial_to_dt(s: float) -> datetime:
    days = math.floor(s)
    frac = s - days
    ms = round(frac * 86400000)
    try:
        return EPOCH + timedelta(days=days, milliseconds=ms)
    except OverflowError:
        from .errors import rt_error
        raise rt_error(6)


def dt_to_serial(dt) -> float:
    if isinstance(dt, date) and not isinstance(dt, datetime):
        dt = datetime(dt.year, dt.month, dt.day)
    delta = dt - EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0


def date_serial(y: int, m: int, d: int) -> float:
    """DateSerial with VBA normalisation (month 13 -> next year, day 0 -> prev month end)."""
    if 0 <= y < 30:
        y += 2000
    elif 30 <= y < 100:
        y += 1900
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    base = datetime(y, m, 1)
    return dt_to_serial(base) + (d - 1)


def time_serial(h: float, n: float, s: float) -> float:
    return (h * 3600 + n * 60 + s) / 86400.0


def _year2(y: int) -> int:
    if y < 30:
        return 2000 + y
    if y < 100:
        return 1900 + y
    return y


def _month_from_name(name: str):
    n = name.lower().rstrip(".")
    if n == "sept":
        n = "sep"
    for i, m in enumerate(MONTHS):
        if n == m or n == m[:3]:
            return i + 1
    return None


_TIME_RE = re.compile(r"^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?\s*([AaPp]\.?[Mm]\.?)?$")


def _parse_time(t: str):
    t = t.strip()
    m = _TIME_RE.match(t)
    if not m:
        return None
    h = int(m.group(1))
    n = int(m.group(2)) if m.group(2) else 0
    s = float(m.group(3)) if m.group(3) else 0.0
    ampm = m.group(4)
    if m.group(2) is None and not ampm:
        return None
    if ampm:
        if h > 12:
            return None
        if ampm[0] in "pP" and h != 12:
            h += 12
        elif ampm[0] in "aA" and h == 12:
            h = 0
    if h > 23 or n > 59 or s >= 60:
        return None
    return time_serial(h, n, s)


def _valid(y, m, d):
    try:
        datetime(y, m, d)
        return True
    except ValueError:
        return False


def _parse_date_part(p: str):
    p = p.strip().rstrip(",")
    if not p:
        return None
    # ISO yyyy-mm-dd or yyyy/mm/dd
    m = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", p)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return (y, mo, d) if _valid(y, mo, d) else None
    m = re.match(r"^(\d{1,2})[-/.](\d{1,2})(?:[-/.](\d{2,4}))?$", p)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        y = _year2(int(m.group(3))) if m.group(3) else datetime.now().year
        if a <= 12 and _valid(y, a, b):      # US m/d/y
            return y, a, b
        if b <= 12 and _valid(y, b, a):      # d/m/y
            return y, b, a
        return None
    # 31-Jan-2024, 31 Jan 2024, 31 January, 2024
    m = re.match(r"^(\d{1,2})[-\s/]+([A-Za-z]+\.?)[-\s/,]*(\d{2,4})?$", p)
    if m:
        mo = _month_from_name(m.group(2))
        if mo:
            y = _year2(int(m.group(3))) if m.group(3) else datetime.now().year
            d = int(m.group(1))
            return (y, mo, d) if _valid(y, mo, d) else None
    # Jan 31, 2024 / January 31 2024 / Jan-2024
    m = re.match(r"^([A-Za-z]+\.?)[-\s/]+(\d{1,2})(?:[,\s/-]+(\d{2,4}))?$", p)
    if m:
        mo = _month_from_name(m.group(1))
        if mo:
            d = int(m.group(2))
            y = _year2(int(m.group(3))) if m.group(3) else datetime.now().year
            return (y, mo, d) if _valid(y, mo, d) else None
    m = re.match(r"^([A-Za-z]+\.?)[-\s/,]+(\d{4})$", p)
    if m:
        mo = _month_from_name(m.group(1))
        if mo:
            return int(m.group(2)), mo, 1
    return None


def parse_date_string(s: str):
    """Return a serial date (float) or None if s is not a recognisable date/time."""
    if not isinstance(s, str):
        return None
    t = s.strip()
    if not t:
        return None
    # strip weekday names: "Monday, January 1, 2024"
    t2 = re.sub(r"^(sun|mon|tue|tues|wed|thu|thur|thurs|fri|sat)[a-z]*,?\s+", "", t, flags=re.I)
    tm = _parse_time(t2)
    if tm is not None:
        return tm
    # split date and time: try progressively
    m = re.match(r"^(.*?)[\sT]+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?\s*(?:[AaPp]\.?[Mm]\.?)?)$", t2)
    if m:
        dp = _parse_date_part(m.group(1))
        tp = _parse_time(m.group(2))
        if dp and tp is not None:
            return date_serial(*dp) + tp
        return None
    dp = _parse_date_part(t2)
    if dp:
        return date_serial(*dp)
    return None

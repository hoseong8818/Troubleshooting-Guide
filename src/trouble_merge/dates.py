"""날짜 읽기와 YYYY-MM-DD 변환 (T020, FR-011)."""

import re
from datetime import date, datetime, timedelta

_TEXT = re.compile(
    r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})"
    r"(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?$"
)
_EXCEL_EPOCH = date(1899, 12, 30)


def _iso(year, month, day):
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return None


def parse_date(value):
    """읽을 수 있으면 'YYYY-MM-DD' 문자열, 읽을 수 없으면 None.

    빈 칸(None, 빈 문자열)은 호출하는 쪽에서 먼저 걸러 낸다.
    두 자리 연도와 MM/DD/YYYY는 읽지 않는다.
    """
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        days = int(value)
        if days < 1 or days > 2958465:
            return None
        return (_EXCEL_EPOCH + timedelta(days=days)).isoformat()
    if isinstance(value, str):
        m = _TEXT.match(value.strip())
        if m:
            return _iso(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

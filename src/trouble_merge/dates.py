"""날짜 읽기와 YYYY-MM-DD 변환 (T020, FR-011)."""

import re
from datetime import date, datetime, timedelta

_TEXT = re.compile(
    r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})"
    r"(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?$"
)
_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})$")  # 8자리 YYYYMMDD (① 표기 차이, T056)
# 날짜로 시작하고 공백 뒤에 다른 글자가 붙은 칸(칸 밀림 의심, T058). 시각만 붙은 것은 _TEXT가 먼저 읽는다.
_DATE_THEN_TEXT = re.compile(r"^(\d{4}[-./]\d{1,2}[-./]\d{1,2}|\d{8})\s+\S")
_EXCEL_EPOCH = date(1899, 12, 30)
DEFAULT_MIN_DATE = "2000-01-01"


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
        whole = not isinstance(value, float) or value.is_integer()
        if whole and 10000000 <= days <= 99999999:  # 8자리 숫자는 일련번호가 아니라 YYYYMMDD로 본다
            return _compact(str(days))
        if days < 1 or days > 2958465:
            return None
        return (_EXCEL_EPOCH + timedelta(days=days)).isoformat()
    if isinstance(value, str):
        s = value.strip()
        m = _TEXT.match(s)
        if m:
            return _iso(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        return _compact(s)
    return None


def _compact(s):
    m = _COMPACT.match(s)
    if not m:
        return None
    return _iso(int(m.group(1)), int(m.group(2)), int(m.group(3)))


def looks_shifted(value):
    """날짜로 시작하고 뒤에 다른 글자가 붙어 있는가(칸 밀림 의심). 앞의 날짜는 실제로 있는 날짜여야 한다."""
    if not isinstance(value, str):
        return False
    s = value.strip()
    m = _DATE_THEN_TEXT.match(s)
    return bool(m) and parse_date(m.group(1)) is not None


def parse_bound(text):
    """명령줄의 날짜 범위 값. 'YYYY-MM-DD'만 받는다. 틀리면 ValueError."""
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", text or ""):
        raise ValueError(text)
    return date.fromisoformat(text).isoformat()

"""컬럼 대응 표와 변환 표 읽기, 양식 판별 (T019)."""

import csv
from pathlib import Path

STANDARD_COLUMNS = ("설비", "일자", "증상", "원인", "조치", "분류")


class ConfigError(Exception):
    """설정 파일 오류. 메시지에는 파일 이름과 줄 번호만 담는다(셀 값 금지)."""


def norm_key(value):
    """변환 표 비교용 키: 앞뒤 공백을 빼고 소문자로 바꾼다(FR-009, FR-018)."""
    return str(value).strip().lower()


def header_name(value):
    return "" if value is None else str(value).strip()


class Format:
    def __init__(self, name):
        self.name = name
        self.columns = []  # (원본 컬럼명, 표준 컬럼)
        self.header_row = 1
        self.sheet = None  # None이면 보이는 시트 전체
        self.invalid = False  # 한 표준 컬럼에 원본 컬럼이 여럿 대응되면 처리하지 않는다

    def source_names(self):
        return [source for source, _ in self.columns]

    def standard_of(self, source):
        for src, std in self.columns:
            if src == source:
                return std
        return None


def load_formats(config_dir):
    path = Path(config_dir) / "column_map.csv"
    formats = {}
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if not row or not any(c.strip() for c in row):
                    continue
                row = [c.strip() for c in row] + [""] * 5
                name, source, standard, header_row, sheet = row[:5]
                if standard not in STANDARD_COLUMNS:
                    raise ConfigError("column_map.csv %d줄: 표준 컬럼이 아님" % reader.line_num)
                fmt = formats.setdefault(name, Format(name))
                fmt.columns.append((source, standard))
                if header_row:
                    fmt.header_row = int(header_row)
                if sheet:
                    fmt.sheet = sheet
    except FileNotFoundError:
        raise ConfigError("column_map.csv 없음")
    except ValueError:
        raise ConfigError("column_map.csv 형식 오류")
    for fmt in formats.values():
        standards = [std for _, std in fmt.columns]
        if len(standards) != len(set(standards)):
            fmt.invalid = True
    return formats


def load_conversions(config_dir):
    """{'설비': {키: 표준 값}, '분류': {키: 표준 값}}. 같은 키가 표준 값을 달리하면 ConfigError."""
    path = Path(config_dir) / "conversion_map.csv"
    tables = {"설비": {}, "분류": {}}
    try:
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if not row or not any(c.strip() for c in row):
                    continue
                row = [c.strip() for c in row] + [""] * 3
                kind, original, standard = row[:3]
                if kind not in tables:
                    raise ConfigError("conversion_map.csv %d줄: 종류가 아님" % reader.line_num)
                key = norm_key(original)
                if key in tables[kind] and tables[kind][key] != standard:
                    raise ConfigError("conversion_map.csv %d줄: 같은 키에 다른 표준 값" % reader.line_num)
                tables[kind][key] = standard
    except FileNotFoundError:
        raise ConfigError("conversion_map.csv 없음")
    return tables


def detect_format(formats, grid, sheet_name):
    """헤더로 양식을 판별한다(FR-017). (양식 또는 None, 사유)를 돌려준다.

    사유: 'ok', 'none'(대응 규칙 없음), 'ambiguous'(양식 모호), 'invalid'(대응 중복)
    """
    best = []
    best_size = -1
    for fmt in formats.values():
        if fmt.sheet is not None and fmt.sheet != sheet_name:
            continue
        if len(grid) < fmt.header_row:
            continue
        header = {header_name(v) for v in grid[fmt.header_row - 1]}
        names = fmt.source_names()
        if not names or not all(n in header for n in names):
            continue
        size = len(names)
        if size > best_size:
            best, best_size = [fmt], size
        elif size == best_size:
            best.append(fmt)
    if not best:
        return None, "none"
    if len(best) > 1:
        return None, "ambiguous"
    if best[0].invalid:
        return None, "invalid"
    return best[0], "ok"

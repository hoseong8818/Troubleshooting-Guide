"""표준 컬럼으로 옮기기, 기타 항목, 분류 변환, 중복 의심, 정렬 (T022).

Phase 3 범위다. '확인 필요' 판정(check.py)은 Phase 4에서 더하므로, 지금은 변환표에 없는
설비·분류는 정리한 원래 값을 그대로 쓰고 '확인 필요' 칸은 비워 둔다.
"""

from .dates import parse_date
from .mapping import header_name, norm_key
from .reader import is_blank

OUTPUT_HEADER = [
    "설비", "일자", "증상", "원인", "조치", "분류", "기타 항목",
    "확인 필요 이유", "원래 값", "중복 의심", "원본 파일", "원본 시트", "원본 행",
]
DUPLICATE_MARK = "중복 의심"


def _text(value):
    if is_blank(value):
        return None
    return str(value).strip()


def convert_rows(file_name, sheet, fmt, conversions):
    """한 시트의 입력 행을 출력 행(사전) 목록으로 바꾼다. 모든 칸이 빈 행은 뺀다(FR-014)."""
    grid = sheet.grid
    header = [header_name(v) for v in grid[fmt.header_row - 1]]

    mapped = {}  # 표준 컬럼 -> 열 위치
    for index, name in enumerate(header):
        standard = fmt.standard_of(name) if name else None
        if standard is not None and standard not in mapped:
            mapped[standard] = index
    others = [(i, name) for i, name in enumerate(header)
              if name and fmt.standard_of(name) is None]

    records = []
    for number, row in enumerate(grid[fmt.header_row:], start=fmt.header_row + 1):
        if all(is_blank(v) for v in row):
            continue

        def cell(standard):
            index = mapped.get(standard)
            return row[index] if index is not None and index < len(row) else None

        equipment = _text(cell("설비"))
        if equipment is not None:
            equipment = conversions["설비"].get(norm_key(equipment), equipment)
        category = _text(cell("분류"))
        if category is not None:
            category = conversions["분류"].get(norm_key(category), category)

        raw_date = cell("일자")
        if is_blank(raw_date):
            date_text = None
        else:
            date_text = parse_date(raw_date) or _text(raw_date)  # 해석 불가 처리는 Phase 4

        parts = []
        for index, name in others:
            value = _text(row[index]) if index < len(row) else None
            if value is not None:
                parts.append("%s: %s" % (name, value))

        records.append({
            "설비": equipment,
            "일자": date_text,
            "증상": _text(cell("증상")),
            "원인": _text(cell("원인")),
            "조치": _text(cell("조치")),
            "분류": category,
            "기타 항목": "; ".join(parts) if parts else None,
            "확인 필요 이유": None,
            "원래 값": None,
            "중복 의심": None,
            "원본 파일": file_name,
            "원본 시트": sheet.name,
            "원본 행": number,
        })
    return records


def mark_duplicates(records):
    """설비·일자·증상이 모두 같은 행에 '중복 의심'을 표시한다. 행은 지우지 않는다(FR-010)."""
    groups = {}
    for record in records:
        if record["설비"] is None or record["증상"] is None:
            continue  # 판단 불가 행은 비교에서 뺀다(Phase 4에서 '확인 필요'로 바뀐다)
        key = (record["설비"], record["일자"], record["증상"])
        groups.setdefault(key, []).append(record)
    for group in groups.values():
        if len(group) > 1:
            for record in group:
                record["중복 의심"] = DUPLICATE_MARK


def sort_records(records):
    return sorted(
        records,
        key=lambda r: (r["설비"] or "", r["일자"] or "", r["원본 파일"], r["원본 행"]),
    )

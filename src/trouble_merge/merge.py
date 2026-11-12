"""표준 컬럼으로 옮기기, 기타 항목, 확인 필요, 중복 의심, 정렬 (T022, T029)."""

from .check import CHECK, judge, text
from .mapping import header_name
from .reader import is_blank

OUTPUT_HEADER = [
    "설비", "일자", "증상", "원인", "조치", "분류", "기타 항목",
    "확인 필요 이유", "원래 값", "중복 의심", "원본 파일", "원본 시트", "원본 행",
]
DUPLICATE_MARK = "중복 의심"


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

        raw = {c: cell(c) for c in ("설비", "일자", "증상", "원인", "조치", "분류")}
        final, reasons, originals = judge(raw, conversions)

        parts = []
        for index, name in others:
            value = text(row[index]) if index < len(row) else None
            if value is not None:
                parts.append("%s: %s" % (name, value))

        record = dict(final)
        record.update({
            "기타 항목": "; ".join(parts) if parts else None,
            "확인 필요 이유": "; ".join(reasons) if reasons else None,
            "원래 값": "; ".join(originals) if originals else None,
            "중복 의심": None,
            "원본 파일": file_name,
            "원본 시트": sheet.name,
            "원본 행": number,
        })
        records.append(record)
    return records


def mark_duplicates(records):
    """설비·일자·증상이 모두 같은 행에 '중복 의심'을 표시한다. 행은 지우지 않는다(FR-010).

    설비, 일자, 증상 중 하나라도 '확인 필요'인 행은 비교에서 뺀다.
    """
    groups = {}
    for record in records:
        if CHECK in (record["설비"], record["일자"], record["증상"]):
            continue
        key = (record["설비"], record["일자"], record["증상"])
        groups.setdefault(key, []).append(record)
    for group in groups.values():
        if len(group) > 1:
            for record in group:
                record["중복 의심"] = DUPLICATE_MARK


def sort_records(records):
    """설비, 일자, 원본 파일, 원본 행 순. 설비가 '확인 필요'인 행은 맨 아래(FR-016)."""
    return sorted(
        records,
        key=lambda r: (
            r["설비"] == CHECK, r["설비"] or "", r["일자"] or "", r["원본 파일"], r["원본 행"],
        ),
    )

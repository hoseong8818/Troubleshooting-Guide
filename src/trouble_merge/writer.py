"""통합 엑셀과 감사 결과 쓰기 (T023). 덮어쓰지 않고, 이미 있으면 실행 시각을 붙인다."""

from pathlib import Path

from openpyxl import Workbook

from .merge import OUTPUT_HEADER

DANGEROUS_PREFIXES = ("=", "+", "-", "@")


def unique_path(out_dir, stem, suffix, stamp):
    out_dir = Path(out_dir)
    path = out_dir / (stem + suffix)
    if not path.exists():
        return path
    path = out_dir / ("%s_%s%s" % (stem, stamp, suffix))
    counter = 2
    while path.exists():
        path = out_dir / ("%s_%s_%d%s" % (stem, stamp, counter, suffix))
        counter += 1
    return path


def write_workbook(records, out_dir, stamp):
    """통합 엑셀을 쓰고 (경로, 수식 방지 접두사를 붙인 셀 수)를 돌려준다(NFR-008)."""
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "통합"
    ws.append(OUTPUT_HEADER)
    prefixed = 0
    for row_number, record in enumerate(records, start=2):
        for column_number, name in enumerate(OUTPUT_HEADER, start=1):
            value = record[name]
            cell = ws.cell(row=row_number, column=column_number)
            if isinstance(value, str):
                if value.startswith(DANGEROUS_PREFIXES):
                    value = "'" + value
                    prefixed += 1
                cell.value = value
                cell.data_type = "s"  # 항상 문자열로 저장
            else:
                cell.value = value
    path = unique_path(out_dir, "merged", ".xlsx", stamp)
    wb.save(path)
    return path, prefixed


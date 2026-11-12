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


def write_audit(out_dir, stamp, summary):
    """감사 결과 텍스트를 쓴다. summary는 이름과 숫자(또는 짧은 문구)만 담는다. 셀 값은 넣지 않는다."""
    lines = ["감사 결과", ""]
    lines.append("입력 행 수: %d" % summary["input_rows"])
    lines.append("출력 행 수: %d" % summary["output_rows"])
    lines.append("행 수 일치: %s" % ("예" if summary["input_rows"] == summary["output_rows"] else "아니오"))
    lines.append("읽은 파일 수: %d" % summary["files_read"])
    lines.append("건너뛰거나 실패한 파일 수: %d" % summary["files_failed"])
    lines.append("대응 규칙이 없거나 양식이 모호한 시트 수: %d" % summary["sheets_unmatched"])
    lines.append("숨긴 시트 수: %d" % summary["hidden_sheets"])
    lines.append("수식 방지 접두사 처리 건수: %d" % summary["prefixed"])
    path = unique_path(out_dir, "audit_" + stamp, ".txt", stamp)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path

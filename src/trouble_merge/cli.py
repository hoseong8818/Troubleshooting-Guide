"""명령줄 진입점 (T023, T036, FR-019). 종료 코드: 0 정상, 1 일부 실패, 2 전부 실패 또는 설정 오류.

로그와 오류는 SafeLog 한 길로만 낸다. 예외는 메시지도 종류도 내지 않는다(셀 값이나 내부 이름이 섞일 수 있다).
"""

import argparse
import sys
import warnings
from datetime import datetime
from pathlib import Path

from .audit import Audit, write_name_map
from .mapping import ConfigError, detect_format, header_name, load_conversions, load_formats
from .merge import convert_rows, mark_duplicates, sort_records
from .reader import is_readable_extension, list_inputs, read_workbook
from .safe_log import SafeLog
from .writer import write_workbook

UNMATCHED_TEXT = {
    "none": "대응 규칙 없음 시트",
    "ambiguous": "양식 모호 시트",
    "invalid": "대응 중복 양식 시트",
}


def build_parser():
    parser = argparse.ArgumentParser(prog="trouble_merge")
    parser.add_argument("--input", nargs="+", required=True, help="폴더 또는 파일")
    parser.add_argument("--config", required=True, help="column_map.csv, conversion_map.csv가 있는 폴더")
    parser.add_argument("--out", required=True, help="출력 폴더")
    parser.add_argument("--show-names", action="store_true",
                        help="로그·감사 결과에 실제 파일명·시트명을 낸다(기본은 마스킹)")
    return parser


def main(argv=None):
    warnings.simplefilter("ignore")  # 라이브러리 경고 문구에 값이 섞이지 않게 한다
    args = build_parser().parse_args(argv)
    try:
        return run(args)
    except Exception:  # 메시지와 종류를 모두 버리고 고정 문구만 낸다
        print("처리 중 예기치 못한 오류", file=sys.stderr)
        return 2


def run(args):
    try:
        formats = load_formats(args.config)
        conversions = load_conversions(args.config)
    except ConfigError as error:
        print("설정 오류: %s" % error, file=sys.stderr)
        return 2

    registered = {name for fmt in formats.values() for name in fmt.source_names()}
    log = SafeLog(registered, show_names=args.show_names)
    audit = Audit(log)

    files = list_inputs(args.input)
    for path in files:
        log.file_label(Path(path).name)  # 입력 순서대로 file_01, file_02 ... 번호를 정한다

    records = []
    unmatched_sheets = 0

    for path in files:
        name = Path(path).name
        if not is_readable_extension(path):
            audit.failed_files.append(name)
            log.log("건너뜀(지원하지 않는 형식)", file=name)
            continue
        try:
            sheets, hidden, formula_missing = read_workbook(path)
        except Exception:  # 예외 종류·메시지 모두 버린다(내부 이름이나 셀 값이 섞일 수 있다)
            audit.failed_files.append(name)
            log.log("열기 실패", file=name)
            continue

        audit.hidden.append((name, hidden))
        audit.formula_missing += formula_missing
        if hidden:
            log.log("숨긴 시트 건너뜀", file=name, count=hidden)

        matched = False
        for sheet in sheets:
            fmt, reason = detect_format(formats, sheet.grid, sheet.name)
            if fmt is None:
                unmatched_sheets += 1
                audit.unmatched.append((UNMATCHED_TEXT[reason], name, sheet.name))
                log.log(UNMATCHED_TEXT[reason], file=name, sheet=sheet.name)
                for position, value in enumerate(sheet.grid[0] if sheet.grid else []):
                    if header_name(value):
                        audit.add_unregistered(name, sheet.name, log.column_label(header_name(value), position))
                continue
            matched = True
            header = [header_name(v) for v in sheet.grid[fmt.header_row - 1]]
            for position, column in enumerate(header):
                if column and fmt.standard_of(column) is None:
                    audit.add_unregistered(name, sheet.name, log.column_label(column, position))

            rows = convert_rows(name, sheet, fmt, conversions)
            audit.input_rows += len(rows)
            for record in rows:
                for standard in record.pop("_blank"):
                    source = fmt.source_of(standard)
                    position = header.index(source) if source in header else 0
                    audit.blank_required.append(
                        (name, sheet.name, record["원본 행"], log.column_label(source, position)))
            records.extend(rows)

        if matched:
            audit.files_read += 1
        else:
            audit.failed_files.append(name)

    if audit.files_read == 0:
        log.log("처리할 수 있는 파일이 없다")
        return 2

    mark_duplicates(records)
    records = sort_records(records)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    _, audit.prefixed = write_workbook(records, args.out, stamp)
    audit.output_rows = len(records)
    audit.check_count = sum(1 for r in records if r["확인 필요 이유"] is not None)
    audit.duplicate_count = sum(1 for r in records if r["중복 의심"] is not None)
    audit.write(args.out, stamp)
    write_name_map(log, args.out, stamp)

    print("완료: 출력 %d행" % len(records))
    return 1 if (audit.failed_files or unmatched_sheets) else 0


if __name__ == "__main__":
    sys.exit(main())

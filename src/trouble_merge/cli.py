"""명령줄 진입점 (T023, FR-019). 종료 코드: 0 정상, 1 일부 실패, 2 전부 실패 또는 설정 오류.

로그와 오류에는 셀 값을 내지 않는다. 예외 메시지는 버리고 고정 문구만 낸다.
(safe_log로 통일하는 일은 T034·T036에서 한다.)
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from .mapping import ConfigError, detect_format, load_conversions, load_formats
from .merge import convert_rows, mark_duplicates, sort_records
from .reader import is_readable_extension, list_inputs, read_workbook
from .writer import write_audit, write_workbook


def build_parser():
    parser = argparse.ArgumentParser(prog="trouble_merge")
    parser.add_argument("--input", nargs="+", required=True, help="폴더 또는 파일")
    parser.add_argument("--config", required=True, help="column_map.csv, conversion_map.csv가 있는 폴더")
    parser.add_argument("--out", required=True, help="출력 폴더")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        formats = load_formats(args.config)
        conversions = load_conversions(args.config)
    except ConfigError as error:
        print("설정 오류: %s" % error, file=sys.stderr)
        return 2

    records = []
    files_read = files_failed = sheets_unmatched = hidden_total = input_rows = 0

    for path in list_inputs(args.input):
        if not is_readable_extension(path):
            files_failed += 1  # .xlsm, .xls 등은 열지 않는다(NFR-007)
            continue
        try:
            sheets, hidden = read_workbook(path)
        except Exception:  # 메시지에 셀 값이 섞일 수 있어 종류만 남기고 버린다
            files_failed += 1
            continue
        hidden_total += hidden
        matched = False
        for sheet in sheets:
            fmt, _reason = detect_format(formats, sheet.grid, sheet.name)
            if fmt is None:
                sheets_unmatched += 1
                continue
            matched = True
            rows = convert_rows(Path(path).name, sheet, fmt, conversions)
            input_rows += len(rows)
            records.extend(rows)
        if matched:
            files_read += 1
        else:
            files_failed += 1

    if files_read == 0:
        print("처리할 수 있는 파일이 없다", file=sys.stderr)
        return 2

    mark_duplicates(records)
    records = sort_records(records)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    _, prefixed = write_workbook(records, args.out, stamp)
    write_audit(args.out, stamp, {
        "input_rows": input_rows,
        "output_rows": len(records),
        "files_read": files_read,
        "files_failed": files_failed,
        "sheets_unmatched": sheets_unmatched,
        "hidden_sheets": hidden_total,
        "prefixed": prefixed,
    })
    print("완료: 출력 %d행" % len(records))
    return 1 if (files_failed or sheets_unmatched) else 0


if __name__ == "__main__":
    sys.exit(main())

"""입력 엑셀 읽기 (T021, FR-012, FR-014).

보이는 시트만 읽고, 병합 셀은 첫 칸 값을 범위 전체에 채우며, 수식 셀은 저장된 계산값만 읽는다.
"""

import zipfile
from pathlib import Path

from openpyxl import load_workbook

READABLE = ".xlsx"
KNOWN_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


class SheetData:
    def __init__(self, name, grid):
        self.name = name
        self.grid = grid  # 행 목록(각 행은 값 목록). grid[0]이 엑셀 1행.


def list_inputs(paths):
    files = []
    for raw in paths:
        p = Path(raw)
        if p.is_dir():
            files.extend(
                sorted(
                    x for x in p.iterdir()
                    if x.is_file() and not x.name.startswith("~$")
                    and x.suffix.lower() in KNOWN_EXTENSIONS
                )
            )
        else:
            files.append(p)
    return files


def is_readable_extension(path):
    return Path(path).suffix == READABLE


MB = 1024 * 1024


def check_size_limits(path, max_file_mb, max_unzipped_mb):
    """파일을 열기 전에 크기 상한을 본다(NFR-007). 넘으면 사유 종류, 아니면 None.

    사유 종류는 감사 결과에 그대로 나가는 고정 문구다. 압축 파일이 아니면 여기서는 넘기고
    열기 단계의 실패로 처리한다.
    """
    try:
        if Path(path).stat().st_size > max_file_mb * MB:
            return "파일 크기"
        with zipfile.ZipFile(path) as z:
            if sum(info.file_size for info in z.infolist()) > max_unzipped_mb * MB:
                return "압축 해제 크기"
    except (OSError, zipfile.BadZipFile):
        return None
    return None


def exceeds_row_limit(sheets, max_rows):
    """보이는 시트 중 헤더를 포함한 행 수가 상한을 넘는 것이 있는가(NFR-007)."""
    return any(len(sheet.grid) > max_rows for sheet in sheets)


def _count_formulas_without_value(path, visible_names, values_wb):
    """저장된 계산값이 없는 수식 셀 수(FR-014). 보이는 시트만 센다."""
    formulas_wb = load_workbook(path, data_only=False)
    count = 0
    try:
        for ws in formulas_wb.worksheets:
            if ws.title not in visible_names:
                continue
            values_ws = values_wb[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.data_type == "f" and values_ws[cell.coordinate].value is None:
                        count += 1
    finally:
        formulas_wb.close()
    return count


def read_workbook(path):
    """(보이는 시트 목록, 숨긴 시트 수, 계산값 없는 수식 셀 수)를 돌려준다.

    열 수 없으면 예외를 그대로 낸다. 호출하는 쪽이 메시지를 버리고 종류만 남긴다.
    """
    wb = load_workbook(path, data_only=True)
    sheets = []
    hidden = 0
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                hidden += 1
                continue
            grid = [list(row) for row in ws.iter_rows(values_only=True)]
            for rng in ws.merged_cells.ranges:
                top_left = ws.cell(rng.min_row, rng.min_col).value
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        if r - 1 < len(grid) and c - 1 < len(grid[r - 1]):
                            grid[r - 1][c - 1] = top_left
            sheets.append(SheetData(ws.title, grid))
        formula_missing = _count_formulas_without_value(path, {s.name for s in sheets}, wb)
    finally:
        wb.close()
    return sheets, hidden, formula_missing


def is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")

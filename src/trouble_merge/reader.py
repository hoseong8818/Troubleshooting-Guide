"""입력 엑셀 읽기 (T021, FR-012, FR-014).

보이는 시트만 읽고, 병합 셀은 첫 칸 값을 범위 전체에 채우며, 수식 셀은 저장된 계산값만 읽는다.
"""

import zipfile
from datetime import datetime
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


def _normalize_excel_value(value):
    """COM이 돌려준 값을 openpyxl 경로와 같은 모양으로 맞춘다(날짜는 시간대 없는 datetime)."""
    if hasattr(value, "year") and hasattr(value, "tzinfo"):
        return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
    return value


def read_workbook_excel(path):
    """설치된 엑셀을 조종해 읽는다(NFR-006). 반환 형식은 read_workbook과 같다.

    파일은 읽기 전용으로 열고, 엑셀은 항상 닫는다. 값은 엑셀이 연 창에서 읽으므로
    엑셀만 열 수 있는 파일도 같은 코드로 읽는다.
    """
    import win32com.client  # pywin32. 이 경로를 쓸 때만 필요하다

    app = win32com.client.DispatchEx("Excel.Application")
    wb = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(str(Path(path).resolve()), 0, True)  # UpdateLinks=0, ReadOnly=True
        sheets = []
        hidden = 0
        formula_missing = 0
        for ws in wb.Worksheets:
            if ws.Visible != -1:  # xlSheetVisible
                hidden += 1
                continue
            used = ws.UsedRange
            top, left = used.Row, used.Column
            rows, cols = used.Rows.Count, used.Columns.Count
            values = used.Value
            formulas = used.Formula
            if rows == 1 and cols == 1:
                values, formulas = ((values,),), ((formulas,),)
            grid = [[None] * (left - 1 + cols) for _ in range(top - 1 + rows)]
            for r in range(rows):
                for c in range(cols):
                    value = values[r][c]
                    grid[top - 1 + r][left - 1 + c] = _normalize_excel_value(value)
                    formula = formulas[r][c]
                    if value is None and isinstance(formula, str) and formula.startswith("="):
                        formula_missing += 1
            if used.MergeCells is not False:  # 병합 칸이 있을 때만 하나씩 확인한다
                for r in range(rows):
                    for c in range(cols):
                        cell = used.Cells(r + 1, c + 1)
                        if cell.MergeCells:
                            area = cell.MergeArea
                            fill = _normalize_excel_value(area.Cells(1, 1).Value)
                            grid[top - 1 + r][left - 1 + c] = fill
            sheets.append(SheetData(ws.Name, grid))
        return sheets, hidden, formula_missing
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        finally:
            app.Quit()


def is_blank(value):
    return value is None or (isinstance(value, str) and value.strip() == "")

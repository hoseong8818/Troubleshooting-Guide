"""T017g: 일자가 빈 행은 같은 파일의 같은 시트 안에서만 중복 의심을 비교한다. 일자가 있는 행은 파일이 달라도 비교한다 (FR-010).

일자가 없으면 서로 다른 파일의 같은 설비·증상을 같은 사건이라 볼 근거가 약해서 파일 사이는 비교하지 않는다.
같은 파일 안의 일자 빈 칸끼리는 기존대로 비교한다(test_duplicates).
"""

import unittest

from openpyxl import Workbook

from tests.support import STD_HEADER_A, CliCase, write_xlsx

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분", "작업자"]
ROW_A = ["PMP01", "소음 발생", "원인", "조치", "베어링"]
ROW_B = ["2026-11-20", "PMP01", "소음 발생", "원인", "조치", "베어링", "작업자1"]
MARK = "중복 의심"


def marked(records):
    return {(r["원본 파일"], r["원본 시트"], r["원본 행"]) for r in records if r["중복 의심"] == MARK}


class UndatedDuplicatesTest(CliCase):
    def folder(self):
        folder = self.tmp / "in"
        folder.mkdir()
        return folder

    def test_undated_rows_in_different_files_are_not_compared(self):
        folder = self.folder()
        write_xlsx(folder / "a.xlsx", STD_HEADER_A, [ROW_A])
        write_xlsx(folder / "b.xlsx", STD_HEADER_A, [ROW_A])
        _, records = self.run_ok([folder])
        self.assertEqual(len(records), 2)
        self.assertEqual(marked(records), set(), "일자 없는 행이 파일을 넘어 중복으로 표시됨")

    def test_undated_rows_in_the_same_file_and_sheet_are_compared(self):
        folder = self.folder()
        write_xlsx(folder / "a.xlsx", STD_HEADER_A, [ROW_A, ROW_A])
        _, records = self.run_ok([folder])
        self.assertEqual(len(marked(records)), 2, "같은 파일 안의 중복이 표시되지 않음")
        self.assertEqual(len(records), 2, "행이 지워짐")

    def test_undated_rows_in_different_sheets_are_not_compared(self):
        wb = Workbook()
        first = wb.active
        first.title = "시트A"
        first.append(STD_HEADER_A)
        first.append(ROW_A)
        second = wb.create_sheet("시트B")
        second.append(STD_HEADER_A)
        second.append(ROW_A)
        path = self.tmp / "two_sheets.xlsx"
        wb.save(path)
        _, records = self.run_ok([path])
        self.assertEqual(len(records), 2)
        self.assertEqual(marked(records), set(), "일자 없는 행이 시트를 넘어 중복으로 표시됨")

    def test_dated_rows_are_still_compared_across_files(self):
        folder = self.folder()
        write_xlsx(folder / "a.xlsx", HEADER_B, [ROW_B])
        write_xlsx(folder / "b.xlsx", HEADER_B, [ROW_B])
        _, records = self.run_ok([folder])
        self.assertEqual({r["원본 파일"] for r in records if r["중복 의심"] == MARK}, {"a.xlsx", "b.xlsx"})

    def test_dated_and_undated_rows_are_never_duplicates_of_each_other(self):
        folder = self.folder()
        write_xlsx(folder / "a.xlsx", STD_HEADER_A, [ROW_A])
        write_xlsx(folder / "b.xlsx", HEADER_B, [ROW_B])
        _, records = self.run_ok([folder])
        self.assertEqual(len(records), 2)
        self.assertEqual(marked(records), set())


if __name__ == "__main__":
    unittest.main()

"""T042b: 보이는 시트는 모두 읽고 시트마다 양식을 판별하며, hidden·veryHidden 시트는 읽지 않는다 (FR-012, FR-017)."""

import unittest

from openpyxl import Workbook

from tests.support import (
    CONFIG, STD_HEADER_A, CliCase, audit_files, by_source_row, read_sheet, result_files, run_cli,
)

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분", "작업자"]


def make_workbook(path, with_unknown_sheet=False):
    wb = Workbook()
    a = wb.active
    a.title = "시트A"
    a.append(STD_HEADER_A)
    a.append(["PMP01", "증상A", "원인A", "조치A", "베어링"])
    b = wb.create_sheet("시트B")
    b.append(HEADER_B)
    b.append(["2026-11-20", "CNV01", "증상B", "원인B", "조치B", "모터", "작업자1"])
    hidden = wb.create_sheet("숨김")
    hidden.append(STD_HEADER_A)
    hidden.append(["PMP01", "LEAK-숨김1", "원인", "조치", "베어링"])
    hidden.sheet_state = "hidden"
    very = wb.create_sheet("깊은숨김")
    very.append(STD_HEADER_A)
    very.append(["PMP01", "LEAK-숨김2", "원인", "조치", "베어링"])
    very.sheet_state = "veryHidden"
    if with_unknown_sheet:
        memo = wb.create_sheet("메모")
        memo.append(["메모"])
        memo.append(["LEAK-메모 내용"])
    wb.save(path)


class MultiSheetTest(CliCase):
    def test_visible_sheets_are_read_each_with_its_own_form(self):
        path = self.tmp / "multi.xlsx"
        make_workbook(path)
        _, records = self.run_ok([path])
        self.assertEqual(len(records), 2)
        self.assertEqual({r["원본 시트"] for r in records}, {"시트A", "시트B"})
        by_sheet = {r["원본 시트"]: r for r in records}
        self.assertEqual(by_sheet["시트A"]["설비"], "PMP01")
        self.assertEqual(by_sheet["시트B"]["설비"], "CNV01")
        self.assertEqual(by_sheet["시트B"]["일자"], "2026-11-20")

    def test_hidden_and_very_hidden_sheets_are_not_read_but_reported(self):
        path = self.tmp / "multi.xlsx"
        make_workbook(path)
        _, records = self.run_ok([path])
        self.assertNotIn("숨김", {r["원본 시트"] for r in records})
        self.assertNotIn("깊은숨김", {r["원본 시트"] for r in records})
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^숨긴 시트 수: 2$")
        self.assertNotIn("LEAK-", text)

    def test_unknown_sheet_is_reported_and_the_rest_is_processed(self):
        path = self.tmp / "multi.xlsx"
        make_workbook(path, with_unknown_sheet=True)
        result = run_cli([path], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        files = result_files(self.out)
        self.assertEqual(len(files), 1)
        _, records = read_sheet(files[0])
        self.assertEqual(len(records), 2)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^대응 규칙 없음 시트: file_01 sheet_\d\d$")
        self.assertNotIn("LEAK-", text + result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()

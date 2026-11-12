"""T041: 양식 2종·파일 5개의 모든 행이 통합 시트 한 장에 모이고 정답과 일치한다 (US4 수용 기준 1, SC-001).

입력은 data/fake/의 A_basic, A_trap, B_basic, B_more, B_trap. 정답은 expected/ALL_expected.xlsx(33행).
"""

import unittest

from tests.support import EXPECTED, FAKE, CliCase, read_sheet

FILES = ["A_basic.xlsx", "A_trap.xlsx", "B_basic.xlsx", "B_more.xlsx", "B_trap.xlsx"]


class MultiFileTest(CliCase):
    def _compare(self, records, header):
        exp_header, exp_records = read_sheet(EXPECTED / "ALL_expected.xlsx")
        self.assertEqual(header, exp_header, "컬럼 이름·순서가 정답과 다름")
        self.assertEqual(len(records), len(exp_records), "출력 행 수가 정답과 다름")
        for index, (got, exp) in enumerate(zip(records, exp_records), start=1):
            for column in exp_header:
                with self.subTest(출력행=index, 컬럼=column):
                    self.assertEqual(got.get(column), exp.get(column))

    def test_five_files_match_expected_row_by_row(self):
        header, records = self.run_ok([FAKE / name for name in FILES])
        self._compare(records, header)

    def test_folder_input_gives_the_same_result(self):
        header, records = self.run_ok([FAKE])
        self._compare(records, header)

    def test_every_file_contributes_rows(self):
        _, records = self.run_ok([FAKE / name for name in FILES])
        self.assertEqual({r["원본 파일"] for r in records}, set(FILES))

    def test_duplicates_are_marked_across_files(self):
        # B_basic 7·8행과 B_more 2행은 설비·일자·증상이 같다(파일이 달라도 비교, FR-010)
        _, records = self.run_ok([FAKE / name for name in FILES])
        marked = {(r["원본 파일"], r["원본 행"]) for r in records if r["중복 의심"] == "중복 의심"}
        self.assertIn(("B_basic.xlsx", 7), marked)
        self.assertIn(("B_basic.xlsx", 8), marked)
        self.assertIn(("B_more.xlsx", 2), marked)


if __name__ == "__main__":
    unittest.main()

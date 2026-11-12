"""T013: A_basic 통합 결과를 정답과 한 행씩 비교한다 (SC-001)."""

import unittest

from tests.support import EXPECTED, CliCase, read_sheet


class CompareTest(CliCase):
    def test_a_basic_matches_expected_row_by_row(self):
        header, records = self.run_basic()
        exp_header, exp_records = read_sheet(EXPECTED / "A_basic_expected.xlsx")

        self.assertEqual(header, exp_header, "컬럼 이름·순서가 정답과 다름")
        self.assertEqual(len(records), len(exp_records), "출력 행 수가 정답과 다름")
        for index, (got, exp) in enumerate(zip(records, exp_records), start=1):
            for column in exp_header:
                with self.subTest(출력행=index, 컬럼=column):
                    self.assertEqual(got.get(column), exp.get(column))


if __name__ == "__main__":
    unittest.main()

"""T017: 설비·일자·증상이 같은 행에 '중복 의심'을 표시하고 행은 지우지 않는다 (FR-010, SC-007).

양식 A는 일자가 비므로 설비·증상이 같은 행이 중복 의심이 된다(합의된 결정).
"""

import unittest

from tests.support import FAKE, CliCase, by_source_row, input_rows

# 설비 PMP01 + 증상 '소음 발생'(공백·대소문자 정리 후 같음)인 원본 행
DUPLICATE_ROWS = {2, 6, 9}


class DuplicatesTest(CliCase):
    def test_duplicates_are_marked_not_removed(self):
        _, records = self.run_basic()
        rows = by_source_row(records)
        for number in DUPLICATE_ROWS:
            with self.subTest(원본행=number):
                self.assertIn(number, rows, "중복 행이 지워짐")
                self.assertEqual(rows[number]["중복 의심"], "중복 의심")

    def test_only_duplicates_are_marked(self):
        _, records = self.run_basic()
        marked = {r["원본 행"] for r in records if r["중복 의심"] == "중복 의심"}
        self.assertEqual(marked, DUPLICATE_ROWS)

    def test_no_row_is_lost(self):
        _, records = self.run_basic()
        self.assertEqual(len(records), len(input_rows(FAKE / "A_basic.xlsx")))


if __name__ == "__main__":
    unittest.main()

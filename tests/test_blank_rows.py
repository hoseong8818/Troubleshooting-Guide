"""T017b: 완전히 빈 행은 입력 행에서 제외된다 (FR-014)."""

import unittest

from tests.support import FAKE, CliCase, input_rows

# A_basic.xlsx에서 모든 칸이 빈 행의 엑셀 행 번호 (tools/make_fake_data.py 참고)
BLANK_ROWS_IN_A_BASIC = {5, 14}


class BlankRowsTest(CliCase):
    def test_blank_rows_are_excluded(self):
        _, records = self.run_basic()
        source_rows = {r["원본 행"] for r in records}
        self.assertTrue(
            source_rows.isdisjoint(BLANK_ROWS_IN_A_BASIC),
            "빈 행이 출력에 들어 있음",
        )

    def test_every_non_blank_row_is_kept(self):
        expected = {number for number, _ in input_rows(FAKE / "A_basic.xlsx")}
        _, records = self.run_basic()
        self.assertEqual({r["원본 행"] for r in records}, expected)


if __name__ == "__main__":
    unittest.main()

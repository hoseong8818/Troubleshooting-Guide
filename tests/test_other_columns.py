"""T015: 표준에 없는 컬럼이 '기타 항목'에 `이름: 값`으로 보관된다 (FR-003, FR-015)."""

import unittest

from tests.support import (
    FAKE, STD_HEADER_A, CliCase, by_source_row, input_rows, make_config, write_xlsx,
)


class OtherColumnsTest(CliCase):
    def test_error_code_kept_for_a_basic(self):
        _, records = self.run_basic()
        rows = by_source_row(records)
        for number, source in input_rows(FAKE / "A_basic.xlsx"):
            code = source["에러코드"]
            expected = ("에러코드: %s" % code) if code else None
            with self.subTest(원본행=number):
                self.assertEqual(rows[number]["기타 항목"], expected)

    def test_multiple_columns_keep_source_order_and_skip_empty(self):
        header = STD_HEADER_A + ["메모가", "메모나"]
        rows = [
            ["PMP01", "증상1", "원인1", "조치1", "베어링", "x1", None],
            ["PMP01", "증상2", "원인2", "조치2", "베어링", "x2", "y2"],
            ["PMP01", "증상3", "원인3", "조치3", "베어링", None, None],
        ]
        path = write_xlsx(self.tmp / "other.xlsx", header, rows)
        config = make_config(self.tmp)
        _, records = self.run_ok([path], config)
        got = by_source_row(records)
        self.assertEqual(got[2]["기타 항목"], "메모가: x1")
        self.assertEqual(got[3]["기타 항목"], "메모가: x2; 메모나: y2")
        self.assertIsNone(got[4]["기타 항목"])


if __name__ == "__main__":
    unittest.main()

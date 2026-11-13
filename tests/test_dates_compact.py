"""T056: 8자리 날짜 YYYYMMDD를 읽는다 (① 표기 차이, FR-011).

양식 D처럼 `20230512`로 적는 것은 오류가 아니라 표기 차이다. '확인 필요'로 보내지 않고 통일한다.
"""

import unittest

from tests.support import CONFIG, CliCase, by_source_row, read_sheet, result_files, run_cli, write_xlsx

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분"]


class CompactDateTest(CliCase):
    def run_rows(self, rows):
        path = write_xlsx(self.tmp / "compact.xlsx", HEADER_B, rows)
        result = run_cli([path], CONFIG, self.out)
        self.assertIn(result.returncode, (0, 1), result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        return by_source_row(records)

    def test_eight_digit_text_and_number_are_read(self):
        got = self.run_rows([
            ["20230512", "PMP01", "증상1", "원인1", "조치1", "베어링"],
            [20230513, "PMP01", "증상2", "원인2", "조치2", "베어링"],
            [" 20230514 ", "PMP01", "증상3", "원인3", "조치3", "베어링"],
        ])
        self.assertEqual(got[2]["일자"], "2023-05-12")
        self.assertEqual(got[3]["일자"], "2023-05-13")
        self.assertEqual(got[4]["일자"], "2023-05-14")
        for row in (2, 3, 4):
            with self.subTest(행=row):
                self.assertIsNone(got[row]["확인 필요 이유"])

    def test_impossible_or_wrong_length_stays_unreadable(self):
        got = self.run_rows([
            ["20231332", "PMP01", "증상1", "원인1", "조치1", "베어링"],  # 13월 32일
            ["2023051", "PMP01", "증상2", "원인2", "조치2", "베어링"],    # 7자리
            ["202305120", "PMP01", "증상3", "원인3", "조치3", "베어링"],  # 9자리
        ])
        for row in (2, 3, 4):
            with self.subTest(행=row):
                self.assertEqual(got[row]["일자"], "확인 필요")
                self.assertIn("일자 해석 불가", got[row]["확인 필요 이유"])


if __name__ == "__main__":
    unittest.main()

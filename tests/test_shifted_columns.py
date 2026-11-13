"""T058: 일자 칸에 날짜와 다른 글자가 같이 있으면 '칸 밀림 의심' (② 값 오류).

수기 작성 중 날짜 칸에 에러코드까지 적으면 뒤 칸이 한 칸씩 앞당겨져 증상·원인·조치가 엉뚱한 값이 된다.
자동으로 되돌리지 않고 일자를 '확인 필요'로 두며 이유 `칸 밀림 의심`과 원래 값을 남긴다.
날짜 뒤의 시각(예: `14:00`)은 칸 밀림이 아니다.
"""

import unittest

from tests.support import CONFIG, CliCase, by_source_row, read_sheet, result_files, run_cli, write_xlsx

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분"]


class ShiftedColumnTest(CliCase):
    def run_rows(self, rows):
        path = write_xlsx(self.tmp / "shifted.xlsx", HEADER_B, rows)
        result = run_cli([path], CONFIG, self.out)
        self.assertIn(result.returncode, (0, 1), result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        return by_source_row(records)

    def test_date_followed_by_other_text_is_flagged_as_shifted(self):
        got = self.run_rows([
            ["2023/05/12 E04", "PMP01", "원인이 들어감", "조치가 들어감", "분류가 들어감", None],
            ["20230512 E04", "PMP01", "증상2", "원인2", "조치2", "베어링"],
            ["2023-05-12 담당 김철수", "PMP01", "증상3", "원인3", "조치3", "베어링"],
        ])
        for row, original in ((2, "2023/05/12 E04"), (3, "20230512 E04"), (4, "2023-05-12 담당 김철수")):
            with self.subTest(행=row):
                self.assertEqual(got[row]["일자"], "확인 필요")
                self.assertIn("칸 밀림 의심", got[row]["확인 필요 이유"])
                self.assertNotIn("일자 해석 불가", got[row]["확인 필요 이유"])
                self.assertIn("일자: " + original, got[row]["원래 값"])

    def test_time_after_date_is_not_shifted(self):
        got = self.run_rows([
            ["2023.05.12 14:00", "PMP01", "증상1", "원인1", "조치1", "베어링"],
        ])
        self.assertEqual(got[2]["일자"], "2023-05-12")
        self.assertIsNone(got[2]["확인 필요 이유"])

    def test_text_that_is_not_a_date_is_still_unreadable(self):
        got = self.run_rows([
            ["미상", "PMP01", "증상1", "원인1", "조치1", "베어링"],
            ["E04 2023/05/12", "PMP01", "증상2", "원인2", "조치2", "베어링"],
        ])
        for row in (2, 3):
            with self.subTest(행=row):
                self.assertEqual(got[row]["일자"], "확인 필요")
                self.assertIn("일자 해석 불가", got[row]["확인 필요 이유"])
                self.assertNotIn("칸 밀림 의심", got[row]["확인 필요 이유"])


if __name__ == "__main__":
    unittest.main()

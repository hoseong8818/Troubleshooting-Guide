"""T059: 원인이 '모름'처럼 정보가 없는 단어면 '확인 필요' (③ 정보 부족).

단어 목록은 conversion_map.csv의 종류 `원인 미상` 줄로 관리한다(표준 값 칸은 비움).
앞뒤 공백·대소문자만 정리한 뒤 정확히 같을 때만 해당한다. 설정 사본에 줄을 덧붙여 쓴다.
"""

import unittest

from tests.support import CONFIG, CliCase, by_source_row, make_config, read_sheet, result_files, run_cli, write_xlsx

HEADER_A = ["설비코드", "증상", "원인", "조치내용", "부품분류"]
WORDS = ["원인 미상,모름,", "원인 미상,원인불명,", "원인 미상,?,", "원인 미상,Unknown,"]


class UnknownCauseTest(CliCase):
    def run_rows(self, rows, config):
        path = write_xlsx(self.tmp / "cause.xlsx", HEADER_A, rows)
        result = run_cli([path], config, self.out)
        self.assertIn(result.returncode, (0, 1), result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        return by_source_row(records)

    def test_listed_words_are_flagged(self):
        config = make_config(self.tmp, conversion_extra=WORDS)
        got = self.run_rows([
            ["PMP01", "증상1", "모름", "조치1", "베어링"],
            ["PMP01", "증상2", " 원인불명 ", "조치2", "베어링"],
            ["PMP01", "증상3", "?", "조치3", "베어링"],
            ["PMP01", "증상4", "unknown", "조치4", "베어링"],
        ], config)
        for row, original in ((2, "모름"), (3, "원인불명"), (4, "?"), (5, "unknown")):
            with self.subTest(행=row):
                self.assertEqual(got[row]["원인"], "확인 필요")
                self.assertIn("원인 미상", got[row]["확인 필요 이유"])
                self.assertIn("원인: " + original, got[row]["원래 값"])

    def test_only_exact_match_counts(self):
        config = make_config(self.tmp, conversion_extra=WORDS)
        got = self.run_rows([
            ["PMP01", "증상1", "모름 (베어링 추정)", "조치1", "베어링"],
            ["PMP01", "증상2", "윤활 부족", "조치2", "베어링"],
        ], config)
        self.assertEqual(got[2]["원인"], "모름 (베어링 추정)")
        self.assertIsNone(got[2]["확인 필요 이유"])
        self.assertEqual(got[3]["원인"], "윤활 부족")

    def test_config_without_words_still_works(self):
        got = self.run_rows([["PMP01", "증상1", "모름", "조치1", "베어링"]], CONFIG)
        self.assertEqual(got[2]["원인"], "모름")
        self.assertIsNone(got[2]["확인 필요 이유"])


if __name__ == "__main__":
    unittest.main()

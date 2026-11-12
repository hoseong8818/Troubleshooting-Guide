"""T014: 입력 행 수와 출력 행 수가 같다 (FR-006, FR-014, SC-002)."""

import unittest

from tests.support import FAKE, CliCase, audit_files, input_rows


class RowCountTest(CliCase):
    def test_output_rows_equal_input_rows(self):
        expected_count = len(input_rows(FAKE / "A_basic.xlsx"))
        _, records = self.run_basic()
        self.assertEqual(len(records), expected_count)

    def test_audit_reports_row_counts(self):
        expected_count = len(input_rows(FAKE / "A_basic.xlsx"))
        self.run_basic()
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1, "감사 결과 파일이 하나가 아님")
        text = audits[0].read_text(encoding="utf-8")
        # 실행 시각 같은 다른 숫자에 걸리지 않도록 '입력', '출력'이 든 줄에서 단어 경계의 숫자를 찾는다.
        for word in ("입력", "출력"):
            with self.subTest(항목=word):
                self.assertRegex(text, r"%s[^\n]*\b%d\b" % (word, expected_count),
                                 "감사 결과에 %s 행 수가 없음" % word)


if __name__ == "__main__":
    unittest.main()

"""T042a: 컬럼 수가 같은 양식이 둘 이상 맞으면 '양식 모호'로 보고하고 그 파일을 건너뛴다 (FR-017)."""

import shutil
import unittest

from tests.support import (
    CONFIG, FAKE, CliCase, audit_files, make_config, run_cli, write_xlsx,
)

HEADER = ["일시", "호기", "현상", "이유", "대처", "구분"]
ROWS = [["2026-11-15", "PMP01", "현상1", "이유1", "대처1", "베어링"]]
TIE_LINES = [
    "X,일시,일자,,", "X,호기,설비,,", "X,현상,증상,,", "X,이유,원인,,", "X,대처,조치,,", "X,구분,분류,,",
    "Y,일시,일자,,", "Y,호기,설비,,", "Y,현상,증상,,", "Y,이유,원인,,", "Y,대처,조치,,", "Y,구분,분류,,",
]


class FormAmbiguousTest(CliCase):
    def _folder(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        write_xlsx(folder / "b_tie.xlsx", HEADER, ROWS)
        return folder

    def test_tie_is_reported_and_skipped(self):
        config = make_config(self.tmp, column_map_extra=TIE_LINES)
        result = run_cli([self._folder()], config, self.out)
        self.assertEqual(result.returncode, 1)
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1)
        text = audits[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^양식 모호 시트: file_02 sheet_\d\d$")
        self.assertRegex(text, r"(?m)^입력 행 수: 11$", "모호한 파일의 행이 입력 행 수에 들어감")

    def test_tied_file_rows_are_not_in_the_output(self):
        from tests.support import read_sheet, result_files
        config = make_config(self.tmp, column_map_extra=TIE_LINES)
        run_cli([self._folder()], config, self.out)
        files = result_files(self.out)
        self.assertEqual(len(files), 1)
        _, records = read_sheet(files[0])
        self.assertEqual({r["원본 파일"] for r in records}, {"a_good.xlsx"})

    def test_without_tie_the_same_file_would_be_unmatched_not_ambiguous(self):
        # 대조: 동점 줄이 없으면 '양식 모호'가 아니라 '대응 규칙 없음'이다
        result = run_cli([self._folder()], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertNotIn("양식 모호", text)
        self.assertIn("대응 규칙 없음 시트", text)


if __name__ == "__main__":
    unittest.main()

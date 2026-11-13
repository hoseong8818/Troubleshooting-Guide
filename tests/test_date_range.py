"""T057: 해석한 날짜가 범위 밖이면 '확인 필요' (② 값 오류, FR-011).

2922년처럼 형식은 맞지만 있을 수 없는 날짜를 그대로 두면 최신 사례처럼 보인다.
행은 빼지 않고 일자를 '확인 필요'로 두며 이유 `일자 범위 밖`과 원래 값을 남긴다.
범위는 `--min-date`(기본 2000-01-01)와 `--max-date`(기본 실행한 날)로 받는다.
"""

import os
import subprocess
import sys
import unittest

from tests.support import CONFIG, ROOT, SRC, CliCase, by_source_row, read_sheet, result_files, write_xlsx

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분"]


def run_tool(inputs, out, *extra):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [sys.executable, "-m", "trouble_merge", "--input", *[str(p) for p in inputs],
           "--config", str(CONFIG), "--out", str(out), *extra]
    return subprocess.run(cmd, cwd=str(ROOT), env=env, capture_output=True,
                          text=True, encoding="utf-8", errors="replace", timeout=120)


ROWS = [
    ["2922-01-29", "PMP01", "증상1", "원인1", "조치1", "베어링"],  # 미래(900년 뒤)
    ["1950-03-01", "PMP01", "증상2", "원인2", "조치2", "베어링"],  # 기본 하한 이전
    ["2023-05-12", "PMP01", "증상3", "원인3", "조치3", "베어링"],  # 정상
    ["2000-01-01", "PMP01", "증상4", "원인4", "조치4", "베어링"],  # 하한 그날(포함)
]


class DateRangeTest(CliCase):
    def run_rows(self, *extra):
        path = write_xlsx(self.tmp / "range.xlsx", HEADER_B, ROWS)
        result = run_tool([path], self.out, *extra)
        self.assertIn(result.returncode, (0, 1), result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        return by_source_row(records)

    def test_default_range_marks_future_and_too_old_dates(self):
        got = self.run_rows()
        for row, original in ((2, "2922-01-29"), (3, "1950-03-01")):
            with self.subTest(행=row):
                self.assertEqual(got[row]["일자"], "확인 필요")
                self.assertIn("일자 범위 밖", got[row]["확인 필요 이유"])
                self.assertIn("일자: " + original, got[row]["원래 값"])
        self.assertEqual(got[4]["일자"], "2023-05-12")
        self.assertIsNone(got[4]["확인 필요 이유"])
        self.assertEqual(got[5]["일자"], "2000-01-01")
        self.assertIsNone(got[5]["확인 필요 이유"])

    def test_row_count_is_kept(self):
        got = self.run_rows()
        self.assertEqual(len(got), len(ROWS))

    def test_options_change_the_range(self):
        got = self.run_rows("--min-date", "1940-01-01", "--max-date", "2023-01-01")
        self.assertEqual(got[3]["일자"], "1950-03-01")
        self.assertIsNone(got[3]["확인 필요 이유"])
        self.assertEqual(got[4]["일자"], "확인 필요")
        self.assertIn("일자 범위 밖", got[4]["확인 필요 이유"])

    def test_bad_option_value_stops_without_traceback(self):
        path = write_xlsx(self.tmp / "range.xlsx", HEADER_B, ROWS)
        result = run_tool([path], self.out, "--min-date", "abc")
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [])
        self.assertNotIn("Traceback", result.stderr)


if __name__ == "__main__":
    unittest.main()

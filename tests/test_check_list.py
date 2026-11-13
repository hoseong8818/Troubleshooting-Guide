"""T060: 감사 결과에 확인 필요 이유별 건수, 출력 폴더에 확인 필요 목록 CSV.

목록은 현장에 돌려주어 원본을 고치게 하는 용도다. 파일·시트는 감사 결과와 같은 마스킹 규칙을 따르고
셀 값은 넣지 않는다(헌법: 로그·출력 텍스트에 셀 값 금지).
"""

import csv
import unittest
from pathlib import Path

from tests.support import CONFIG, CliCase, audit_files, run_cli, write_xlsx

HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분"]
ROWS = [
    ["2023-05-12", None, "LEAK-증상1", "LEAK-원인1", "LEAK-조치1", "베어링"],      # 2행: 설비 코드 빈 칸
    ["2023-05-12", "ZZZ99", "LEAK-증상2", "LEAK-원인2", "LEAK-조치2", "베어링"],   # 3행: 설비 코드 미등록
    ["LEAK-미상", "PMP01", "LEAK-증상3", "LEAK-원인3", "LEAK-조치3", "베어링"],    # 4행: 일자 해석 불가
    ["2023-05-13", "PMP01", "LEAK-증상4", "LEAK-원인4", "LEAK-조치4", "베어링"],   # 5행: 정상
    ["LEAK-미상", None, "LEAK-증상5", "LEAK-원인5", "LEAK-조치5", "베어링"],       # 6행: 이유 두 개
]


def check_lists(out):
    return sorted(Path(out).glob("check_list_*.csv"))


class CheckListTest(CliCase):
    def setUp(self):
        super().setUp()
        path = write_xlsx(self.tmp / "LEAK-파일.xlsx", HEADER_B, ROWS)
        self.result = run_cli([path], CONFIG, self.out)
        self.assertIn(self.result.returncode, (0, 1), self.result.stderr[-300:])

    def test_audit_counts_each_reason(self):
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^확인 필요 이유별 건수: 설비 코드 빈 칸 2$")
        self.assertRegex(text, r"(?m)^확인 필요 이유별 건수: 설비 코드 미등록 1$")
        self.assertRegex(text, r"(?m)^확인 필요 이유별 건수: 일자 해석 불가 2$")
        self.assertRegex(text, r"(?m)^확인 필요 건수: 4$")

    def test_check_list_has_one_line_per_flagged_row_without_cell_values(self):
        files = check_lists(self.out)
        self.assertEqual(len(files), 1, "check_list_*.csv가 하나가 아님")
        raw = files[0].read_bytes()
        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"), "UTF-8 BOM이 아님")
        with open(files[0], encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        self.assertEqual(rows[0], ["파일", "시트", "행", "이유"])
        body = rows[1:]
        self.assertEqual([r[2] for r in body], ["2", "3", "4", "6"])
        self.assertEqual({r[0] for r in body}, {"file_01"})
        self.assertEqual({r[1] for r in body}, {"sheet_01"})
        reasons = {r[2]: r[3] for r in body}
        self.assertEqual(reasons["2"], "설비 코드 빈 칸")
        self.assertEqual(reasons["3"], "설비 코드 미등록")
        self.assertEqual(reasons["4"], "일자 해석 불가")
        self.assertIn("설비 코드 빈 칸", reasons["6"])
        self.assertIn("일자 해석 불가", reasons["6"])
        self.assertNotIn("LEAK", raw.decode("utf-8-sig"))

    def test_existing_check_list_is_not_overwritten(self):
        self.assertEqual(len(check_lists(self.out)), 1, "check_list_*.csv가 하나가 아님")
        first = check_lists(self.out)[0]
        before = first.read_bytes()
        path = write_xlsx(self.tmp / "again.xlsx", HEADER_B, ROWS[:1])
        run_cli([path], CONFIG, self.out)
        self.assertEqual(first.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()

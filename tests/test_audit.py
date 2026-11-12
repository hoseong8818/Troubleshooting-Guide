"""T031: 감사 결과가 규칙대로 나온다 (FR-008, US3, SC-008).

감사 결과의 줄 형식은 이 테스트가 정한다(사람이 검토해 확정).
- 파일명·시트명은 기본이 마스킹(`file_01`, `sheet_01`)이다(NFR-009). 번호는 입력 파일을 정렬한 순서다.
- 컬럼은 컬럼 대응 표에 등록된 이름만 그대로 쓰고, 나머지는 `열 2(B)`처럼 위치로 쓴다.
- 셀 값은 나오지 않는다.
"""

import shutil
import unittest

from tests.support import (
    CONFIG, FAKE, STD_HEADER_A, CliCase, audit_files, run_cli, write_xlsx,
)


class AuditTest(CliCase):
    def audit_text(self):
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1, "감사 결과 파일이 하나가 아님")
        return audits[0].read_text(encoding="utf-8")

    def test_counts_and_match_for_a_basic(self):
        self.run_basic()
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^입력 행 수: 11$")
        self.assertRegex(text, r"(?m)^출력 행 수: 11$")
        self.assertRegex(text, r"(?m)^행 수 일치: 예$")
        self.assertRegex(text, r"(?m)^확인 필요 건수: 0$")

    def test_unregistered_column_is_reported_by_position_not_name(self):
        # A_basic의 2번째 열(B)은 에러코드. 대응 표에 없는 이름이라 위치로만 나온다.
        self.run_basic()
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^대응 규칙 없는 컬럼: file_01 sheet_01 열 2\(B\)$")
        self.assertNotIn("에러코드", text)

    def test_trap_check_count_blank_required_values_and_hidden_sheets(self):
        self.run_ok([FAKE / "A_trap.xlsx"])
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^확인 필요 건수: 6$")
        # 필수값이 빈 칸: 행 번호와 (등록된) 컬럼명
        for row, column in ((2, "설비코드"), (5, "증상"), (5, "조치내용"), (6, "원인"), (7, "설비코드")):
            with self.subTest(행=row, 컬럼=column):
                self.assertRegex(
                    text, r"(?m)^필수값 빈 칸: file_01 sheet_01 행 %d 컬럼 %s$" % (row, column))
        # hidden 1개 + veryHidden 1개
        self.assertRegex(text, r"(?m)^숨긴 시트 수: 2$")

    def test_failed_file_is_reported_by_masked_name(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        (folder / "b_broken.xlsx").write_bytes(b"not a zip file")
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^실패한 파일: .*file_02")
        self.assertNotIn("b_broken", text)
        self.assertRegex(text, r"(?m)^입력 행 수: 11$")

    def test_formulas_without_saved_value_are_counted(self):
        # openpyxl이 쓴 수식 셀에는 저장된 계산값이 없다. 빈 칸으로 보고 건수를 감사 결과에 낸다(FR-014).
        rows = [
            ["PMP01", "=1+1", "원인1", "조치1", "베어링"],
            ["PMP01", "=SUM(1,2)", "원인2", "조치2", "베어링"],
            ["PMP01", "정상", "원인3", "조치3", "베어링"],
        ]
        path = write_xlsx(self.tmp / "formula.xlsx", STD_HEADER_A, rows)
        result = run_cli([path], CONFIG, self.out)
        self.assertEqual(result.returncode, 0)
        self.assertRegex(self.audit_text(), r"(?m)^수식 값 없음 건수: 2$")


if __name__ == "__main__":
    unittest.main()

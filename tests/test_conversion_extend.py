"""T026: 변환 표에 한 줄을 더하면 해당 행이 다음 실행부터 자동 처리된다 (FR-004, SC-006).

임시 사본의 변환 표만 고친다. 원본 config/는 건드리지 않는다.
"""

import unittest

from tests.support import FAKE, CliCase, by_source_row, make_config


class ConversionExtendTest(CliCase):
    def _check_count(self, records):
        return sum(1 for r in records if r["확인 필요 이유"] is not None)

    def test_adding_lines_removes_the_check(self):
        trap = [FAKE / "A_trap.xlsx"]

        # 1회차: 기본 변환 표. A_trap의 6행이 모두 확인 필요다.
        _, first = self.run_ok(trap, make_config(self.tmp / "c1"))
        rows1 = by_source_row(first)
        self.assertEqual(self._check_count(first), 6)
        self.assertIn("설비 코드 미등록", rows1[3]["확인 필요 이유"])
        self.assertIn("분류 미등록", rows1[4]["확인 필요 이유"])

        # 2회차: 미등록 설비 한 줄과 미등록 분류 한 줄을 더한다(키는 대소문자·공백이 달라도 된다).
        config2 = make_config(
            self.tmp / "c2",
            conversion_extra=["설비,  zz99 ,ZZ99", "분류,LEAK-분류,밸브"],
        )
        self.out = self.tmp / "out2"
        _, second = self.run_ok(trap, config2)
        rows2 = by_source_row(second)

        self.assertEqual(rows2[3]["설비"], "ZZ99")
        self.assertIsNone(rows2[3]["확인 필요 이유"], "설비를 등록했는데도 확인 필요가 남음")
        self.assertEqual(rows2[4]["분류"], "밸브")
        self.assertIsNone(rows2[4]["확인 필요 이유"], "분류를 등록했는데도 확인 필요가 남음")
        self.assertEqual(self._check_count(second), 4)
        self.assertEqual(len(second), len(first), "행 수가 달라짐")


if __name__ == "__main__":
    unittest.main()

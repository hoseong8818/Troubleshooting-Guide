"""T017b: 병합 셀은 범위의 첫 칸 값을 범위 전체에 채워 읽는다 (FR-014).

양식 A 가짜 데이터에는 병합 셀이 없으므로 임시 입력을 만들어 시험한다.
"""

import unittest

from tests.support import STD_HEADER_A, CliCase, by_source_row, make_config, write_xlsx


class MergedCellsTest(CliCase):
    def test_merged_value_fills_the_whole_range(self):
        rows = [
            ["PMP01", "증상1", "원인1", "조치1", "베어링"],
            ["CNV01", "증상2", "원인2", "조치2", "모터"],
            [None, "증상3", "원인3", "조치3", "모터"],  # A3:A4 병합 -> 3,4행이 모두 CNV01
            ["PRS04", "증상4", "원인4", "조치4", "밸브"],
        ]
        path = write_xlsx(self.tmp / "merged.xlsx", STD_HEADER_A, rows, merges=["A3:A4"])
        config = make_config(self.tmp)
        _, records = self.run_ok([path], config)
        got = by_source_row(records)
        self.assertEqual(len(records), 4)
        self.assertEqual(got[3]["설비"], "CNV01")
        self.assertEqual(got[4]["설비"], "CNV01")
        self.assertEqual(got[5]["설비"], "PRS04")
        self.assertIsNone(got[4]["확인 필요 이유"], "병합으로 채운 칸이 확인 필요가 됨")


if __name__ == "__main__":
    unittest.main()

"""T025: 판단이 안 되는 행은 지우지 않고 '확인 필요'로 남긴다 (FR-005, FR-013, FR-015, SC-005).

- A_trap: 설비 코드 빈 칸·미등록, 분류 미등록, 필수값 빈 행, 사유가 여러 개인 행, 숨긴 시트.
  결과를 expected/A_trap_expected.xlsx와 한 행씩 비교한다.
- 날짜 해석 불가는 양식 A에 일자 컬럼이 없어서 임시 입력(양식 T)으로 시험한다.
"""

import unittest

from tests.support import (
    EXPECTED, FAKE, CliCase, by_source_row, make_config, read_sheet, string_cells, write_xlsx,
)

CHECK = "확인 필요"

DATE_HEADER = ["설비코드", "발생일자", "증상", "원인", "조치내용", "부품분류"]
DATE_COLUMN_MAP_EXTRA = [
    "T,설비코드,설비,,",
    "T,발생일자,일자,,",
    "T,증상,증상,,",
    "T,원인,원인,,",
    "T,조치내용,조치,,",
    "T,부품분류,분류,,",
]


class TrapCompareTest(CliCase):
    def test_a_trap_matches_expected_row_by_row(self):
        header, records = self.run_ok([FAKE / "A_trap.xlsx"])
        exp_header, exp_records = read_sheet(EXPECTED / "A_trap_expected.xlsx")

        self.assertEqual(header, exp_header, "컬럼 이름·순서가 정답과 다름")
        self.assertEqual(len(records), len(exp_records), "출력 행 수가 정답과 다름")
        for index, (got, exp) in enumerate(zip(records, exp_records), start=1):
            for column in exp_header:
                with self.subTest(출력행=index, 컬럼=column):
                    self.assertEqual(got.get(column), exp.get(column))

    def test_no_row_is_lost_and_hidden_sheets_are_not_read(self):
        _, records = self.run_ok([FAKE / "A_trap.xlsx"])
        self.assertEqual(len(records), 6, "보이는 시트의 6행이 모두 있어야 함")
        self.assertEqual({r["원본 시트"] for r in records}, {"시트1"})
        self.assertNotIn("LEAK-숨김", string_cells(self.result_path))

    def test_every_check_row_has_reason_and_original_value(self):
        _, records = self.run_ok([FAKE / "A_trap.xlsx"])
        for record in records:
            with self.subTest(원본행=record["원본 행"]):
                self.assertIsNotNone(record["확인 필요 이유"])
                self.assertIsNotNone(record["원래 값"])

    def test_check_equipment_rows_are_at_the_bottom(self):
        _, records = self.run_ok([FAKE / "A_trap.xlsx"])
        flags = [r["설비"] == CHECK for r in records]
        self.assertEqual(flags, sorted(flags), "설비가 확인 필요인 행이 맨 아래가 아님")

    def test_check_equipment_rows_are_not_marked_duplicate(self):
        # 원본 2행과 7행은 증상이 같지만 설비가 확인 필요라 비교에서 빠진다(FR-010)
        _, records = self.run_ok([FAKE / "A_trap.xlsx"])
        rows = by_source_row(records)
        self.assertIsNone(rows[2]["중복 의심"])
        self.assertIsNone(rows[7]["중복 의심"])


class DateCheckTest(CliCase):
    def _run(self, rows):
        path = write_xlsx(self.tmp / "dates.xlsx", DATE_HEADER, rows)
        config = make_config(self.tmp, column_map_extra=DATE_COLUMN_MAP_EXTRA)
        _, records = self.run_ok([path], config)
        return by_source_row(records)

    def test_unreadable_dates_become_check(self):
        got = self._run([
            ["PMP01", "11/07/2026", "d1", "원인", "조치", "베어링"],  # MM/DD/YYYY는 읽지 않는다
            ["PMP01", "26.11.07", "d2", "원인", "조치", "베어링"],    # 두 자리 연도는 읽지 않는다
            ["PMP01", "abc", "d3", "원인", "조치", "베어링"],
        ])
        for number, original in ((2, "11/07/2026"), (3, "26.11.07"), (4, "abc")):
            with self.subTest(원본행=number):
                self.assertEqual(got[number]["일자"], CHECK)
                self.assertEqual(got[number]["확인 필요 이유"], "일자 해석 불가")
                self.assertEqual(got[number]["원래 값"], "일자: " + original)

    def test_empty_date_is_not_check_and_readable_date_is_kept(self):
        got = self._run([
            ["PMP01", None, "d4", "원인", "조치", "베어링"],
            ["PMP01", "2026-11-07", "d5", "원인", "조치", "베어링"],
        ])
        self.assertIsNone(got[2]["일자"])
        self.assertIsNone(got[2]["확인 필요 이유"], "빈 일자는 확인 필요가 아님(FR-013)")
        self.assertEqual(str(got[3]["일자"]), "2026-11-07")
        self.assertIsNone(got[3]["확인 필요 이유"])

    def test_multiple_reasons_follow_column_order(self):
        got = self._run([
            ["ZZ99", "abc", "d6", "원인", "조치", "베어링"],
        ])
        self.assertEqual(got[2]["확인 필요 이유"], "설비 코드 미등록; 일자 해석 불가")
        self.assertEqual(got[2]["원래 값"], "설비: ZZ99; 일자: abc")

    def test_check_dates_are_excluded_from_duplicates(self):
        got = self._run([
            ["PMP01", "11/07/2026", "dup", "원인", "조치", "베어링"],
            ["PMP01", "11/07/2026", "dup", "원인", "조치", "베어링"],
            ["PMP01", "2026-11-07", "dup2", "원인", "조치", "베어링"],
            ["PMP01", "2026-11-07", "dup2", "원인", "조치", "베어링"],
        ])
        self.assertIsNone(got[2]["중복 의심"], "일자가 확인 필요인 행이 중복 비교에 들어감")
        self.assertIsNone(got[3]["중복 의심"])
        self.assertEqual(got[4]["중복 의심"], "중복 의심")
        self.assertEqual(got[5]["중복 의심"], "중복 의심")


if __name__ == "__main__":
    unittest.main()

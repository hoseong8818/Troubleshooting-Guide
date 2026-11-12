"""T017a: 설비 코드와 변환 표 키의 공백·대소문자 정리 (FR-009, FR-018)."""

import unittest

from tests.support import (
    FAKE, STD_HEADER_A, CliCase, by_source_row, make_config, result_files,
    run_cli, write_xlsx,
)


class NormalizeTest(CliCase):
    def test_equipment_codes_are_trimmed_and_case_folded(self):
        _, records = self.run_basic()
        rows = by_source_row(records)
        # 원본 3행 ' PMP01 '(공백), 4행 'cnv01'(소문자), 9행 'pmp01 '(소문자+공백)
        self.assertEqual(rows[3]["설비"], "PMP01")
        self.assertEqual(rows[4]["설비"], "CNV01")
        self.assertEqual(rows[9]["설비"], "PMP01")

    def test_category_keys_ignore_case_and_spaces(self):
        # 원본 9행의 분류 ' Bearing'과 7행의 'SENSOR'가 표준 이름으로 바뀐다
        _, records = self.run_basic()
        rows = by_source_row(records)
        self.assertEqual(rows[9]["분류"], "베어링")
        self.assertEqual(rows[7]["분류"], "센서")
        self.assertEqual(rows[4]["분류"], "모터")

    def test_conversion_key_in_table_is_normalized_too(self):
        # 변환 표 키가 '  Foo  '(공백·대소문자)여도 입력 'FOO'와 일치한다
        path = write_xlsx(
            self.tmp / "keys.xlsx", STD_HEADER_A,
            [["PMP01", "증상", "원인", "조치", "FOO"]],
        )
        config = make_config(self.tmp, conversion_extra=["분류,  Foo  ,모터"])
        _, records = self.run_ok([path], config)
        self.assertEqual(records[0]["분류"], "모터")

    def test_conflicting_conversion_keys_stop_the_run(self):
        # '모타'가 이미 '모터'로 등록돼 있는데 '센서'로 또 등록하면 실행을 멈춘다
        config = make_config(self.tmp, conversion_extra=["분류,모타,센서"])
        result = run_cli([FAKE / "A_basic.xlsx"], config, self.out)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [], "충돌인데 통합 엑셀이 만들어짐")


if __name__ == "__main__":
    unittest.main()

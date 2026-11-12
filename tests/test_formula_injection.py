"""T017d: `=`, `+`, `-`, `@`로 시작하는 입력 값이 출력에서 수식 셀이 되지 않고 `'`가 붙는다 (NFR-008, SC-009)."""

import unittest

from tests.support import CliCase, by_source_row, formula_cells, string_cells

DANGEROUS_PREFIXES = ("=", "+", "-", "@")


class FormulaInjectionTest(CliCase):
    def test_no_formula_cells_in_output(self):
        self.run_basic()
        self.assertEqual(formula_cells(self.result_path), [], "수식으로 저장된 셀이 있음")

    def test_no_string_cell_starts_with_dangerous_character(self):
        self.run_basic()
        bad = [v for v in string_cells(self.result_path) if v.startswith(DANGEROUS_PREFIXES)]
        self.assertEqual(bad, [], "접두사 없이 위험 문자로 시작하는 문자열 셀이 있음")

    def test_prefixed_values_for_each_character(self):
        _, records = self.run_basic()
        rows = by_source_row(records)
        self.assertEqual(rows[10]["증상"], "'=1+1")
        self.assertEqual(rows[11]["원인"], "'+1 과부하")
        self.assertEqual(rows[12]["조치"], "'-2 분해 청소")
        self.assertEqual(rows[13]["조치"], "'@퓨즈 교체")


if __name__ == "__main__":
    unittest.main()

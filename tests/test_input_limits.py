"""T017e: .xlsm·.xls는 건너뛰고 나머지는 처리한다 (NFR-007, SC-010).

크기·행 수 상한 시험은 아직 없다. 상한 숫자와 설정 파일 이름을 사람이 정한 뒤(T003b) 추가한다.
"""

import shutil
import unittest

from tests.support import (
    CONFIG, FAKE, CliCase, audit_files, input_rows, read_sheet, result_files, run_cli,
)


class InputLimitsTest(CliCase):
    def _prepare_folder(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "good.xlsx")
        shutil.copy(FAKE / "A_basic.xlsx", folder / "macro.xlsm")  # 형식은 xlsx지만 확장자가 .xlsm
        (folder / "old.xls").write_bytes(b"\xd0\xcf\x11\xe0 not a real xls")
        return folder

    def test_xlsm_and_xls_are_skipped_and_the_rest_is_processed(self):
        folder = self._prepare_folder()
        result = run_cli([folder], CONFIG, self.out)
        self.assertIn(result.returncode, (0, 1), "종료 코드가 0이나 1이 아님")
        files = result_files(self.out)
        self.assertEqual(len(files), 1, "통합 엑셀이 하나가 아님")
        _, records = read_sheet(files[0])
        expected = len(input_rows(FAKE / "A_basic.xlsx"))
        self.assertEqual(len(records), expected, ".xlsm 행이 섞였거나 good.xlsx가 처리되지 않음")

    def test_audit_is_written(self):
        folder = self._prepare_folder()
        run_cli([folder], CONFIG, self.out)
        self.assertEqual(len(audit_files(self.out)), 1, "감사 결과 파일이 하나가 아님")


if __name__ == "__main__":
    unittest.main()

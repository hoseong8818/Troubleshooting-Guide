"""T043: 열 수 없는 파일은 건너뛰고 나머지를 처리하며(종료 코드 1), 전부 실패하면 통합 파일 없이 멈춘다(종료 코드 2) (FR-019, 스토리 1·3)."""

import shutil
import unittest

from tests.support import CONFIG, FAKE, CliCase, read_sheet, result_files, run_cli


class PartialFailureTest(CliCase):
    def test_broken_file_is_skipped_and_the_rest_is_processed(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        (folder / "b_broken.xlsx").write_bytes(b"not a zip file")
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        files = result_files(self.out)
        self.assertEqual(len(files), 1)
        _, records = read_sheet(files[0])
        self.assertEqual(len(records), 11)

    def test_all_files_failing_stops_without_an_integrated_file(self):
        folder = self.tmp / "in"
        folder.mkdir()
        (folder / "a_broken.xlsx").write_bytes(b"not a zip file")
        (folder / "b_old.xls").write_bytes(b"\xd0\xcf\x11\xe0 not a real xls")
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [], "전부 실패했는데 통합 엑셀이 만들어짐")

    def test_missing_input_path_is_a_failure_not_a_crash(self):
        result = run_cli([self.tmp / "none.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [])
        self.assertNotIn("Traceback", result.stderr)

    def test_good_files_alone_exit_zero(self):
        result = run_cli([FAKE / "A_basic.xlsx", FAKE / "B_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 0)


if __name__ == "__main__":
    unittest.main()

"""전부 실패해도 감사 결과는 남긴다: 통합 엑셀은 없고, 실패한 파일이 감사 결과에 나온다 (US3-6, 실패하는 경우)."""

import unittest

from tests.support import CONFIG, CliCase, audit_files, result_files, run_cli


class AllFailedAuditTest(CliCase):
    def _broken_folder(self):
        folder = self.tmp / "in"
        folder.mkdir()
        (folder / "a_broken.xlsx").write_bytes(b"not a zip file")
        (folder / "b_old.xls").write_bytes(b"\xd0\xcf\x11\xe0 not a real xls")
        return folder

    def test_audit_is_written_and_lists_the_failed_files(self):
        result = run_cli([self._broken_folder()], CONFIG, self.out)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [], "전부 실패했는데 통합 엑셀이 만들어짐")
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1, "전부 실패했는데 감사 결과가 없음")
        text = audits[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^실패한 파일: file_01, file_02$")
        self.assertRegex(text, r"(?m)^읽은 파일 수: 0$")
        self.assertRegex(text, r"(?m)^출력 행 수: 0$")

    def test_name_map_is_written_so_masked_names_can_be_resolved(self):
        run_cli([self._broken_folder()], CONFIG, self.out)
        self.assertEqual(len(list(self.out.glob("names_*.txt"))), 1)

    def test_missing_input_path_is_listed_as_failed(self):
        result = run_cli([self.tmp / "none.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [])
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1)
        self.assertRegex(audits[0].read_text(encoding="utf-8"), r"(?m)^실패한 파일: file_01$")
        self.assertNotIn("Traceback", result.stderr)


if __name__ == "__main__":
    unittest.main()

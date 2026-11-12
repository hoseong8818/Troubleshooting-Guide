"""T043a: 출력 파일이 이미 있으면 덮어쓰지 않고 실행 시각을 붙인 새 이름으로 만든다 (FR-019, 엣지 케이스).

통합 파일 이름은 정하지 않았다. 같은 출력 폴더에 두 번 실행해서 파일 수와 내용으로 확인한다.
"""

import unittest

from tests.support import CONFIG, FAKE, CliCase, audit_files, result_files, run_cli, sha256


class OutputNameTest(CliCase):
    def test_second_run_does_not_overwrite_the_first(self):
        first = run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(first.returncode, 0)
        files = result_files(self.out)
        self.assertEqual(len(files), 1)
        first_path = files[0]
        first_hash = sha256(first_path)

        second = run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(second.returncode, 0)
        files = result_files(self.out)
        self.assertEqual(len(files), 2, "두 번째 실행이 첫 결과를 덮어씀")
        self.assertTrue(first_path.exists())
        self.assertEqual(sha256(first_path), first_hash, "첫 결과가 바뀜")
        self.assertEqual(len({p.name for p in files}), 2)

    def test_audit_files_are_not_overwritten_either(self):
        run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(len(audit_files(self.out)), 2)
        self.assertEqual(len(list(self.out.glob("names_*.txt"))), 2)


if __name__ == "__main__":
    unittest.main()

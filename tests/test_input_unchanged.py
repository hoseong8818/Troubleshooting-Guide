"""T017c: 실행 전후 입력 파일의 해시가 같다 (NFR-003)."""

import unittest

from tests.support import FAKE, CliCase, sha256


class InputUnchangedTest(CliCase):
    def test_input_files_are_not_modified(self):
        inputs = [FAKE / "A_basic.xlsx", FAKE / "A_trap.xlsx"]
        before = {p.name: sha256(p) for p in inputs}

        # 실행이 실제로 성공해야 한다. 실행이 아예 안 돼도 해시가 같아서 통과하는 것을 막는다.
        self.run_ok([FAKE / "A_basic.xlsx"])

        after = {p.name: sha256(p) for p in inputs}
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()

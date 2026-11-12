"""전부 실패하면 감사 결과에 '처리된 파일 없음'을 분명히 적고, 정상 실행에는 적지 않는다 (US3).

전부 실패하면 입력 0행, 출력 0행이라 '행 수 일치: 예'가 나온다. 정상 처리로 오해하지 않게 한 줄을 더 낸다.
"""

import unittest

from tests.support import CONFIG, FAKE, CliCase, audit_files, run_cli

NOTICE = r"(?m)^처리 결과: 처리된 파일 없음$"


class AllFailedNoticeTest(CliCase):
    def test_notice_is_written_when_every_file_failed(self):
        folder = self.tmp / "in"
        folder.mkdir()
        (folder / "a_broken.xlsx").write_bytes(b"not a zip file")
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 2)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, NOTICE)
        self.assertRegex(text, r"(?m)^읽은 파일 수: 0$")

    def test_notice_is_not_written_on_a_normal_run(self):
        result = run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 0)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertNotRegex(text, r"처리된 파일 없음")

    def test_notice_is_not_written_when_only_some_files_failed(self):
        folder = self.tmp / "in"
        folder.mkdir()
        (folder / "a_broken.xlsx").write_bytes(b"not a zip file")
        (folder / "b_good.xlsx").write_bytes((FAKE / "A_basic.xlsx").read_bytes())
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertNotRegex(text, r"처리된 파일 없음")


if __name__ == "__main__":
    unittest.main()

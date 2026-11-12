"""T032: 로그·stderr·감사 결과에 셀 값, 예외 메시지, 등록 안 된 이름이 나오지 않는다 (NFR-002, NFR-009, SC-004, SC-011).

가짜 데이터의 표식 값은 `LEAK-`로 시작한다(A_trap의 증상·분류, 숨긴 시트의 값).
"""

import os
import shutil
import subprocess
import sys
import unittest

from tests.support import (
    CONFIG, FAKE, ROOT, SRC, STD_HEADER_A, CliCase, audit_files, make_config, run_cli,
    write_xlsx,
)

BAD_WORDS = ("zip", "badzipfile", "openpyxl", "traceback", "exception")


def collect(result, out):
    """stdout, stderr, 감사 결과 텍스트를 하나로 합친다. 이름 대응표 파일은 제외한다."""
    text = result.stdout + "\n" + result.stderr
    for audit in audit_files(out):
        text += "\n" + audit.read_text(encoding="utf-8")
    return text


class LogSafeTest(CliCase):
    def test_normal_run_has_no_marker_values(self):
        result = run_cli([FAKE / "A_trap.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 0)
        self.assertNotIn("LEAK-", collect(result, self.out))

    def test_error_run_has_no_marker_values_or_exception_messages(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_trap.xlsx", folder / "a_trap.xlsx")
        (folder / "b_broken.xlsx").write_bytes(b"LEAK-garbage not a zip file")
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        text = collect(result, self.out)
        self.assertNotIn("LEAK-", text)
        for word in BAD_WORDS:
            with self.subTest(단어=word):
                self.assertNotIn(word, text.lower(), "예외 메시지나 내부 이름이 그대로 나옴")

    def test_config_error_message_has_no_cell_or_key_values(self):
        config = make_config(self.tmp, conversion_extra=["분류,모타,센서"])
        result = run_cli([FAKE / "A_basic.xlsx"], config, self.out)
        self.assertEqual(result.returncode, 2)
        for value in ("모타", "센서"):
            self.assertNotIn(value, result.stdout + result.stderr)

    def test_file_sheet_and_unregistered_column_names_are_masked(self):
        result = run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 0)
        text = collect(result, self.out)
        for name in ("A_basic", "시트1", "에러코드"):
            with self.subTest(이름=name):
                self.assertNotIn(name, text)
        self.assertIn("file_01", text)

    def test_names_map_file_is_written_next_to_the_audit(self):
        run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        maps = sorted(self.out.glob("names_*.txt"))
        self.assertEqual(len(maps), 1, "이름 대응표 파일이 하나가 아님")
        content = maps[0].read_text(encoding="utf-8")
        self.assertIn("file_01", content)
        self.assertIn("A_basic.xlsx", content)
        self.assertIn("sheet_01", content)
        self.assertIn("시트1", content)

    def test_show_names_option_prints_real_names(self):
        env = dict(os.environ)
        env["PYTHONPATH"] = str(SRC)
        env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            [sys.executable, "-m", "trouble_merge", "--input", str(FAKE / "A_basic.xlsx"),
             "--config", str(CONFIG), "--out", str(self.out), "--show-names"],
            cwd=str(ROOT), env=env, capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=120,
        )
        self.assertEqual(result.returncode, 0)
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1)
        self.assertIn("A_basic.xlsx", audits[0].read_text(encoding="utf-8"))

    def test_title_row_mistaken_for_header_does_not_leak(self):
        # 1행이 제목 문장이라 헤더로 오인되는 파일. 제목 문장이 로그·감사 결과에 나가면 안 된다.
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        write_xlsx(
            folder / "b_title.xlsx",
            ["LEAK-제목 문장은 컬럼명이 아니다"],
            [STD_HEADER_A, ["PMP01", "증상", "원인", "조치", "베어링"]],
        )
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        text = collect(result, self.out)
        self.assertNotIn("LEAK-", text)
        self.assertRegex(text, r"(?m)^대응 규칙 없음 시트: file_02 sheet_01$")


if __name__ == "__main__":
    unittest.main()

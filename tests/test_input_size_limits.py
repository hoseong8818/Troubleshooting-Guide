"""T017f: 크기·행 수 상한을 넘은 파일은 건너뛰고 감사 결과에 보고, 나머지는 처리한다 (NFR-007, SC-010).

상한은 명령줄 옵션으로 받는다: --max-file-mb(파일 크기), --max-unzipped-mb(압축 해제 후 크기), --max-rows(시트당 행 수).
기본값은 50, 200, 100000이다. 테스트는 임시 폴더에 작은 파일을 만들고 상한을 그보다 작게 줘서 시험한다.
행 수는 헤더를 포함한 시트의 행 수다(끝쪽 빈 서식 행도 센다. 메모리를 막는 것이 목적이다).
감사 결과에는 파일 표시명과 사유 종류만 나오고 셀 값과 실제 파일명은 나오지 않는다.
"""

import os
import subprocess
import sys
import unittest
import zipfile

from tests.support import (
    CONFIG, FAKE, ROOT, SRC, STD_HEADER_A, CliCase, audit_files, read_sheet, result_files, write_xlsx,
)

MB = 1024 * 1024


def run_with_limits(inputs, out, *extra):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [
        sys.executable, "-m", "trouble_merge",
        "--input", *[str(p) for p in inputs],
        "--config", str(CONFIG),
        "--out", str(out),
        *extra,
    ]
    return subprocess.run(
        cmd, cwd=str(ROOT), env=env, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=120,
    )


def rows_for(count, symptom="증상"):
    return [["PMP01", "%s%d" % (symptom, i), "원인", "조치", "베어링"] for i in range(count)]


def unzipped_size(path):
    with zipfile.ZipFile(path) as z:
        return sum(info.file_size for info in z.infolist())


class SizeLimitCase(CliCase):
    def folder_with(self, **files):
        folder = self.tmp / "in"
        folder.mkdir()
        for name, rows in files.items():
            write_xlsx(folder / (name + ".xlsx"), STD_HEADER_A, rows)
        return folder

    def audit_text(self):
        return audit_files(self.out)[0].read_text(encoding="utf-8")


class RowLimitTest(SizeLimitCase):
    def test_file_over_the_row_limit_is_skipped_and_the_rest_is_processed(self):
        folder = self.folder_with(a_small=rows_for(3), b_big=rows_for(30, symptom="LEAK-큰파일"))
        result = run_with_limits([folder], self.out, "--max-rows", "10")
        self.assertEqual(result.returncode, 1, result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual(len(records), 3)
        self.assertEqual({r["원본 파일"] for r in records}, {"a_small.xlsx"})
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^상한 초과 파일: file_02 행 수$")
        self.assertRegex(text, r"(?m)^입력 행 수: 3$", "건너뛴 파일의 행이 입력 행 수에 들어감")

    def test_limit_is_inclusive_at_the_boundary(self):
        # 헤더 1행 + 데이터 3행 = 시트 4행
        folder = self.folder_with(a_small=rows_for(3))
        ok = run_with_limits([folder], self.out, "--max-rows", "4")
        self.assertEqual(ok.returncode, 0, "상한과 같은 행 수는 통과해야 함")
        over = run_with_limits([folder], self.tmp / "out2", "--max-rows", "3")
        self.assertEqual(over.returncode, 2, "상한보다 1행 많으면 건너뛰어야 함")

    def test_nothing_leaks_into_audit_or_output(self):
        folder = self.folder_with(a_small=rows_for(3), b_secret=rows_for(30, symptom="LEAK-비밀"))
        result = run_with_limits([folder], self.out, "--max-rows", "10")
        everything = self.audit_text() + result.stdout + result.stderr
        self.assertNotIn("LEAK-", everything)
        self.assertNotIn("b_secret", everything, "마스킹되지 않은 파일명이 나옴")


class FileSizeLimitTest(SizeLimitCase):
    def test_file_over_the_size_limit_is_skipped(self):
        folder = self.folder_with(a_small=rows_for(3), b_big=rows_for(5000))
        small = (folder / "a_small.xlsx").stat().st_size
        big = (folder / "b_big.xlsx").stat().st_size
        self.assertLess(small * 2, big, "시험 준비: 큰 파일이 충분히 커야 함")
        limit_mb = ((small + big) / 2) / MB
        result = run_with_limits([folder], self.out, "--max-file-mb", repr(limit_mb))
        self.assertEqual(result.returncode, 1, result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual({r["원본 파일"] for r in records}, {"a_small.xlsx"})
        self.assertRegex(self.audit_text(), r"(?m)^상한 초과 파일: file_02 파일 크기$")


class UnzippedLimitTest(SizeLimitCase):
    def test_small_file_that_unzips_large_is_skipped(self):
        bomb_rows = [["PMP01", "%d-" % i + "A" * 30000, "원인", "조치", "베어링"] for i in range(40)]
        folder = self.folder_with(a_small=rows_for(3), b_bomb=bomb_rows)
        small_unzipped = unzipped_size(folder / "a_small.xlsx")
        bomb_zip = (folder / "b_bomb.xlsx").stat().st_size
        bomb_unzipped = unzipped_size(folder / "b_bomb.xlsx")
        self.assertLess(bomb_zip * 5, bomb_unzipped, "시험 준비: 압축률이 충분히 높아야 함")
        limit_mb = ((small_unzipped + bomb_unzipped) / 2) / MB
        result = run_with_limits([folder], self.out, "--max-unzipped-mb", repr(limit_mb))
        self.assertEqual(result.returncode, 1, result.stderr[-300:])
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual({r["원본 파일"] for r in records}, {"a_small.xlsx"})
        self.assertRegex(self.audit_text(), r"(?m)^상한 초과 파일: file_02 압축 해제 크기$")


class AllOverTheLimitTest(SizeLimitCase):
    def test_all_files_over_the_limit_stop_without_an_integrated_file(self):
        folder = self.folder_with(a_big=rows_for(30), b_big=rows_for(40))
        result = run_with_limits([folder], self.out, "--max-rows", "10")
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [], "전부 상한 초과인데 통합 엑셀이 만들어짐")
        text = self.audit_text()
        self.assertRegex(text, r"(?m)^상한 초과 파일: file_01 행 수$")
        self.assertRegex(text, r"(?m)^상한 초과 파일: file_02 행 수$")
        self.assertRegex(text, r"(?m)^처리 결과: 처리된 파일 없음$")


class DefaultLimitTest(SizeLimitCase):
    def test_default_limits_do_not_skip_normal_files(self):
        result = run_with_limits([FAKE / "A_basic.xlsx", FAKE / "B_basic.xlsx"], self.out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        self.assertNotIn("상한 초과", self.audit_text())


if __name__ == "__main__":
    unittest.main()

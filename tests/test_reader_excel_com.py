"""T017h: --open-with excel 옵션은 설치된 엑셀을 조종해 파일을 열고 같은 표준 컬럼으로 통합한다 (NFR-006, NFR-007).

이 PC에 Excel과 pywin32가 있어야 엑셀 경로 테스트가 돈다. 둘 중 하나라도 없으면 그 테스트는 건너뛴다(실패 아님).
엑셀 조종 경로는 파일을 openpyxl로 직접 해석하지 않고 설치된 엑셀이 연 창에서 값을 읽는다. 그래서 회사 DRM처럼
엑셀과 DRM 에이전트만 열 수 있는 파일도 같은 코드로 읽을 수 있다. 이 테스트는 DRM이 없는 가짜 데이터로
엑셀 경로가 기존 openpyxl 경로와 같은 결과를 내는지만 본다. 실제 DRM 파일로 여는 것은 회사 승인 후 사람이 확인한다.
"""

import os
import subprocess
import sys
import unittest

from openpyxl import Workbook

from tests.support import (
    CONFIG, EXPECTED, FAKE, ROOT, SRC, STD_HEADER_A, CliCase, audit_files, by_source_row, make_config,
    read_sheet, result_files, write_xlsx,
)


def run_with_backend(inputs, out, config=CONFIG, backend="excel"):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [
        sys.executable, "-m", "trouble_merge",
        "--input", *[str(p) for p in inputs],
        "--config", str(config),
        "--out", str(out),
        "--open-with", backend,
    ]
    return subprocess.run(
        cmd, cwd=str(ROOT), env=env, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=240,
    )


def excel_is_available():
    try:
        import win32com.client
    except ImportError:
        return False
    try:
        app = win32com.client.DispatchEx("Excel.Application")
    except Exception:
        return False
    try:
        app.Quit()
    except Exception:
        pass
    return True


class ExcelBackendTest(CliCase):
    @classmethod
    def setUpClass(cls):
        if not excel_is_available():
            raise unittest.SkipTest("pywin32 또는 설치된 Excel이 없어 엑셀 경로 테스트를 건너뜀")

    def test_a_basic_matches_the_expected_rows_like_the_default_path(self):
        result = run_with_backend([FAKE / "A_basic.xlsx"], self.out)
        self.assertEqual(result.returncode, 0, result.stderr[-500:])
        header, records = read_sheet(result_files(self.out)[0])
        exp_header, exp_records = read_sheet(EXPECTED / "A_basic_expected.xlsx")
        self.assertEqual(header, exp_header)
        self.assertEqual(len(records), len(exp_records), "출력 행 수가 정답과 다름")
        for index, (got, exp) in enumerate(zip(records, exp_records), start=1):
            for column in exp_header:
                with self.subTest(출력행=index, 컬럼=column):
                    self.assertEqual(got.get(column), exp.get(column))

    def test_merged_value_fills_the_whole_range(self):
        rows = [
            ["PMP01", "증상1", "원인1", "조치1", "베어링"],
            ["CNV01", "증상2", "원인2", "조치2", "모터"],
            [None, "증상3", "원인3", "조치3", "모터"],  # A3:A4 병합 -> 3,4행이 모두 CNV01
            ["PRS04", "증상4", "원인4", "조치4", "밸브"],
        ]
        path = write_xlsx(self.tmp / "merged.xlsx", STD_HEADER_A, rows, merges=["A3:A4"])
        config = make_config(self.tmp)
        result = run_with_backend([path], self.out, config)
        self.assertEqual(result.returncode, 0, result.stderr[-500:])
        _, records = read_sheet(result_files(self.out)[0])
        got = by_source_row(records)
        self.assertEqual(len(records), 4)
        self.assertEqual(got[3]["설비"], "CNV01")
        self.assertEqual(got[4]["설비"], "CNV01")
        self.assertIsNone(got[4]["확인 필요 이유"], "병합으로 채운 칸이 확인 필요가 됨")

    def test_hidden_and_very_hidden_sheets_are_skipped_and_counted(self):
        wb = Workbook()
        visible = wb.active
        visible.title = "시트A"
        visible.append(STD_HEADER_A)
        visible.append(["PMP01", "증상A", "원인A", "조치A", "베어링"])
        hidden = wb.create_sheet("숨김")
        hidden.append(STD_HEADER_A)
        hidden.append(["PMP01", "LEAK-숨김1", "원인", "조치", "베어링"])
        hidden.sheet_state = "hidden"
        very = wb.create_sheet("깊은숨김")
        very.append(STD_HEADER_A)
        very.append(["PMP01", "LEAK-숨김2", "원인", "조치", "베어링"])
        very.sheet_state = "veryHidden"
        path = self.tmp / "hidden.xlsx"
        wb.save(path)

        result = run_with_backend([path], self.out)
        self.assertEqual(result.returncode, 0, result.stderr[-500:])
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual({r["원본 시트"] for r in records}, {"시트A"})
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^숨긴 시트 수: 2$")
        self.assertNotIn("LEAK-", text + result.stdout + result.stderr)

    def test_no_excel_process_is_left_behind_and_input_is_unchanged(self):
        import hashlib
        before = hashlib.sha256((FAKE / "A_basic.xlsx").read_bytes()).hexdigest()
        run_with_backend([FAKE / "A_basic.xlsx"], self.out)
        after = hashlib.sha256((FAKE / "A_basic.xlsx").read_bytes()).hexdigest()
        self.assertEqual(before, after, "엑셀로 열고 닫은 뒤 입력 파일이 바뀜(읽기 전용으로 열어야 함)")


class OpenWithOptionTest(CliCase):
    """엑셀이 없어도 도는 테스트: 옵션 값 검사."""

    def test_unknown_backend_is_rejected_without_a_traceback(self):
        result = run_with_backend([FAKE / "A_basic.xlsx"], self.out, backend="nonsense")
        self.assertEqual(result.returncode, 2)
        self.assertEqual(result_files(self.out), [])
        self.assertNotIn("Traceback", result.stderr)

    def test_default_backend_is_still_openpyxl(self):
        # 옵션을 주지 않으면 지금까지와 같은 경로를 쓴다(엑셀이 없어도 도는 경로)
        from tests.support import run_cli
        result = run_cli([FAKE / "A_basic.xlsx"], CONFIG, self.out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])


if __name__ == "__main__":
    unittest.main()

"""테스트 공용 도구 (Phase 3).

설계 규칙
- 도구는 외부 프로세스(`python -m trouble_merge ...`)로 실행한다. 구현이 아직 없어도
  ImportError가 아니라 '종료 코드·출력 파일 불일치'로 실패한다(T018).
- 출력은 항상 임시 폴더에 쓴다. `out/`을 쓰지 않는다.
- 통합 파일 이름과 시트 이름은 명세에 없다. 출력 폴더의 유일한 .xlsx를 결과로 보고 첫 시트를 읽는다.
- 감사 결과는 출력 폴더의 `audit_*.txt`로 본다.
- 이 파일의 값은 가짜 데이터에서만 온다. 실제 데이터를 쓰지 않는다.
"""

import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "src"
FAKE = ROOT / "data" / "fake"
EXPECTED = ROOT / "expected"
CONFIG = ROOT / "config"

STD_HEADER_A = ["설비코드", "증상", "원인", "조치내용", "부품분류"]


def run_cli(inputs, config, out):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [
        sys.executable, "-m", "trouble_merge",
        "--input", *[str(p) for p in inputs],
        "--config", str(config),
        "--out", str(out),
    ]
    return subprocess.run(
        cmd, cwd=str(ROOT), env=env, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=120,
    )


def result_files(out):
    return sorted(Path(out).glob("*.xlsx"))


def audit_files(out):
    return sorted(Path(out).glob("audit_*.txt"))


def _clean(value):
    return None if value == "" else value


def read_sheet(path):
    """첫 시트를 (헤더 목록, 행 사전 목록)으로 읽는다. 빈 문자열은 None으로 본다."""
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header = [h for h in rows[0]]
    records = [dict(zip(header, [_clean(v) for v in r])) for r in rows[1:]]
    return header, records


def input_rows(path):
    """입력 엑셀의 첫 시트에서 (엑셀 행 번호, 행 사전) 중 빈 행이 아닌 것만 돌려준다."""
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    found = []
    for number, r in enumerate(rows[1:], start=2):
        if any(v is not None for v in r):
            found.append((number, dict(zip(header, r))))
    return found


def formula_cells(path):
    """결과 엑셀에서 수식으로 저장된 셀의 좌표 목록."""
    wb = load_workbook(path)
    found = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    found.append(cell.coordinate)
    return found


def string_cells(path):
    """결과 엑셀의 첫 시트 데이터 영역(헤더 아래)의 문자열 셀 값 목록."""
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    values = []
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str):
                values.append(cell.value)
    return values


def by_source_row(records):
    return {r["원본 행"]: r for r in records}


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def make_config(tmp, column_map_extra=(), conversion_extra=()):
    """config/의 사본을 만들고 줄을 덧붙인다. 원본 config는 건드리지 않는다."""
    dest = Path(tmp) / "config"
    shutil.copytree(CONFIG, dest)
    for name, extra in (("column_map.csv", column_map_extra),
                        ("conversion_map.csv", conversion_extra)):
        if not extra:
            continue
        path = dest / name
        data = path.read_bytes()
        if not data.endswith(b"\n"):
            data += b"\n"
        data += ("\n".join(extra) + "\n").encode("utf-8")
        path.write_bytes(data)
    return dest


def write_xlsx(path, header, rows, merges=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "시트1"
    ws.append(list(header))
    for row in rows:
        ws.append(list(row))
    for rng in merges:
        ws.merge_cells(rng)
    wb.save(path)
    return Path(path)


class CliCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)
        self.out = self.tmp / "out"

    def run_ok(self, inputs, config=CONFIG):
        """도구를 실행해 종료 코드 0과 결과 파일 하나를 확인하고 (헤더, 행 목록)을 돌려준다."""
        result = run_cli(inputs, config, self.out)
        self.assertEqual(
            result.returncode, 0,
            "종료 코드가 0이 아님. stderr 끝부분: " + result.stderr[-300:],
        )
        files = result_files(self.out)
        self.assertEqual(len(files), 1, "출력 폴더의 통합 엑셀이 하나가 아님: %d개" % len(files))
        self.result_path = files[0]
        return read_sheet(files[0])

    def run_basic(self):
        return self.run_ok([FAKE / "A_basic.xlsx"])

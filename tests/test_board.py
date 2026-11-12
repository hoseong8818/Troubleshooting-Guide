"""T052: 통합 엑셀을 QnA 게시판 HTML 한 장으로 만든다 (데모, 헌법 원칙 1).

계약
- 실행: `python -m trouble_merge.board --merged <통합 xlsx> --out <폴더>` -> `<폴더>/board.html`
- 게시글 데이터는 `<script type="application/json" id="board-data">` 안의 JSON 배열이다.
  게시글 하나는 통합 엑셀 한 행이고 키는 통합 시트의 헤더 그대로다(빈 칸은 null).
- 외부 주소 요청, 외부 스크립트, 외부 스타일이 없다. 검색은 브라우저 안에서만 한다.
- 이 파일의 값은 가짜 데이터에서만 온다.
"""

import json
import os
import re
import subprocess
import sys
import unittest

from tests.support import CONFIG, ROOT, SRC, CliCase, run_cli, write_xlsx

DATA_RE = re.compile(r'<script[^>]*id="board-data"[^>]*>(.*?)</script>', re.DOTALL)


def run_board(merged, out):
    env = dict(os.environ)
    env["PYTHONPATH"] = str(SRC) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    cmd = [sys.executable, "-m", "trouble_merge.board", "--merged", str(merged), "--out", str(out)]
    return subprocess.run(
        cmd, cwd=str(ROOT), env=env, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=60,
    )


def posts_of(html):
    match = DATA_RE.search(html)
    if match is None:
        raise AssertionError("board-data JSON을 찾지 못함")
    return json.loads(match.group(1))


class BoardTest(CliCase):
    def setUp(self):
        super().setUp()
        self.header, self.records = self.run_basic()
        self.board_out = self.tmp / "board_out"

    def make(self):
        result = run_board(self.result_path, self.board_out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        path = self.board_out / "board.html"
        self.assertTrue(path.exists(), "board.html이 만들어지지 않음")
        return path.read_text(encoding="utf-8")

    def test_one_post_per_merged_row_with_same_values(self):
        posts = posts_of(self.make())
        self.assertEqual(len(posts), len(self.records), "게시글 수가 통합 행 수와 다름")
        for index, (post, record) in enumerate(zip(posts, self.records), start=1):
            with self.subTest(게시글=index):
                self.assertEqual(post, record)

    def test_no_external_requests(self):
        html = self.make()
        self.assertNotRegex(html, r"https?://")
        self.assertNotRegex(html, r"(?i)<script[^>]+\bsrc\s*=")
        self.assertNotRegex(html, r"(?i)<link\b")
        self.assertNotRegex(html, r"(?i)\b(fetch|XMLHttpRequest|WebSocket)\b")

    def test_has_search_box_category_filter_and_no_result_message(self):
        html = self.make()
        self.assertRegex(html, r'(?i)<input[^>]+type="search"')
        self.assertIn("분류", html)
        self.assertIn("해당 사례 없음", html)

    def test_check_needed_and_duplicate_marks_are_kept_in_data(self):
        posts = posts_of(self.make())
        self.assertEqual(
            [p["확인 필요 이유"] is not None for p in posts],
            [r["확인 필요 이유"] is not None for r in self.records],
        )
        self.assertEqual(
            [p["중복 의심"] is not None for p in posts],
            [r["중복 의심"] is not None for r in self.records],
        )

    def test_script_text_in_cells_cannot_break_out_of_the_data_block(self):
        evil = "</script><script>alert(1)</script>"
        row = ["PMP01", evil, "원인1", "조치1", "베어링"]
        path = write_xlsx(self.tmp / "evil.xlsx", ["설비코드", "증상", "원인", "조치내용", "부품분류"], [row])
        out = self.tmp / "evil_out"
        self.assertEqual(run_cli([path], CONFIG, out).returncode, 0)
        merged = sorted(out.glob("*.xlsx"))[0]
        result = run_board(merged, self.tmp / "evil_board")
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        html = (self.tmp / "evil_board" / "board.html").read_text(encoding="utf-8")
        self.assertNotIn("</script><script>alert", html)
        posts = posts_of(html)
        self.assertEqual(len(posts), 1)
        self.assertIn("alert(1)", posts[0]["증상"])

    def test_existing_board_is_not_overwritten(self):
        self.make()
        first = (self.board_out / "board.html").read_bytes()
        result = run_board(self.result_path, self.board_out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        self.assertEqual((self.board_out / "board.html").read_bytes(), first, "기존 board.html이 바뀜")
        self.assertEqual(len(list(self.board_out.glob("board*.html"))), 2)

    def test_missing_merged_file_stops_without_traceback(self):
        result = run_board(self.tmp / "none.xlsx", self.board_out)
        self.assertEqual(result.returncode, 2)
        self.assertFalse((self.board_out / "board.html").exists())
        self.assertNotIn("Traceback", result.stderr)

if __name__ == "__main__":
    unittest.main()

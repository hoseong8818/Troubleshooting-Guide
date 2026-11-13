"""T065: 게시판에 '새 사례 등록' 시안 화면이 있다 (데모).

앞으로 엑셀 대신 화면에서 입력하면 수기 오류를 입력 단계에서 막을 수 있다는 것을 보여 주는 시안이다.
설비와 분류는 게시글에 있는 값 목록에서 고르고, 일자는 날짜 입력(미래 불가), 증상·원인·조치는 필수다.
저장하거나 전송하지 않는다(`<form>`에 action 없음, 외부 요청 없음).
"""

import re
import unittest

from tests.support import FAKE, CliCase
from tests.test_board import run_board

FORM_RE = re.compile(r'<form\b[^>]*id="new-form"[^>]*>(.*?)</form>', re.DOTALL)


class BoardFormTest(CliCase):
    def setUp(self):
        super().setUp()
        self.run_ok([FAKE / "A_basic.xlsx"])
        out = self.tmp / "board_out"
        result = run_board(self.result_path, out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        self.html = (out / "board.html").read_text(encoding="utf-8")

    def form(self):
        match = FORM_RE.search(self.html)
        if match is None:
            raise AssertionError('id="new-form"인 <form>이 없음')
        return match

    def test_has_new_case_entry(self):
        self.assertIn("새 사례 등록", self.html)
        self.form()

    def test_form_never_submits_anywhere(self):
        tag = re.search(r'<form\b[^>]*id="new-form"[^>]*>', self.html).group(0) if 'id="new-form"' in self.html else ""
        self.assertTrue(tag, "등록 양식이 없음")
        self.assertNotRegex(tag, r"\baction\s*=")
        self.assertNotRegex(tag, r"\bmethod\s*=\s*\"?post", "전송 방식이 지정됨")

    def test_choices_instead_of_typing_for_equipment_and_category(self):
        body = self.form().group(1)
        self.assertRegex(body, r'<select\b[^>]*id="f-equipment"')
        self.assertRegex(body, r'<select\b[^>]*id="f-category"')

    def test_date_input_and_required_text_fields(self):
        body = self.form().group(1)
        self.assertRegex(body, r'<input\b[^>]*type="date"[^>]*id="f-date"|<input\b[^>]*id="f-date"[^>]*type="date"')
        for field in ("f-symptom", "f-cause", "f-action"):
            with self.subTest(칸=field):
                self.assertRegex(body, r'<textarea\b[^>]*id="%s"[^>]*\brequired\b|<textarea\b[^>]*\brequired\b[^>]*id="%s"' % (field, field))

    def test_says_it_is_a_mockup_that_does_not_save(self):
        self.assertIn("시안", self.html)
        self.assertIn("저장되지 않습니다", self.html)


if __name__ == "__main__":
    unittest.main()

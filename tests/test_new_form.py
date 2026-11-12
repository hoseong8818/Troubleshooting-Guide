"""T042: 새 양식은 코드 수정 없이 대응 표에 줄만 추가하면 처리된다. 규칙이 없으면 감사에 '대응 규칙 없음' (US4 수용 기준 2·3, FR-007).

임시 사본의 대응 표만 고친다. 원본 config/는 건드리지 않는다.
"""

import shutil
import unittest

from tests.support import (
    CONFIG, FAKE, CliCase, audit_files, by_source_row, make_config, run_cli, write_xlsx,
)

NEW_HEADER = ["일시", "호기", "현상", "이유", "대처", "구분"]
NEW_ROWS = [
    ["2026-11-15", "PMP01", "현상1", "이유1", "대처1", "베어링"],
    ["2026-11-16", "CNV01", "현상2", "이유2", "대처2", "모터"],
]
NEW_FORM_LINES = [
    "C,일시,일자,,",
    "C,호기,설비,,",
    "C,현상,증상,,",
    "C,이유,원인,,",
    "C,대처,조치,,",
    "C,구분,분류,,",
]


class NewFormTest(CliCase):
    def test_unknown_form_is_reported_not_silently_dropped(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        write_xlsx(folder / "b_new.xlsx", NEW_HEADER, NEW_ROWS)
        result = run_cli([folder], CONFIG, self.out)
        self.assertEqual(result.returncode, 1)
        audits = audit_files(self.out)
        self.assertEqual(len(audits), 1)
        text = audits[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^대응 규칙 없음 시트: file_02 sheet_\d\d$")

    def test_adding_table_lines_makes_the_new_form_work(self):
        path = write_xlsx(self.tmp / "new.xlsx", NEW_HEADER, NEW_ROWS)
        config = make_config(self.tmp, column_map_extra=NEW_FORM_LINES)
        _, records = self.run_ok([path], config)
        got = by_source_row(records)
        self.assertEqual(len(records), 2)
        self.assertEqual(got[2]["설비"], "PMP01")
        self.assertEqual(got[2]["일자"], "2026-11-15")
        self.assertEqual(got[3]["분류"], "모터")
        self.assertIsNone(got[2]["확인 필요 이유"])

    def test_new_form_and_existing_forms_merge_into_one_sheet(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        write_xlsx(folder / "b_new.xlsx", NEW_HEADER, NEW_ROWS)
        config = make_config(self.tmp, column_map_extra=NEW_FORM_LINES)
        _, records = self.run_ok([folder], config)
        self.assertEqual(len(records), 11 + 2)
        self.assertEqual({r["원본 파일"] for r in records}, {"a_good.xlsx", "b_new.xlsx"})


if __name__ == "__main__":
    unittest.main()

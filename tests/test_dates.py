"""T016: 날짜의 여러 형식이 `YYYY-MM-DD`로 통일된다 (FR-011, SC-003).

양식 A에는 일자 컬럼이 없어서 가짜 파일을 쓰지 않는다. 테스트가 임시 폴더에
일자 컬럼이 있는 입력과 임시 대응 표(양식 T)를 만들어 시험한다.
"""

import unittest
from datetime import date, datetime

from tests.support import CliCase, by_source_row, make_config, write_xlsx

SERIAL_20261107 = (date(2026, 11, 7) - date(1899, 12, 30)).days

HEADER = ["설비코드", "발생일자", "증상", "원인", "조치내용", "부품분류"]
COLUMN_MAP_EXTRA = [
    "T,설비코드,설비,,",
    "T,발생일자,일자,,",
    "T,증상,증상,,",
    "T,원인,원인,,",
    "T,조치내용,조치,,",
    "T,부품분류,분류,,",
]

# (증상 이름, 일자 칸의 원래 값)
CASES = [
    ("날짜형 셀", datetime(2026, 11, 7)),
    ("시각이 붙은 값", datetime(2026, 11, 7, 13, 45, 10)),
    ("엑셀 일련번호", SERIAL_20261107),
    ("하이픈 문자열", "2026-11-07"),
    ("점 문자열", "2026.11.07"),
    ("슬래시 문자열", "2026/11/7"),
]


class DatesTest(CliCase):
    def test_all_formats_become_iso_dates(self):
        rows = [["PMP01", value, name, "원인", "조치", "베어링"] for name, value in CASES]
        path = write_xlsx(self.tmp / "dates.xlsx", HEADER, rows)
        config = make_config(self.tmp, column_map_extra=COLUMN_MAP_EXTRA)
        _, records = self.run_ok([path], config)
        got = by_source_row(records)
        for index, (name, _) in enumerate(CASES):
            with self.subTest(형식=name):
                self.assertEqual(str(got[index + 2]["일자"]), "2026-11-07")


if __name__ == "__main__":
    unittest.main()

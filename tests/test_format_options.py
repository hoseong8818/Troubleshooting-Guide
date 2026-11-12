"""T051a: 컬럼 대응 표의 헤더 행·시트 이름 지정, 대응 중복 양식, 컬럼명 30자 제한 (FR-017, NFR-009).

구현이 먼저 있었던 동작을 뒤늦게 시험한다. 임시 사본의 대응 표만 고치고 원본 config/는 건드리지 않는다.
"""

import shutil
import unittest

from openpyxl import Workbook

from tests.support import (
    FAKE, CliCase, audit_files, make_config, read_sheet, result_files, run_cli, write_xlsx,
)

HEADER_D = ["일시D", "호기D", "현상D", "이유D", "대처D", "구분D"]
ROW_D = ["2026-11-15", "PMP01", "현상1", "이유1", "대처1", "베어링"]


def lines_for(name, header, header_row="", sheet=""):
    standards = ["일자", "설비", "증상", "원인", "조치", "분류"]
    return ["%s,%s,%s,%s,%s" % (name, src, std, header_row, sheet) for src, std in zip(header, standards)]


class HeaderRowTest(CliCase):
    def _title_row_file(self):
        # 1행은 제목, 2행이 헤더, 3행부터 데이터
        return write_xlsx(self.tmp / "title.xlsx", ["제목행"], [HEADER_D, ROW_D])

    def test_header_row_option_reads_the_header_below_a_title_row(self):
        config = make_config(self.tmp, column_map_extra=lines_for("D", HEADER_D, header_row="2"))
        _, records = self.run_ok([self._title_row_file()], config)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["설비"], "PMP01")
        self.assertEqual(records[0]["일자"], "2026-11-15")
        self.assertEqual(records[0]["원본 행"], 3, "원본 행은 엑셀 행 번호여야 함")

    def test_without_the_option_the_same_file_is_unmatched(self):
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        shutil.copy(self._title_row_file(), folder / "b_title.xlsx")
        config = make_config(self.tmp, column_map_extra=lines_for("D", HEADER_D))
        result = run_cli([folder], config, self.out)
        self.assertEqual(result.returncode, 1)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^대응 규칙 없음 시트: file_02 sheet_\d\d$")


class SheetNameTest(CliCase):
    def test_sheet_option_reads_only_the_named_sheet(self):
        wb = Workbook()
        target = wb.active
        target.title = "대상시트"
        target.append(HEADER_D)
        target.append(ROW_D)
        other = wb.create_sheet("다른시트")
        other.append(HEADER_D)
        other.append(["2026-11-16", "CNV01", "현상2", "이유2", "대처2", "모터"])
        path = self.tmp / "two.xlsx"
        wb.save(path)

        config = make_config(self.tmp, column_map_extra=lines_for("E", HEADER_D, sheet="대상시트"))
        result = run_cli([path], config, self.out)
        self.assertEqual(result.returncode, 1, "이름이 다른 시트는 '대응 규칙 없음'으로 보고되어야 함")
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual({r["원본 시트"] for r in records}, {"대상시트"})
        self.assertEqual(len(records), 1)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^대응 규칙 없음 시트: file_01 sheet_\d\d$")


class DuplicateMappingTest(CliCase):
    def test_form_with_two_sources_for_one_standard_column_is_reported_and_skipped(self):
        header = HEADER_D + ["호기D2"]
        extra = lines_for("F", HEADER_D) + ["F,호기D2,설비,,"]
        folder = self.tmp / "in"
        folder.mkdir()
        shutil.copy(FAKE / "A_basic.xlsx", folder / "a_good.xlsx")
        write_xlsx(folder / "b_dup.xlsx", header, [ROW_D + ["PMP02"]])
        config = make_config(self.tmp, column_map_extra=extra)
        result = run_cli([folder], config, self.out)
        self.assertEqual(result.returncode, 1)
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^대응 중복 양식 시트: file_02 sheet_\d\d$")
        _, records = read_sheet(result_files(self.out)[0])
        self.assertEqual({r["원본 파일"] for r in records}, {"a_good.xlsx"})


class ColumnNameLengthTest(CliCase):
    def test_registered_column_names_are_cut_at_30_characters_in_the_audit(self):
        long_name = "EQUIPMENT_" + "X" * 25  # 35자
        self.assertEqual(len(long_name), 35)
        header = [HEADER_D[0], long_name] + HEADER_D[2:]
        path = write_xlsx(self.tmp / "long.xlsx", header, [[ROW_D[0], None] + ROW_D[2:]])
        config = make_config(self.tmp, column_map_extra=lines_for("G", header))
        result = run_cli([path], config, self.out)
        self.assertEqual(result.returncode, 0, result.stderr[-300:])
        text = audit_files(self.out)[0].read_text(encoding="utf-8")
        self.assertRegex(text, r"(?m)^필수값 빈 칸: file_01 sheet_\d\d 행 2 컬럼 %s$" % long_name[:30])
        self.assertNotIn(long_name, text + result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()

"""안전한 로그 (T034, NFR-002, NFR-009).

이 모듈이 로그를 내는 유일한 길이다. 받는 정보는 종류, 파일명, 시트명, 행 번호, 컬럼명, 건수뿐이다.
- 자유 문장이나 셀 값을 받는 인자가 없다.
- 파일명·시트명은 기본이 마스킹(file_01, sheet_01)이고, 실제 이름은 사람이 옵션으로 켠다.
- 컬럼명은 컬럼 대응 표에 등록된 이름만 그대로 내고(30자까지), 나머지는 위치(열 3(C))로만 낸다.
"""

import sys

from openpyxl.utils import get_column_letter

MAX_NAME = 30


class SafeLog:
    def __init__(self, registered_columns, show_names=False, stream=None):
        self.registered = set(registered_columns)
        self.show_names = show_names
        self.stream = stream if stream is not None else sys.stderr
        self._files = {}
        self._sheets = {}

    # ---- 이름 표시 ----
    def file_label(self, name):
        if self.show_names:
            return name
        index = self._files.setdefault(name, len(self._files) + 1)
        return "file_%02d" % index

    def sheet_label(self, name):
        if self.show_names:
            return name
        index = self._sheets.setdefault(name, len(self._sheets) + 1)
        return "sheet_%02d" % index

    def column_label(self, name, position):
        """등록된 이름이면 이름(30자까지), 아니면 위치(position은 0부터)로 낸다."""
        text = "" if name is None else str(name).strip()
        if text and text in self.registered:
            return text[:MAX_NAME]
        return "열 %d(%s)" % (position + 1, get_column_letter(position + 1))

    # ---- 로그 ----
    def log(self, kind, file=None, sheet=None, row=None, column=None, count=None):
        """column은 column_label()이 만든 문자열만 받는다."""
        parts = [kind]
        if file is not None:
            parts.append("파일=" + self.file_label(file))
        if sheet is not None:
            parts.append("시트=" + self.sheet_label(sheet))
        if row is not None:
            parts.append("행=%d" % row)
        if column is not None:
            parts.append("컬럼=" + column)
        if count is not None:
            parts.append("건수=%d" % count)
        print(" ".join(parts), file=self.stream)

    # ---- 이름 대응표 ----
    def name_map_lines(self):
        lines = ["%s = %s" % (label, name) for name, label in self._labelled(self._files, "file_%02d")]
        lines += ["%s = %s" % (label, name) for name, label in self._labelled(self._sheets, "sheet_%02d")]
        return lines

    @staticmethod
    def _labelled(table, pattern):
        return [(name, pattern % index) for name, index in sorted(table.items(), key=lambda kv: kv[1])]

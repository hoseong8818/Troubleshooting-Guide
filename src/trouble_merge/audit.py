"""감사 결과 텍스트 (T035, FR-008, US3).

한국어 텍스트 한 장. 파일명·시트명·컬럼명은 SafeLog가 정한 표시(마스킹·위치)로만 쓰고 셀 값은 넣지 않는다.
"""

from .writer import unique_path


class Audit:
    def __init__(self, log):
        self.log = log
        self.input_rows = 0
        self.output_rows = 0
        self.check_count = 0
        self.duplicate_count = 0
        self.files_read = 0
        self.failed_files = []       # 파일명
        self.unmatched = []          # (사유 문구, 파일명, 시트명)
        self.unregistered = []       # (파일명, 시트명, 컬럼 표시)
        self.blank_required = []     # (파일명, 시트명, 행 번호, 컬럼 표시)
        self.hidden = []             # (파일명, 숨긴 시트 수)
        self.formula_missing = 0
        self.prefixed = 0

    def add_unregistered(self, file, sheet, label):
        entry = (file, sheet, label)
        if entry not in self.unregistered:
            self.unregistered.append(entry)

    def render(self):
        log = self.log
        lines = ["감사 결과", ""]
        lines.append("입력 행 수: %d" % self.input_rows)
        lines.append("출력 행 수: %d" % self.output_rows)
        lines.append("행 수 일치: %s" % ("예" if self.input_rows == self.output_rows else "아니오"))
        lines.append("확인 필요 건수: %d" % self.check_count)
        lines.append("중복 의심 건수: %d" % self.duplicate_count)
        lines.append("읽은 파일 수: %d" % self.files_read)
        if self.failed_files:
            lines.append("실패한 파일: " + ", ".join(log.file_label(f) for f in self.failed_files))
        else:
            lines.append("실패한 파일: 없음")
        for reason, file, sheet in self.unmatched:
            lines.append("%s: %s %s" % (reason, log.file_label(file), log.sheet_label(sheet)))
        for file, sheet, label in self.unregistered:
            lines.append("대응 규칙 없는 컬럼: %s %s %s" % (log.file_label(file), log.sheet_label(sheet), label))
        for file, sheet, row, label in self.blank_required:
            lines.append("필수값 빈 칸: %s %s 행 %d 컬럼 %s" % (
                log.file_label(file), log.sheet_label(sheet), row, label))
        lines.append("숨긴 시트 수: %d" % sum(n for _, n in self.hidden))
        for file, count in self.hidden:
            if count:
                lines.append("숨긴 시트: %s %d개" % (log.file_label(file), count))
        lines.append("수식 값 없음 건수: %d" % self.formula_missing)
        lines.append("수식 방지 접두사 처리 건수: %d" % self.prefixed)
        return "\n".join(lines) + "\n"

    def write(self, out_dir, stamp):
        path = unique_path(out_dir, "audit_" + stamp, ".txt", stamp)
        path.write_text(self.render(), encoding="utf-8")
        return path


def write_name_map(log, out_dir, stamp):
    """마스킹 이름과 실제 이름의 대응표를 감사 결과 옆에 따로 쓴다(NFR-009)."""
    path = unique_path(out_dir, "names_" + stamp, ".txt", stamp)
    path.write_text("\n".join(log.name_map_lines()) + "\n", encoding="utf-8")
    return path

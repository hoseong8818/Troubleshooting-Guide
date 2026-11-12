"""5개 파일(양식 A 2개 + 양식 B 3개)의 통합 정답 초안을 expected/ALL_expected.xlsx로 만든다 (T040 초안).

사용: 저장소 루트에서 `python tools/make_expected_multi_draft.py` (사람이 자기 터미널에서 실행)

- 입력: data/fake/ 의 A_basic, A_trap, B_basic, B_more, B_trap (각 파일의 보이는 시트 '시트1')
- 값은 손으로 계산해 적었다. 구현 코드와 무관하다. 만든 뒤 사람이 엑셀에서 입력 파일과 한 행씩 대조해 확정한다.
- 이미 있으면 덮어쓰지 않고 중단한다.

정렬(FR-016): 설비 -> 일자(빈 칸이 먼저, '확인 필요'가 날짜 뒤) -> 원본 파일 -> 원본 행.
설비가 '확인 필요'인 행은 맨 아래다. 총 33행(A_basic 11 + A_trap 6 + B_basic 6 + B_more 4 + B_trap 6).
중복 의심 6행: A_basic의 PMP01 '소음 발생' 3행(일자 빈 칸끼리), B의 PMP01 2026-11-07 '소음 발생' 3행(파일이 달라도).
A와 B 사이는 일자가 달라 중복이 아니다.

확인 필요 판정 문구와 표기는 expected-A-draft.md에서 확정한 것을 그대로 쓴다.
"""

import sys
from pathlib import Path

from openpyxl import Workbook

HEADER = [
    "설비", "일자", "증상", "원인", "조치", "분류", "기타 항목",
    "확인 필요 이유", "원래 값", "중복 의심", "원본 파일", "원본 시트", "원본 행",
]
N = None
CHK = "확인 필요"
DUP = "중복 의심"
SHEET = "시트1"


def R(eq, day, symptom, cause, action, category, other, reason, original, dup, file, row):
    return [eq, day, symptom, cause, action, category, other, reason, original, dup, file, SHEET, row]


ROWS = [
    # ---- CNV01 ----
    R("CNV01", N, "벨트 이탈", "장력 부족", "장력 조정", "모터", "에러코드: E003", N, N, N, "A_basic.xlsx", 4),
    R("CNV01", N, "'=1+1", "센서 오염", "센서 청소", "센서", "에러코드: E007", N, N, N, "A_basic.xlsx", 10),
    R("CNV01", N, "정지", "퓨즈 단선", "'@퓨즈 교체", "모터", "에러코드: E010", N, N, N, "A_basic.xlsx", 13),
    R("CNV01", N, CHK, "과열", CHK, "모터", "에러코드: E104",
      "증상 빈 칸; 조치 빈 칸", "증상: (빈 칸); 조치: (빈 칸)", N, "A_trap.xlsx", 5),
    R("CNV01", "2026-11-04", "벨트 처짐", "장력 저하", "벨트 장력 조정", "모터", "작업자: 작업자2", N, N, N, "B_basic.xlsx", 3),
    R("CNV01", "2026-11-08", "벨트 이탈", "롤러 파손", "롤러 교체", "모터", "작업자: 작업자2", N, N, N, "B_more.xlsx", 3),
    R("CNV01", CHK, "정지", "퓨즈 단선", "퓨즈 교체", "모터", N,
      "일자 해석 불가", "일자: 26.11.07", N, "B_trap.xlsx", 3),
    # ---- MIX03 ----
    R("MIX03", N, "온도 상승", "냉각수 부족", "냉각수 보충", "센서", N, N, N, N, "A_basic.xlsx", 7),
    R("MIX03", N, "교반 정지", "'+1 과부하", "재기동", "모터", "에러코드: E008", N, N, N, "A_basic.xlsx", 11),
    R("MIX03", N, CHK, "과열", CHK, "모터", N,
      "증상 빈 칸; 조치 빈 칸", "증상: (빈 칸); 조치: (빈 칸)", N, "B_trap.xlsx", 6),
    R("MIX03", "2026-11-05", "교반 지연", "모터 과열", "냉각 후 재가동", "모터", N, N, N, N, "B_basic.xlsx", 4),
    R("MIX03", "2026-11-09", "온도 상승", "냉각수 부족", "냉각수 보충", "센서", N, N, N, N, "B_more.xlsx", 4),
    # ---- PMP01 ----
    R("PMP01", N, "소음 발생", "베어링 마모", "베어링 교체", "베어링", "에러코드: E001", N, N, DUP, "A_basic.xlsx", 2),
    R("PMP01", N, "진동 증가", "축 정렬 불량", "축 재정렬", "모터", "에러코드: E002", N, N, N, "A_basic.xlsx", 3),
    R("PMP01", N, "소음 발생", "윤활 부족", "그리스 주입", "베어링", "에러코드: E004", N, N, DUP, "A_basic.xlsx", 6),
    R("PMP01", N, "소음 발생", "체결 풀림", "볼트 조임", "베어링", "에러코드: E006", N, N, DUP, "A_basic.xlsx", 9),
    R("PMP01", N, "경보", "센서 오염", "센서 청소", CHK, "에러코드: E103",
      "분류 미등록", "분류: LEAK-분류", N, "A_trap.xlsx", 4),
    R("PMP01", "2026-11-03", "누수 발생", "패킹 노후", "패킹 교체", "밸브", "작업자: 작업자1", N, N, N, "B_basic.xlsx", 2),
    R("PMP01", "2026-11-07", "소음 발생", "베어링 마모", "베어링 교체", "베어링", "작업자: 작업자1", N, N, DUP, "B_basic.xlsx", 7),
    R("PMP01", "2026-11-07", "소음 발생", "윤활 불량", "그리스 보충", "베어링", "작업자: 작업자4", N, N, DUP, "B_basic.xlsx", 8),
    R("PMP01", "2026-11-07", "소음 발생", "체결 풀림", "볼트 조임", "베어링", "작업자: 작업자5", N, N, DUP, "B_more.xlsx", 2),
    R("PMP01", CHK, "경보", "센서 오염", "센서 청소", "센서", "작업자: 작업자1",
      "일자 해석 불가", "일자: 11/07/2026", N, "B_trap.xlsx", 2),
    # ---- PRS04 ----
    R("PRS04", N, "압력 저하", "밸브 누설", "밸브 교체", "밸브", "에러코드: E005", N, N, N, "A_basic.xlsx", 8),
    R("PRS04", N, "압력 상승", "밸브 고착", "'-2 분해 청소", "밸브", "에러코드: E009", N, N, N, "A_basic.xlsx", 12),
    R("PRS04", "2026-11-06", "압력 헌팅", "벨브 마모", "벨브 교체", "밸브", "작업자: 작업자3", N, N, N, "B_basic.xlsx", 6),
    R("PRS04", "2026-11-10", "압력 상승", "밸브 고착", "분해 청소", "밸브", "작업자: 작업자3", N, N, N, "B_more.xlsx", 5),
    # ---- 설비가 '확인 필요'인 행 (맨 아래) ----
    R(CHK, N, "누유", "씰 파손", "씰 교체", "밸브", "에러코드: E101",
      "설비 코드 빈 칸", "설비: (빈 칸)", N, "A_trap.xlsx", 2),
    R(CHK, N, "LEAK-정지", "원인 불명", "재기동", "모터", "에러코드: E102",
      "설비 코드 미등록", "설비: ZZ99", N, "A_trap.xlsx", 3),
    R(CHK, N, "소음", CHK, "점검", CHK, "에러코드: E105",
      "설비 코드 미등록; 원인 빈 칸; 분류 미등록", "설비: ZZ98; 원인: (빈 칸); 분류: 미등록분류", N, "A_trap.xlsx", 6),
    R(CHK, N, "누유", "씰 노후", "점검", "밸브", "에러코드: E106",
      "설비 코드 빈 칸", "설비: (빈 칸)", N, "A_trap.xlsx", 7),
    R(CHK, "2026-11-11", "누유", "씰 파손", "씰 교체", "밸브", "작업자: 작업자2",
      "설비 코드 빈 칸", "설비: (빈 칸)", N, "B_trap.xlsx", 4),
    R(CHK, "2026-11-11", "소음", "원인 불명", "점검", CHK, "작업자: 작업자3",
      "설비 코드 미등록; 분류 미등록", "설비: ZZ97; 분류: LEAK-부품구분", N, "B_trap.xlsx", 5),
    R(CHK, CHK, "LEAK-정지B", CHK, "점검", "밸브", "작업자: 작업자4",
      "설비 코드 미등록; 일자 해석 불가; 원인 빈 칸", "설비: ZZ96; 일자: abc; 원인: (빈 칸)", N, "B_trap.xlsx", 7),
]


def main():
    out = Path(__file__).resolve().parent.parent / "expected" / "ALL_expected.xlsx"
    if out.exists():
        print("중단: 이미 있는 파일은 덮어쓰지 않는다 -> ALL_expected.xlsx")
        return 1
    wb = Workbook()
    ws = wb.active
    ws.title = "통합"
    ws.append(HEADER)
    for r_idx, row in enumerate(ROWS, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"
    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "xb") as f:
        wb.save(f)
    print("만듦: expected/ALL_expected.xlsx (%d행)" % len(ROWS))
    return 0


if __name__ == "__main__":
    sys.exit(main())

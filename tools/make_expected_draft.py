"""양식 A 정답 초안 엑셀 두 개를 expected/에 만든다 (T010 초안).

사용: 저장소 루트에서 `python tools/make_expected_draft.py` (사람이 자기 터미널에서 실행)

- 값은 docs/proposals/expected-A-draft.md의 표를 그대로 옮긴 것이다. 구현 코드와 무관하게 손으로 적은 값이다.
- 만든 뒤 사람이 엑셀에서 한 행씩 검토해 확정한다. 확정 전에는 '초안'이다.
- 이미 있는 파일은 덮어쓰지 않는다. 하나라도 있으면 아무것도 쓰지 않고 중단한다.
- 빈 칸은 진짜 빈 셀이다('(빈 칸)'은 '원래 값' 열 안의 글자일 때만 쓴다).
- 통합 시트의 시트 이름은 명세에 없어 '통합'으로 임시 지정했다.
"""

import sys
from pathlib import Path

from openpyxl import Workbook

HEADER = [
    "설비", "일자", "증상", "원인", "조치", "분류", "기타 항목",
    "확인 필요 이유", "원래 값", "중복 의심", "원본 파일", "원본 시트", "원본 행",
]

N = None  # 빈 셀
DUP = "중복 의심"
CHK = "확인 필요"

# 순서: 설비, 일자, 증상, 원인, 조치, 분류, 기타 항목, 이유, 원래 값, 중복 의심, 파일, 시트, 행
BASIC = [
    ["CNV01", N, "벨트 이탈", "장력 부족", "장력 조정", "모터", "에러코드: E003", N, N, N, "A_basic.xlsx", "시트1", 4],
    ["CNV01", N, "'=1+1", "센서 오염", "센서 청소", "센서", "에러코드: E007", N, N, N, "A_basic.xlsx", "시트1", 10],
    ["CNV01", N, "정지", "퓨즈 단선", "'@퓨즈 교체", "모터", "에러코드: E010", N, N, N, "A_basic.xlsx", "시트1", 13],
    ["MIX03", N, "온도 상승", "냉각수 부족", "냉각수 보충", "센서", N, N, N, N, "A_basic.xlsx", "시트1", 7],
    ["MIX03", N, "교반 정지", "'+1 과부하", "재기동", "모터", "에러코드: E008", N, N, N, "A_basic.xlsx", "시트1", 11],
    ["PMP01", N, "소음 발생", "베어링 마모", "베어링 교체", "베어링", "에러코드: E001", N, N, DUP, "A_basic.xlsx", "시트1", 2],
    ["PMP01", N, "진동 증가", "축 정렬 불량", "축 재정렬", "모터", "에러코드: E002", N, N, N, "A_basic.xlsx", "시트1", 3],
    ["PMP01", N, "소음 발생", "윤활 부족", "그리스 주입", "베어링", "에러코드: E004", N, N, DUP, "A_basic.xlsx", "시트1", 6],
    ["PMP01", N, "소음 발생", "체결 풀림", "볼트 조임", "베어링", "에러코드: E006", N, N, DUP, "A_basic.xlsx", "시트1", 9],
    ["PRS04", N, "압력 저하", "밸브 누설", "밸브 교체", "밸브", "에러코드: E005", N, N, N, "A_basic.xlsx", "시트1", 8],
    ["PRS04", N, "압력 상승", "밸브 고착", "'-2 분해 청소", "밸브", "에러코드: E009", N, N, N, "A_basic.xlsx", "시트1", 12],
]

TRAP = [
    ["CNV01", N, CHK, "과열", CHK, "모터", "에러코드: E104",
     "증상 빈 칸; 조치 빈 칸", "증상: (빈 칸); 조치: (빈 칸)", N, "A_trap.xlsx", "시트1", 5],
    ["PMP01", N, "경보", "센서 오염", "센서 청소", CHK, "에러코드: E103",
     "분류 미등록", "분류: LEAK-분류", N, "A_trap.xlsx", "시트1", 4],
    [CHK, N, "누유", "씰 파손", "씰 교체", "밸브", "에러코드: E101",
     "설비 코드 빈 칸", "설비: (빈 칸)", N, "A_trap.xlsx", "시트1", 2],
    [CHK, N, "LEAK-정지", "원인 불명", "재기동", "모터", "에러코드: E102",
     "설비 코드 미등록", "설비: ZZ99", N, "A_trap.xlsx", "시트1", 3],
    [CHK, N, "소음", CHK, "점검", CHK, "에러코드: E105",
     "설비 코드 미등록; 원인 빈 칸; 분류 미등록", "설비: ZZ98; 원인: (빈 칸); 분류: 미등록분류", N, "A_trap.xlsx", "시트1", 6],
    [CHK, N, "누유", "씰 노후", "점검", "밸브", "에러코드: E106",
     "설비 코드 빈 칸", "설비: (빈 칸)", N, "A_trap.xlsx", "시트1", 7],
]


def build(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "통합"
    ws.append(HEADER)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"
    return wb


def main():
    out_dir = Path(__file__).resolve().parent.parent / "expected"
    targets = {
        "A_basic_expected.xlsx": BASIC,
        "A_trap_expected.xlsx": TRAP,
    }

    existing = [name for name in targets if (out_dir / name).exists()]
    if existing:
        print("중단: 이미 있는 파일은 덮어쓰지 않는다 -> " + ", ".join(existing))
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    for name, rows in targets.items():
        with open(out_dir / name, "xb") as f:
            build(rows).save(f)
        print("만듦: expected/" + name + " (" + str(len(rows)) + "행)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

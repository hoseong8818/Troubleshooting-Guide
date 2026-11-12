"""양식 A 가짜 입력 엑셀 두 개를 data/fake/에 만든다 (T009).

사용: 저장소 루트에서 `python tools/make_fake_data.py` (사람이 자기 터미널에서 실행)

- 이미 있는 파일은 덮어쓰지 않는다. 하나라도 있으면 아무것도 쓰지 않고 중단한다(FR-020).
- 컬럼명은 가칭이다(docs/columns-A.md). 실제 이름으로 바꿀 때는 HEADER만 고치고 다시 만든다.
- 일자 컬럼은 없다(양식 A는 일자를 비워 둠). 날짜 형식 시험 행은 이 양식에 넣지 않는다.
- 값은 모두 지어낸 것이다. 실제 데이터를 쓰지 않는다.

가짜 값과 변환 표(config/conversion_map.csv, T011)의 관계 — 사람이 표에 넣을 줄 제안:
  설비 등록: PMP01, CNV01, MIX03, PRS04
  설비 미등록: ZZ98, ZZ99 (A_trap 전용)
  분류 등록: 베어링, 모터, 센서, 밸브 (표준 값은 자기 자신)
  분류 변환(표기가 다른 값 -> 표준 값): 모타 -> 모터, sensor -> 센서, bearing -> 베어링
    (A_basic의 입력 `모타`, `SENSOR`, ` Bearing`이 대소문자·공백 정리 후 이 줄과 일치)
  분류 미등록: 분류미등록, LEAK-분류, 미등록분류 (A_trap 전용)
"""

import sys
from pathlib import Path

from openpyxl import Workbook

HEADER = ["설비코드", "에러코드", "증상", "원인", "조치내용", "부품분류"]

# A_basic: 정상 행, 공백·대소문자, 완전히 빈 행, 설비·증상이 같은 행(중복 의심),
# 에러코드 빈 칸, 수식으로 시작하는 값(=, +, -, @).
BASIC_ROWS = [
    ["PMP01", "E001", "소음 발생", "베어링 마모", "베어링 교체", "베어링"],
    [" PMP01 ", "E002", "진동 증가", "축 정렬 불량", "축 재정렬", "모터"],
    ["cnv01", "E003", "벨트 이탈", "장력 부족", "장력 조정", "모타"],
    [None, None, None, None, None, None],
    ["PMP01", "E004", "소음 발생", "윤활 부족", "그리스 주입", "베어링"],
    ["MIX03", None, "온도 상승", "냉각수 부족", "냉각수 보충", "SENSOR"],
    ["PRS04", "E005", "압력 저하", "밸브 누설", "밸브 교체", "밸브"],
    ["pmp01 ", "E006", " 소음 발생 ", "체결 풀림", "볼트 조임", " Bearing"],
    ["CNV01", "E007", "=1+1", "센서 오염", "센서 청소", "센서"],
    ["MIX03", "E008", "교반 정지", "+1 과부하", "재기동", "모터"],
    ["PRS04", "E009", "압력 상승", "밸브 고착", "-2 분해 청소", "밸브"],
    ["CNV01", "E010", "정지", "퓨즈 단선", "@퓨즈 교체", "모터"],
    [None, None, None, None, None, None],
]

# A_trap: 설비 코드 빈 칸·미등록, 분류 미등록, 필수값 빈 행, 사유가 여러 개인 행,
# 설비가 '확인 필요'인 중복 후보 행, 표식 값 LEAK-.
TRAP_ROWS = [
    [None, "E101", "누유", "씰 파손", "씰 교체", "밸브"],
    ["ZZ99", "E102", "LEAK-정지", "원인 불명", "재기동", "모터"],
    ["PMP01", "E103", "경보", "센서 오염", "센서 청소", "LEAK-분류"],
    ["CNV01", "E104", None, "과열", None, "모터"],
    ["ZZ98", "E105", "소음", None, "점검", "미등록분류"],
    [None, "E106", "누유", "씰 노후", "점검", "밸브"],
]

HIDDEN_ROWS = [["PMP01", "E901", "LEAK-숨김", "숨긴 시트 값", "숨긴 시트 조치", "베어링"]]


def fill(ws, rows):
    ws.append(HEADER)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"  # '='로 시작해도 수식이 아니라 문자열로 저장


def make_basic():
    wb = Workbook()
    ws = wb.active
    ws.title = "시트1"
    fill(ws, BASIC_ROWS)
    return wb


def make_trap():
    wb = Workbook()
    ws = wb.active
    ws.title = "시트1"
    fill(ws, TRAP_ROWS)
    hidden = wb.create_sheet("숨김시트")
    fill(hidden, HIDDEN_ROWS)
    hidden.sheet_state = "hidden"
    very = wb.create_sheet("깊은숨김")
    fill(very, HIDDEN_ROWS)
    very.sheet_state = "veryHidden"
    return wb


def main():
    out_dir = Path(__file__).resolve().parent.parent / "data" / "fake"
    targets = {"A_basic.xlsx": make_basic, "A_trap.xlsx": make_trap}

    existing = [name for name in targets if (out_dir / name).exists()]
    if existing:
        print("중단: 이미 있는 파일은 덮어쓰지 않는다 -> " + ", ".join(existing))
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    for name, build in targets.items():
        with open(out_dir / name, "xb") as f:  # 'x': 이미 있으면 실패
            build().save(f)
        print("만듦: data/fake/" + name)
    return 0


if __name__ == "__main__":
    sys.exit(main())

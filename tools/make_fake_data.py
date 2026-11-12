"""가짜 입력 엑셀을 data/fake/에 만든다 (T009, T039).

사용 (저장소 루트에서, 사람이 자기 터미널에서 실행):
  python tools/make_fake_data.py        # 모든 묶음(A, B)
  python tools/make_fake_data.py B      # 양식 B 묶음만 (A 파일이 이미 있을 때)

- 이미 있는 파일은 덮어쓰지 않는다. 요청한 묶음의 파일이 하나라도 있으면 아무것도 쓰지 않고 중단한다(FR-020).
  그래서 양식 A를 이미 만들었다면 `B`를 지정해서 새 묶음만 만든다.
- 컬럼명은 가칭이다(docs/columns-A.md, docs/columns-B.md). 실제 이름으로 바꿀 때는 HEADER만 고치고 다시 만든다.
- 값은 모두 지어낸 것이다. 실제 데이터를 쓰지 않는다.

가짜 값과 설정 표의 관계 — 사람이 표에 넣을 줄 제안:
  [config/column_map.csv] 양식 B
    B,발생일,일자,,
    B,장비번호,설비,,
    B,고장내용,증상,,
    B,고장원인,원인,,
    B,처리결과,조치,,
    B,부품구분,분류,,
  [config/conversion_map.csv]
    설비 등록: PMP01, CNV01, MIX03, PRS04
    설비 미등록: ZZ96, ZZ97, ZZ98, ZZ99 (trap 전용)
    분류 등록: 베어링, 모터, 센서, 밸브 (표준 값은 자기 자신)
    분류 변환(표기가 다른 값 -> 표준 값):
      모타 -> 모터, sensor -> 센서, bearing -> 베어링   (양식 A, 이미 있음)
      motor -> 모터, 벨브 -> 밸브                        (양식 B, 새로 추가)
    분류 미등록: 분류미등록, LEAK-분류, 미등록분류, LEAK-부품구분 (trap 전용)
"""

import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

HEADER_A = ["설비코드", "에러코드", "증상", "원인", "조치내용", "부품분류"]
HEADER_B = ["발생일", "장비번호", "고장내용", "고장원인", "처리결과", "부품구분", "작업자"]

N = None
SERIAL_20261105 = (date(2026, 11, 5) - date(1899, 12, 30)).days  # 엑셀 일련번호

# ---- 양식 A -------------------------------------------------------------------------
# A_basic: 정상 행, 공백·대소문자, 완전히 빈 행, 설비·증상이 같은 행(중복 의심),
# 에러코드 빈 칸, 수식으로 시작하는 값(=, +, -, @), 표기가 다른 분류.
A_BASIC_ROWS = [
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
# 설비가 '확인 필요'인 중복 후보 행, 표식 값 LEAK-, 숨긴 시트.
A_TRAP_ROWS = [
    [None, "E101", "누유", "씰 파손", "씰 교체", "밸브"],
    ["ZZ99", "E102", "LEAK-정지", "원인 불명", "재기동", "모터"],
    ["PMP01", "E103", "경보", "센서 오염", "센서 청소", "LEAK-분류"],
    ["CNV01", "E104", None, "과열", None, "모터"],
    ["ZZ98", "E105", "소음", None, "점검", "미등록분류"],
    [None, "E106", "누유", "씰 노후", "점검", "밸브"],
]

A_HIDDEN_ROWS = [["PMP01", "E901", "LEAK-숨김", "숨긴 시트 값", "숨긴 시트 조치", "베어링"]]

# ---- 양식 B -------------------------------------------------------------------------
# B_basic: 날짜 형식(날짜형, 시각 붙은 값, 일련번호, 점·슬래시 문자열), 빈 행,
# 같은 파일 안의 중복(설비 PMP01 + 2026-11-07 + 소음 발생), 표기가 다른 분류와 설비 코드.
B_BASIC_ROWS = [
    [datetime(2026, 11, 3), "PMP01", "누수 발생", "패킹 노후", "패킹 교체", "밸브", "작업자1"],
    [datetime(2026, 11, 4, 9, 30), " CNV01", "벨트 처짐", "장력 저하", "벨트 장력 조정", "MOTOR", "작업자2"],
    [SERIAL_20261105, "MIX03", "교반 지연", "모터 과열", "냉각 후 재가동", "Motor", None],
    [None, None, None, None, None, None, None],
    ["2026.11.06", "PRS04", "압력 헌팅", "벨브 마모", "벨브 교체", "벨브", "작업자3"],
    ["2026/11/7", "pmp01", "소음 발생", "베어링 마모", "베어링 교체", "Bearing", "작업자1"],
    ["2026-11-07", "PMP01", "소음 발생", "윤활 불량", "그리스 보충", "베어링", "작업자4"],
    [None, None, None, None, None, None, None],
]

# B_more: B_basic과 설비·일자·증상이 같은 행(파일이 달라도 중복 의심, FR-010), 분류 변환.
B_MORE_ROWS = [
    ["2026-11-07", "PMP01", "소음 발생", "체결 풀림", "볼트 조임", "베어링", "작업자5"],
    [datetime(2026, 11, 8), "CNV01", "벨트 이탈", "롤러 파손", "롤러 교체", "모터", "작업자2"],
    [datetime(2026, 11, 9), "MIX03", "온도 상승", "냉각수 부족", "냉각수 보충", "SENSOR", None],
    [datetime(2026, 11, 10), "PRS04", "압력 상승", "밸브 고착", "분해 청소", "밸브", "작업자3"],
]

# B_trap: 읽을 수 없는 날짜, 설비 빈 칸·미등록, 분류 미등록, 필수값 빈 행,
# 사유가 여러 개인 행, 일자 빈 칸(확인 필요 아님), 표식 값 LEAK-, 숨긴 시트.
B_TRAP_ROWS = [
    ["11/07/2026", "PMP01", "경보", "센서 오염", "센서 청소", "센서", "작업자1"],
    ["26.11.07", "CNV01", "정지", "퓨즈 단선", "퓨즈 교체", "모터", None],
    [datetime(2026, 11, 11), None, "누유", "씰 파손", "씰 교체", "밸브", "작업자2"],
    [datetime(2026, 11, 11), "ZZ97", "소음", "원인 불명", "점검", "LEAK-부품구분", "작업자3"],
    [None, "MIX03", None, "과열", None, "모터", None],
    ["abc", "ZZ96", "LEAK-정지B", None, "점검", "밸브", "작업자4"],
]

B_HIDDEN_ROWS = [["2026-11-12", "PMP01", "LEAK-숨김B", "숨긴 시트 값", "숨긴 시트 조치", "베어링", "작업자9"]]


def fill(ws, header, rows):
    ws.append(header)
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.value = value
            if isinstance(value, str):
                cell.data_type = "s"  # '='로 시작해도 수식이 아니라 문자열로 저장


def make_book(header, rows, hidden_rows=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "시트1"
    fill(ws, header, rows)
    if hidden_rows is not None:
        hidden = wb.create_sheet("숨김시트")
        fill(hidden, header, hidden_rows)
        hidden.sheet_state = "hidden"
        very = wb.create_sheet("깊은숨김")
        fill(very, header, hidden_rows)
        very.sheet_state = "veryHidden"
    return wb


GROUPS = {
    "A": {
        "A_basic.xlsx": lambda: make_book(HEADER_A, A_BASIC_ROWS),
        "A_trap.xlsx": lambda: make_book(HEADER_A, A_TRAP_ROWS, A_HIDDEN_ROWS),
    },
    "B": {
        "B_basic.xlsx": lambda: make_book(HEADER_B, B_BASIC_ROWS),
        "B_more.xlsx": lambda: make_book(HEADER_B, B_MORE_ROWS),
        "B_trap.xlsx": lambda: make_book(HEADER_B, B_TRAP_ROWS, B_HIDDEN_ROWS),
    },
}


def main(argv):
    choice = argv[1].upper() if len(argv) > 1 else "ALL"
    if choice == "ALL":
        names = list(GROUPS)
    elif choice in GROUPS:
        names = [choice]
    else:
        print("사용: python tools/make_fake_data.py [A|B|all]")
        return 2

    targets = {}
    for group in names:
        targets.update(GROUPS[group])

    out_dir = Path(__file__).resolve().parent.parent / "data" / "fake"
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
    sys.exit(main(sys.argv))

"""설비 트러블슈팅 예시 데이터(전부 가상, 사람이 손으로 적은 것처럼 일부러 흐트러뜨림) 생성기.

- 설비 40종, 총 600대, 2015-01 ~ 2026-11
- 설비 종류마다 파일 2개(2015~2020, 2021~현재) = 80개, 파일당 50~100행(중복 행 포함)
- 첫 시트 '인덱스'(컬럼 대응 + 설비 종류별 에러코드 정의표), 둘째 시트 데이터
- 에러코드마다 증상·부품분류가 하나씩 정해져 있고 코드 번호 순서는 설비 종류마다 다르다
- 양식(컬럼명·날짜 표기) 4가지가 설비 종류마다 하나씩 배정된다
- 함정: 원인 공란/불명, 터무니없는 날짜, 날짜 칸에 에러코드를 같이 적어 한 칸씩 앞당겨진 행,
  설비코드·에러코드 표기 흔들림, 미등록 에러코드, 분류·증상 공란, 날짜 '미상', 중복 행

실행(저장소 루트에서): python <이 파일>   결과는 ./data/sample (이미 있으면 중단)
"""
import csv
import random
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

SEED = 20150101
OUT = Path.cwd() / "data" / "sample"
END = date(2026, 11, 10)

# (설비 종류, 코드 접두어, 부품 4개)
TYPES = [
    ("사출성형기", "INJ", ["스크류", "히터밴드", "금형", "유압펌프"]),
    ("프레스", "PRS", ["클러치", "브레이크", "금형", "슬라이드"]),
    ("CNC선반", "LTH", ["스핀들", "척", "공구대", "절삭유펌프"]),
    ("머시닝센터", "MCT", ["스핀들", "ATC", "이송축", "절삭유펌프"]),
    ("연삭기", "GRD", ["연삭휠", "스핀들", "드레서", "냉각수펌프"]),
    ("용접로봇", "WLD", ["토치", "와이어피더", "로봇암", "티칭펜던트"]),
    ("도장로봇", "PNT", ["스프레이건", "도료펌프", "로봇암", "배기팬"]),
    ("컨베이어", "CNV", ["구동모터", "벨트", "롤러", "감지센서"]),
    ("지게차", "FLT", ["배터리", "마스트", "체인", "유압실린더"]),
    ("천장크레인", "CRN", ["호이스트", "와이어로프", "주행모터", "리밋스위치"]),
    ("에어컴프레서", "CMP", ["압축기", "에어드라이어", "압력스위치", "오일필터"]),
    ("냉동기", "CHL", ["압축기", "증발기", "냉매배관", "순환펌프"]),
    ("냉각탑", "CLT", ["팬모터", "충진재", "순환펌프", "수위센서"]),
    ("보일러", "BLR", ["버너", "급수펌프", "압력계", "연소제어기"]),
    ("진공펌프", "VAC", ["로터", "펌프오일", "베어링", "진공게이지"]),
    ("건조로", "OVN", ["히터", "순환팬", "온도센서", "도어"]),
    ("열처리로", "HTF", ["히터", "열전대", "이송컨베이어", "가스공급장치"]),
    ("세척기", "WSH", ["노즐", "순환펌프", "필터", "히터"]),
    ("포장기", "PKG", ["실링바", "필름공급부", "감지센서", "구동모터"]),
    ("라벨러", "LBL", ["프린트헤드", "롤러", "감지센서", "구동모터"]),
    ("비전검사기", "VIS", ["카메라", "조명", "렌즈", "검사프로그램"]),
    ("X-ray검사기", "XRY", ["X선관", "검출기", "컨베이어", "차폐문"]),
    ("3차원측정기", "CMM", ["프로브", "에어베어링", "스케일", "컨트롤러"]),
    ("SMT마운터", "SMT", ["흡착노즐", "피더", "헤드", "비전카메라"]),
    ("리플로우오븐", "RFL", ["히터존", "컨베이어체인", "순환팬", "온도센서"]),
    ("스크린프린터", "SPR", ["스퀴지", "스텐실", "정렬카메라", "구동축"]),
    ("AOI검사기", "AOI", ["카메라", "조명", "이송축", "검사프로그램"]),
    ("디스펜서", "DSP", ["니들", "시린지", "이송축", "압력조절기"]),
    ("레이저마킹기", "LSR", ["레이저소스", "갈바노미러", "냉각기", "집속렌즈"]),
    ("레이저절단기", "LCT", ["레이저소스", "절단헤드", "보조가스라인", "이송축"]),
    ("와이어컷방전기", "EDM", ["와이어", "가공액펌프", "이송축", "전원부"]),
    ("3D프린터", "TDP", ["노즐", "히팅베드", "압출기", "스텝모터"]),
    ("믹서", "MIX", ["교반기", "감속기", "메커니컬씰", "구동모터"]),
    ("반응기", "RCT", ["교반기", "재킷", "온도센서", "제어밸브"]),
    ("원심분리기", "CNT", ["로터", "베어링", "구동모터", "진동센서"]),
    ("충진기", "FIL", ["충진노즐", "계량기", "서보모터", "실린더"]),
    ("팔레타이저", "PLT", ["그리퍼", "로봇암", "컨베이어", "안전센서"]),
    ("AGV", "AGV", ["배터리", "구동바퀴", "라이다", "충전접점"]),
    ("승강기", "ELV", ["권상기", "승강장도어", "제어반", "로프"]),
    ("비상발전기", "GEN", ["엔진", "발전기본체", "AVR", "냉각수라인"]),
]

# 에러코드 카탈로그: (부품분류, 증상, [원인], [조치]). {c}는 설비 종류별 부품명.
GENERIC = [
    ("통신", "통신 케이블 불량", ["커넥터 접촉 불량", "케이블 단선", "케이블 피복 손상"], ["케이블 교체", "커넥터 재체결"]),
    ("통신", "이더넷 통신 타임아웃", ["스위치 포트 불량", "IP 설정 충돌", "랜케이블 불량"], ["랜케이블 교체", "IP 재설정", "스위치 포트 변경"]),
    ("통신", "시리얼 통신 에러", ["통신 파라미터 불일치", "노이즈 유입"], ["파라미터 재설정", "실드 케이블 적용"]),
    ("통신", "필드버스 연결 끊김", ["터미네이터 불량", "슬레이브 모듈 불량"], ["터미네이터 교체", "모듈 교체"]),
    ("통신", "무선 통신 불안정", ["AP 신호 간섭", "안테나 불량"], ["채널 변경", "안테나 교체"]),
    ("PC", "제어 PC 부팅 불량", ["SSD 불량", "메인보드 불량", "전원 어댑터 불량"], ["SSD 교체", "PC 교체", "어댑터 교체"]),
    ("PC", "제어 PC 프리징", ["메모리 부족", "백그라운드 프로그램 충돌"], ["재부팅 및 프로세스 정리", "메모리 증설"]),
    ("PC", "HMI 화면 멈춤", ["터치패널 불량", "그래픽 드라이버 오류"], ["패널 교체", "드라이버 재설치"]),
    ("PC", "프로그램 오류 팝업 발생", ["소프트웨어 버그", "설정 파일 손상"], ["프로그램 재설치", "백업 복원"]),
    ("PC", "레시피 불러오기 실패", ["레시피 파일 손상", "저장 경로 접근 불가"], ["백업 복원", "경로 권한 재설정"]),
    ("PC", "데이터 저장 실패", ["저장 공간 부족", "DB 연결 오류"], ["불필요 데이터 정리", "DB 서비스 재시작"]),
    ("PC", "라이선스 만료 경고", ["라이선스 기간 만료"], ["라이선스 갱신"]),
    ("전원", "메인 전원 차단", ["차단기 트립", "입력 전압 이상"], ["차단기 리셋 후 모니터링", "입력 전원 점검"]),
    ("전원", "UPS 경보 발생", ["UPS 배터리 수명 도래", "UPS 과부하"], ["UPS 배터리 교체", "부하 분산"]),
    ("전원", "SMPS 출력 불량", ["SMPS 내부 소자 불량", "과부하"], ["SMPS 교체"]),
    ("전원", "과전류 알람 발생", ["배선 단락", "부하 증가"], ["배선 점검 및 보수", "부하 점검"]),
    ("전원", "접지 불량 알람", ["접지선 풀림", "접지 저항 상승"], ["접지 재체결", "접지 보강"]),
    ("제어", "PLC 모듈 오류", ["I/O 모듈 불량", "CPU 모듈 불량"], ["모듈 교체"]),
    ("제어", "인터락 알람", ["도어 스위치 접점 불량", "안전 릴레이 불량"], ["스위치 교체", "릴레이 교체"]),
    ("제어", "비상정지 회로 이상", ["E-STOP 접점 불량", "배선 접촉 불량"], ["접점 교체", "배선 재체결"]),
    ("제어", "파라미터 이상 알람", ["파라미터 오설정", "백업 배터리 방전"], ["파라미터 재설정", "백업 배터리 교체"]),
]
PER_COMP = [
    ("기구", "{c} 이상 소음 발생", ["{c} 마모", "윤활 부족", "이물질 유입"], ["{c} 교체", "그리스 주입", "청소 후 재조립"]),
    ("기구", "{c} 마모 과다", ["수명 도래", "가동량 증가"], ["{c} 교체"]),
    ("기구", "{c} 파손 확인", ["과부하", "충돌"], ["{c} 교체", "{c} 수리"]),
    ("기구", "{c} 체결 풀림", ["진동에 의한 풀림", "체결 토크 부족"], ["{c} 재체결 및 정렬"]),
    ("구동", "{c} 구동 불가", ["모터 소손", "드라이버 불량", "배선 단선"], ["모터 교체", "드라이버 교체", "배선 재결선"]),
    ("구동", "{c} 과부하 알람", ["기계적 간섭", "부하 증가"], ["간섭 제거", "부하 점검"]),
    ("구동", "{c} 속도 편차 발생", ["엔코더 불량", "게인 설정 오류"], ["엔코더 교체", "게인 재설정"]),
    ("센서", "{c} 위치 센서 미감지", ["센서 감도 저하", "센서 위치 틀어짐"], ["센서 위치 재조정", "센서 교체"]),
    ("센서", "{c} 센서 오감지", ["이물질 부착", "노이즈 유입"], ["센서 청소", "접지 보강"]),
    ("유공압", "{c} 압력 저하", ["실(Seal) 손상", "필터 막힘"], ["실 교체", "필터 교체"]),
    ("유공압", "{c} 누유·누기 발생", ["호스 균열", "피팅 풀림"], ["호스 교체", "피팅 재체결"]),
    ("온도", "{c} 온도 편차 이상", ["열전대 불량", "히터 열화"], ["열전대 교체", "히터 교체"]),
    ("온도", "{c} 과열 알람", ["냉각 불량", "제어기 설정 오류"], ["냉각계통 점검", "설정 재조정"]),
    ("소모품", "{c} 수명 소진 알람", ["정기 교체 주기 초과"], ["{c} 교체"]),
    ("소모품", "{c} 교체 주기 도래", ["가동량 증가"], ["{c} 교체", "소모품 보충 후 교체"]),
]


def build_error_table(rng, comps):
    """설비 종류 하나의 에러코드 정의표: [(코드, 증상, 부품분류, 원인들, 조치들)].
    같은 증상이라도 코드 번호는 설비 종류마다 제각각이다."""
    pool = list(GENERIC)
    for c in comps:
        pool += [(cat, sy.format(c=c), [x.format(c=c) for x in ca], [x.format(c=c) for x in ac])
                 for cat, sy, ca, ac in PER_COMP]
    rng.shuffle(pool)
    pool = pool[:rng.randint(24, 40)]
    return [(f"E{i:02d}", sy, cat, ca, ac) for i, (cat, sy, ca, ac) in enumerate(pool, start=1)]


# 양식 4종: 컬럼명(표준 순서: 설비, 일자, 에러코드, 증상, 원인, 조치, 분류, 기타1, 기타2), 날짜 표기
FORMS = {
    "A": dict(cols=["설비코드", "발생일자", "에러코드", "증상", "원인", "조치내용", "부품분류", "담당자", "정지시간(분)"],
              date="cell", sheet="트러블이력"),
    "B": dict(cols=["호기", "일시", "Err No.", "현상", "고장원인", "처리내용", "부품구분", "작업자", "비용(원)"],
              date="%Y.%m.%d %H:%M", sheet="고장기록"),
    "C": dict(cols=["Equipment ID", "Date", "Error Code", "Problem", "Root Cause", "Action Taken", "Part Category", "Engineer", "Downtime(h)"],
              date="%Y/%m/%d", sheet="Trouble_Log"),
    "D": dict(cols=["설비번호", "접수일", "알람코드", "이상내용", "원인분석", "대책", "부품유형", "담당", "관련부품"],
              date="%Y%m%d", sheet="이력"),
}
STD = ["설비", "일자", "에러코드", "증상", "원인", "조치", "분류", "기타 항목", "기타 항목"]
STD_NOTE = ["설비 코드 기준", "날짜는 YYYY-MM-DD로 통일", "설비 종류별 에러코드표(아래) 기준",
            "에러코드표의 증상과 일치", "", "", "부품분류(통신·PC·전원 등)",
            "표에 없는 컬럼 → 기타 항목", "표에 없는 컬럼 → 기타 항목"]
FORM_OF = "ABCD"

SURNAMES = list("김이박최정강조윤장임한오서신권황안송류홍")
GIVEN = ["민준", "서연", "지호", "하윤", "도윤", "지우", "현우", "수빈", "준서", "유진", "태양", "예은", "성민", "다은", "재현"]
PERIODS = [("2015-2020", date(2015, 1, 1), date(2020, 12, 31)),
           ("2021-2026", date(2021, 1, 1), END)]
UNKNOWN_CAUSE = ["", "알수없음", "원인불명", "미확인", "원인 파악 못함", "부품 교체 후 정상 작동", "모름", "?"]


def make_row(rng, table, weights, code, when, form_key):
    ec, symptom, cat, causes, actions = rng.choices(table, weights)[0]
    cause, action = rng.choice(causes), rng.choice(actions)
    who = rng.choice(SURNAMES) + rng.choice(GIVEN)
    minutes = rng.choice([20, 30, 45, 60, 90, 120, 180, 240, 480, 720])
    if form_key == "A":
        extra = (who, minutes)
    elif form_key == "B":
        extra = (who, minutes * rng.randint(800, 4000) // 10 * 10)
    elif form_key == "C":
        extra = (who, round(minutes / 60, 1))
    else:
        extra = (who, symptom.split()[0])
    return [code, when, ec, symptom, cause, action, cat, *extra]


def date_text(v):
    return v.isoformat() if isinstance(v, date) else str(v)


def make_messy(rng, data):
    """사람이 손으로 쓴 것처럼 일부 행을 흐트러뜨린다.
    행 위치: 0 설비, 1 일자, 2 에러코드, 3 증상, 4 원인, 5 조치, 6 분류."""
    out = []
    for row in data:
        row = list(row)
        if "교체" in str(row[5]):  # 부품만 갈았더니 나았다 → 원인을 안 적는 경우가 많다
            if rng.random() < 0.10:
                row[4] = rng.choice(UNKNOWN_CAUSE)
        elif rng.random() < 0.02:
            row[4] = rng.choice(UNKNOWN_CAUSE)
        r = rng.random()
        if r < 0.007:  # 터무니없는 날짜
            s = date_text(row[1])
            row[1] = str(int(s[:4]) + rng.choice([60, 100, 900])) + s[4:]
        elif r < 0.014:  # 날짜 칸에 에러코드까지 같이 적어서, 뒤 칸이 한 칸씩 앞당겨짐
            row = [row[0], f"{date_text(row[1])} {row[2]}"] + row[3:] + [None]
        elif r < 0.020:
            row[1] = "미상"
        elif r < 0.027:
            row[0] = None
        elif r < 0.039 and row[0]:  # 설비코드 표기 흔들림
            row[0] = rng.choice([row[0].lower(), row[0] + " ", row[0].replace("-", ""), " " + row[0]])
        elif r < 0.047:  # 에러코드 표기 흔들림
            row[2] = rng.choice([row[2].lower(), "E " + row[2][1:], row[2][1:], row[2] + " "])
        elif r < 0.051:  # 정의표에 없는 에러코드
            row[2] = "E99"
        elif r < 0.057:
            row[6] = rng.choice([None, "기타", " "])
        elif r < 0.060:
            row[3] = None
        out.append(row)
        if rng.random() < 0.006:  # 같은 내용을 두 번 적은 행
            out.append(list(row))
    return out


def fmt_date(when, key):
    mode = FORMS[key]["date"]
    return when.date() if mode == "cell" else when.strftime(mode)


def build_index(ws, tname, prefix, units, form_key, period, rows, sheet, table):
    cols = FORMS[form_key]["cols"]
    ws.append(["원본 컬럼명", "표준 컬럼", "비고"])
    for c, s, n in zip(cols, STD, STD_NOTE):
        ws.append([c, s, n])
    ws["E1"], ws["F1"] = "항목", "값"
    meta = [("설비 종류", tname), ("설비 코드 범위", f"{prefix}-001 ~ {prefix}-{units:03d}"),
            ("대상 기간", period), ("데이터 시트명", sheet), ("양식", form_key),
            ("데이터 행 수", rows),
            ("날짜 표기", "실제 날짜 셀" if FORMS[form_key]["date"] == "cell" else FORMS[form_key]["date"])]
    for i, (k, v) in enumerate(meta, start=2):
        ws[f"E{i}"], ws[f"F{i}"] = k, v
    start = max(len(cols), len(meta)) + 4  # 에러코드표 머리글 행
    ws.cell(start - 1, 1, "에러코드 정의 (이 설비 종류 전용)")
    for j, h in enumerate(["에러코드", "증상", "부품분류"], start=1):
        ws.cell(start, j, h)
    for i, (ec, sy, cat, _, _) in enumerate(table, start=start + 1):
        ws.cell(i, 1, ec), ws.cell(i, 2, sy), ws.cell(i, 3, cat)
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDEBF7")
    ws.cell(start - 1, 1).font = bold
    for cell in ("A1", "B1", "C1", "E1", "F1", f"A{start}", f"B{start}", f"C{start}"):
        ws[cell].font, ws[cell].fill = bold, fill
    for i, w in enumerate([18, 30, 30, 3, 16, 26], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def main():
    if OUT.exists():
        sys.exit(f"이미 있습니다. 덮어쓰지 않고 중단합니다: {OUT}")
    OUT.mkdir(parents=True)
    rng = random.Random(SEED)

    unit_counts = [10] * 20 + [20] * 20
    rng.shuffle(unit_counts)
    assert sum(unit_counts) == 600 and len(TYPES) == 40

    equipment_rows, file_rows, error_rows, total = [], [], [], 0
    for idx, ((tname, prefix, comps), units) in enumerate(zip(TYPES, unit_counts), start=1):
        form_key = FORM_OF[(idx - 1) % 4]
        codes = [f"{prefix}-{n:03d}" for n in range(1, units + 1)]
        equipment_rows += [[c, tname] for c in codes]
        table = build_error_table(rng, comps)
        weights = [1 / (i + 1) ** 0.7 for i in range(len(table))]
        rng.shuffle(weights)  # 자주 나는 코드는 종류마다 다르다
        error_rows += [[tname, ec, sy, cat] for ec, sy, cat, _, _ in table]
        for pname, d0, d1 in PERIODS:
            n = rng.randint(50, 100)
            span = (d1 - d0).days
            days = sorted(rng.randint(0, span) for _ in range(n))
            data = []
            for off in days:
                when = datetime.combine(d0 + timedelta(days=off), datetime.min.time()) + timedelta(
                    hours=rng.randint(6, 22), minutes=rng.choice([0, 10, 20, 30, 40, 50]))
                row = make_row(rng, table, weights, rng.choice(codes), when, form_key)
                row[1] = fmt_date(when, form_key)
                data.append(row)
            data = make_messy(rng, data)
            n = len(data)

            wb = Workbook()
            ws_idx = wb.active
            ws_idx.title = "인덱스"
            sheet = FORMS[form_key]["sheet"]
            build_index(ws_idx, tname, prefix, units, form_key, pname, n, sheet, table)
            ws = wb.create_sheet(sheet)
            ws.append(FORMS[form_key]["cols"])
            for c in ws[1]:
                c.font = Font(bold=True)
            for row in data:
                ws.append(row)
            for i, w in enumerate([14, 18, 10, 30, 24, 26, 12, 12, 14], start=1):
                ws.column_dimensions[get_column_letter(i)].width = w
            if FORMS[form_key]["date"] == "cell":
                for r in range(2, ws.max_row + 1):
                    if isinstance(ws.cell(r, 2).value, date):
                        ws.cell(r, 2).number_format = "yyyy-mm-dd"
            fname = f"{idx:02d}_{tname}_{pname}.xlsx"
            wb.save(OUT / fname)
            file_rows.append([fname, tname, form_key, pname, n])
            total += n

    for name, header, rows in (
        ("설비목록.csv", ["설비코드", "설비종류"], equipment_rows),
        ("에러코드표.csv", ["설비종류", "에러코드", "증상", "부품분류"], error_rows),
        ("파일목록.csv", ["파일명", "설비종류", "양식", "기간", "행수"], file_rows),
    ):
        with open(OUT / name, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
    print(f"설비 종류 {len(TYPES)}, 설비 {len(equipment_rows)}대, 파일 {len(file_rows)}개, "
          f"데이터 행 {total}, 에러코드 정의 {len(error_rows)}개")


if __name__ == "__main__":
    main()

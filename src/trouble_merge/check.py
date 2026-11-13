"""'확인 필요' 판정과 이유·원래 값 문구 (T028, FR-005, FR-009, FR-011, FR-013, FR-015).

판단이 안 되는 행은 빼지 않고 해당 칸에 '확인 필요'를 쓰고, 이유와 원래 값을 남긴다.
이유와 원래 값은 설비, 일자, 증상, 원인, 조치, 분류 순서로 '; '로 잇는다.
"""

from datetime import date

from .dates import DEFAULT_MIN_DATE, looks_shifted, parse_date
from .mapping import UNKNOWN_CAUSE, norm_key
from .reader import is_blank

CHECK = "확인 필요"
EMPTY_NOTE = "(빈 칸)"
REQUIRED_TEXT = ("증상", "원인", "조치")


def text(value):
    """빈 칸이면 None, 아니면 앞뒤 공백을 뺀 문자열."""
    if is_blank(value):
        return None
    return str(value).strip()


def default_date_range():
    """기본 날짜 범위(T057): 2000-01-01부터 실행한 해의 12월 31일까지.

    상한을 '실행한 날'로 하면 같은 해의 예정·입력 지연 날짜(며칠 뒤)까지 막혀서, 해의 끝으로 둔다.
    2922년이나 60년 뒤 같은 터무니없는 날짜는 그대로 걸린다.
    """
    return DEFAULT_MIN_DATE, "%04d-12-31" % date.today().year


def judge(raw, conversions, date_range=None):
    """한 행의 여섯 표준 컬럼을 판정한다.

    raw: {'설비','일자','증상','원인','조치','분류'} -> 입력 칸의 원래 값
    date_range: ('YYYY-MM-DD' 하한, 상한). None이면 기본 범위.
    돌려주는 값: (최종 칸 값 사전, 이유 목록, 원래 값 목록)
    """
    low, high = date_range or default_date_range()
    final = {}
    reasons = []
    originals = []

    equipment = text(raw["설비"])
    if equipment is None:
        final["설비"] = CHECK
        reasons.append("설비 코드 빈 칸")
        originals.append("설비: " + EMPTY_NOTE)
    else:
        standard = conversions["설비"].get(norm_key(equipment))
        if standard is None:
            final["설비"] = CHECK
            reasons.append("설비 코드 미등록")
            originals.append("설비: " + equipment)
        else:
            final["설비"] = standard

    date_value = raw["일자"]
    if is_blank(date_value):
        final["일자"] = None  # 일자가 빈 것 자체는 확인 필요가 아니다(FR-013)
    else:
        iso = parse_date(date_value)
        reason = None
        if iso is None:
            # 날짜 뒤에 다른 글자가 붙어 있으면 칸이 밀린 것으로 의심한다(② 값 오류, T058)
            reason = "칸 밀림 의심" if looks_shifted(date_value) else "일자 해석 불가"
        elif not (low <= iso <= high):
            reason = "일자 범위 밖"  # 형식은 맞지만 있을 수 없는 날짜(② 값 오류, T057)
        if reason is None:
            final["일자"] = iso
        else:
            final["일자"] = CHECK
            reasons.append(reason)
            originals.append("일자: " + (text(date_value) or str(date_value)))

    unknown_cause = conversions.get(UNKNOWN_CAUSE, {})
    for column in REQUIRED_TEXT:
        value = text(raw[column])
        if value is None:
            final[column] = CHECK
            reasons.append(column + " 빈 칸")
            originals.append("%s: %s" % (column, EMPTY_NOTE))
        elif column == "원인" and norm_key(value) in unknown_cause:
            final[column] = CHECK  # '모름'처럼 정보가 없는 원인(③ 정보 부족, T059)
            reasons.append("원인 미상")
            originals.append("원인: " + value)
        else:
            final[column] = value

    category = text(raw["분류"])
    if category is None:
        final["분류"] = None  # 분류가 빈 것은 확인 필요가 아니다(FR-015)
    else:
        standard = conversions["분류"].get(norm_key(category))
        if standard is None:
            final["분류"] = CHECK
            reasons.append("분류 미등록")
            originals.append("분류: " + category)
        else:
            final["분류"] = standard

    return final, reasons, originals

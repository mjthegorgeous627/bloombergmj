"""
자동화가 오늘 돌아도 되는 날인지 판단하는 공용 모듈 (2026-09-13 신설).

발단: SAP_Morning_Routine(매일 08:30 Windows 작업 스케줄러 트리거)이 요일
구분 없이 매일(DaysInterval=1, 요일 제한 없음) 실행되고 있었다. evening_
shutdown.py가 매일 19:00에 루프를 끄고 stop_automation.flag를 남겨두더라도,
다음날 아침 morning_routine.py가 무조건 startup.py를 다시 띄우고 그 안의
main.py run_loop() 진입부가 STOP_FILE을 무조건 지워버린다(main.py 참고) -
그래서 주말/공휴일 아침에도 그대로 SAP에 자동 접속해 무인 상태로 돌다가
에러 팝업이 쌓이는 사고가 실제로 발생함(사용자 실측, 2026-09-13: 휴일에
와보니 그동안 자동화가 돌고 있어 에러창이 떠 있었음).

이 모듈은 "오늘 자동으로 뭔가를 새로 시작/재시작해도 되는가"를 한 곳에서
판단해서, 자동 기동/재기동 경로 세 곳(morning_routine.py의 최초 기동,
standalone watchdog.py, workbench_app.py의 _sap_loop_watchdog_thread)이
전부 같은 기준을 쓰게 한다. evening_shutdown.py는 손대지 않음 - 매일
꺼주는 동작 자체는 요일과 무관하게 항상 안전(무언가 켜져 있으면 끄고,
없으면 조용히 아무 일도 안 함)하다.

주의: 이건 "자동으로 시작/재시작하는 것"만 막는다. workbench 페이지의
"SAP 시작" 버튼 등 사용자의 명시적 수동 동작은 이 판단과 무관하게 항상
그대로 작동한다(의도적으로 안 건드림 - 급한 일로 주말에 직접 켜야 하는
경우까지 막을 이유는 없음).

공휴일 판단은 holidays 패키지(KR)를 쓴다 - 설날/추석 등 음력 공휴일과
대체공휴일까지 계산해준다(pip install holidays로 설치됨, 2026-09-13).
연도를 미리 지정하지 않아도 조회 시점에 필요한 연도를 자동으로 채워주므로
매년 갱신할 필요가 없다. 다만 그 해에 정부가 임시로 지정하는 "임시공휴일"
(선거일 등, 사전 예측 불가능한 것)은 라이브러리 업데이트 전까지는 못 잡을
수 있다 - 이런 예외적인 날은 launcher.py/workbench의 "SAP 자동루프 중지"를
전날 미리 눌러두는 수동 대응이 필요하다.
"""

import logging
from datetime import date

logger = logging.getLogger(__name__)

try:
    import holidays as _holidays_lib
    _KR_HOLIDAYS = _holidays_lib.KR()
except Exception as _exc:  # pragma: no cover - holidays 패키지 자체가 없거나 깨진 경우
    _KR_HOLIDAYS = None
    logger.warning(f"holidays 패키지 로드 실패 - 공휴일 판단 불가, 주말만 걸러짐: {_exc}")

_WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토", "일"]


def skip_reason(today=None):
    """오늘 자동 기동/재기동을 건너뛰어야 하는 이유 문자열. 건너뛸 필요
    없으면 None. today는 테스트/수동 확인용 - 보통 생략(오늘 날짜 사용)."""
    d = today or date.today()
    if d.weekday() >= 5:  # 5=토, 6=일
        return f"주말({_WEEKDAY_NAMES[d.weekday()]}요일)"
    if _KR_HOLIDAYS is not None:
        try:
            if d in _KR_HOLIDAYS:
                return f"한국 공휴일({_KR_HOLIDAYS.get(d)})"
        except Exception as exc:
            logger.warning(f"공휴일 조회 실패 (무시하고 진행): {exc}")
    return None


def is_automation_day(today=None):
    """True면 오늘 자동 기동/재기동해도 되는 평일·비공휴일."""
    return skip_reason(today) is None


def kr_holiday_labels(years):
    """주어진 연도들의 한국 공휴일 {ISO 날짜: 이름} dict (2026-09-13 workbench
    달력 빨간날 표시용으로 추가). holidays 패키지가 없으면 빈 dict - 그래도
    주말 표시 자체는 workbench 쪽에서 JS의 Date.getDay()로 따로 계산하므로
    영향 없다."""
    if _KR_HOLIDAYS is None:
        return {}
    try:
        hs = _holidays_lib.KR(years=list(years))
        return {d.isoformat(): name for d, name in hs.items()}
    except Exception as exc:
        logger.warning(f"공휴일 목록 조회 실패 (무시하고 빈 목록 반환): {exc}")
        return {}

"""
NWBC '+' 새 탭 버튼이 자동화 실행 중 가끔 멈추는 버그(NWBC 8.00 클라이언트 자체 문제,
BLOOMBERG_HANDOFF.md Session 3 참고)를 우회해서, SAP GUI Scripting API로 직접 새 세션을
연다. 자동화가 4개 세션을 처음 만들 때 쓰는 것과 동일한 방식(session.createSession())이라
NWBC UI의 새 탭 로직과는 다른 내부 경로를 탄다.

사용법:
  python open_session.py                  → 빈 새 세션만 열기
  python open_session.py va02              → 새 세션 열고 VA02 진입
  python open_session.py order 67086095    → 새 세션 열고, 오더번호 앞자리로
                                              VA02(6*)/VA03(그 외) 자동판별해서
                                              그 오더를 직접 열기 (workbench
                                              "오더 바로가기" 버튼이 이걸 씀 -
                                              데이터 수집/카카오전송 없이 화면만
                                              띄워서 사용자가 직접 보거나 고칠
                                              용도, manual_order_handler.py의
                                              run_manual_order()과는 다름)
"""

import sys
import time
import logging
from pathlib import Path

from sap_handler import get_scripting_engine

BASE_DIR = Path(__file__).resolve().parent

# workbench_app.py가 이 스크립트를 CREATE_NO_WINDOW로 백그라운드 실행하므로
# (sap_ops.sap_open_session/sap_open_order 참고) 콘솔이 없어 기본 stderr
# 로깅은 아무도 못 본다 - 다른 자동화 스크립트들과 같은 방식으로 파일에도
# 남긴다.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.FileHandler(BASE_DIR / "open_session.log", encoding="utf-8"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# SAP GUI 연결당 세션 수 하드 리밋 (zrec_handler.py/sn_relo_handler.py와 동일).
# 자동루프 고정 4개 + ZREC류 세션까지 이미 여러 개가 항상 떠있는 상태라, 이
# 체크 없이 무작정 createSession()을 시도하면 SAP가 에러 없이 그냥 먹통이
# 되는 경우가 많다(2026-09 실측: "바로가기"를 반복 클릭한 뒤 SAP 창이
# 응답 없음 상태로 굳는 문제).
MAX_SESSIONS = 6


def _create_new_session(conn):
    """conn.Children(0).createSession() 후 실제로 새로 생긴 세션을 찾아 반환.
    "세션 개수가 늘었다 → 마지막 child가 새 세션"이라는 가정은, 방금 닫힌
    세션의 SessionNumber가 재사용될 때 엉뚱한 기존 세션에 명령을 잘못
    보내는 실제 사고를 낸 적이 있다 (zrec_handler.py의 2026-08-18 실측
    사고 기록 참고 - ZIH08 세션이 통째로 ZREC로 덮어써짐). 생성 전/후
    SessionNumber 집합을 비교해서 진짜 새로 생긴 번호만 새 세션으로
    인정하는, zrec_handler.py와 동일한 안전한 방식을 쓴다."""
    if conn.Children.Count >= MAX_SESSIONS:
        raise RuntimeError(
            f"SAP GUI 세션이 이미 최대({MAX_SESSIONS}개)입니다 - "
            "새 세션을 열려면 SAP에서 안 쓰는 세션을 먼저 닫으세요."
        )
    before = {conn.Children(i).Info.SessionNumber for i in range(conn.Children.Count)}
    conn.Children(0).createSession()
    deadline = time.time() + 10
    while time.time() < deadline:
        time.sleep(0.5)
        for i in range(conn.Children.Count):
            s = conn.Children(i)
            if s.Info.SessionNumber not in before:
                return s
    return None


def open_new_session(tcode=None):
    app = get_scripting_engine()
    conn = app.Children(0)
    before = conn.Children.Count
    logger.info(f"현재 세션 수: {before}개 → 새 세션 생성 중...")

    try:
        new_session = _create_new_session(conn)
    except RuntimeError as e:
        logger.error(str(e))
        return False

    if new_session is None:
        logger.error("새 세션이 시간 내에 나타나지 않음 (생성 실패 또는 지연)")
        return False

    logger.info(f"새 세션 생성 완료 (세션 수: {before} → {conn.Children.Count})")

    if tcode:
        time.sleep(1)
        try:
            new_session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
            new_session.findById("wnd[0]").sendVKey(0)
            logger.info(f"{tcode} 진입 완료")
        except Exception as e:
            logger.warning(f"{tcode} 진입 실패 (세션 자체는 열림): {e}")

    return True


def open_order(order_num):
    """새 세션을 열고 오더번호 앞자리로 VA02(6*)/VA03(그 외)을 자동판별해서
    그 오더 화면으로 직접 들어간다 - manual_order_handler._open_order()와
    같은 판별규칙·필드ID(ctxtVBAK-VBELN)·Information 팝업 처리를 쓰지만,
    저기서 뒤따르는 항목 추출/workbench 반영/카카오 전송은 전혀 안 한다
    (사용자가 그냥 오더를 눈으로 보거나 직접 고치고 싶을 때 쓰는 순수
    바로가기)."""
    order_num = str(order_num).strip()
    if not order_num:
        logger.error("오더번호가 비어 있습니다.")
        return False

    tcode = "VA02" if order_num.startswith("6") else "VA03"

    app = get_scripting_engine()
    conn = app.Children(0)
    before = conn.Children.Count
    logger.info(f"오더 {order_num} → {tcode} 새 세션에서 열기 (현재 세션 수: {before}개)")

    try:
        new_session = _create_new_session(conn)
    except RuntimeError as e:
        logger.error(str(e))
        return False

    if new_session is None:
        logger.error("새 세션이 시간 내에 나타나지 않음 (생성 실패 또는 지연)")
        return False

    time.sleep(1)
    try:
        new_session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
        new_session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)
        new_session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order_num
        new_session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)

        # manual_order_handler._open_order()와 동일: "special handling notes"
        # 메모가 걸린 오더는 이 시점에 Information 팝업(wnd[1])이 떠서 오더
        # 화면을 가릴 수 있음 - 있으면 Enter로 닫아준다.
        try:
            popup = new_session.findById("wnd[1]")
            if popup.Text == "Information":
                popup.sendVKey(0)
                time.sleep(1)
        except Exception:
            pass

        title = new_session.findById("wnd[0]").Text
        logger.info(f"오더 {order_num} ({tcode}) 진입 완료: {title}")
    except Exception as e:
        logger.warning(f"오더 {order_num} ({tcode}) 진입 실패 (세션 자체는 열림): {e}")
        return False

    return True


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1].lower() == "order":
        ok = open_order(sys.argv[2])
    else:
        tcode_arg = sys.argv[1] if len(sys.argv) > 1 else None
        ok = open_new_session(tcode_arg)
    sys.exit(0 if ok else 1)

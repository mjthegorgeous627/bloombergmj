"""
SAP 자동화 시작 스크립트.

사용법:
  1. SAP 로그인 (창 1개만 열린 상태)
  2. python startup.py

자동으로:
  - SAP 세션 4개 설정 (VL06O / VL10G / ZRMA RLKR / ZRMA Q2)
  - Excel 파일 열기
  - 누락 오더 catchup
  - 이후 10분마다 자동 실행
"""

import win32com.client
import time
import os
import sys
import logging
from datetime import datetime

# ── 로깅 설정 ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)

from config import EXCEL_PATH, REFRESH_INTERVAL_MINUTES
from sap_handler import navigate_to_vl06o_list
from vl10g_handler import navigate_to_vl10g
from zrma_handler import navigate_to_zrma_q
from main import run_once, run_loop


def get_sap_connection():
    """SAP GUI 연결. 로그인 안 된 경우 안내 후 종료."""
    try:
        sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        conn = sap.Children(0)
        if conn.Children.Count == 0:
            print("\n[오류] SAP 세션이 없습니다. SAP에 로그인 후 다시 실행하세요.")
            sys.exit(1)
        return sap, conn
    except Exception as e:
        print(f"\n[오류] SAP GUI에 연결할 수 없습니다: {e}")
        print("SAP GUI를 실행하고 로그인 후 다시 실행하세요.")
        sys.exit(1)


def setup_sap_sessions(conn):
    """SAP 세션 4개 열고 각 화면 설정."""
    current = conn.Children.Count
    logger.info(f"현재 SAP 세션 수: {current}개")

    # 4개 미만이면 추가 생성
    while conn.Children.Count < 4:
        try:
            conn.Children(0).createSession()
            time.sleep(2.5)
            logger.info(f"세션 추가 → 현재 {conn.Children.Count}개")
        except Exception as e:
            logger.error(f"세션 생성 실패: {e}")
            break

    if conn.Children.Count < 4:
        logger.warning(f"세션이 {conn.Children.Count}개뿐 - 계속 진행")

    # ── 세션0: VL06O ────────────────────────────────────────────
    try:
        sess0 = conn.Children(0)
        logger.info("세션0 VL06O 설정 중...")
        navigate_to_vl06o_list(sess0)
        logger.info("세션0 VL06O ✓")
    except Exception as e:
        logger.error(f"세션0 VL06O 설정 실패: {e}")

    # ── 세션1: VL10G ────────────────────────────────────────────
    try:
        sess1 = conn.Children(1)
        logger.info("세션1 VL10G 설정 중...")
        navigate_to_vl10g(sess1)
        logger.info("세션1 VL10G ✓")
    except Exception as e:
        logger.error(f"세션1 VL10G 설정 실패: {e}")

    # ── 세션2: ZRMA_Q RLKR ──────────────────────────────────────
    try:
        sess2 = conn.Children(2)
        logger.info("세션2 ZRMA RLKR 설정 중...")
        navigate_to_zrma_q(sess2, 'rlkr', 'next_month')
        logger.info("세션2 ZRMA RLKR ✓")
    except Exception as e:
        logger.error(f"세션2 ZRMA RLKR 설정 실패: {e}")

    # ── 세션3: ZRMA_Q Q2 ────────────────────────────────────────
    try:
        sess3 = conn.Children(3)
        logger.info("세션3 ZRMA Q2 설정 중...")
        navigate_to_zrma_q(sess3, 'q2', 'year_end')
        logger.info("세션3 ZRMA Q2 ✓")
    except Exception as e:
        logger.error(f"세션3 ZRMA Q2 설정 실패: {e}")

    logger.info("SAP 4개 세션 설정 완료")


def open_excel():
    """Excel 파일 열기."""
    try:
        if os.path.exists(EXCEL_PATH):
            os.startfile(EXCEL_PATH)
            time.sleep(4)
            logger.info(f"Excel 열기 완료: {EXCEL_PATH}")
        else:
            logger.warning(f"Excel 파일 없음: {EXCEL_PATH}")
    except Exception as e:
        logger.error(f"Excel 열기 실패: {e}")


if __name__ == "__main__":
    print("=" * 55)
    print("  SAP 배송 자동화 시작")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 55)

    # 1. SAP 연결 확인
    sap, conn = get_sap_connection()
    logger.info(f"SAP 연결 성공 (세션 {conn.Children.Count}개)")

    # 2. SAP 4개 세션 화면 설정
    setup_sap_sessions(conn)

    # 3. Excel 열기
    open_excel()

    # 4. 누락 오더 catchup (Excel 기준)
    logger.info("누락 오더 확인 중 (catchup)...")
    run_once(excel_only=True)

    # 5. 10분마다 자동 반복
    logger.info(f"자동 반복 시작 ({REFRESH_INTERVAL_MINUTES}분 간격)")
    run_loop()

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
import subprocess
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

# SAP Logon 실행 파일 경로
SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"


def get_sap_connection():
    """SAP GUI 연결. 없으면 saplogon.exe 실행 후 대기."""
    # SAP GUI가 실행 중인지 확인
    for attempt in range(2):
        try:
            sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            conn = sap.Children(0)
            return sap, conn
        except Exception:
            if attempt == 0:
                logger.info("SAP GUI 미실행 → saplogon.exe 실행 중...")
                try:
                    subprocess.Popen(SAPLOGON_PATH)
                    time.sleep(5)
                except Exception as e:
                    logger.error(f"saplogon.exe 실행 실패: {e}")
                    print("SAP GUI를 직접 실행 후 다시 시도하세요.")
                    sys.exit(1)

    print("[오류] SAP GUI 연결 실패. 직접 실행 후 다시 시도하세요.")
    sys.exit(1)


def sap_login(conn):
    """
    SAP 로그인 화면이 떠 있으면 자동 로그인.
    이미 로그인된 세션이 있으면 스킵.
    """
    try:
        from credentials import SAP_CLIENT, SAP_USER, SAP_PASSWORD, SAP_LANGUAGE
    except ImportError:
        logger.warning("credentials.py 없음 - 수동 로그인 필요")
        return

    if not SAP_USER or not SAP_PASSWORD:
        logger.warning("credentials.py에 ID/PW 미입력 - 수동 로그인 필요")
        return

    # 연결된 세션이 없으면 로그인 화면 없음 → 대기
    if conn.Children.Count == 0:
        logger.info("SAP 연결 대기 중 (시스템 선택 후 로그인 화면 뜰 때까지)...")
        for _ in range(30):  # 최대 30초 대기
            time.sleep(1)
            if conn.Children.Count > 0:
                break
        if conn.Children.Count == 0:
            logger.error("로그인 화면 대기 시간 초과")
            return

    sess = conn.Children(0)
    title = sess.findById("wnd[0]").Text

    # 로그인 화면 확인 ("SAP" 타이틀이거나 빈 경우)
    if "SAP" not in title and title.strip():
        logger.info(f"이미 로그인됨: {title}")
        return

    logger.info("SAP 로그인 중...")
    try:
        wnd = sess.findById("wnd[0]")
        sess.findById("wnd[0]/usr/txtRSYST-MANDT").text = SAP_CLIENT
        sess.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
        sess.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASSWORD
        sess.findById("wnd[0]/usr/txtRSYST-LANGU").text  = SAP_LANGUAGE
        wnd.sendVKey(0)  # Enter
        time.sleep(3)

        # 중복 로그인 팝업 처리 (다른 세션 종료 후 계속)
        try:
            popup = sess.findById("wnd[1]")
            popup_text = popup.Text
            if "logon" in popup_text.lower() or "session" in popup_text.lower():
                # "Continue with this logon" 선택 (라디오버튼 또는 버튼)
                for fid in [
                    "wnd[1]/usr/radMULTI_LOGON_OPT2",  # 기존 세션 유지
                    "wnd[1]/usr/radMULTI_LOGON_OPT1",  # 기존 세션 종료
                ]:
                    try:
                        sess.findById(fid).select()
                        break
                    except Exception:
                        continue
                sess.findById("wnd[1]").sendVKey(0)
                time.sleep(2)
        except Exception:
            pass  # 팝업 없음 = 정상

        logger.info("SAP 로그인 완료")
    except Exception as e:
        logger.error(f"SAP 로그인 실패: {e}")


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

    # 1. SAP GUI 연결 (없으면 saplogon.exe 자동 실행)
    sap, conn = get_sap_connection()

    # 2. 자동 로그인 (로그인 화면인 경우)
    sap_login(conn)
    logger.info(f"SAP 세션 수: {conn.Children.Count}개")

    # 3. SAP 4개 세션 화면 설정
    setup_sap_sessions(conn)

    # 4. Excel 열기
    open_excel()

    # 5. 누락 오더 catchup (Excel 기준)
    logger.info("누락 오더 확인 중 (catchup)...")
    run_once(excel_only=True)

    # 6. 10분마다 자동 반복
    logger.info(f"자동 반복 시작 ({REFRESH_INTERVAL_MINUTES}분 간격)")
    run_loop()

"""
SAP 자동화 시작 스크립트.

사용법:
  python startup.py

자동으로:
  - SAP 실행 + 자동 로그인 (sapshcut.exe 사용)
  - SAP 세션 4개 설정 (VL06O / VL10G / ZRMA RLKR / ZRMA Q2)
  - Excel 파일 열기
  - 누락 오더 catchup
  - 이후 10분마다 자동 실행
"""

import win32com.client
import win32gui
import win32con
import subprocess
import ctypes
import threading
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

SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"


def _get_scripting_engine():
    """SAP GUI Scripting Engine 객체 반환. 없으면 None."""
    try:
        return win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except Exception:
        return None


def _is_login_screen(sess):
    """
    로그인 화면 여부 판단.
    사용자 ID 입력 필드(txtRSYST-BNAME)가 있으면 로그인 화면.
    없으면 이미 로그인된 상태.
    """
    try:
        sess.findById("wnd[0]/usr/txtRSYST-BNAME")
        return True
    except Exception:
        return False


def _dismiss_popup(sess):
    """wnd[1] 팝업(System Status 등)이 열려 있으면 닫기."""
    try:
        sess.findById("wnd[1]").sendVKey(12)  # F12 = 닫기/취소
        time.sleep(0.5)
    except Exception:
        pass


def _find_window_by_title(partial_title):
    """창 제목에 partial_title이 포함된 최상위 창 핸들 반환."""
    result = []
    def cb(hwnd, _):
        if partial_title.lower() in win32gui.GetWindowText(hwnd).lower() and win32gui.IsWindowVisible(hwnd):
            result.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return result[0] if result else None


def _click_ok_on_saplogon_popup():
    """
    'saplogon' 제목의 보안 경고 팝업에서 확인 버튼 클릭.
    '스크립트가 GUI에 연결되고 있습니다' 팝업 자동 처리.
    """
    hwnd = _find_window_by_title("saplogon")
    if not hwnd:
        return False
    # 모든 Button 컨트롤 수집 후 첫 번째 클릭 (확인 버튼)
    buttons = []
    def find_btns(h, _):
        if win32gui.GetClassName(h) == "Button":
            buttons.append(h)
        return True
    try:
        win32gui.EnumChildWindows(hwnd, find_btns, None)
    except Exception:
        pass
    if buttons:
        # SetForegroundWindow 없이 직접 BM_CLICK 전송
        win32gui.SendMessage(buttons[0], win32con.BM_CLICK, 0, 0)
        logger.info("SAP 스크립팅 보안 팝업 → 확인 클릭")
        return True
    return False


def _get_scripting_engine():
    """
    SAP GUI Scripting Engine 연결 (1회 호출).
    GetObject가 블로킹되는 동안 별도 스레드에서 보안 팝업 자동 클릭.
    """
    stop_flag = threading.Event()

    def popup_clicker():
        while not stop_flag.is_set():
            _click_ok_on_saplogon_popup()
            time.sleep(0.3)

    t = threading.Thread(target=popup_clicker, daemon=True)
    t.start()
    try:
        engine = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        return engine
    except Exception:
        return None
    finally:
        stop_flag.set()


def _send_key(vk_code):
    """가상 키 코드로 키 입력 전송."""
    user32 = ctypes.windll.user32
    user32.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(vk_code, 0, 0x0002, 0)


def _send_enter():
    _send_key(0x0D)  # Enter


def _handle_rp1_login_dialog(user, password):
    """
    'RP1 BEST CORE SYSTEM' 커스텀 로그인 팝업 자동 처리.
    사용자 이름(BLOOMBERG 지우고 실제 ID 입력) + 비밀번호 입력 → 로그온 클릭.
    반환: True(처리됨) / False(팝업 없음)
    """
    hwnd = _find_window_by_title("RP1 BEST CORE SYSTEM")
    if not hwnd:
        return False

    # Edit 컨트롤 수집 (순서: 사용자이름, 비밀번호)
    edits = []
    def find_edits(h, _):
        if win32gui.GetClassName(h) == "Edit":
            edits.append(h)
        return True
    try:
        win32gui.EnumChildWindows(hwnd, find_edits, None)
    except Exception:
        pass

    if len(edits) < 2:
        logger.warning(f"RP1 로그인 팝업: Edit 컨트롤 {len(edits)}개 발견 (2개 필요)")
        return False

    # 사용자 이름 필드: 기존값 지우고 SAP_USER 입력
    win32gui.SendMessage(edits[0], 0x000C, 0, user)   # WM_SETTEXT
    # 비밀번호 필드 입력
    win32gui.SendMessage(edits[1], 0x000C, 0, password)

    # 로그온 버튼 클릭 (첫 번째 Button = 로그온(L))
    buttons = []
    def find_btns(h, _):
        if win32gui.GetClassName(h) == "Button":
            buttons.append(h)
        return True
    try:
        win32gui.EnumChildWindows(hwnd, find_btns, None)
    except Exception:
        pass

    if buttons:
        win32gui.SendMessage(buttons[0], win32con.BM_CLICK, 0, 0)
        logger.info(f"RP1 로그인 팝업 → 사용자({user}) 입력 후 로그온 클릭")
        return True

    return False


def launch_sap_and_connect():
    """
    SAP 자동 실행 + 연결.

    1. 이미 로그인된 SAP 세션이 있으면 그대로 사용.
    2. SAP Logon 패드 실행 → Enter로 마지막 시스템 접속.
    3. 로그인 완료까지 대기 (스크립팅 보안 팝업 자동 처리).
    4. 로그인 화면이 남아있으면 credentials로 수동 로그인.

    반환: (sap, conn)
    """
    try:
        from credentials import SAP_CLIENT, SAP_USER, SAP_PASSWORD, SAP_LANGUAGE
    except ImportError:
        logger.error("credentials.py 없음 - SAP 자동 실행 불가")
        sys.exit(1)

    # ── 1. 이미 로그인된 세션 있으면 바로 사용 ────────────────
    sap = _get_scripting_engine()
    if sap and sap.Children.Count > 0:
        conn = sap.Children(0)
        if conn.Children.Count > 0:
            sess = conn.Children(0)
            _dismiss_popup(sess)
            if not _is_login_screen(sess):
                title = sess.findById("wnd[0]").Text
                logger.info(f"기존 SAP 세션 사용 중: {title}")
                return sap, conn

    # ── 2. .sap 바로가기 파일로 직접 접속 ────────────────────
    try:
        from credentials import SAP_SHORTCUT_PATH
    except ImportError:
        SAP_SHORTCUT_PATH = None

    if SAP_SHORTCUT_PATH and os.path.exists(SAP_SHORTCUT_PATH):
        logger.info(f"SAP 바로가기 실행: {SAP_SHORTCUT_PATH}")
        os.startfile(SAP_SHORTCUT_PATH)
    else:
        logger.warning(f"SAP 바로가기 파일 없음: {SAP_SHORTCUT_PATH}")
        logger.info("SAP Logon 실행 중...")
        subprocess.Popen(SAPLOGON_PATH)

    # ── 3. 'RP1 BEST CORE SYSTEM' 로그인 팝업 대기 및 자동 처리 ──
    logger.info("RP1 로그인 팝업 대기 중...")
    rp1_handled = False
    for _ in range(20):
        time.sleep(1)
        if _handle_rp1_login_dialog(SAP_USER, SAP_PASSWORD):
            rp1_handled = True
            break

    if not rp1_handled:
        logger.warning("RP1 로그인 팝업을 찾지 못함 - 수동 로그인 필요할 수 있음")

    # ── 4. SAP Scripting Engine 연결 (1회 호출) ──────────────
    logger.info("SAP Scripting Engine 연결 중...")
    time.sleep(5)  # 로그인 후 SAP 연결 완료 대기
    sap = _get_scripting_engine()

    if not sap:
        logger.error("SAP Scripting Engine 연결 실패")
        sys.exit(1)

    # ── 5. 로그인 완료까지 대기 (엔진 재사용) ────────────────
    for elapsed in range(60):
        time.sleep(1)
        try:
            if sap.Children.Count == 0:
                if elapsed % 10 == 9:
                    logger.info(f"  SAP 연결 대기 중... ({elapsed+1}초)")
                continue
            conn = sap.Children(0)
            if conn.Children.Count == 0:
                continue
            sess = conn.Children(0)
            _dismiss_popup(sess)
            if not _is_login_screen(sess):
                title = sess.findById("wnd[0]").Text
                logger.info(f"SAP 로그인 완료 ({elapsed+1}초): {title}")
                return sap, conn
        except Exception:
            pass
        if elapsed % 10 == 9:
            logger.info(f"  로그인 대기 중... ({elapsed+1}초)")

    logger.error("SAP 연결 완전 실패")
    sys.exit(1)


def _try_manual_login(conn, client, user, password, language):
    """
    로그인 화면이 떠 있을 때 수동으로 필드 채워서 로그인.
    sapshcut.exe가 로그인을 완료하지 못한 경우 폴백.
    """
    if conn.Children.Count == 0:
        logger.error("로그인할 세션 없음")
        return

    sess = conn.Children(0)
    try:
        title = sess.findById("wnd[0]").Text
    except Exception:
        return

    if title.strip() and "SAP" not in title:
        logger.info(f"이미 로그인됨: {title}")
        return

    logger.info("수동 로그인 시도 중...")
    # 표준 SAP 로그인 필드 시도
    field_sets = [
        # 표준 SAP 로그인
        ("wnd[0]/usr/txtRSYST-MANDT", "wnd[0]/usr/txtRSYST-BNAME",
         "wnd[0]/usr/pwdRSYST-BCODE", "wnd[0]/usr/txtRSYST-LANGU"),
    ]
    for mandt_fid, user_fid, pw_fid, lang_fid in field_sets:
        try:
            sess.findById(mandt_fid).text = client
            sess.findById(user_fid).text  = user
            sess.findById(pw_fid).text    = password
            sess.findById(lang_fid).text  = language
            sess.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            logger.info("수동 로그인 완료")
            break
        except Exception:
            continue

    # 중복 로그인 팝업 처리
    try:
        popup = sess.findById("wnd[1]")
        popup_text = popup.Text.lower()
        if "logon" in popup_text or "session" in popup_text:
            for fid in ["wnd[1]/usr/radMULTI_LOGON_OPT2", "wnd[1]/usr/radMULTI_LOGON_OPT1"]:
                try:
                    sess.findById(fid).select()
                    break
                except Exception:
                    continue
            sess.findById("wnd[1]").sendVKey(0)
            time.sleep(2)
    except Exception:
        pass


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

    # 1. SAP 자동 실행 + 로그인
    sap, conn = launch_sap_and_connect()
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

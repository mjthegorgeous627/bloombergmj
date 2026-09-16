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
from config import EXCEL_PATH, REFRESH_INTERVAL_MINUTES, LOG_FILE

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# ── 로깅 설정 ──────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ],
    force=True,
)
logger = logging.getLogger(__name__)

from sap_handler import (
    navigate_to_vl06o_list,
    save_session_map,
    handle_multi_logon_popup,
    graceful_close_all_sap_connections,
)
from vl10g_handler import navigate_to_vl10g
from zrma_handler import navigate_to_zrma_q
from main import run_once, run_loop

SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"


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
    """wnd[1] 팝업(System Status 등)이 열려 있으면 닫기.

    2026-09-14: "License Information for Multiple Logons"(다중 로그온 충돌)
    팝업은 F12로 안 닫혀서 로그인이 영원히 멈춰서는 원인이었다 - 먼저 전용
    핸들러로 처리 시도하고, 그게 아니면 기존 F12 처리로 넘어간다."""
    try:
        if handle_multi_logon_popup(sess):
            return
    except Exception:
        pass
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


def _show_sap_gui_window():
    """SAP GUI 메인 창(4탭이 들어있는 창)이 안 보이거나 크기가 깨진 상태면
    강제로 정상 크기로 보이게 한다. 2026-08-18 실측 - 잠금 화면 상태에서
    startup.py가 SAP GUI를 띄우면(모닝루틴이 8:30에 자동으로 하는 경우)
    그 창이 생성은 되지만 IsWindowVisible=False 상태로 남아서, 화면 잠금이
    풀린 뒤에도 계속 안 보임 - 원격 데스크톱(구글 리모트)으로 접속해도
    SAP 탭을 클릭할 대상 자체가 없어서 "SAP 조작이 안 된다"는 증상으로
    나타남(자동화 자체는 Scripting API로 동작하니 영향 없음 - 사람이
    직접 보고 클릭할 때만 문제). SetForegroundWindow와 달리 ShowWindow는
    포그라운드 권한이 필요 없어서 잠금 상태에서도 동작하는 것을 실측으로
    확인함.

    주의(중요, 같은 날 실측으로 잡은 버그): ShowWindow만 호출하면
    IsWindowVisible은 True가 되지만, 숨겨져 있는 동안 창이 한 번도
    제대로 된 크기를 받은 적이 없어서 실제 크기가 200x75 픽셀 같은 작은
    크기로 남아있을 수 있음 - 제목+아이콘만 있는 작은 알림창처럼 보여서
    오히려 더 헷갈리는 새 팝업처럼 보였음(사용자가 스크린샷으로 실측
    확인). 그래서 SetWindowPos로 먼저 정상 크기를 강제해준 다음 최대화까지
    한다."""
    try:
        found = []

        def cb(hwnd, _):
            cls = win32gui.GetClassName(hwnd)
            if cls.startswith("SAP GUI for Windows"):
                found.append(hwnd)
            return True

        win32gui.EnumWindows(cb, None)
        for hwnd in found:
            visible = win32gui.IsWindowVisible(hwnd)
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            too_small = (right - left) < 600 or (bottom - top) < 400
            if not visible or too_small:
                if too_small:
                    win32gui.SetWindowPos(hwnd, 0, 50, 50, 1400, 900, 0x0040)  # SWP_SHOWWINDOW
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                logger.info(
                    f"SAP GUI 창 강제로 보이게/정상 크기로 복구 (hwnd={hwnd}, "
                    f"이전 상태: visible={visible} rect=[{left},{top},{right},{bottom}])"
                )
    except Exception as e:
        logger.warning(f"SAP GUI 창 보이기 처리 실패 (무시하고 계속): {e}")


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
        from sap_handler import get_scripting_engine
        return get_scripting_engine()
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


def _run_with_timeout(fn, timeout_sec):
    """fn()을 데몬 스레드에서 실행하고 timeout_sec 안에 안 끝나면 (False, None).
    SAP GUI Scripting COM 호출은 자체 타임아웃이 없어서, 원격 NWBC/SapGuiServer가
    죽어있으면(예: 주말 내내 idle이던 세션이 백엔드 타임아웃/VPN 끊김으로
    응답불능이 되는 경우) 그냥 영원히 블록된다 - 예외를 던지는 게 아니라
    "응답 없음"이라 try/except로는 절대 못 잡는다.

    2026-08-30 실사고: 이 실패 유형이 정확히 이 함수가 감싸는 "기존 세션
    재사용" 체크(launch_sap_and_connect() 1단계) 안에서 터졌다. 워치독이
    30분마다 python 프로세스를 강제 재시작해도, 새로 뜬 startup.py가 이
    체크에서 또 그 죽은 세션을 발견하고 다시 블록되는 걸 반복 - 그래서
    일요일 오전부터 월요일 아침까지 20번 넘게 재시작해도 하나도 안 풀렸다.
    타임아웃으로 감싸서 "N초 안에 응답 없음"을 "죽음"으로 취급해야 그
    자리에서 실제로 죽은 프로세스를 찾아 죽이고 새로 시작할 수 있다.

    타임아웃이 나도 이 스레드 자체를 강제로 죽일 방법은 없다(daemon=True라
    최소한 프로세스 종료는 안 막음) - 대부분은 이 함수를 부른 직후
    _kill_stuck_sap_gui_processes()가 그 죽은 NWBC/SapGuiServer를 실제로
    죽여버리므로, 블록돼 있던 COM 호출도 곧 에러로 풀려나며 스레드가
    알아서 끝난다."""
    result = {}

    def runner():
        try:
            result["value"] = fn()
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=runner, daemon=True)
    t.start()
    t.join(timeout_sec)
    if t.is_alive():
        return False, None
    if "error" in result:
        raise result["error"]
    return True, result.get("value")


def _kill_stuck_sap_gui_processes():
    """실제 SAP 프론트엔드 프로세스(NWBC/NwbcProcessAgent/SapGuiServer,
    구형 saplogon 포함)를 강제 종료. python.exe는 절대 안 건드림 - 워치독의
    기존 taskkill(main.py/startup.py 대상)은 이 프로세스들을 전혀 안
    건드려서, 진짜 죽은 SAP GUI는 그대로 남아있고 재시작마다 그 좀비를
    다시 붙잡는 게 2026-08-30 사고의 근본 원인이었다. 죽일 게 없어도
    안전하게 아무 일도 안 함.

    2026-09-14: 강제종료(Stop-Process) 직전에 정상 로그오프를 먼저 시도한다
    (타임아웃 8초로 감싸서, 이미 완전히 죽어 응답 없는 경우엔 그냥 넘어가고
    바로 강제종료로 진행 - graceful_close_all_sap_connections() 자체가
    막혀있는 스크립팅 엔진을 부를 수 있어서 _run_with_timeout 없이 직접
    부르면 여기서도 영원히 블록될 수 있음). 로그오프 없이 그냥 죽이면 SAP
    백엔드에 로그온이 남아 다음 로그인 시도가 "License Information for
    Multiple Logons" 팝업과 충돌하는 게 근본 원인이었다."""
    try:
        _run_with_timeout(graceful_close_all_sap_connections, timeout_sec=8)
    except Exception as e:
        logger.info(f"정상 로그오프 시도 중 예외 (무시하고 강제종료로 진행): {e}")

    ps_cmd = (
        "Get-Process -Name NWBC,NwbcProcessAgent,SapGuiServer,saplogon "
        "-ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15,
        )
        logger.info("응답 없는 SAP GUI 프로세스(NWBC/SapGuiServer 등) 강제 종료")
    except Exception as e:
        logger.warning(f"SAP GUI 프로세스 강제 종료 실패 (무시하고 계속 진행): {e}")
    time.sleep(2)


def _try_reuse_existing_session():
    """launch_sap_and_connect() 1단계의 실제 체크 로직 - _run_with_timeout()로
    감싸기 위해 분리. 재사용 가능한 세션을 찾으면 (sap, conn, title), 아니면
    None을 반환한다(예외 없이 정상적으로 "없음"인 경우)."""
    sap = _get_scripting_engine()
    if sap and sap.Children.Count > 0:
        conn = sap.Children(0)
        if conn.Children.Count > 0:
            sess = conn.Children(0)
            _dismiss_popup(sess)
            if not _is_login_screen(sess):
                title = sess.findById("wnd[0]").Text
                return sap, conn, title
    return None


def launch_sap_and_connect():
    """
    SAP 자동 실행 + 연결.

    1. 이미 로그인된 SAP 세션이 있으면 그대로 사용 (단, 10초 안에 응답이
       없으면 죽은 세션으로 판단 - 아래 _run_with_timeout 주석 참고 - 실제
       NWBC/SapGuiServer 프로세스를 강제 종료하고 2번으로 넘어간다).
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

    # ── 1. 이미 로그인된 세션 있으면 바로 사용 (10초 응답 타임아웃) ──
    try:
        ok, reuse_result = _run_with_timeout(_try_reuse_existing_session, timeout_sec=10)
    except Exception as e:
        logger.warning(f"기존 세션 재사용 체크 중 예외 (무시하고 새로 시작): {e}")
        ok, reuse_result = True, None
    if not ok:
        logger.warning("기존 SAP 세션이 10초 넘게 응답 없음 → 죽은 세션으로 판단, 강제 종료 후 새로 시작")
        _kill_stuck_sap_gui_processes()
    elif reuse_result:
        sap, conn, title = reuse_result
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

    sessions = {}

    # ── 세션0: VL06O ────────────────────────────────────────────
    try:
        sess0 = conn.Children(0)
        logger.info("세션0 VL06O 설정 중...")
        navigate_to_vl06o_list(sess0)
        sessions[0] = sess0
        logger.info("세션0 VL06O ✓")
    except Exception as e:
        logger.error(f"세션0 VL06O 설정 실패: {e}")

    # ── 세션1: VL10G ────────────────────────────────────────────
    try:
        sess1 = conn.Children(1)
        logger.info("세션1 VL10G 설정 중...")
        navigate_to_vl10g(sess1)
        sessions[1] = sess1
        logger.info("세션1 VL10G ✓")
    except Exception as e:
        logger.error(f"세션1 VL10G 설정 실패: {e}")

    # ── 세션2: ZRMA_Q RLKR ──────────────────────────────────────
    try:
        sess2 = conn.Children(2)
        logger.info("세션2 ZRMA RLKR 설정 중...")
        navigate_to_zrma_q(sess2, 'rlkr', 'next_month', apply_layout=True)
        sessions[2] = sess2
        logger.info("세션2 ZRMA RLKR ✓")
    except Exception as e:
        logger.error(f"세션2 ZRMA RLKR 설정 실패: {e}")

    # ── 세션3: ZRMA_Q Q2 ────────────────────────────────────────
    try:
        sess3 = conn.Children(3)
        logger.info("세션3 ZRMA Q2 설정 중...")
        navigate_to_zrma_q(sess3, 'q2', 'year_end', apply_layout=True)
        sessions[3] = sess3
        logger.info("세션3 ZRMA Q2 ✓")
    except Exception as e:
        logger.error(f"세션3 ZRMA Q2 설정 실패: {e}")

    # 지금 위치[0..3]가 확실히 올바른 순서인 시점에, SAP가 부여한 고정
    # SessionNumber를 저장해둔다. 나중에 NWBC 탭이 중간에 삽입/삭제되어
    # 위치가 밀려도 get_sap_session()이 이 매핑으로 올바른 세션을 다시 찾는다.
    try:
        session_map = {str(idx): sess.Info.SessionNumber for idx, sess in sessions.items()}
        if session_map:
            save_session_map(session_map)
    except Exception as e:
        logger.warning(f"세션 매핑 생성 실패: {e}")

    logger.info("SAP 4개 세션 설정 완료")


def open_excel():
    """Excel 파일 열기 - 이미 열려 있으면(자동화용 xlwings 인스턴스 포함)
    다시 열지 않는다.

    2026-09-01 실사고: 무조건 os.startfile()로 열다 보니, write_orders_to_excel()의
    xlwings 인스턴스가 이미 이 파일을 열어둔 상태에서 SAP 재로그인/워치독
    재시작/강제 재시작이 있을 때마다(최근 안정성 작업으로 재시작이 잦아짐)
    중복 Excel 프로세스가 하나씩 더 생겼다. 두 번째 인스턴스가 뜨면 Excel이
    "Excel에서 동일한 이름을 가진 두 개의 통합 문서를 동시에 열 수 없습니다"
    확인 팝업을 띄우는데, 이게 화면 구석(사람 눈에 안 띄는 위치)에서 응답
    없이 쌓여 있으면 write_orders_to_excel()의 COM 호출이 그 팝업이 닫힐
    때까지 무기한 블록된다 - "SAP 수집은 되는데 workbench/엑셀에 하나도
    안 올라온다"는 문의의 실제 원인 중 하나로 실측 확인됨(오더 551060846,
    67085142 둘 다 이 팝업 때문에 몇 분~몇 시간씩 멈춰있었다).

    Excel은 이름만 같으면(경로가 달라도) 동시에 못 연다 - 실측으로 확인된
    실제 사례: 전혀 다른 파일인 `C:\\Users\\bloomberg\\Documents\\2026
    배송장.xlsx`가 별도 인스턴스에 열려 있었는데도 이름이 같다는 이유로
    진짜 대상 파일(`C:\\1\\배송장\\2026 배송장.xlsx`)과 충돌했다. 그래서
    fullname이 아니라 name만 보고 "이미 열려 있다"고 판단하면, 이런
    이름만 겹치는 딴 파일을 우리 대상 파일로 착각해 스킵해버리는 오탐이
    난다 - fullname까지 비교해서 진짜 우리 파일인지 확인하고, 이름만
    겹치는 다른 파일이 발견되면 (스킵하는 대신) 큰 소리로 경고만 남기고
    진행한다 - 그래도 결국 같은 충돌 팝업이 뜨겠지만, 최소한 로그에 원인이
    명확히 남는다."""
    try:
        if not os.path.exists(EXCEL_PATH):
            logger.warning(f"Excel 파일 없음: {EXCEL_PATH}")
            return

        target_name = os.path.basename(EXCEL_PATH)
        target_full = os.path.normcase(os.path.abspath(EXCEL_PATH))
        try:
            import xlwings as xw
            for app in xw.apps:
                for book in app.books:
                    if book.name != target_name:
                        continue
                    if os.path.normcase(os.path.abspath(book.fullname)) == target_full:
                        logger.info(f"Excel 이미 열려 있어 재실행 스킵 (중복 방지): {EXCEL_PATH}")
                        return
                    logger.warning(
                        f"경고: 이름은 같지만 다른 파일이 이미 열려 있음 ({book.fullname}) - "
                        f"Excel이 이것 때문에 '{EXCEL_PATH}'를 못 열 수 있음. 그 파일을 닫아야 함."
                    )
        except Exception as e:
            logger.warning(f"Excel 열림 여부 확인 실패 (무시하고 진행): {e}")

        os.startfile(EXCEL_PATH)
        time.sleep(4)
        logger.info(f"Excel 열기 완료: {EXCEL_PATH}")
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

    # (2026-08-18: 여기 있던 _show_sap_gui_window() 호출은 제거함 - 잘못된
    # 진단으로 추가했던 것. "SAP GUI for Windows 800" 창을 4탭 세션 창으로
    # 착각했는데, 실측해보니 그건 SapGuiServer.exe 내부용 빈 컨테이너
    # 창이었고, 강제로 보이게/최대화하니 빈 흰색 창이 화면 전체를 덮어서
    # 오히려 리모트 조작을 더 방해함(사용자 스크린샷 2개로 실측 확인,
    # 원인 재조사 필요 - 진짜 4탭 창이 뭔지 아직 특정 못함). 함수 정의는
    # 남겨두되 자동으로는 절대 호출하지 않는다.)

    # 4. (2026-09-02 제거: 사용자 요청) 공유파일(EXCEL_PATH)이 이제 매일
    # SAP 자동수집을 직접 안 받는 "마감된 데이터만 담는 아카이브"로 역할이
    # 바뀌면서, 더 이상 아침마다 화면에 띄워둘 실시간 파일이 아니게 됐다
    # (delivery_excel_board_sync_bridge 메모리 2026-09-01 3번째 후속 참고).
    # 사용자가 실제로 아침에 이 파일이 화면에 열려있는 걸 발견하고 "안 열어도
    # 된다"고 지적함 - open_excel() 정의는 남겨두되(수동 필요시 대비)
    # 자동으로는 더 이상 호출하지 않는다. 아래 catchup(run_once(excel_only=
    # True))은 get_workbook()이 필요하면 COM으로 알아서 여는 것과 무관.
    # open_excel()

    # 5. 누락 오더 catchup
    #
    # 2026-09-02 버그 수정: excel_only=True(Excel 기록만 기준, JSON 무시)를
    # 계속 쓰고 있었는데, 이건 "매일 SAP 수집을 직접 받던" 구 아키텍처
    # 전제였다(바로 위 4번 주석 참고, 2026-09-01 폐기됨) - 공유 Excel이
    # 이제 마감(확정)된 날에만 데이터가 들어가는 아카이브가 되면서,
    # get_existing_order_numbers()로 "오늘 이미 처리됨"을 걸러내던 이
    # catchup은 오늘 하루 동안 워크벤치로만 반영된 오더를 전부 "Excel에
    # 아직 없음=미처리"로 오판했다. 그 결과 startup.py가 다시 실행될
    # 때마다(모닝루틴, 워치독 재시작, 수동 재시작 전부 포함) 이미 정상
    # 처리·카카오 전송까지 끝난 오더를 처음부터 다시 수집해 카카오를
    # 중복 발송했다 - 2026-09-02 실측: 오더 7851395가 이 버그로 재시작
    # 직후 카카오 중복 발송됨. 이제 그냥 run_once()를 쓴다 - JSON
    # (order_tracker.json)이 모든 수집 경로(run_manual_order/_run_zrma/
    # _write_and_notify)가 공통으로 갱신하는 진짜 진행 상태이고, JSON이
    # 비어있는 예외 상황(초기화/삭제 복구용)은 run_once() 안에서 이미
    # Excel 스캔으로 자동 복구한다(_run_once_impl 참고) - excel_only가
    # 주려던 안전망을 그대로 유지하면서 "매번 재실행" 오탐만 없앤다.
    logger.info("누락 오더 확인 중 (catchup)...")
    run_once()

    # 6. 10분마다 자동 반복
    logger.info(f"자동 반복 시작 ({REFRESH_INTERVAL_MINUTES}분 간격)")
    run_loop()

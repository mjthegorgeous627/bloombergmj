"""
아침 자동 시작 루틴 - Windows 작업 스케줄러에서 매일 08:30에 실행 (사용자 요청,
2026-08-14). 컴퓨터가 잠금 상태여도 실행되도록 "사용자가 로그온한 경우에만
실행"(Interactive Token) 방식으로 등록한다 - 잠금은 로그오프가 아니라서
백그라운드 예약 작업은 계속 돈다.

한다:
  1. SAP 실행 + 로그인 + 4세션 셋업 + 자동 루프 시작 (startup.py) - "1. SAP
     Start & Loop" 버튼과 동일한 동작. 이미 떠 있으면 건너뜀.
  2. workbench 대시보드 실행 (workbench_app.py). 이미 떠 있으면 건너뜀.
  3. (제거됨, 2026-08-18 - 아래 사고 참고) Excel 배송장 파일 열기.
  3-2. (신설, 2026-08-28) 작업표시줄 트레이 아이콘 실행 (tray_icon.py).
     workbench가 이미 떠있어서 2번이 건너뛰어지면 브라우저 탭이 자동으로
     안 열리는데(=아침에 workbench가 안 보이는 것처럼 느껴짐), 트레이
     아이콘은 항상 눈에 보이므로 "지금 돌고 있나?"를 확인할 수단이 된다.
  4. Bloomberg 터미널 실행 - 로그인 창을 띄우고 ID칸에 미리 msuh33까지
     입력해둔다(사용자 확인, 2026-08-14 - 실제 로그인창에서 입력 테스트
     후 확인됨). B-Unit 인증 자체는 하드웨어 보안키/생체인증 기반이라
     소프트웨어로 자동화할 수 없다(= 자동화하지 않는 게 아니라 원천적으로
     불가능. 그게 B-Unit이 존재하는 이유이기도 함) - 인증은 사용자가 직접.

     주의(중요): ID 자동입력은 win32 keybd_event(시스템 전역 키보드 입력
     시뮬레이션)를 쓰는데, Windows는 보안상 잠금 화면(별도의 보안
     데스크톱)에는 이런 입력 시뮬레이션이 전달되지 않게 막아놓는다. 터미널
     "실행"(프로세스 시작) 자체는 잠긴 상태에서도 되지만, ID 자동입력은
     잠금이 풀려있어야 확실히 동작한다 - 8:30에 아직 잠긴 상태라면 터미널
     로그인창은 뜨는데 ID는 자동으로 안 채워질 수 있다(그럴 땐 사용자가
     6글자만 직접 치면 됨 - 큰 손해는 아님).

절대 하지 않는 것: 이미 켜져 있는 프로세스를 중복 실행하지 않는다 (각 단계
전에 프로세스 목록을 확인) - 사용자가 이미 손으로 SAP/workbench를 켜놨어도
안전하게 건너뛴다. main.py/workbench_app.py 두 개가 같은 SAP GUI Scripting
세션이나 같은 라이브 Excel 파일을 동시에 건드리면 실제로 사고가 났던 적
있음(see project_sap_order_collection_gaps 2026-08-14 세션의 ZRX 67076714
경합 사고) - 그래서 중복 실행 방지를 제일 먼저 확인한다.

2026-08-18 사고 및 open_excel() 제거 이유: 첫 실제 08:30 무인 실행에서
컴퓨터가 1시간 넘게 먹통이 됨. automation.log 확인 결과 - 08:30에 SAP
루프는 이미 실행 중이었고(건너뜀), 이 함수가 os.startfile()로 같은
'2026 배송장.xlsx'를 별도로 열었는데, 그 파일은 그 순간 main.py의
xlwings COM 자동화가 실시간으로 쓰고 있던 바로 그 파일 - 두 프로세스가
같은 Excel COM 세션을 동시에 건드리면서 어딘가에 응답 대기 팝업(파일
잠금/읽기전용 경고 등)이 뜬 것으로 보이는데, 화면이 잠겨있어 아무도
그걸 눌러줄 수 없어 무한 대기가 됨. 그 여파로 main.py의 SAP 루프
자체도 08:30 직후로 automation.log가 끊겼고, 30분 후 workbench_app.py
워치독이 자동 재시작을 시도했지만 그것마저 SAP RP1 로그인 팝업 자동입력이
잠금화면에서 막혀서(Bloomberg ID 자동입력과 같은 제약) 못 풀림 - 결국
09:44에 사용자가 직접 전원을 껐다 켜서 복구. 8/14 ZRX 67076714 경합
사고와 같은 유형(락 없이 같은 라이브 Excel 파일에 여러 프로세스가 동시
접근)인데, 이번엔 무인 상태라 아무도 못 풀어서 피해가 훨씬 컸음. 그래서
Excel 여는 단계를 통째로 제거함 - main.py/startup.py는 자기 프로세스
안에서 필요할 때 알아서 열므로 이 루틴이 따로 열어줄 필요가 원래 없었음.
"""

import ctypes
import logging
import os
import subprocess
import sys
import time

import win32gui

from config import LOG_FILE
from credentials import PORTAL_PASSWORD, PORTAL_USER
from holiday_check import skip_reason

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MORNING_LOG_FILE = os.path.join(BASE_DIR, "morning_routine.log")
BLOOMBERG_LNK = (
    r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Bloomberg\Bloomberg.lnk"
)
BLOOMBERG_LOGIN_WINDOW_TITLE = "BLOOMBERG: Login"
# 사용자 확인, 2026-08-14: Bloomberg 터미널 로그인 ID/비번은 Bloomberg
# Vendor Portal(credentials.py)과 동일 - 별도로 저장하지 않고 그대로 재사용.
BLOOMBERG_ID = PORTAL_USER
BLOOMBERG_PASSWORD = PORTAL_PASSWORD

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(MORNING_LOG_FILE, encoding="utf-8"),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),  # automation.log에도 남겨서 한 곳에서 확인 가능
    ],
)
logger = logging.getLogger(__name__)


def _pids_matching(*patterns):
    """main.py/startup.py/workbench_app.py 등 특정 스크립트를 실행 중인
    python(w).exe PID 목록. workbench_app.py의 _running_sap_loop_pids()와
    같은 방식(라이브 프로세스 목록 조회, PID 파일 아님 - 죽은 프로세스가
    남긴 stale 파일 걱정 없음)."""
    cond = " -or ".join(f"$_.CommandLine -match '{p}'" for p in patterns)
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        f"Where-Object {{ $_.Name -in @('python.exe','pythonw.exe') -and ({cond}) }} | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
        )
    except Exception as e:
        logger.warning(f"프로세스 조회 실패 (무시하고 계속): {e}")
        return []
    return [int(p.strip()) for p in out.splitlines() if p.strip().isdigit()]


def start_sap_loop():
    if _pids_matching(r"main\.py", r"startup\.py"):
        logger.info("SAP 루프 이미 실행 중 - 건너뜀")
        return
    subprocess.Popen(
        [sys.executable, "startup.py"], cwd=BASE_DIR,
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )
    logger.info("SAP 시작 & 루프 실행")


def start_workbench():
    """2026-08-31 수정: 예전엔 workbench_app.py가 이미 떠있으면 여기서
    그냥 건너뛰었는데, 그러면 브라우저 탭이 아예 안 열려서 "아침에
    workbench가 안 보인다"는 증상으로 이어졌다(2026-08-28 워크벤치
    중복인스턴스 수정 때 이미 예견했던 문제 - 실제로 2026-08-31(월) 아침
    재현: 주말 내내 SAP 세션이 죽어있었지만 workbench_app.py 서버 자체는
    안 죽고 계속 떠있었으므로 08:30 모닝루틴이 이 조건에 걸려 조용히
    건너뛰었고, 사용자는 브라우저에 아무 탭도 없어 "workbench가 안
    떠있다"고 오판함). 이제 배송장 엑셀이 백업 용도가 되면서 아침에
    workbench를 실제로 보는 게 맞다는 사용자 확인(2026-08-31)에 따라,
    이미 떠있어도 항상 새로 실행한다 - workbench_app.py 자신의 단일
    인스턴스 뮤텍스(serve())가 새 서버를 안 띄우고 기존 서버의 브라우저
    탭만 여는 걸 보장하므로 중복 실행 위험은 없다(트레이 아이콘의
    "Workbench 열기" 메뉴와 완전히 같은 방식)."""
    subprocess.Popen(
        [sys.executable, "workbench_app.py"], cwd=BASE_DIR,
        creationflags=subprocess.CREATE_NEW_CONSOLE,
    )
    logger.info("workbench 실행 (이미 떠있으면 브라우저 탭만 열림)")


def start_tray_icon():
    """2026-08-28 신설: workbench가 이미 떠있어서 위 start_workbench()가
    건너뛰어도(=오늘 아침엔 브라우저 탭이 자동으로 안 열림), 트레이 아이콘은
    항상 눈에 보이는 상태 표시 수단이므로 별도로 켠다. tray_icon.py 자신의
    단일 인스턴스 뮤텍스가 중복 실행을 막아주므로 여기서 프로세스 목록을
    따로 확인할 필요는 없지만, 매일 쓸데없이 새 프로세스를 만들지 않도록
    똑같이 먼저 확인한다."""
    if _pids_matching(r"tray_icon\.py"):
        logger.info("트레이 아이콘 이미 실행 중 - 건너뜀")
        return
    subprocess.Popen(
        [sys.executable, "tray_icon.py"], cwd=BASE_DIR,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    logger.info("트레이 아이콘 실행")


def _type_text(text):
    """실제 키보드 입력처럼 시스템 전역으로 문자를 하나씩 타이핑한다 (CEF/
    크롬 기반 렌더링 창은 WM_SETTEXT로 특정 컨트롤을 직접 지정할 수 없어서
    -  Bloomberg 로그인창 실측으로 확인, 2026-08-14 - startup.py의
    _send_key()와 같은 방식). 잠금 화면에서는 Windows가 이 입력을 막는다.

    VkKeyScanW의 상위 바이트(shift 상태)를 무시하면 대문자가 소문자로
    입력된다 - 실제로 비밀번호("Qmffnaqjrm1234")의 앞글자 대문자 Q가
    q로 들어가서 로그인 에러가 난 걸 사용자가 실측으로 확인함
    (2026-08-14). 그래서 Shift가 필요한 문자는 Shift를 누른 채로 입력."""
    user32 = ctypes.windll.user32
    VK_SHIFT = 0x10
    for ch in text:
        scan = user32.VkKeyScanW(ord(ch))
        vk = scan & 0xFF
        need_shift = bool((scan >> 8) & 1)
        if need_shift:
            user32.keybd_event(VK_SHIFT, 0, 0, 0)
            time.sleep(0.02)
        user32.keybd_event(vk, 0, 0, 0)
        time.sleep(0.03)
        user32.keybd_event(vk, 0, 0x0002, 0)
        if need_shift:
            time.sleep(0.02)
            user32.keybd_event(VK_SHIFT, 0, 0x0002, 0)
        time.sleep(0.03)


def _find_window_by_title(title_substr, timeout=25):
    deadline = time.time() + timeout
    while time.time() < deadline:
        result = []

        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd) and title_substr in win32gui.GetWindowText(hwnd):
                result.append(hwnd)
            return True

        win32gui.EnumWindows(cb, None)
        if result:
            return result[0]
        time.sleep(1)
    return None


def _press_tab():
    user32 = ctypes.windll.user32
    VK_TAB = 0x09
    user32.keybd_event(VK_TAB, 0, 0, 0)
    time.sleep(0.05)
    user32.keybd_event(VK_TAB, 0, 0x0002, 0)
    time.sleep(0.3)


def _force_foreground(hwnd):
    """SetForegroundWindow(hwnd)를 그냥 부르면 Windows가 "포그라운드를
    가진 적 없는 백그라운드 프로세스가 임의로 창을 앞으로 가져오는 것"을
    보안상 막아서 조용히 실패한다 - "0. 아침 루틴 강제 실행" 버튼으로
    실행했을 때(workbench_app.py가 백그라운드에서 띄운 프로세스) 실제로
    이 에러로 ID/비번 입력이 통째로 스킵됐음(사용자 실측, 2026-08-14: 로그에
    'SetForegroundWindow', 'No error message is available'). 터미널에서
    직접 실행했을 때는 이 제약에 안 걸려서 그때는 됐던 것.

    표준 우회법: Alt 키를 한 번 눌렀다 떼면(다른 아무 동작도 안 함) Windows의
    포그라운드 잠금이 풀린다 - 방금 실제 문제가 났던 것과 같은 상황(버튼으로
    띄운 백그라운드 프로세스)에서 실측으로 확인됨."""
    user32 = ctypes.windll.user32
    VK_MENU = 0x12  # Alt
    user32.keybd_event(VK_MENU, 0, 0, 0)
    time.sleep(0.03)
    user32.keybd_event(VK_MENU, 0, 0x0002, 0)
    time.sleep(0.1)
    win32gui.SetForegroundWindow(hwnd)


def open_bloomberg_login():
    """Bloomberg 터미널 실행 → 로그인창이 뜨면 ID/비번칸까지 채워두고
    멈춘다 (Enter/로그인 버튼은 절대 안 누름 - B-Unit 인증은 사용자가
    직접). ID/비번은 Bloomberg Vendor Portal과 동일하다고 사용자가 확인
    (2026-08-14, credentials.PORTAL_USER/PORTAL_PASSWORD 재사용). 이미
    로그인돼 있으면(로그인창이 안 뜸) 조용히 넘어간다 - 실패로 취급 안 함."""
    if not os.path.exists(BLOOMBERG_LNK):
        logger.warning(f"Bloomberg 바로가기를 찾을 수 없음: {BLOOMBERG_LNK}")
        return
    try:
        os.startfile(BLOOMBERG_LNK)
        logger.info("Bloomberg 터미널 실행")
    except Exception as e:
        logger.warning(f"Bloomberg 실행 실패: {e}")
        return

    hwnd = _find_window_by_title(BLOOMBERG_LOGIN_WINDOW_TITLE, timeout=25)
    if not hwnd:
        logger.info("Bloomberg 로그인창 감지 안 됨 (이미 로그인돼 있거나 시간 초과) - ID/비번 자동입력 생략")
        return
    try:
        _force_foreground(hwnd)
        # 창 제목("BLOOMBERG: Login")은 안의 CEF/크롬 콘텐츠(실제 로그인
        # 폼)가 다 그려지기 전에 먼저 뜬다 - 0.5초만 기다리고 치면 폼이
        # 아직 준비 안 돼서 입력이 씹힘(사용자 실측, 2026-08-14: 몇 분간
        # 열려있던 창엔 정상 입력됐지만 방금 뜬 창엔 ID/비번이 안 들어감).
        # 넉넉히 기다린다.
        time.sleep(4)
        _type_text(BLOOMBERG_ID)   # ID칸 (창 열리면 기본 포커스)
        _press_tab()               # 비번칸으로 이동
        _type_text(BLOOMBERG_PASSWORD)
        logger.info("Bloomberg 로그인창 ID/비번 자동입력 완료 - B-Unit 인증은 직접 진행 (Enter는 안 누름)")
    except Exception as e:
        logger.warning(f"Bloomberg ID/비번 자동입력 실패 (직접 입력 필요): {e}")


def main():
    # 2026-09-13: Windows 작업 스케줄러(SAP_Morning_Routine)는 요일 구분 없이
    # 매일 08:30에 이 스크립트를 실행한다 - 주말/한국 공휴일에도 그대로 SAP에
    # 자동 접속해 무인 상태로 돌다가 에러 팝업이 쌓이는 사고가 실제로 있었음
    # (holiday_check.py 모듈 독스트링 참고). 여기서 걸러내면 SAP 루프/워크벤치/
    # 트레이/Bloomberg 로그인 자동입력까지 오늘 자동으로 시작되는 것 전부가
    # 함께 막힌다.
    reason = skip_reason()
    if reason:
        logger.info(f"=== 아침 자동 루틴 건너뜀 ({reason} - 자동 기동 안 함) ===")
        return
    logger.info("=== 아침 자동 루틴 시작 ===")
    start_sap_loop()
    time.sleep(2)  # SAP 프로세스 창 생성 여유
    start_workbench()
    time.sleep(2)
    start_tray_icon()
    time.sleep(2)
    open_bloomberg_login()
    logger.info("=== 아침 자동 루틴 완료 (Bloomberg B-Unit 로그인은 직접 진행) ===")


if __name__ == "__main__":
    main()

"""작업표시줄 트레이 아이콘 - workbench/SAP 자동루프가 지금 돌고 있는지를
브라우저를 열지 않고도 한눈에 확인하고, 켜고/끌 수 있게 하는 상시 감시자.

배경 (2026-08-28): workbench는 콘솔 창도 없는 백그라운드 서버라, 지금까지
"떠 있나?"를 확인할 유일한 방법이 브라우저 탭을 열어보는 것뿐이었다. 그런데
그 브라우저 탭 자체가 "workbench_app.py가 새로 시작될 때만" 자동으로
열리므로(serve()의 webbrowser.open 참고), 서버는 백그라운드에서 멀쩡히
돌고 있는데 탭만 안 열려있는 상황이 실제로 발생했다. 사용자는 이걸 "꺼졌다"고
오판하고 workbench.bat을 손으로 한 번 더 실행 → 서버가 중복으로 뜨는 사고로
이어졌다(단발성 확인 수단 자체가 없어서 생긴 문제).

이 트레이 아이콘은:
  - 15초마다 workbench의 기존 REST API(/api/sap/status)를 조회해서 상태를
    색으로 보여준다 (초록=SAP 루프 정상, 노랑=workbench는 켜져있지만 SAP
    루프 꺼짐, 빨강=workbench 자체가 응답 없음)
  - 우클릭 메뉴로 "Workbench 열기"(=workbench_app.py 실행 - 이미 떠있으면
    workbench_app.py 자신의 단일 인스턴스 잠금이 새 서버를 안 띄우고 기존
    걸 브라우저로만 열어주므로 중복 걱정 없이 그냥 눌러도 됨), SAP 루프
    시작/중지(기존 /api/sap/start_loop, stop_loop 재사용 - 그쪽에 이미
    중복 방지 로직이 있음), 문제가 생겼을 때 쓰는 "강제 재시작"을 제공한다.
  - 이 트레이 프로세스 자신도 중복 실행되지 않도록 Win32 named mutex로
    막는다(workbench_app.py의 2026-08-28 수정과 같은 방식).

새 서버/프로세스를 직접 만들거나 죽이는 로직은 최대한 안 만들고 기존
workbench_app.py의 API와 프로세스 규칙을 그대로 재사용한다 - SAP 루프
중복 방지(sap_start_loop/sap_stop_loop)와 workbench 자체 중복 방지(단일
인스턴스 뮤텍스)는 이미 그쪽에 구현돼 있으므로, 여기서 같은 걸 다시
구현하면 두 군데가 서로 다른 판단을 할 위험만 생긴다.
"""

import json
import logging
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

import win32api
import win32event
import winerror
from PIL import Image, ImageDraw
from pystray import Icon, Menu, MenuItem

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "tray_icon.log"
WORKBENCH_URL = "http://127.0.0.1:8765"
POLL_INTERVAL_SEC = 15
HTTP_TIMEOUT_SEC = 3

_TRAY_MUTEX_NAME = "Global\\MJSuh_Automation_Tray_SingleInstance"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger("tray_icon")


def _pythonw_exe():
    """콘솔 창 없이 백그라운드로 띄우기 위한 pythonw.exe 경로 - workbench_
    watchdog_check.py의 같은 이름 함수와 동일한 로직."""
    candidate = Path(sys.executable).with_name("pythonw.exe")
    return str(candidate) if candidate.exists() else sys.executable


def _fetch_sap_status():
    """workbench의 기존 /api/sap/status를 그대로 조회. 반환값:
      ("ok", pids)          - workbench 응답함, SAP 루프 실행 중
      ("loop_stopped", [])  - workbench 응답함, SAP 루프는 꺼짐
      ("workbench_down", None) - workbench 자체가 응답 없음(연결 실패/타임아웃)
    """
    try:
        with urllib.request.urlopen(f"{WORKBENCH_URL}/api/sap/status", timeout=HTTP_TIMEOUT_SEC) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if data.get("loopRunning"):
            return "ok", data.get("pids", [])
        return "loop_stopped", []
    except Exception:
        return "workbench_down", None


def _make_icon_image(color):
    """64x64 투명 배경에 색칠된 원 하나 - 상태를 색으로만 구분하는 가장
    단순한 표시. 텍스트/글자는 트레이 아이콘 크기(보통 16~24px 축소 표시)에서
    안 보이므로 넣지 않는다."""
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 60, 60), fill=color)
    return img


ICON_OK = _make_icon_image((40, 180, 90, 255))          # 초록
ICON_LOOP_STOPPED = _make_icon_image((235, 175, 30, 255))  # 노랑
ICON_DOWN = _make_icon_image((215, 60, 55, 255))         # 빨강


def _run_bg(fn, *args):
    """메뉴 클릭 콜백에서 느린 작업(subprocess/HTTP)을 트레이 UI 스레드를
    막지 않고 실행."""
    threading.Thread(target=fn, args=args, daemon=True).start()


def _open_workbench(icon=None, item=None):
    """workbench_app.py를 그대로 실행한다. 이미 떠있으면 그쪽 단일 인스턴스
    잠금이 새 서버 없이 브라우저 탭만 열어주고, 안 떠있으면 새로 켜지면서
    자기가 알아서 브라우저를 연다 - 어느 경우든 이 호출 하나로 충분하고
    중복이 생기지 않는다."""
    try:
        subprocess.Popen(
            [_pythonw_exe(), "workbench_app.py"], cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.info("Workbench 열기 요청")
    except Exception:
        logger.exception("Workbench 열기 실패")


def _post_sap(path):
    try:
        req = urllib.request.Request(f"{WORKBENCH_URL}/api/sap{path}", method="POST", data=b"")
        with urllib.request.urlopen(req, timeout=10) as resp:
            logger.info(f"{path} 응답: {resp.read().decode('utf-8')}")
    except Exception:
        logger.exception(f"{path} 요청 실패 (workbench가 꺼져있을 수 있음)")


def _start_sap_loop(icon=None, item=None):
    _post_sap("/start_loop")


def _stop_sap_loop(icon=None, item=None):
    _post_sap("/stop_loop")


def _force_restart_workbench(icon=None, item=None):
    """문제가 생겨서 그냥 껐다 켜고 싶을 때 쓰는 비상 버튼. workbench_app.py
    프로세스를 전부 taskkill한 뒤(단일 인스턴스 뮤텍스는 프로세스 종료 시
    OS가 자동으로 풀어준다) 하나만 새로 띄운다."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'workbench_app\\.py' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        pids = [p.strip() for p in out.splitlines() if p.strip().isdigit()]
        for pid in pids:
            subprocess.run(
                ["taskkill", "/F", "/PID", pid],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        if pids:
            time.sleep(2)
        subprocess.Popen(
            [_pythonw_exe(), "workbench_app.py"], cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.warning(f"Workbench 강제 재시작 (이전 PID: {pids or '없음'})")
    except Exception:
        logger.exception("Workbench 강제 재시작 실패")


def _exit_tray(icon, item):
    icon.stop()


def _poll_loop(icon):
    while True:
        state, pids = _fetch_sap_status()
        if state == "ok":
            icon.icon = ICON_OK
            icon.title = f"자동화 정상 작동 중 (SAP 루프 PID {', '.join(map(str, pids))})"
        elif state == "loop_stopped":
            icon.icon = ICON_LOOP_STOPPED
            icon.title = "Workbench는 켜져있지만 SAP 루프는 꺼져있음"
        else:
            icon.icon = ICON_DOWN
            icon.title = "Workbench가 응답 없음 (꺼져있거나 문제 발생)"
        time.sleep(POLL_INTERVAL_SEC)


def main():
    mutex = win32event.CreateMutex(None, False, _TRAY_MUTEX_NAME)
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        logger.info("트레이 아이콘이 이미 실행 중 - 중복 실행 안 함")
        return

    menu = Menu(
        MenuItem("Workbench 열기", lambda i, it: _run_bg(_open_workbench)),
        Menu.SEPARATOR,
        MenuItem("SAP 루프 시작", lambda i, it: _run_bg(_start_sap_loop)),
        MenuItem("SAP 루프 중지", lambda i, it: _run_bg(_stop_sap_loop)),
        Menu.SEPARATOR,
        MenuItem("Workbench 강제 재시작 (문제 있을 때)", lambda i, it: _run_bg(_force_restart_workbench)),
        Menu.SEPARATOR,
        MenuItem("트레이 종료", _exit_tray),
    )
    icon = Icon("mjsuh_automation", ICON_DOWN, "상태 확인 중...", menu)
    threading.Thread(target=_poll_loop, args=(icon,), daemon=True).start()
    logger.info("트레이 아이콘 시작")
    icon.run()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        logger.exception("tray_icon.py 최상위에서 처리되지 않은 예외로 종료")
        raise

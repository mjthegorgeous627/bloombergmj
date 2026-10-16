"""폰 크기 /m 창을 오른쪽 모니터에 띄운다 (2026-10-16 사용자 요청).

사용자는 Chrome 원격 데스크톱으로 폰에서 PC 화면을 본다. 폰에서 보기 좋게 workbench PC 화면(오른쪽 모니터) 위에
폰 비율의 Chrome 창(/m)을 하나 띄워 두고, 폰에서 그 영역만 확대해서 쓴다.
- 왼쪽 모니터 = SAP/터미널/Portal, 오른쪽 모니터 = 배송장/workbench 이므로 항상 "가장 오른쪽 모니터"에 둔다.
- 창이 이미 떠 있으면 새로 만들지 않고 제 위치/크기로만 되돌린다(중복 방지).
- 포커스를 빼앗지 않는다(SAP/Bloomberg 로그인 자동입력 중일 수 있음).
- 실패해도 아침 루틴을 막지 않는다.
단독 실행: python mobile_window.py   (모바일창.bat)
"""
import logging
import os
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import win32api
import win32con
import win32gui

logger = logging.getLogger(__name__)

URL = "http://127.0.0.1:8765/m"
TITLE = "Bloomberg Mobile"          # templates_v2/mobile 페이지 제목(workbench_app.render_mobile_m_page)
WIDTH = 570                          # 갤럭시 S26 원격 데스크톱 화면(상단 도구줄 제외, 비율 약 0.55)에 맞춘 너비 - 높이는 모니터 작업영역 전체. 460일 때 폰에서 좌우가 28% 남았음(2026-10-16 사용자 스크린샷)
SIDE = "right"                       # 창을 모니터의 어느 쪽에 붙일지: right | left
PROFILE_DIR = Path(__file__).resolve().parent / "_m_chrome_profile"   # 다른 Chrome(Portal 등)과 완전히 분리
CHROME_CANDIDATES = [
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


def _target_rect():
    """가장 오른쪽 모니터의 작업영역(작업표시줄 제외)에 폰 비율 창이 들어갈 좌표 (left, top, width, height)."""
    mons = []
    for h, _, _ in win32api.EnumDisplayMonitors():
        mi = win32api.GetMonitorInfo(h)
        mons.append(mi["Work"])
    left, top, right, bottom = max(mons, key=lambda r: r[0])
    x = right - WIDTH if SIDE == "right" else left
    return x, top, WIDTH, bottom - top


def _find_window():
    found = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and TITLE in win32gui.GetWindowText(hwnd):
            found.append(hwnd)
    win32gui.EnumWindows(cb, None)
    return found[0] if found else None


def _wait_server(timeout=60):
    end = time.time() + timeout
    while time.time() < end:
        try:
            with urllib.request.urlopen(URL, timeout=5) as r:
                if r.status == 200:
                    return True
        except Exception:
            time.sleep(2)
    return False


def _place(hwnd):
    x, y, w, h = _target_rect()
    if win32gui.IsIconic(hwnd) or win32gui.GetWindowPlacement(hwnd)[1] == win32con.SW_SHOWMAXIMIZED:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    # SWP_NOACTIVATE: 위치/크기만 바꾸고 포커스는 건드리지 않는다
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW)


def open_mobile_window():
    """반환: True(창이 제 위치에 있음) / False(실패)"""
    try:
        hwnd = _find_window()
        if not hwnd:
            if not _wait_server():
                logger.warning("모바일 창: workbench(/m)가 응답하지 않아 건너뜀")
                return False
            chrome = next((c for c in CHROME_CANDIDATES if c.exists()), None)
            if not chrome:
                logger.warning("모바일 창: Chrome을 찾지 못함")
                return False
            x, y, w, h = _target_rect()
            subprocess.Popen(
                [str(chrome), f"--app={URL}", f"--user-data-dir={PROFILE_DIR}",
                 f"--window-position={x},{y}", f"--window-size={w},{h}",
                 "--no-first-run", "--no-default-browser-check"],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            end = time.time() + 30
            while time.time() < end and not hwnd:
                time.sleep(0.5)
                hwnd = _find_window()
        if not hwnd:
            logger.warning("모바일 창: 창이 뜨지 않음")
            return False
        _place(hwnd)
        logger.info(f"모바일 창 배치 완료: {_target_rect()}")
        return True
    except Exception:
        logger.warning("모바일 창 열기 실패(무시)", exc_info=True)
        return False


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    sys.exit(0 if open_mobile_window() else 1)

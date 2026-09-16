"""
자동화가 도는 동안 사용자 작업 창의 키보드 포커스를 뺏기지 않게 하는 가드.

문제 (2026-08-24, 사용자 실측): main.py 루프가 SAP GUI Scripting으로
VL06O/VL10G/ZRMA_Q 세션들을 조작하는 동안, SAP 창이 반복적으로 활성화되면서
사용자가 그 순간 쓰고 있던 다른 창(엑셀, 폴더, 다른 SAP 창 등)의 포커스를
뺏어간다. sap_handler.py/zrma_handler.py 등 실제 스크래핑 코드에는
SetForegroundWindow 같은 호출이 전혀 없음(코드 검색으로 확인) - 이건 우리
파이썬 코드가 거는 게 아니라 SAP GUI Scripting 자체의 내부 동작으로 보인다.

해결: Win32 LockSetForegroundWindow(LSFW_LOCK)을 걸면, "사용자의 실제 클릭"이
아닌 "다른 프로세스의 프로그램적 SetForegroundWindow 호출"만 Windows가 거부한다
- SAP 창을 화면 밖으로 숨기거나(오프스크린) 아예 다른 로그인 세션으로
격리하는(계정 분리) 방법과 달리, 창은 그대로 보이고 사용자가 직접 클릭해서
SAP 4개 세션을 확인/조작하는 것도 전혀 막히지 않는다 - 오직 "자동화 스스로
포커스를 채가는 것"만 차단한다(사용자 요구사항: "SAP 창 4개를 중간중간 직접
확인은 해야 함" - 완전 격리는 그래서 부적합했다).

불확실한 점(2026-08-24 작성 당시): SAP GUI Scripting이 실제로 SetForegroundWindow를
통해 활성화되는 게 맞는지는 라이브 SAP 없이 검증 불가 - 만약 다른 경로(예: 입력 큐
직접 조작)로 활성화된다면 이 락으로는 못 막을 수 있다. 다만 적용/해제가 저렴하고
부작용이 거의 없는 API라 시도해볼 가치가 있다(안 되면 그냥 효과가 없을 뿐).

2026-09-03 추가 (사용자 실측): 위 우려가 실제로 맞았다 - LockSetForegroundWindow
호출 자체는 automation.log에 실패 경고 0건으로 매 사이클 성공하는데도, 20분
자동루프가 돌 때 여전히 SAP 창이 사용자 작업 창의 포커스를 뺏어간다고 확인됨. 즉
SAP GUI Scripting은 이 락이 막는 SetForegroundWindow 경로로 창을 활성화하는 게
아닌 것으로 보인다(정확히 어떤 경로인지는 여전히 미확인 - 라이브 진단이 더
필요하면 그때 밝히면 됨). 원천 차단이 안 되니 **막는 대신 되돌리는** 보강을
추가했다: ForegroundLock 진입 시점의 foreground 창을 기억해뒀다가, 종료 시점에
foreground가 SAP류 프로세스면(=실제로 뺏긴 경우에만) 원래 창으로 스냅back한다.
사용자가 사이클 도중 직접 다른 창(SAP가 아닌)으로 옮겨간 경우는 "지금 foreground가
SAP인지"부터 확인하는 `_restore_previous()`가 자연히 건드리지 않는다. 기존
LockSetForegroundWindow 호출은 부작용 없고 일부 케이스엔 도움이 될 수 있어 그대로
유지 - 이번 추가는 순수 보강이며 `ForegroundLock`의 공개 인터페이스(`with
ForegroundLock():`)는 전혀 바뀌지 않았으므로 main.py 등 호출부는 손댈 필요 없음.
"""

import ctypes
import logging
import os

import win32api
import win32con
import win32gui
import win32process

logger = logging.getLogger(__name__)

_LSFW_LOCK = 1
_LSFW_UNLOCK = 2

_user32 = ctypes.WinDLL("user32", use_last_error=True)

# 자동화가 활성화시키는 SAP 관련 창의 실행파일 이름들 - 사이클 종료 시점에
# "지금 foreground가 SAP 창인지" 판단할 때만 쓴다(복원 여부를 결정하는 용도이지,
# 무언가를 막는 용도가 아니다).
_SAP_PROCESS_NAMES = {"nwbc.exe", "sapgui.exe", "saplogon.exe", "sapguiserver.exe"}


def _priming_tap():
    # Alt 키를 한 번 누르는 걸 시뮬레이션하면(실제로 화면엔 아무 효과 없음)
    # Windows가 "방금 이 프로세스가 입력을 받았다"고 인식해서 foreground 권한을
    # 일시적으로 내준다 - SetForegroundWindow/LockSetForegroundWindow 둘 다 이
    # 프라이밍 직후에 호출해야 성공한다(2026-08-24 실측). print_nlbl.py가 이미
    # 쓰던 검증된 방법.
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)


def _lock():
    # 2026-08-24 실측: main.py처럼 사용자가 직접 클릭한 적 없는 백그라운드
    # 콘솔 프로세스는 foreground 활성화 권한이 없어 LockSetForegroundWindow를
    # 조건 없이 그냥 호출하면 ERROR_ACCESS_DENIED(5)로 조용히 실패한다(반환값
    # 0, 예외는 안 남 - 그래서 반환값을 꼭 확인해야 함). 프라이밍 없이는 매번
    # 실패, 프라이밍 후에는 매번 성공(실측 확인). 이 함수는 절대 예외를 밖으로
    # 내보내면 안 된다 - 락 하나 실패했다고 SAP 루프 전체가 죽으면 본말전도이므로
    # 통째로 try/except.
    try:
        _priming_tap()
        ok = _user32.LockSetForegroundWindow(_LSFW_LOCK)
        if not ok:
            logger.warning(
                f"LockSetForegroundWindow(LOCK) 실패 (GetLastError={ctypes.get_last_error()}) "
                "- 이번 사이클은 포커스 보호 없이 진행됨"
            )
    except Exception as e:
        logger.warning(f"포커스 락 획득 중 예외 (무시하고 계속): {e}")


def _unlock():
    # UNLOCK은 실측상 프라이밍 없이도 항상 성공(자기가 건 락을 자기가 푸는
    # 것이라 별도 권한이 필요 없는 것으로 보임). _lock()과 같은 이유로 통째로
    # try/except.
    try:
        ok = _user32.LockSetForegroundWindow(_LSFW_UNLOCK)
        if not ok:
            logger.warning(f"LockSetForegroundWindow(UNLOCK) 실패 (GetLastError={ctypes.get_last_error()})")
    except Exception as e:
        logger.warning(f"포커스 락 해제 중 예외 (무시): {e}")


def _foreground_process_name():
    """지금 foreground 창을 소유한 프로세스의 exe 파일명(소문자, 확장자 포함)을
    반환. 뭔가 실패하면 None - 판단을 못 하면 복원을 시도하지 않는 게 안전하므로
    (엉뚱한 창을 SAP로 오판해서 복원해버리는 것보단 아무것도 안 하는 게 낫다)."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return None
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        try:
            path = win32process.GetModuleFileNameEx(handle, 0)
        finally:
            win32api.CloseHandle(handle)
        return os.path.basename(path).lower()
    except Exception:
        return None


def _restore_previous(prev_hwnd):
    # 2026-09-03 추가: LockSetForegroundWindow가 실제로는 SAP의 활성화 경로를
    # 못 막는다는 게 실측으로 확인돼서(위 모듈 docstring 참고), "막기" 대신
    # "사이클 끝나면 되돌리기"로 보강. prev_hwnd는 사이클 시작 시점 foreground
    # 창 - 지금(=사이클 끝) foreground가 SAP류 프로세스일 때만 되돌린다. 그
    # 사이 사용자가 SAP가 아닌 다른 창으로 직접 옮겨갔다면 _foreground_process_name()이
    # SAP가 아닌 걸 반환하므로 자연히 아무것도 안 하고 넘어간다 - 사용자의 그
    # 다음 작업을 억지로 되돌리는 사고를 막기 위한 안전장치.
    if not prev_hwnd:
        return
    try:
        if not win32gui.IsWindow(prev_hwnd):
            return
        if _foreground_process_name() not in _SAP_PROCESS_NAMES:
            return
        _priming_tap()
        ok = win32gui.SetForegroundWindow(prev_hwnd)
        if not ok:
            logger.warning("사이클 종료 후 포커스 복원 실패 (무시)")
    except Exception as e:
        logger.warning(f"포커스 복원 중 예외 (무시): {e}")


class ForegroundLock:
    """with ForegroundLock(): ... - 이 블록 동안 다른 프로세스(SAP GUI 포함)의
    프로그램적 창 활성화를 막는다. try/finally로 반드시 해제 - 안에서 예외가
    나도 락이 계속 걸린 채로 남지 않는다(계속 걸려있으면 다른 자동화 스크립트의
    SetForegroundWindow 호출도 실패하기 시작하므로 - morning_routine.py의 Bloomberg
    ID 자동입력이 이미 겪은 것과 같은 종류의 실패, 치명적이진 않지만 누적되면
    안 좋다). 2026-09-03: 락만으로는 SAP의 실제 활성화 경로를 못 막는 게 확인돼서,
    블록이 끝날 때 foreground가 SAP 창이면 블록 시작 전 창으로 되돌리는 보강을
    추가함 (_restore_previous 참고) - 공개 인터페이스(with 사용법)는 그대로다."""

    def __enter__(self):
        try:
            self._prev_hwnd = win32gui.GetForegroundWindow()
        except Exception:
            self._prev_hwnd = None
        _lock()
        return self

    def __exit__(self, exc_type, exc, tb):
        _unlock()
        _restore_previous(getattr(self, "_prev_hwnd", None))
        return False

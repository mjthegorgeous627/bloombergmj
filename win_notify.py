"""
윈도우 토스트 알림 - 새 오더 발견 시 PC 화면에 소리+팝업.

카카오 "나에게 보내기"(memo API)는 카카오 서버 정책상 푸시 알림이
잘 안 뜨는 경우가 많아 보조 채널로 추가. 별도 설정 불필요.
"""

import logging
import subprocess

from winotify import Notification, audio
import winotify

logger = logging.getLogger(__name__)

# 2026-09-04: 자동루프 도중 검은 콘솔 창이 순간순간 뜨는 증상의 원인을 찾음 -
# winotify가 토스트를 띄울 때마다 내부적으로 powershell.exe를 새로 실행하는데
# (winotify/__init__.py의 _run_ps), STARTUPINFO로 창을 숨기려고는 하지만
# CREATE_NO_WINDOW 플래그가 빠져있어 콘솔이 아예 안 만들어지진 않고 순간
# 깜빡인다. winotify는 이 프로젝트 전체에서 이 파일에서만 쓰므로, 설치된
# 패키지 파일을 직접 고치는 대신 여기서만 winotify.subprocess.Popen을
# CREATE_NO_WINDOW를 강제로 끼워넣는 래퍼로 바꿔치기한다 - winotify가 만드는
# powershell 프로세스에만 영향을 주고, 이 프로젝트의 다른 subprocess 호출은
# 전혀 건드리지 않는다(다른 파일은 각자 자기 모듈에서 import한 진짜
# subprocess.Popen을 그대로 쓰므로 무관).
_orig_popen = winotify.subprocess.Popen


def _no_window_popen(*args, **kwargs):
    kwargs["creationflags"] = kwargs.get("creationflags", 0) | subprocess.CREATE_NO_WINDOW
    return _orig_popen(*args, **kwargs)


winotify.subprocess.Popen = _no_window_popen


def send_windows_notification(title, message, duration="long"):
    """
    윈도우 액션 센터 토스트 알림 전송.
    반환: True(성공) / False(실패, 무시하고 계속 진행 가능)
    """
    try:
        toast = Notification(
            app_id="배송장 워크벤치",
            title=title or "새 오더 도착",
            msg=message or "",
            duration=duration,
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
        return True
    except Exception as e:
        logger.warning(f"[윈도우 알림] 전송 실패: {e}")
        return False

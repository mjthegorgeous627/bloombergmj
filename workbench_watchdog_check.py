"""
workbench_app.py 자체가 죽었는지 감시하는 외부 감시자.

문제: workbench_app.py 안에 있는 _sap_loop_watchdog_thread()는 main.py(SAP
자동화 루프)가 죽었을 때만 재시작해준다 - workbench_app.py 자기 자신이
죽으면 그 워치독 스레드도 함께 죽으므로 아무도 지켜봐주지 않는다("워치독을
누가 지켜보는가" 문제, 2026-08-24 구조 점검에서 발견). 이 스크립트는
workbench_app.py 프로세스 '밖'에서 Windows 작업 스케줄러가 주기적으로(권장:
5분마다) 실행해서 살아있는지만 확인하고, 없으면 재시작한다 - workbench_app.py
자체가 어떤 이유로든 죽어도 다음 체크 때 자동 복구되게 하는 게 목적이다.

설계: 한 번 실행하고 끝나는 스크립트다(무한루프로 계속 떠있는 감시 프로세스가
아님) - "지켜보는 프로세스를 계속 띄워두면 그 프로세스 자체가 죽었을 때
또 아무도 안 지켜본다"는 문제를 다시 반복하지 않으려고, 항상 떠있는 시스템
서비스인 작업 스케줄러가 반복 실행을 책임지게 했다. 작업 스케줄러 자체가
죽는 경우까지는 이 설계로 막을 수 없지만, 그건 이미 morning_routine.py(매일
08:30 자동 실행)도 똑같이 의존하고 있는 전제라 새로운 위험은 아니다.

등록 방법 (1회):
  schtasks /Create /TN "Workbench_Watchdog" /TR "\"<pythonw.exe 경로>\" \"<이 파일 경로>\"" ^
    /SC MINUTE /MO 5 /RL LIMITED /F
  또는 PowerShell의 New-ScheduledTask* cmdlet으로 등록(RepetitionInterval 5분,
  RepetitionDuration 매우 김, LogonType Interactive - morning_routine.py의
  SAP_Morning_Routine 작업과 같은 이유로 Interactive: 잠금 화면에서도 백그라운드
  예약 작업 자체는 계속 돌아야 하므로).
"""

import logging
import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "workbench_watchdog.log"


def _pythonw_exe():
    """pythonw.exe 경로 - 콘솔 창 없이 워크벤치를 띄우기 위함(2026-08-25).
    표준 CPython 설치 구조상 python.exe와 같은 폴더에 있다; 못 찾으면(임베디드
    배포 등 드문 경우) 현재 인터프리터로 그냥 폴백."""
    candidate = Path(sys.executable).with_name("pythonw.exe")
    return str(candidate) if candidate.exists() else sys.executable

# workbench_app.py 서버가 정상적으로 켜져있어야 할 시간대. main.py 루프의
# 워치독(WATCHDOG_BUSINESS_HOUR_START/END, 8~20시)보다 살짝 넓게 잡아서
# 아침 일찍/야근 대비 - 이 시간대 밖에서는 재시작을 시도하지 않는다(사용자가
# 의도적으로 꺼둔 것일 수 있으므로 밤에 갑자기 콘솔 창이 뜨는 걸 방지).
BUSINESS_HOUR_START = 7
BUSINESS_HOUR_END = 21

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
logger = logging.getLogger("workbench_watchdog")


def _workbench_running():
    """workbench_app.py가 서버 모드로 떠있는지 확인. --import-rows-json 같은
    1회성 CLI 호출(main.py/manual_order_handler.py가 workbench.db 반영용으로
    수시로 띄우는 짧은 서브프로세스)도 workbench_app.py를 실행하므로, 그것까지
    "서버가 떠있다"고 오판하지 않도록 --import로 시작하는 커맨드라인은 제외한다."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'workbench_app\\.py' -and "
        "$_.CommandLine -notmatch '--import' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
        )
    except Exception as exc:
        # 조회 자체가 실패했을 때 "죽었다"고 오판해 재시작을 시도하면, 사실은
        # 멀쩡히 떠있는 workbench 옆에 중복 서버가 하나 더 생겨 포트 충돌만
        # 낼 수 있다 - 판단 불가 상황에서는 아무것도 안 하는 쪽이 더 안전하다.
        logger.warning(f"프로세스 조회 실패 (판단 보류, 이번 체크는 건너뜀): {exc}")
        return True
    return bool(out.strip())


def _tray_running():
    """tray_icon.py가 떠있는지 확인. 상태를 보여주는 게 목적인 프로세스라
    _workbench_running()과 달리 --import류 제외 처리는 필요 없다."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'tray_icon\\.py' } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
        )
    except Exception as exc:
        logger.warning(f"트레이 아이콘 프로세스 조회 실패 (판단 보류, 이번 체크는 건너뜀): {exc}")
        return True
    return bool(out.strip())


def _restart_tray():
    """workbench와 별개로 죽어있으면 조용히 재시작 - 트레이 자신이 유일한
    상태 확인 수단이라 죽어있는 걸 사용자가 알아챌 방법이 없으므로, 이건
    워치독 재시작 경고 토스트 없이 조용히 되살린다(workbench 재시작과
    달리 이 자체는 자동화 동작에 영향 없는 표시용 프로세스라서)."""
    try:
        proc = subprocess.Popen(
            [_pythonw_exe(), "tray_icon.py"], cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.info(f"tray_icon.py 재시작 완료 (PID {proc.pid})")
    except Exception:
        logger.exception("tray_icon.py 재시작 실패")


def main():
    hour = datetime.now().hour
    if not (BUSINESS_HOUR_START <= hour < BUSINESS_HOUR_END):
        return

    if not _tray_running():
        logger.warning("tray_icon.py가 죽어있음을 감지 → 재시작 시도")
        _restart_tray()

    if _workbench_running():
        return

    logger.warning("workbench_app.py가 죽어있음을 감지 → 재시작 시도")
    # 콘솔 창이 사라진 뒤로(2026-08-25) 이게 죽었다가 살아난 걸 사용자가 알 수
    # 있는 유일한 신호이므로, 재시작 '전에' 먼저 눈에 띄게 알린다 - serve()가
    # 시작하면서 띄우는 평범한 "시작됨" 토스트는 이 뒤에 이어서 따로 뜬다.
    try:
        from win_notify import send_windows_notification
        send_windows_notification(
            "Workbench가 죽어있어 자동 재시작합니다", "잠시 후 대시보드를 새로고침해보세요.", duration="long",
        )
    except Exception:
        logger.warning("재시작 알림 토스트 실패(무시)", exc_info=True)
    try:
        proc = subprocess.Popen(
            [_pythonw_exe(), "workbench_app.py"], cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.info(f"workbench_app.py 재시작 완료 (PID {proc.pid})")
    except Exception:
        logger.exception("workbench_app.py 재시작 실패")


if __name__ == "__main__":
    main()

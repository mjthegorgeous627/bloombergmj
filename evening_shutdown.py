"""
저녁 강제종료 루틴 - Windows 작업 스케줄러에서 매일 19:00에 실행 (사용자 요청,
2026-08-31). morning_routine.py와 대칭되는 스크립트.

발단: 2026-08-28(금) 저녁 SAP 창을 닫고 퇴근했는데, 백그라운드의 자동화
루프(main.py)는 계속 돌면서 주말 내내 idle 상태이던 SAP 세션을 찔러봤고,
2026-08-30(일) 09:34 VL06O 새로고침 도중 응답이 끊겼다. SAP GUI Scripting
COM 호출은 자체 타임아웃이 없어서 그 자리에서 영원히 블록됐고, workbench의
내부 워치독(_sap_loop_watchdog_thread)이 30분마다 python 프로세스를 강제
재시작했지만 - taskkill이 죽이는 건 main.py/startup.py(python.exe)뿐이고
실제 SAP 프론트엔드(NWBC.exe/SapGuiServer.exe)는 별개 프로세스라 전혀 안
죽어서 그대로 좀비로 남아있었다. 게다가 새로 뜬 startup.py는 "이미 로그인된
세션이 있으면 재사용"하는 로직 때문에 매번 그 죽은 세션을 다시 붙잡고
다시 블록됨 - 그래서 일요일 오전부터 월요일 아침까지 20번 넘게 자동
재시작해도 하나도 안 풀렸다(그 재사용 로직 자체는 launch_sap_and_connect()에
타임아웃을 추가해 별도로 고침, see startup.py의 _run_with_timeout).

이 스크립트는 그 근본 원인과 별개로, 애초에 "퇴근 후 아무도 안 보는
긴 유휴 시간 동안 세션을 계속 열어두지 않는다"는 예방 차원의 안전장치다 -
평일 저녁~다음날 아침의 약 13시간 idle 구간에서 같은 클래스의 백엔드
타임아웃/네트워크 끊김이 재발할 기회 자체를 차단한다.

한다:
  1. STOP_AUTOMATION_FLAG 생성 - workbench_app.py의 내부 SAP 루프 워치독
     (_sap_loop_watchdog_thread)이 이 파일이 있으면 재시작을 완전히
     건너뛰므로, 아래 2번에서 강제종료한 직후 워치독이 "로그 무갱신"을
     감지하고 곧바로 되살리는 걸 막는다. morning_routine.py가 다음날
     아침 SAP 루프를 다시 시작할 때(startup.py→main.py의 run_loop()
     최초 진입부) 이 플래그를 지운다 - 기존 "SAP 루프 시작" 버튼과
     완전히 같은 관례(workbench_app.py의 sap_start_loop() 참고).
  2. main.py/startup.py(SAP 자동화 루프) 프로세스 강제 종료.
  3. 실제 SAP 프론트엔드 프로세스(NWBC/NwbcProcessAgent/SapGuiServer,
     구형 saplogon 포함) 강제 종료 - 위 배경 설명대로 이게 핵심.
  4. Excel(2026 배송장.xlsx)을 열어놓은 EXCEL.EXE는 굳이 안 건드림 -
     최신 상태로 저장된 채 열려있는 정도는 다음날 문제 없고, 강제 종료
     시 저장 안 된 변경사항이 있으면 오히려 위험.

절대 안 하는 것: workbench_app.py, tray_icon.py는 그대로 둔다. 이제
배송장 엑셀은 백업 용도이고 workbench가 실제 확인 수단이므로(사용자
확인, 2026-08-31), 저녁에도 대시보드 자체는 계속 조회 가능해야 한다 -
이 루틴이 끄는 건 "SAP에 붙어서 자동으로 도는 부분"뿐이다.

등록 방법 (1회, morning_routine.py와 동일한 관례 - Interactive 토큰,
LogonType Interactive, RunLevel Limited):
  schtasks /Create /TN "SAP_Evening_Shutdown" ^
    /TR "\"<pythonw.exe 경로>\" \"<이 파일 경로>\"" ^
    /SC DAILY /ST 19:00 /RL LIMITED /F
"""

import logging
import os
import subprocess
import threading

from config import LOG_FILE
from sap_handler import graceful_close_all_sap_connections

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STOP_AUTOMATION_FLAG = os.path.join(BASE_DIR, "stop_automation.flag")
EVENING_LOG_FILE = os.path.join(BASE_DIR, "evening_shutdown.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(EVENING_LOG_FILE, encoding="utf-8"),
        logging.FileHandler(LOG_FILE, encoding="utf-8"),  # automation.log에도 남겨 한 곳에서 확인 가능
    ],
)
logger = logging.getLogger(__name__)


def _pids_matching(*patterns):
    """morning_routine.py의 동일 함수와 같은 방식 - 살아있는 프로세스
    목록 조회(PID 파일 아님)."""
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


def set_stop_flag():
    """workbench_app.py 내부 워치독이 곧바로 되살리지 못하게 먼저 세운다 -
    sap_stop_loop()과 완전히 같은 관례."""
    try:
        with open(STOP_AUTOMATION_FLAG, "w", encoding="utf-8"):
            pass
        logger.info("stop_automation.flag 생성 (워치독 자동재시작 방지)")
    except Exception as e:
        logger.warning(f"stop_automation.flag 생성 실패 (무시하고 계속): {e}")


def stop_sap_loop():
    pids = _pids_matching(r"main\.py", r"startup\.py")
    if not pids:
        logger.info("SAP 자동루프 실행 중 아님 - 건너뜀")
        return
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
            )
        except Exception as e:
            logger.warning(f"PID {pid} 종료 실패 (무시): {e}")
    logger.info(f"SAP 자동루프 종료 완료 (PID {', '.join(map(str, pids))})")


def _graceful_logoff_with_timeout(timeout_sec=8):
    """강제종료 전에 정상 로그오프를 먼저 시도 - 타임아웃으로 감싸서
    SAP가 이미 완전히 응답 없는 상태여도 여기서 막히지 않게 한다
    (startup.py의 _run_with_timeout과 같은 이유, 같은 방식).

    2026-09-14: 이 스크립트가 매일 19:00에 로그오프 없이 그냥 프로세스를
    죽여온 게, SAP 백엔드에 유령 로그온을 남겨 다음 로그인 시도가
    "License Information for Multiple Logons" 팝업과 충돌하는 근본
    원인으로 실측 확인됨."""
    result = {}

    def runner():
        try:
            result["value"] = graceful_close_all_sap_connections()
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=runner, daemon=True)
    t.start()
    t.join(timeout_sec)
    if t.is_alive():
        logger.info("정상 로그오프 응답 없음 (타임아웃) - 강제종료로 진행")
        return False
    if "error" in result:
        logger.info(f"정상 로그오프 실패 (무시): {result['error']}")
        return False
    if result.get("value"):
        logger.info("정상 로그오프 완료")
    else:
        logger.info("정상 로그오프 생략 (연결된 SAP 세션 없음)")
    return result.get("value", False)


def stop_sap_gui():
    """실제 SAP 프론트엔드 프로세스 강제 종료 - startup.py의
    _kill_stuck_sap_gui_processes()와 같은 대상. python.exe는 안 건드림."""
    _graceful_logoff_with_timeout()

    ps_cmd = (
        "Get-Process -Name NWBC,NwbcProcessAgent,SapGuiServer,saplogon "
        "-ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15,
        )
        logger.info("SAP GUI 프로세스(NWBC/SapGuiServer 등) 종료 완료")
    except Exception as e:
        logger.warning(f"SAP GUI 프로세스 종료 실패 (무시): {e}")


def main():
    logger.info("=== 저녁 자동 종료 루틴 시작 ===")
    set_stop_flag()
    stop_sap_loop()
    stop_sap_gui()
    logger.info("=== 저녁 자동 종료 루틴 완료 (workbench/트레이는 그대로 둠) ===")


if __name__ == "__main__":
    main()

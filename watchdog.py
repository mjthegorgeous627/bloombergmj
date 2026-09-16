"""
자동루프 워치독.

main.py의 20분 자동루프(startup.py로 시작)가 SAP COM 쪽 문제로 조용히
죽거나(프로세스 자체가 사라짐) 멈춰버리는 경우, automation.log가 일정
시간 이상 갱신되지 않는 것으로 감지해서 startup.py를 자동으로 재시작한다.

사용법:
  python watchdog.py

동작 원칙:
  - 사용자가 런처의 "자동화 중지"를 눌러 stop_automation.flag가 있으면
    개입하지 않는다 (의도적으로 꺼둔 상태 존중).
  - 업무 시간(08:00~20:00)에만 개입한다 — 밤새 방치돼도 재로그인을
    반복 시도하지 않도록.
  - 재시작 직후에는 쿨다운(기본 8분) 동안 재판단하지 않는다 — SAP
    재로그인 + 4세션 셋업 시간을 확보하기 위함.
  - 최근 1시간 내 재시작이 3회 이상 반복되면(=재시작해도 계속 죽는
    상황) 쿨다운을 30분으로 늘려서 SAP 로그인 시도를 무한 반복하지
    않는다.
  - startup.py 자체가 "이미 로그인된 세션 있으면 재사용" 로직을
    가지고 있어(launch_sap_and_connect), SAP가 살아있는데 우리
    파이썬 프로세스만 죽은 경우에도 안전하게 재시작 가능.
"""

import os
import sys
import time
import subprocess
import logging
from datetime import datetime
from pathlib import Path

from config import LOG_FILE, REFRESH_INTERVAL_MINUTES
from holiday_check import skip_reason

BASE_DIR = Path(__file__).resolve().parent
STOP_FILE = BASE_DIR / "stop_automation.flag"
WATCHDOG_LOG = BASE_DIR / "watchdog.log"

CHECK_INTERVAL_SEC = 60
STALE_MINUTES = REFRESH_INTERVAL_MINUTES + 10          # 30분 무갱신 → 죽은 것으로 판단
RESTART_COOLDOWN_SEC = 8 * 60                           # 재시작 직후 재판단 유예
LONG_COOLDOWN_SEC = 30 * 60                             # 반복 실패 시 확대 유예
FLAP_WINDOW_SEC = 60 * 60                                # 이 시간 내 재시작이
FLAP_THRESHOLD = 3                                       # 이 횟수 이상이면 반복 실패로 판단
BUSINESS_HOUR_START = 8
BUSINESS_HOUR_END = 20

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(WATCHDOG_LOG, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
    force=True,
)
logger = logging.getLogger(__name__)


def _log_stale_seconds():
    if not os.path.exists(LOG_FILE):
        return None
    return time.time() - os.path.getmtime(LOG_FILE)


def _append_to_automation_log(message):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as fh:
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            fh.write(f"{ts} [WARNING] [Watchdog] {message}\n")
    except Exception as exc:
        logger.warning(f"automation.log 기록 실패: {exc}")


def _find_loop_pids():
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "($_.CommandLine -match 'main\\.py' -or $_.CommandLine -match 'startup\\.py') } | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
        )
    except Exception as exc:
        logger.warning(f"프로세스 조회 실패: {exc}")
        return []
    return [p.strip() for p in out.splitlines() if p.strip().isdigit()]


def _kill_stray_loop_processes():
    pids = _find_loop_pids()
    my_pid = os.getpid()
    for pid in pids:
        if int(pid) == my_pid:
            continue
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", pid],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
            )
            logger.info(f"강제 종료: PID {pid}")
        except Exception as exc:
            logger.warning(f"PID {pid} 종료 실패: {exc}")
    return pids


def _restart_loop():
    logger.warning("자동루프 무응답 감지 → 재시작 시도")
    _append_to_automation_log(f"automation.log {STALE_MINUTES}분 이상 무갱신 감지 → startup.py 재시작")

    killed = _kill_stray_loop_processes()
    if killed:
        time.sleep(3)

    if STOP_FILE.exists():
        try:
            STOP_FILE.unlink()
        except Exception:
            pass

    try:
        proc = subprocess.Popen(
            [sys.executable, "startup.py"],
            cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
        logger.info(f"startup.py 재실행됨 (PID {proc.pid})")
    except Exception as exc:
        logger.error(f"startup.py 재실행 실패: {exc}")


def run():
    logger.info(
        f"워치독 시작 (기준: {STALE_MINUTES}분 무갱신 시 재시작, "
        f"{CHECK_INTERVAL_SEC}초마다 점검, 업무시간 {BUSINESS_HOUR_START}~{BUSINESS_HOUR_END}시)"
    )

    restart_history = []
    cooldown_until = 0.0

    while True:
        time.sleep(CHECK_INTERVAL_SEC)
        now = time.time()

        if STOP_FILE.exists():
            continue

        # 2026-09-13: 평소엔 evening_shutdown.py가 매일 19:00에 STOP_FILE을
        # 남겨두고 morning_routine.py가 다음날 아침에만 지우므로 이 워치독은
        # 주말/공휴일에 자동으로 STOP_FILE이 그대로 남아 개입 안 하는 게
        # 보통이다. 그래도 그 플래그가 어떤 이유로든 없어진 채로 주말/공휴일
        # 08~20시를 맞으면 "무갱신"으로 오판해 SAP를 무인 상태로 다시 깨울 수
        # 있으므로 여기서도 직접 한 번 더 걸러낸다 (holiday_check.py 참고).
        if skip_reason():
            continue

        hour = datetime.now().hour
        if not (BUSINESS_HOUR_START <= hour < BUSINESS_HOUR_END):
            continue

        if now < cooldown_until:
            continue

        stale = _log_stale_seconds()
        if stale is None or stale <= STALE_MINUTES * 60:
            continue

        restart_history = [t for t in restart_history if now - t < FLAP_WINDOW_SEC]
        _restart_loop()
        restart_history.append(now)

        if len(restart_history) >= FLAP_THRESHOLD:
            logger.warning(
                f"최근 {FLAP_WINDOW_SEC // 60}분 내 재시작 {len(restart_history)}회 "
                f"→ 반복 실패로 판단, 유예를 {LONG_COOLDOWN_SEC // 60}분으로 확대"
            )
            cooldown_until = now + LONG_COOLDOWN_SEC
        else:
            cooldown_until = now + RESTART_COOLDOWN_SEC


if __name__ == "__main__":
    run()

"""SAP automation loop control (start/stop/status/watchdog) and per-order
SAP triggers (ZREC receive/commit, S/N relocation, session/order jump,
morning routine).

Split out of workbench_app.py on 2026-09-16 - see db.py's module docstring
and BLOOMBERG_HANDOFF.md. Pure code motion.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

from config import REFRESH_INTERVAL_MINUTES
from db import BASE_DIR, LOG_FILE, add_event, connect, ensure_db, logger, now_text
from holiday_check import skip_reason
from sap_handler import kill_stuck_sap_gui_processes

def _running_sap_loop_pids():
    """PIDs of any python process currently running main.py or startup.py as
    the SAP loop, regardless of whether launcher.py or this dashboard started
    it. Mirrors watchdog.py's own detection query so both agree on what
    counts as "a loop is already running" - a live process-list check instead
    of a lock file, so there's nothing to go stale if a process ever dies
    without cleaning up after itself."""
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
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        return []
    return [int(p.strip()) for p in out.splitlines() if p.strip().isdigit()]


# ---------------------------------------------------------------------------
# In-process SAP loop watchdog - ported from the old standalone watchdog.py.
#
# watchdog.py ran as its own separate console process, started manually via
# launcher.py's "워치독 시작" button, and auto-restarted startup.py whenever
# automation.log went 30+ minutes without a new line (main.py/startup.py's
# SAP GUI Scripting/COM session hanging or dying silently - a real, recurring
# failure mode: confirmed live in automation.log, e.g. 2026-08-05 09:02-09:56,
# a 54-minute silent hang - "VL06O 진입 중..." then nothing until an RPC
# failure - that only recovered because the user happened to notice on the
# workbench dashboard and clicked "SAP Start & Loop" by hand). watchdog.py
# itself stopped running sometime after 2026-07-24 (its own log has exactly
# one line, from that day, and no live python.exe process is running it) -
# nothing since then has started/monitored/restarted it, so every SAP hang
# since has needed a manual restart, which reads as "SAP randomly goes down"
# from the workbench side even though the underlying SAP GUI flakiness
# itself isn't new (the same class of COM error shows up as far back as
# March). Since workbench is the thing actually kept running/open all day
# now, the same restart logic lives here instead, tied to workbench_app.py's
# own process lifetime - a daemon thread started once from serve() - rather
# than a separate console window nobody's watching. watchdog.py is left in
# place and still usable by hand; this doesn't replace it, it just means the
# safety net no longer depends on remembering to click that button.
# ---------------------------------------------------------------------------

WATCHDOG_STALE_MINUTES = REFRESH_INTERVAL_MINUTES + 10   # 30분 무갱신 → 죽은 것으로 판단
WATCHDOG_CHECK_INTERVAL_SEC = 60
WATCHDOG_RESTART_COOLDOWN_SEC = 8 * 60                    # 재시작 직후 재판단 유예
WATCHDOG_LONG_COOLDOWN_SEC = 30 * 60                      # 반복 실패 시 확대 유예
WATCHDOG_FLAP_WINDOW_SEC = 60 * 60
WATCHDOG_FLAP_THRESHOLD = 3
WATCHDOG_BUSINESS_HOUR_START = 8
WATCHDOG_BUSINESS_HOUR_END = 20
WATCHDOG_CONSEC_FAILURE_CYCLES = 2   # 연속 전세션 연결 실패 사이클 수 → mtime staleness와 무관하게 재시작
STOP_AUTOMATION_FLAG = BASE_DIR / "stop_automation.flag"

# 2026-09-17: 사이클 "진행 중" 도중 멈추는 새 증상(get_sap_session() 등 SAP GUI
# Scripting 호출이 예외 없이 그냥 영원히 블로킹) 전용 감지. main.py가 cycle_state.json에
# running/idle을 기록한다 - 사이클 사이 정상 유휴 시간(최대 20분)과 달리 "running"
# 상태로 오래 멈춰있는 건 정상적으로 있을 수 없으므로 훨씬 짧은 기준으로 잡아도
# 오탐이 없다. WATCHDOG_STALE_MINUTES(30분)보다 훨씬 빠르게 잡는 게 목적.
WATCHDOG_CYCLE_HANG_SECONDS = 3 * 60
CYCLE_STATE_FILE = BASE_DIR / "cycle_state.json"


def _watchdog_cycle_hung_seconds():
    """cycle_state.json이 "running" 상태로 얼마나 오래 멈춰있는지(초). 유휴
    상태이거나 파일이 없거나 형식이 깨졌으면 None (모두 "안 멈췄음"으로 취급)."""
    try:
        state = json.loads(CYCLE_STATE_FILE.read_text(encoding="utf-8"))
        if state.get("status") != "running":
            return None
        since = datetime.fromisoformat(state["since"])
        return (datetime.now() - since).total_seconds()
    except Exception:
        return None


def _watchdog_log_stale_seconds():
    if not LOG_FILE.exists():
        return None
    return time.time() - os.path.getmtime(LOG_FILE)


# 2026-08-21 실측: startup.py 프로세스가 캐시된(죽은) SAP GUI Scripting COM
# 엔진을 붙잡은 채로 계속 도는 상태(NWBC가 사라진 뒤에도 매 사이클 4세션
# 전부 "연결 실패" 에러만 남기고 "이번 실행 완료"까지 찍음)에서는 로그
# 파일이 매 사이클(기본 20분)마다 계속 갱신되므로 위 mtime 기반 staleness
# 판정이 절대 걸리지 않는다 - 워치독이 "살아있다"고 오판해 하루 종일
# 재시작을 안 함. 그래서 mtime과 별개로, 로그 "내용"을 봐서 최근 사이클들이
# 연속으로 전부 실패였는지도 판정한다.
def _watchdog_recent_cycle_texts(n):
    """automation.log에서 최근 n개 사이클(각 '실행 시작:' 구분)의 텍스트를
    오래된 순으로 반환. 파일이 커질 수 있으니 끝에서 일정 크기만 읽는다."""
    try:
        with LOG_FILE.open("rb") as fh:
            fh.seek(0, os.SEEK_END)
            size = fh.tell()
            fh.seek(max(0, size - 200_000))
            data = fh.read()
        text = data.decode("utf-8", errors="replace")
    except Exception:
        return []
    parts = text.split("실행 시작:")
    cycles = ["실행 시작:" + p for p in parts[1:]]
    return cycles[-n:] if cycles else []


def _watchdog_cycle_all_sessions_failed(cycle_text):
    """한 사이클 안에서 자동 루프가 실제로 도는 세션이 전부 '연결 실패'로
    끝났는지 판단. 2026-08-26부터 자동 루프는 세션0/1(VL06O/VL10G)만 돌고
    세션2/3(ZRMA RLKR/Q2)은 main.py의 AUTO_SCAN_ZRMA=False로 빠졌으므로
    (그 세션들 몫의 '연결 실패:' 로그 자체가 안 생김), 임계값도 4가 아니라
    2 - AUTO_SCAN_ZRMA를 다시 켜면 여기도 4로 되돌려야 한다."""
    return cycle_text.count("연결 실패:") >= 2


def _watchdog_consecutive_total_failures():
    """가장 최근에 완료된("이번 실행 완료" 있는) 사이클들부터 거꾸로 훑어서,
    연속으로 전세션 실패인 사이클이 몇 개인지 센다(중간에 하나라도 성공이
    섞이면 거기서 멈춤). 아직 진행 중인 마지막 사이클은 제외."""
    cycles = _watchdog_recent_cycle_texts(WATCHDOG_CONSEC_FAILURE_CYCLES + 1)
    completed = [c for c in cycles if "이번 실행 완료" in c]
    count = 0
    for c in reversed(completed):
        if _watchdog_cycle_all_sessions_failed(c):
            count += 1
        else:
            break
    return count


def _watchdog_append_log(message):
    try:
        with LOG_FILE.open("a", encoding="utf-8") as fh:
            fh.write(f"{now_text()} [WARNING] [Workbench-Watchdog] {message}\n")
    except Exception:
        pass


def _watchdog_restart_loop(reason):
    """2026-09-17 실측(스크린샷, "License Information for Multiple Logons"
    팝업에 걸려 로그인 화면에 멈춰선 SAP): 여기서는 python 프로세스(startup.py)
    만 taskkill하고 실제 SAP 프론트엔드(NWBC/SapGuiServer)는 그대로 뒀다 -
    startup.py 자체의 재연결 로직(launch_sap_and_connect())은 이미 2026-08-30
    사고 이후 죽은 SAP GUI를 직접 찾아 죽이는 kill_stuck_sap_gui_processes()를
    쓰도록 고쳐졌는데, 이 워치독(automation.log 30분 무갱신 감지 시 재시작하는
    경로)만 그 수정을 안 받고 예전 방식 그대로였다. 그 결과: python만 죽고
    실제로 멈춰있던 SAP GUI는 좀비로 남아 SAP 백엔드에 로그온 상태를 유지,
    재시작된 startup.py의 새 로그인 시도가 그 좀비와 충돌해 License
    Information for Multiple Logons 팝업에 다시 멈춰서는 걸 반복했다.
    kill_stuck_sap_gui_processes()를 여기서도 호출해 launch_sap_and_connect()
    가 하는 것과 같은 정리를 거치게 한다."""
    pids = _running_sap_loop_pids()
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass
    if pids:
        time.sleep(3)
    try:
        kill_stuck_sap_gui_processes()
    except Exception as exc:
        _watchdog_append_log(f"SAP GUI 프로세스 정리 실패 (무시하고 재시작 계속): {exc}")
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    # 죽은 이전 프로세스가 남긴 "running" 상태가 그대로면, 새로 뜬 프로세스가
    # 자기 첫 사이클을 시작하기 전(세션 4개 재설정에 보통 1분 내외 걸림) 그 낡은
    # running 상태를 자기 것으로 오인해 워치독이 곧바로 다시 "멈췄다"고 오판할 수
    # 있다 - WATCHDOG_RESTART_COOLDOWN_SEC(8분)이 이미 막아주지만, 상태도 같이
    # 지워 이중으로 방지한다.
    if CYCLE_STATE_FILE.exists():
        try:
            CYCLE_STATE_FILE.unlink()
        except Exception:
            pass
    try:
        proc = subprocess.Popen(
            [sys.executable, "startup.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
        )
        _watchdog_append_log(f"{reason} → startup.py 재시작 (PID {proc.pid})")
    except Exception as exc:
        _watchdog_append_log(f"startup.py 재시작 실패: {exc}")


def _sap_loop_watchdog_thread():
    """Runs forever in a daemon thread started once from serve(). Same
    detect-and-restart rules as the old watchdog.py (see module docstring
    above) - respects a manual stop_automation.flag, only acts during
    business hours, and backs off if restarting keeps not fixing it."""
    restart_history = []
    cooldown_until = 0.0
    while True:
        time.sleep(WATCHDOG_CHECK_INTERVAL_SEC)
        now = time.time()

        if STOP_AUTOMATION_FLAG.exists():
            continue
        # 2026-09-13: watchdog.py의 동일 지점과 같은 이유(holiday_check.py
        # 참고) - STOP_AUTOMATION_FLAG가 무슨 이유로든 없어진 채로 주말/공휴일
        # 업무시간을 맞아도 여기서 한 번 더 걸러 무인 재기동을 막는다.
        if skip_reason():
            continue
        hour = datetime.now().hour
        if not (WATCHDOG_BUSINESS_HOUR_START <= hour < WATCHDOG_BUSINESS_HOUR_END):
            continue
        if now < cooldown_until:
            continue

        stale = _watchdog_log_stale_seconds()
        is_stale = stale is not None and stale > WATCHDOG_STALE_MINUTES * 60
        consec_fail = _watchdog_consecutive_total_failures()
        is_flatlined = consec_fail >= WATCHDOG_CONSEC_FAILURE_CYCLES
        cycle_hung_sec = _watchdog_cycle_hung_seconds()
        is_cycle_hung = cycle_hung_sec is not None and cycle_hung_sec > WATCHDOG_CYCLE_HANG_SECONDS
        if not (is_stale or is_flatlined or is_cycle_hung):
            continue

        if is_cycle_hung:
            reason = f"사이클 진행 중 {int(cycle_hung_sec)}초째 무응답 감지 (SAP GUI Scripting 블로킹 추정)"
        elif is_stale:
            reason = f"automation.log {WATCHDOG_STALE_MINUTES}분 이상 무갱신 감지"
        else:
            reason = f"연속 {consec_fail}회 전세션 SAP 연결 실패 감지"

        restart_history = [t for t in restart_history if now - t < WATCHDOG_FLAP_WINDOW_SEC]
        _watchdog_restart_loop(reason)
        restart_history.append(now)

        if len(restart_history) >= WATCHDOG_FLAP_THRESHOLD:
            cooldown_until = now + WATCHDOG_LONG_COOLDOWN_SEC
        else:
            cooldown_until = now + WATCHDOG_RESTART_COOLDOWN_SEC


def sap_status():
    pids = _running_sap_loop_pids()
    return {"loopRunning": bool(pids), "pids": pids}


def sap_start_loop():
    """"1. SAP Start & Loop" - refuses to start a second loop (from either
    front-end) on top of one that's already running, instead of silently
    launching a process that will fight the existing one over the same 4 SAP
    sessions.

    2026-08-25: startup.py used to run in its own CREATE_NEW_CONSOLE window,
    and closing that window was the only way to stop the loop (the message
    below used to say "그 창을 먼저 종료하세요"). Now that every spawned
    script runs windowless (see sap_stop_loop()), a start here also clears
    STOP_AUTOMATION_FLAG - otherwise a loop stopped via the "중지" button
    would come back up but the watchdog thread would stay silently disabled
    forever (it skips entirely whenever that flag file exists, see
    _sap_loop_watchdog_thread())."""
    pids = _running_sap_loop_pids()
    if pids:
        return {
            "started": False,
            "message": f"이미 SAP 자동루프가 실행 중입니다 (PID {', '.join(map(str, pids))}). "
                       "새로 시작하려면 먼저 중지하세요.",
        }
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    proc = subprocess.Popen(
        [sys.executable, "startup.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def sap_stop_loop():
    """"SAP 자동루프 중지" (2026-08-25) - startup.py/main.py가 더 이상
    CREATE_NEW_CONSOLE 창으로 안 뜨면서, 그 창을 닫는 게 유일한 중지 수단이던
    게 없어져서 새로 추가한 버튼. STOP_AUTOMATION_FLAG를 남겨(워치독이 곧바로
    되살리지 못하게) 지금 떠있는 루프 프로세스를 taskkill로 종료한다. 다시
    시작하려면 sap_start_loop()/"1. SAP Start & Loop"가 이 플래그를 지운다."""
    pids = _running_sap_loop_pids()
    STOP_AUTOMATION_FLAG.write_text("", encoding="utf-8")
    if not pids:
        return {"stopped": False, "message": "실행 중인 SAP 자동루프가 없습니다."}
    killed = []
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            killed.append(pid)
        except Exception:
            pass
    return {"stopped": True, "message": f"SAP 자동루프 중지됨 (PID {', '.join(map(str, killed))})"}


def workbench_force_restart():
    """2026-08-28 신설: 트레이 아이콘의 "강제 재시작"과 완전히 동일한 동작을
    workbench 페이지 안에서도 쓸 수 있게 함(사용자가 "트레이에 있는 걸 여기서도
    바꿀 수 있냐" 요청). 지금 이 요청을 처리하는 프로세스 자신을 죽여야 하므로,
    죽이는 동작은 이 프로세스가 살아있는 채로 직접 하지 않고 - 응답을 먼저
    보내야 하고, taskkill이 자기 자신을 겨냥하면 응답 전송이 끊길 수 있음 -
    별도의 분리된(detached) PowerShell 한 줄짜리 헬퍼를 띄워서 0.8초 뒤에
    실행하게 위임한다. 그 헬퍼가 workbench_app.py 프로세스를 전부 taskkill한
    뒤(단일 인스턴스 뮤텍스는 프로세스 종료 시 OS가 자동으로 풀어줌) 하나만
    새로 띄운다."""
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    pythonw = str(pythonw) if pythonw.exists() else sys.executable
    ps_cmd = (
        "Start-Sleep -Milliseconds 800; "
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'workbench_app\\.py' } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; "
        "Start-Sleep -Milliseconds 500; "
        f"Start-Process -FilePath '{pythonw}' -ArgumentList 'workbench_app.py' -WorkingDirectory '{BASE_DIR}'"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.warning("Workbench 강제 재시작 요청 (페이지에서 트리거)")
        return {"restarting": True, "message": "1~2초 후 재시작됩니다. 잠시 후 이 페이지를 새로고침하세요."}
    except Exception as exc:
        logger.exception("Workbench 강제 재시작 트리거 실패")
        return {"restarting": False, "message": f"재시작 트리거 실패: {exc}"}


def sap_run_all():
    """"2. SAP 전체 조회" - if a loop is already running somewhere, just signal
    it via the same run_now.flag file launcher.py's "즉시 조회" button uses
    (no new process touching the shared sessions). If nothing is running,
    spawn bare `main.py` (no --once) - it does an immediate full pass through
    all 4 sessions first thing inside run_loop(), then keeps looping every
    REFRESH_INTERVAL_MINUTES on its own, same as "1. SAP Start & Loop" minus
    the SAP-launch/session-setup steps (this assumes the 4 sessions already
    exist, same assumption --once made)."""
    pids = _running_sap_loop_pids()
    if pids:
        (BASE_DIR / "run_now.flag").write_text("", encoding="utf-8")
        return {"mode": "signal", "message": "실행 중인 자동루프에 즉시 조회 신호를 보냈습니다."}
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    proc = subprocess.Popen(
        [sys.executable, "main.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {
        "mode": "spawned_loop", "pid": proc.pid,
        "message": "실행 중인 루프가 없어 전체 조회 후 자동 루프를 새로 시작했습니다.",
    }


SESSION_LABELS = {0: "VL06O", 1: "VL10G", 2: "ZRMA RLKR", 3: "ZRMA Q2"}


def sap_run_session(idx):
    """세션 하나만 즉시 조회 - launcher.py(0minjung.bat)의 "특정 창만 조회"
    버튼과 완전히 같은 메커니즘(run_now_N.flag)을 워크벤치 대시보드에도
    노출한다(2026-08-26, 사용자가 launcher.py 대신 워크벤치를 쓴다고 요청).
    now.py도 이 파일을 씀 - 셋 다 같은 트리거를 공유. 세션2/3(ZRMA RLKR/Q2)은
    main.py의 AUTO_SCAN_ZRMA=False로 20분 자동루프에서 빠졌으니, 그 두 개를
    확인하고 싶을 때 이 버튼이 사실상 유일한 수단이다.

    run_loop()가 이 플래그를 최대 5초 주기로만 확인하므로(sleep(5) 루프 안),
    자동루프 자체가 안 돌고 있으면 플래그를 써도 아무도 안 지켜봐서 조용히
    무시된다 - sap_run_all()의 "루프 없으면 새로 띄움" 폴백과 달리 여기선
    그렇게 안 한다(세션 하나만 보자고 새 main.py 루프를 통째로 띄우는 건
    과함) - 대신 루프가 없다고 먼저 알려준다."""
    if idx not in SESSION_LABELS:
        return {"started": False, "message": "세션 번호는 0~3이어야 합니다."}
    if not _running_sap_loop_pids():
        return {
            "started": False,
            "message": "실행 중인 SAP 자동루프가 없습니다. 먼저 '1. SAP Start & Loop'를 누르세요.",
        }
    (BASE_DIR / f"run_now_{idx}.flag").write_text("", encoding="utf-8")
    return {"started": True, "message": f"세션{idx} ({SESSION_LABELS[idx]}) 단독 조회 요청됨 (최대 5초 내 시작)"}


def sap_run_order(order_no):
    """"3. 오더번호 반영" - opens its own dedicated SAP session (see
    manual_order_handler._open_new_sap_session), so it's safe to run alongside
    a loop in general. The one exception (ZRX orders briefly reading VL06O
    session 0 for its OBD) is a narrow, low-frequency race that launcher.py's
    equivalent button doesn't guard against either - not worth extra friction
    here for a window this small."""
    order_no = (order_no or "").strip()
    if not order_no:
        return {"started": False, "message": "오더번호를 입력하세요."}
    proc = subprocess.Popen(
        [sys.executable, "order.py", order_no], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def _zrec_item_rows(item_ids):
    """체크된 item_ids에 대한 order_items+orders 조인 행. ZREC 관련 두
    함수(zrec_lookup_batch/zrec_commit_one)와 sn_relo_one()이 공유."""
    if not item_ids:
        return []
    ensure_db()
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT i.*, o.order_no AS o_order_no, o.order_type AS o_order_type, o.id AS o_id "
            f"FROM order_items i JOIN orders o ON o.id = i.order_id "
            f"WHERE i.id IN ({placeholders})",
            item_ids,
        ).fetchall()
    by_id = {r["id"]: dict(r) for r in rows}
    # 원래 요청한 순서(item_ids)대로 반환 - IN절은 순서를 보장하지 않음
    return [by_id[i] for i in item_ids if i in by_id]


def sn_relo_one(item_id):
    """회수 오더 수집 시 SAP상 S/N이 없어서('X') 워크벤치에 X로 찍혔던
    항목 하나를 대상으로 한다 - 사용자가 실제로 회수한 S/N을 이미
    item.serial에 수동 입력해둔 상태라고 가정하고, 그 값을 VA02로 오더를
    열어 Technical Objects > Serial Numbers에 입력·저장 시도한다
    (2026-08-19, 사용자 요청). sn_relo_handler.py를 서브프로세스로 실행 -
    _zrec_run_receive()와 같은 '서브프로세스 하나 = SAP 세션 하나 열고
    쓰고 닫기' 패턴 그대로.

    저장 성공(status='saved')은 그 S/N의 firm/cust가 오더와 일치한다는
    뜻 - 이 항목은 이제 일반 회수 항목과 똑같은 상태이니, 이어서 기존
    "5. ZREC 처리" 버튼으로 처리하면 된다(자동으로 이어붙이지 않음 - 이
    함수는 오더 반영 성공/실패 판정까지만 책임진다). 실패(status='error')는
    firm/cust 불일치로 추정 - paper relo 프로세스가 필요하다는 뜻이라
    sn_relo_error에 사유를 남겨 대시보드에 표시한다. 'skipped'(오더에
    이미 - 워크벤치가 모르던 - 다른 S/N이 들어있던 경우)도 에러는 아니지만
    워크벤치 기록과 다를 수 있어 확인이 필요하므로 마찬가지로 표시한다."""
    items = _zrec_item_rows([item_id])
    if not items:
        raise RuntimeError("item not found")
    item = items[0]

    serial = (item["serial"] or "").strip()
    order_no = (item["o_order_no"] or "").strip()
    order_label = f"{item['o_order_type'] or ''} {order_no}".strip()
    if not serial or serial.upper() == "X":
        raise RuntimeError("실제 회수된 S/N을 먼저 워크벤치에 입력해야 합니다 (현재 S/N 칸이 비어있거나 X)")
    if not order_no:
        raise RuntimeError("오더번호가 없어 처리할 수 없습니다")

    # portal_error와 같은 이유 - 재시도를 시작하는 순간 즉시 비워서, 이번
    # 시도가 끝나기 전까지는 화면에 지난 실패가 안 남아있게 한다.
    with connect() as con:
        con.execute("UPDATE order_items SET sn_relo_error='' WHERE id=?", (item_id,))

    args = [sys.executable, str(BASE_DIR / "sn_relo_handler.py"), "fill",
            "--order", order_no, "--serial", serial]
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=45, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-15:]) if lines else "(출력 없음)"
        with connect() as con:
            con.execute("UPDATE order_items SET sn_relo_error=? WHERE id=?", (f"실행 실패: {detail}", item_id))
        raise RuntimeError(f"S/N 오더 반영 실패 (결과 파싱 불가): {detail}")

    payload = json.loads(json_line[len("RESULT_JSON:"):])
    payload["itemId"] = item_id
    payload["orderLabel"] = order_label

    with connect() as con:
        if payload.get("status") == "saved":
            con.execute("UPDATE order_items SET sn_relo_error='' WHERE id=?", (item_id,))
            add_event(con, item["o_id"], f"S/N '{serial}' 오더 반영 성공: {payload.get('message', '')}")
        else:
            con.execute(
                "UPDATE order_items SET sn_relo_error=? WHERE id=?",
                (payload.get("message") or "알 수 없는 실패", item_id),
            )
            add_event(con, item["o_id"], f"S/N '{serial}' 오더 반영 실패({payload.get('status')}): {payload.get('message', '')}")

    return payload


def zrec_lookup_batch(item_ids):
    """"ZREC 준비" 1단계 (todo c, 2026-08-19) - 체크된 회수 항목들의
    시리얼을 전부 모아서 ZIH08에 **한 번에** 조회한다. 사용자의 실제 하루
    업무 방식("오늘 회수된 모든 시리얼을 한번에 넣고 확인한 뒤, ZREC는
    하나씩") 그대로 - 이전 버전(zrec_handler.py의 prepare 서브커맨드)은
    항목마다 ZIH08을 매번 새로 열어 조회했는데, 그러면 ZIH08 세션을 계속
    재사용/재조회하게 되어 느리고 실제 업무 방식과도 다름. 여기서는
    조회만 하고 ZREC 화면은 전혀 건드리지 않는다 - 실제로 채우고 접수하는
    건 zrec_commit_one()이 항목별로 담당(2026-08-27부터 채우기+클릭이
    한 호출로 통합됨)."""
    items = _zrec_item_rows(item_ids)
    serial_items = [
        i for i in items
        if i["mode"] != "material_only" and (i["serial"] or "").strip()
    ]
    serials = [i["serial"].strip() for i in serial_items]

    ih08_by_serial = {}
    if serials:
        args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "lookup", "--serials", *serials]
        child_env = os.environ.copy()
        child_env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            args, cwd=str(BASE_DIR), env=child_env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
            timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        output = result.stdout or ""
        json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
        if not json_line:
            lines = [ln for ln in output.strip().splitlines() if ln.strip()]
            detail = "\n".join(lines[-10:]) if lines else "(출력 없음)"
            raise RuntimeError(f"ZIH08 조회 실패: {detail}")
        ih08_by_serial = json.loads(json_line[len("RESULT_JSON:"):])

    plans = []
    for item in items:
        serial = (item["serial"] or "").strip()
        material = (item["material"] or "").strip()
        qty = item["qty"] or 1
        non_serial = (item["mode"] == "material_only") or not serial
        order_no = item["o_order_no"] or ""
        order_label = f"{item['o_order_type'] or ''} {order_no}".strip()
        warnings = []
        ih08_info = {}

        if non_serial:
            if not material:
                warnings.append("Material 정보가 없어 ZREC 준비를 할 수 없습니다.")
            filled_order = order_no
            warnings.append("Non-serialized 품목이라 ZIH08 대조를 건너뛰었습니다 - Reference Document를 직접 확인하세요.")
        else:
            ih08_info = ih08_by_serial.get(serial, {})
            matched_order = ih08_info.get("matched_order", "")
            if matched_order:
                filled_order = matched_order
                if order_no and matched_order != order_no:
                    warnings.append(
                        f"workbench 오더번호({order_no})와 ZIH08 매칭 오더(KDAUF={matched_order})가 다릅니다 - "
                        "ZIH08 매칭값을 사용합니다. 반드시 직접 확인하세요."
                    )
            else:
                filled_order = order_no
                warnings.append("ZIH08에서 매칭된 Sales Order(KDAUF)를 찾지 못했습니다 - workbench 오더번호로 대체했습니다. 반드시 직접 확인하세요.")
            if ih08_info.get("plant") == "6507" and ih08_info.get("location") == "0052":
                warnings.append("ZIH08 조회 결과 이 시리얼은 이미 Plant 6507/Location 0052로 접수되어 있습니다 - 중복 접수 주의.")

        plans.append({
            "itemId": item["id"],
            "orderLabel": order_label,
            "serial": serial,
            "material": material,
            "qty": qty,
            "description": item.get("description") or "",
            "nonSerial": non_serial,
            "filledOrder": filled_order,
            "ih08": ih08_info,
            "warnings": warnings,
        })

    return {"plans": plans}


def _zrec_run_receive(item_id, order_no, commit, timeout):
    """zrec_handler.py의 receive 서브커맨드를 서브프로세스로 실행 -
    commit=False면 화면만 채우고 멈춤(dry-run, 수동 디버깅용으로만 남겨둠 -
    workbench는 2026-08-27부터 항상 commit=True로만 호출), commit=True면
    채우기부터 검증, Receive Equipment 클릭까지 한 번에 실행한다.
    zrec_commit_one()의 실행부."""
    items = _zrec_item_rows([item_id])
    if not items:
        raise RuntimeError("item not found")
    item = items[0]

    serial = (item["serial"] or "").strip()
    material = (item["material"] or "").strip()
    qty = item["qty"] or 1
    non_serial = (item["mode"] == "material_only") or not serial
    if non_serial and not material:
        raise RuntimeError("Material 정보가 없어 ZREC 처리를 할 수 없습니다 (Non-serialized인데 Material도 없음)")
    if not non_serial and not serial:
        raise RuntimeError("Serial 정보가 없어 ZREC 처리를 할 수 없습니다")
    order_no = (order_no or item["o_order_no"] or "").strip()
    if not order_no:
        raise RuntimeError("오더번호가 없어 ZREC 처리를 할 수 없습니다")

    args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "receive", "--order", order_no]
    if non_serial:
        args += ["--material", material, "--qty", str(qty), "--non-serial"]
    else:
        args += ["--serial", serial]
    if commit:
        args.append("--commit")

    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-15:]) if lines else "(출력 없음)"
        action = "Receive Equipment 클릭" if commit else "ZREC 채우기"
        raise RuntimeError(f"{action} 실패 (결과 파싱 불가): {detail}")

    payload = json.loads(json_line[len("RESULT_JSON:"):])
    payload["itemId"] = item_id
    payload["orderLabel"] = f"{item['o_order_type'] or ''} {item['o_order_no'] or ''}".strip()
    with connect() as con:
        verb = "ZREC 실제 접수(Receive Equipment 클릭)" if commit else "ZREC 준비"
        add_event(con, item["o_id"], f"{verb} 완료: " + " ".join(args[2:]))
    return payload


def zrec_commit_one(item_id, order_no):
    """"ZREC 준비" 2단계 (todo c, 2026-08-19 - 사용자 확정: 확인 팝업에서
    "확인"을 누르면 사람이 SAP에서 직접 누르는 대신 자동으로 Receive
    Equipment를 클릭) - ZREC 화면에 Plant/Reference Document/Serial(또는
    Non-serialized Material+수량)을 채우고, 화면값을 재확인(검증)한 뒤 곧바로
    Receive Equipment를 클릭해 성공/실패 팝업까지 한 번에 처리한다.

    2026-08-27 이전에는 이 채우기를 별도의 zrec_fill_one() 호출(dry-run)로
    먼저 한 번 하고, 이 함수가 세션을 리셋해서 "안전하게" 똑같은 값을
    다시 채운 뒤 클릭했다. 하지만 배치 확인창(runZrecPrepare 2단계)이
    ZIH08 조회 결과만으로 이미 사람 승인을 받은 뒤 이 함수가 항목마다
    중간 팝업 없이 곧바로 호출되므로, fill_one이 채운 화면을 사람이 보고
    판단하는 순간이 실제로는 없었다 - 세션 리셋+동일 값 재입력은 안전
    효과 없이 SAP 왕복만 두 배로 내는 낭비였다(사용자 지적, 2026-08-27).
    지금은 이 함수 하나가 채우기부터 클릭까지 한 세션에서 끝낸다 -
    안전장치는 그대로 유지: _validate_zrec_fields()가 클릭 직전 화면값을
    재확인하고(엉뚱한 값으로 접수 방지), 클릭 후에는 아래처럼 상태바로
    1차 판정한다.

    **주의**: zrec_handler.py의 성공/실패 판정(_popup_text 문구 매칭)은
    실사용에서 매번 오판(실제 성공을 실패로 표시)하는 게 확인됐다
    (2026-08-19). 상태바 MessageType 우선 판정으로 고쳤지만 그 판정
    자체도 참고용일 뿐이니 결과의 success 값을 곧이곧대로 믿지 말 것 -
    실제 done 처리 여부는 이어지는 zrec_verify_batch()의 ZIH08 재조회
    (Plant/Location 6507/0052 확인)로 최종 결정된다."""
    return _zrec_run_receive(item_id, order_no, commit=True, timeout=30)


def zrec_verify_batch(item_ids):
    """"ZREC 준비" 4단계(완료 확인) - 방금 Receive Equipment를 누른
    항목들의 시리얼을 ZIH08로 재조회해서 Plant/Location이 6507/0052로
    들어왔는지 확인한다 (todo c, 2026-08-19: 사용자 요청 - "zih08 조회후
    6507 0052 들어왔는지 조회 확인"). Non-serialized 항목은 ZIH08이
    시리얼 기반이라 확인 대상에서 제외 - Non-serialized는 ZIH08/시리얼
    개념이 없어 애초에 이 확인 방법이 적용 안 됨."""
    items = _zrec_item_rows(item_ids)
    serial_items = [i for i in items if i["mode"] != "material_only" and (i["serial"] or "").strip()]
    serials = [i["serial"].strip() for i in serial_items]
    if not serials:
        return {"results": []}

    args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "verify", "--serials", *serials]
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-10:]) if lines else "(출력 없음)"
        raise RuntimeError(f"ZREC 완료 확인 실패: {detail}")
    by_serial = json.loads(json_line[len("RESULT_JSON:"):])

    results = []
    for item in serial_items:
        serial = item["serial"].strip()
        row = by_serial.get(serial, {})
        results.append({
            "itemId": item["id"],
            "orderLabel": f"{item['o_order_type'] or ''} {item['o_order_no'] or ''}".strip(),
            "serial": serial,
            "receivedOk": bool(row.get("received_ok")),
            "plant": row.get("plant", ""),
            "location": row.get("location", ""),
        })
    return {"results": results}


def sap_open_session(tcode):
    """"4. New SAP session open" - just asks SAP to spawn a new window; never
    touches sessions 0-3's screens, safe to run anytime including mid-loop
    (open_session.py already anticipated this - session lookups go through
    the SessionNumber map, not raw index, precisely so an extra session
    appearing mid-run doesn't shift anything)."""
    args = [sys.executable, "open_session.py"]
    tcode = (tcode or "").strip()
    if tcode:
        args.append(tcode)
    proc = subprocess.Popen(args, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW)
    return {"started": True, "pid": proc.pid}


def sap_open_order(order_no):
    """"오더 바로가기" - 새 SAP 세션을 열어 오더번호 앞자리로 VA02(6*)/VA03
    (그 외)을 자동판별해서 그 오더 화면으로 바로 들어간다(open_session.py의
    open_order()). "3. 오더번호 반영"(sap_run_order, order.py)과 달리 항목
    추출/workbench 반영/카카오 전송을 전혀 안 하는 순수 화면 바로가기라
    아무 때나(자동루프 중에도) 안전하게 눌러도 된다 - sap_open_session()과
    같은 이유로 세션 0~3을 건드리지 않음."""
    order_no = (order_no or "").strip()
    if not order_no:
        return {"started": False, "message": "오더번호를 입력하세요."}
    proc = subprocess.Popen(
        [sys.executable, "open_session.py", "order", order_no], cwd=str(BASE_DIR),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def _run_hidden_with_toast(args, label, cwd=None):
    """2026-08-25: replaces the CREATE_NEW_CONSOLE pattern several one-shot
    scripts used purely so their own print()'d success/failure text was
    visible immediately (morning_routine.py / portal_login.py / print_zpl_file.py
    - see each caller's docstring below for why). Runs windowless, captures
    stdout+stderr instead of letting them go to a console that no longer
    exists, and reports the outcome as a Windows toast once the process
    exits - a background watcher thread, same shape as run_pod_update()'s."""
    proc = subprocess.Popen(
        args, cwd=str(cwd or BASE_DIR),
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
        encoding="utf-8", errors="replace", creationflags=subprocess.CREATE_NO_WINDOW,
    )

    def watcher():
        output = proc.stdout.read() if proc.stdout else ""
        code = proc.wait()
        try:
            from win_notify import send_windows_notification
            if code == 0:
                send_windows_notification(f"{label} 완료", "", duration="short")
            else:
                lines = [ln.strip() for ln in output.strip().splitlines() if ln.strip()]
                tail = lines[-1] if lines else f"exit {code}"
                send_windows_notification(f"{label} 실패", tail, duration="long")
        except Exception:
            logger.warning(f"{label} 결과 토스트 실패(무시)", exc_info=True)

    threading.Thread(target=watcher, daemon=True).start()
    return proc.pid


def sap_run_morning_routine():
    """"0. 아침 루틴 강제 실행" - morning_routine.py를 지금 바로 수동으로
    돌린다. 매일 08:30 Windows 작업 스케줄러(SAP_Morning_Routine 작업)가
    자동으로 실행해주지만, 혹시 그게 안 돌았을 때를 대비한 수동 대체
    버튼(사용자 요청, 2026-08-14). morning_routine.py 자체가 이미 켜져
    있는 건 건너뛰는 로직이라 SAP 루프/workbench가 이미 떠 있어도 안전.
    2026-08-25: 예전엔 콘솔을 띄워서 뭐가 실행됐는지 눈으로 바로 확인했지만,
    이제 창 없이 돌리는 대신 끝나면 토스트로 결과를 알린다(morning_routine.py
    자체 로그는 morning_routine.log에 그대로 남음)."""
    pid = _run_hidden_with_toast([sys.executable, "morning_routine.py"], "아침 루틴")
    return {"started": True, "pid": pid}



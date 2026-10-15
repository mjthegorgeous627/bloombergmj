"""블룸버그 터미널 조작(티켓 확인)과 SAP 자동루프가 서로 화면/입력을 뺏지 않게 하는 공용 잠금 (2026-10-15).

문제: 티켓 확인은 마우스/키보드로 터미널을 앞으로 가져와 클릭·타이핑하고, SAP 루프(main.py)는 사이클마다 SAP 창을
앞으로 가져온다. 동시에 돌면 입력이 엉뚱한 창으로 가거나(2026-10-12 사고) 서로의 화면 읽기가 깨진다.

규칙:
- 티켓 확인이 시작되기 전: SAP 사이클이 'running'이면 끝날 때까지 기다린 뒤 잠금 파일(terminal_busy.json)을 만든다.
- SAP 루프는 사이클(run_once / 개별 세션)을 시작하기 전에 wait_until_free()로 잠금이 풀릴 때까지 기다린다.
- 잠금은 STALE_SECONDS 넘게 갱신(heartbeat)이 없으면 죽은 것으로 보고 무시한다(프로세스가 죽어도 루프가 영원히 막히지 않게).
"""
import json
import os
import time
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCK_FILE = os.path.join(BASE_DIR, "terminal_busy.json")
CYCLE_STATE_FILE = os.path.join(BASE_DIR, "cycle_state.json")
STALE_SECONDS = 120          # heartbeat가 이만큼 끊기면 잠금은 죽은 것
CYCLE_STALE_SECONDS = 30 * 60  # 사이클이 이만큼 'running'이면 멈춘 것으로 보고 기다리지 않음


def _read_lock():
    try:
        with open(LOCK_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


def lock_active():
    """살아 있는 터미널 잠금이 있으면 그 내용(dict), 없거나 죽었으면 None."""
    info = _read_lock()
    if not info:
        return None
    if time.time() - float(info.get("heartbeat", 0)) > STALE_SECONDS:
        return None
    return info


def wait_until_free(max_wait=20 * 60, poll=5, log=None):
    """SAP 루프 쪽: 터미널 잠금이 풀릴 때까지 대기. max_wait 넘기면 그냥 진행(루프가 영원히 막히지 않게)."""
    t0 = time.time()
    announced = False
    while True:
        info = lock_active()
        if not info:
            return True
        if not announced and log:
            log(f"터미널 티켓 확인 중({info.get('owner', '?')}) - SAP 사이클 대기")
            announced = True
        if time.time() - t0 > max_wait:
            if log:
                log("터미널 잠금 대기 초과 - 그대로 진행")
            return False
        time.sleep(poll)


def _cycle_running():
    try:
        with open(CYCLE_STATE_FILE, "r", encoding="utf-8") as fh:
            st = json.load(fh)
        if st.get("status") != "running":
            return False
        since = datetime.fromisoformat(st.get("since"))
        return (datetime.now() - since).total_seconds() < CYCLE_STALE_SECONDS
    except Exception:
        return False


class TerminalGate:
    """티켓 확인 쪽: with TerminalGate('티켓 확인') as g: ... g.beat() 를 주기적으로 호출.
    진입 시 SAP 사이클이 끝날 때까지 기다리고(최대 wait_cycle초), 다른 터미널 작업이 이미 있으면 RuntimeError."""

    def __init__(self, owner="티켓 확인", wait_cycle=15 * 60):
        self.owner = owner
        self.wait_cycle = wait_cycle

    def _write(self):
        tmp = LOCK_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"owner": self.owner, "pid": os.getpid(), "heartbeat": time.time(),
                       "since": datetime.now().isoformat()}, fh)
        os.replace(tmp, LOCK_FILE)

    def beat(self):
        self._write()

    def __enter__(self):
        other = lock_active()
        if other and other.get("pid") != os.getpid():
            raise RuntimeError(f"이미 다른 터미널 작업이 진행 중입니다({other.get('owner')}).")
        self._write()                      # 먼저 잠금을 걸어 새 사이클이 시작되지 못하게 한다
        t0 = time.time()
        while _cycle_running():            # 이미 도는 사이클은 끝날 때까지 기다림
            if time.time() - t0 > self.wait_cycle:
                self.__exit__(None, None, None)
                raise RuntimeError("SAP 사이클이 너무 오래 진행 중이라 터미널 작업을 시작하지 못했습니다.")
            self._write()
            time.sleep(3)
        time.sleep(3)                      # 루프가 잠금 확인 직전에 사이클을 막 시작한 경우 대비
        while _cycle_running():
            self._write()
            time.sleep(3)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            os.remove(LOCK_FILE)
        except OSError:
            pass
        return False

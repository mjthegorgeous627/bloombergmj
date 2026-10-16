"""workbench "티켓 확인" 버튼이 실행하는 러너 (2026-10-15).

흐름: 상태 'waiting' -> 터미널 잠금 획득(SAP 사이클이 끝날 때까지 대기, terminal_gate.py) -> 'running' ->
백그라운드 Claude(claude -p)가 CHECK_PROMPT.md 절차대로 수신함을 읽고 보드 반영 '제안'을 만든다 -> 'done'.
보드/DB는 수정하지 않는다(1단계). 결과: tickets/runs/latest_report.md, 상태: tickets/runs/status.json.

실행: python tickets/ticket_check_run.py   (workbench 서버가 숨김 프로세스로 띄움)
"""
import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
sys.path.insert(0, BASE)
from terminal_gate import TerminalGate  # noqa: E402
import ticket_store  # noqa: E402

RUNS = os.path.join(HERE, "runs")
STATUS = os.path.join(RUNS, "status.json")
REPORT = os.path.join(RUNS, "latest_report.md")
CLAUDE_EXE = os.path.join(os.environ.get("APPDATA", ""), "npm", "node_modules", "@anthropic-ai", "claude-code", "bin", "claude.exe")
IMG_DIR = r"C:\1\티켓관리"
MAX_SECONDS = 30 * 60
ALLOWED = ["Read", "Bash(python -X utf8 tickets/ticket_tools.py:*)"]


def set_status(state, message="", **extra):
    os.makedirs(RUNS, exist_ok=True)
    data = {"state": state, "message": message, "updated": datetime.now().isoformat(timespec="seconds")}
    data.update(extra)
    tmp = STATUS + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)
    os.replace(tmp, STATUS)


def current_status():
    try:
        with open(STATUS, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {"state": "idle", "message": ""}


def main():
    started = datetime.now().isoformat(timespec="seconds")
    forced = []                       # 사용자가 번호로 직접 지정한 티켓(처리했던 것이어도 다시 '처리할 것'으로 올린다)
    set_status("waiting", "SAP 자동루프 사이클이 끝나기를 기다리는 중", started=started)
    try:
        with TerminalGate("티켓 확인") as gate:
            set_status("running", "터미널에서 수신함/티켓을 읽는 중 - 이 동안 마우스·키보드를 건드리지 마세요", started=started)
            stop = threading.Event()

            def beat():
                while not stop.wait(20):
                    try:
                        gate.beat()
                    except Exception:
                        pass
            threading.Thread(target=beat, daemon=True).start()
            try:
                with open(os.path.join(HERE, "CHECK_PROMPT.md"), "r", encoding="utf-8") as fh:
                    prompt = fh.read()
                # 이미 저장소에 있는 티켓 표(마지막 글 번호/시각/처리 여부)를 넘겨서, 새 글이 없는 티켓은 열지 않게 한다
                try:
                    with open(os.path.join(HERE, "PLAYBOOK.md"), "r", encoding="utf-8") as fp_:
                        prompt += chr(10) * 2 + fp_.read()
                except Exception:
                    pass
                known = ticket_store.known_summary()
                NL0 = chr(10)
                prompt += (NL0 * 2 + "## 이미 아는 티켓 (저장소 기준)" + NL0
                           + (NL0.join(known) if known else "(없음)") + NL0
                           + "수신함 알림 줄의 시각이 위 표의 '마지막 글' 시각보다 새로울 때만 그 티켓을 연다. 같거나 오래됐으면 열지 않는다(이미 본 글)." + NL0)
                # 선택 기능: tickets/force.txt에 SDSK 번호(줄/쉼표 구분)가 있으면 읽음 여부와 무관하게 그 티켓들을 연다(1회용, 읽고 나면 삭제)
                fp = os.path.join(HERE, "force.txt")
                if os.path.exists(fp):
                    with open(fp, "r", encoding="utf-8") as ff:
                        nums = [x for x in ff.read().replace(",", " ").split() if x.isdigit()]
                    os.remove(fp)
                    if nums:
                        forced.extend(nums)
                        NL = chr(10)
                        prompt += (NL * 2 + "## 이번 실행의 추가 지시(1회)" + NL
                                   + "사용자가 아래 SDSK 티켓 번호를 직접 지정했다. 수신함은 보지 않고, '이미 아는 티켓' 표의 '열지 않는다' 규칙은 이번에는 무시한다. "
                                   "각 번호마다 `read <번호>`로 본문 전체(#1 글 포함)를 읽어 정리한다(번호로 직접 들어가며 `open`은 쓰지 않는다). "
                                   "이미 아는 티켓이어도 이번에는 #1 정리글과 흐름 이해에 필요한 글, 알아야 할 정보를 모두 posts에 넣고, 이전에 본 글은 new:false, 이번에 새로 알게 된 최신 글만 new:true로 한다. 티켓 번호: " + ", ".join(nums)
                                   + ". 일반 메일은 이번에는 열지 않는다." + NL)

                cmd = [CLAUDE_EXE, "-p", "--permission-mode", "dontAsk", "--output-format", "text",
                       "--add-dir", IMG_DIR, "--allowedTools"] + ALLOWED
                proc = subprocess.run(cmd, input=prompt, capture_output=True, text=True, encoding="utf-8",
                                      cwd=BASE, timeout=MAX_SECONDS, creationflags=subprocess.CREATE_NO_WINDOW)
                out = (proc.stdout or "").strip()
                if proc.returncode != 0 or not out:
                    raise RuntimeError(f"Claude 실행 실패(코드 {proc.returncode}): {(proc.stderr or out)[-500:]}")
            finally:
                stop.set()
        os.makedirs(RUNS, exist_ok=True)
        with open(REPORT, "w", encoding="utf-8") as fh:
            fh.write(out)
        with open(os.path.join(RUNS, f"report_{datetime.now():%Y%m%d_%H%M%S}.md"), "w", encoding="utf-8") as fh:
            fh.write(out)
        name = f"report_{datetime.now():%Y%m%d_%H%M%S}.md"
        try:
            stats = ticket_store.ingest(ticket_store.extract_json(out), name)
            ticket_store.force_pending(forced)
        except Exception as e:
            raise RuntimeError(f"결과를 저장소에 넣지 못했습니다(원문은 latest_report.md): {e}")
        set_status("done", f"완료 (새 티켓 {stats['new']} / 업데이트 {stats['updated']} / 변화 없음 {stats['unchanged']})",
                   started=started, finished=datetime.now().isoformat(timespec="seconds"), stats=stats)
    except subprocess.TimeoutExpired:
        set_status("error", f"{MAX_SECONDS // 60}분 안에 끝나지 않아 중단했습니다", started=started)
    except Exception as e:
        set_status("error", str(e)[:600], started=started)


if __name__ == "__main__":
    from terminal_gate import lock_active
    if lock_active():          # 이미 진행 중이면 상태 파일을 건드리지 않고 종료
        raise SystemExit("이미 터미널 작업(티켓 확인)이 진행 중입니다.")
    main()

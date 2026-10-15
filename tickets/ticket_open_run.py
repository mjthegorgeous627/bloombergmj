"""workbench 카드의 SDSK 번호를 눌렀을 때 터미널에서 그 티켓을 여는 작은 실행기 (2026-10-15).
사용: python -X utf8 tickets/ticket_open_run.py 1339712425
SAP 루프와 충돌하지 않게 TerminalGate로 터미널 잠금을 잡고, 'SDSK 번호'를 명령창에 입력해 티켓 화면으로 이동만 한다(읽기 전용)."""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, BASE)


def main():
    num = re.sub(r"\D", "", sys.argv[1] if len(sys.argv) > 1 else "")
    if not num:
        raise SystemExit("티켓 번호가 없습니다.")
    from terminal_gate import TerminalGate
    import term_ops as t
    with TerminalGate("티켓 열기"):
        t.goto(f"SDSK {num}", wait=6)


if __name__ == "__main__":
    from terminal_gate import lock_active
    if lock_active():
        raise SystemExit(0)
    main()

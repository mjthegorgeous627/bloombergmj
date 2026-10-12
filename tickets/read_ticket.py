"""SDSK 티켓 본문 전체를 캡처하는 읽기 전용 도구 (TICKET_MANUAL 5-5, 12-3).

사용: python read_ticket.py "SDSK 1331133687"
- 터미널(왼쪽 모니터 전체화면)에서 티켓을 열고, 맨 아래(최신 글)에서 위(1번 글)로 올라가며
  글 영역을 캡처해 C:\\1\\티켓관리\\cases\\<티켓번호>\\frame_01.png ... 로 저장한다.
- 클로드가 이 이미지들을 읽고 판단한다. 이 스크립트는 판단/쓰기를 하지 않는다.
- 수신함 줄은 클릭하지 않는다(읽음 처리 방지). 터미널 조작이므로 SAP 자동 루프와 동시에 돌리지 말 것.
"""
import os, re, sys, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import term_ops as t
from PIL import Image, ImageChops

CASES_DIR = r"C:\1\티켓관리\cases"
BOX = (180, 360, 1390, 980)     # Key Details 글 영역
WHEEL_XY = (770, 650)
STEP_NOTCHES = 20               # 실측: 20칸 ≈ 470px (영역 높이 620px 이하라 겹침 있음)
MAX_FRAMES = 80


def read_ticket(ticket):
    num = re.sub(r"\D", "", ticket)
    if not num:
        raise SystemExit("티켓번호가 없습니다.")
    out = os.path.join(CASES_DIR, f"SDSK_{num}")
    os.makedirs(out, exist_ok=True)
    for f in os.listdir(out):
        if f.startswith("frame_"):
            os.remove(os.path.join(out, f))
    t.goto(f"SDSK {num}", wait=8)
    t.shot(os.path.join(out, "header.png"))          # 요약/상태/Rep/Country 등 전체 화면
    def scroll_up():
        t.ensure_terminal()                           # 매 스크롤 전 포커스 재확인(다른 창이 앞이면 예외로 중단)
        for _ in range(STEP_NOTCHES // 5):
            t.wheel(*WHEEL_XY, 5); time.sleep(0.3)

    # 이미 열려 있던 티켓은 스크롤 위치를 기억하므로, 먼저 맨 아래(최신 글)로 내린 뒤 위로 올라간다
    t.ensure_terminal()
    probe, still = None, 0
    for _ in range(40):                               # 반응이 느려 고정 횟수로는 부족 -> 화면이 2번 연속 안 변할 때까지
        t.ensure_terminal()
        for _ in range(STEP_NOTCHES // 5):
            t.wheel(*WHEEL_XY, -5); time.sleep(0.2)
        time.sleep(1.5)
        cur = t.grab_box(BOX)
        if probe is not None and ImageChops.difference(probe, cur).getbbox() is None:
            still += 1
            if still >= 2:
                break
        else:
            still = 0
        probe = cur
    prev, saved, i = None, [], 0
    time.sleep(0.5)
    while i < MAX_FRAMES:
        path = os.path.join(out, f"frame_{i + 1:02d}.png")
        t.shot(path, BOX)
        img = Image.open(path).convert("RGB")
        if prev is not None and ImageChops.difference(prev, img).getbbox() is None:
            # 안 움직임: 터미널 반응이 늦을 수 있어(실측: 첫 휠이 무시됨) 최대 3번 더 시도 후에야 "맨 위"로 본다
            moved = False
            for _ in range(3):
                scroll_up(); time.sleep(2.0)
                t.shot(path, BOX)
                img = Image.open(path).convert("RGB")
                if ImageChops.difference(prev, img).getbbox() is not None:
                    moved = True
                    break
            if not moved:
                os.remove(path)                       # 끝까지 안 움직임 = 맨 위 도착
                break
        saved.append(path); prev = img; i += 1
        scroll_up(); time.sleep(1.2)
    return out, saved


if __name__ == "__main__":
    folder, frames = read_ticket(sys.argv[1] if len(sys.argv) > 1 else "")
    print(folder)
    print(f"{len(frames)} frames (frame_01 = 최신 글 쪽, 번호가 커질수록 위로 = 오래된 글)")

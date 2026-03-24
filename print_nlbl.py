"""
NiceLabel .nlbl 파일 출력 스크립트.
ZebraDesigner 3에서 파일 열고 Ctrl+P로 자동 출력.

사용법:
  python print_nlbl.py kb           -> KB Spare 1매 출력
  python print_nlbl.py kb 3         -> KB Spare 3매 출력
  python print_nlbl.py kb5          -> KB5 Spare 1매 출력
  python print_nlbl.py kb5 2        -> KB5 Spare 2매 출력
  python print_nlbl.py [파일경로] [매수]
"""
import sys
import subprocess
import time
import win32gui
import win32api
import win32con

ZEBRA_EXE = r"C:\Program Files\Zebra Technologies\ZebraDesigner 3\bin.net\ZebraDesigner.exe"

LABELS = {
    'kb':  r"C:\Users\bloomberg\Desktop\라벨\KB Spare.nlbl",
    'kb5': r"C:\Users\bloomberg\Desktop\라벨\KB5 Spare.nlbl",
}


def find_zebra_window():
    result = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and 'ZebraDesigner' in win32gui.GetWindowText(hwnd):
            result.append(hwnd)
    win32gui.EnumWindows(cb, None)
    return result[0] if result else None


def print_nlbl(filepath, copies=1):
    print(f"[Zebra] {filepath} x{copies}매 출력 중...")

    already_open = find_zebra_window() is not None

    # ZebraDesigner로 파일 열기 — 최소화 상태로 시작 (화면 안 가림)
    si = subprocess.STARTUPINFO()
    si.dwFlags = subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 7  # SW_SHOWMINNOACTIVE: 최소화, 포커스 뺏지 않음
    subprocess.Popen([ZEBRA_EXE, filepath], startupinfo=si)

    # 창 뜰 때까지 대기 (최대 20초)
    hwnd = None
    for _ in range(40):
        time.sleep(0.5)
        hwnd = find_zebra_window()
        if hwnd:
            break

    if not hwnd:
        print("[Zebra] ZebraDesigner 창을 찾을 수 없음")
        return False

    # 파일 로딩 대기 (새로 열었으면 더 기다림)
    time.sleep(3 if not already_open else 1.5)

    # 창 활성화
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.BringWindowToTop(hwnd)
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.3)

    # Ctrl+P를 copies만큼 반복 (다이얼로그 없이 바로 출력됨)
    for i in range(copies):
        win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        win32api.keybd_event(ord('P'), 0, 0, 0)
        win32api.keybd_event(ord('P'), 0, win32con.KEYEVENTF_KEYUP, 0)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.5)

    time.sleep(1.5)
    print(f"[Zebra] {copies}매 출력 완료")

    # 창 닫기 (새로 열었을 때만)
    if not already_open:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        time.sleep(0.5)
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    return True


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("사용법:")
        print("  python print_nlbl.py kb       -> KB Spare 출력")
        print("  python print_nlbl.py kb5      -> KB5 Spare 출력")
        print("  python print_nlbl.py [파일경로]")
        sys.exit(1)

    arg = sys.argv[1].lower()
    filepath = LABELS.get(arg, sys.argv[1])
    copies = int(sys.argv[2]) if len(sys.argv) >= 3 else 1
    print_nlbl(filepath, copies)

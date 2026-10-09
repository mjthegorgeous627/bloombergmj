"""Bloomberg 알림창 감지 방법 조사용 (읽기 전용, 아무것도 조작하지 않음).

bplus64(Bloomberg 터미널) 프로세스의 '보이는 최상위 창'을 1초마다 훑어서
새로 생기거나 사라진 창을 로그로 남긴다. 우상단 'Msg:S. Notification' 알림이
별도 창으로 뜨는지, 제목/클래스/위치가 뭔지 알아내는 게 목적.

사용: python probe_notify_windows.py [로그경로] [유지시간(초)]
"""
import sys, time, ctypes, datetime
import win32gui, win32process, win32con

LOG = sys.argv[1] if len(sys.argv) > 1 else r"C:\1\티켓관리\_probe\notify_windows.log"
DURATION = int(sys.argv[2]) if len(sys.argv) > 2 else 8 * 3600

import os
os.makedirs(os.path.dirname(LOG), exist_ok=True)


def proc_name(pid):
    k32 = ctypes.windll.kernel32
    h = k32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not h:
        return str(pid)
    try:
        buf = ctypes.create_unicode_buffer(520)
        n = ctypes.c_uint(520)
        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)):
            return os.path.basename(buf.value)
        return str(pid)
    finally:
        k32.CloseHandle(h)


def snapshot():
    snap = {}

    def cb(h, _):
        if not win32gui.IsWindowVisible(h):
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(h)
            name = proc_name(pid)
            if "bplus" not in name.lower() and "blp" not in name.lower():
                return
            ex = win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE)
            snap[h] = dict(
                title=win32gui.GetWindowText(h),
                cls=win32gui.GetClassName(h),
                rect=win32gui.GetWindowRect(h),
                topmost=bool(ex & win32con.WS_EX_TOPMOST),
                toolwin=bool(ex & win32con.WS_EX_TOOLWINDOW),
                proc=name,
            )
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    return snap


def log(msg):
    line = f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(line + "\n")


prev = snapshot()
log(f"START baseline={len(prev)} windows")
for h, w in prev.items():
    log(f"  BASE {h:#x} {w}")
end = time.time() + DURATION
while time.time() < end:
    time.sleep(1)
    cur = snapshot()
    for h in cur.keys() - prev.keys():
        log(f"NEW  {h:#x} {cur[h]}")
    for h in prev.keys() - cur.keys():
        log(f"GONE {h:#x} {prev[h]['title']!r}")
    for h in cur.keys() & prev.keys():
        if cur[h]["title"] != prev[h]["title"]:
            log(f"TITLE {h:#x} {prev[h]['title']!r} -> {cur[h]['title']!r}")
    prev = cur
log("END")

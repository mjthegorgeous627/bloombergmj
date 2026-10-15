"""Bloomberg 터미널 조작 도우미 (마우스/키보드/화면캡처). 좌표는 터미널이 왼쪽 모니터(1920x1080) 전체화면일 때 기준.

주의(2026-10-09 실측):
- SendInput 유니코드 방식 타이핑은 터미널이 무시함 -> vk_type()(가상키+스캔코드) 사용
- 명령창 클릭 위치 (360,85), AI Summary 버튼 (70,321)
- 메시지 목록 줄을 클릭하면 노랑(안읽음)이 회색으로 바뀜 -> 목록 영역 클릭 금지, 티켓은 goto('SDSK 번호')로 직접 이동
"""
import ctypes, time, sys, win32api, win32con, win32clipboard
from PIL import ImageGrab
ctypes.windll.user32.SetProcessDPIAware()
U = ctypes.windll.user32

def click(x, y):
    win32api.SetCursorPos((x, y)); time.sleep(0.25)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0); time.sleep(0.06)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0); time.sleep(0.25)

class KI(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort), ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
class IN(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [("ki", KI)]
    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.c_ulong), ("u", _U)]

def _send(vk=0, ch=None, up=False):
    if ch is not None:
        f = 0x4 | (0x2 if up else 0)
        i = IN(type=1, ki=KI(0, ord(ch), f, 0, None))
    else:
        i = IN(type=1, ki=KI(vk, 0, 0x2 if up else 0, 0, None))
    U.SendInput(1, ctypes.byref(i), ctypes.sizeof(i))

def type_text(s):
    for c in s:
        _send(ch=c); _send(ch=c, up=True); time.sleep(0.03)

def key(vk):
    _send(vk); _send(vk, up=True); time.sleep(0.1)

def ctrl(vk):
    _send(0x11); _send(vk); _send(vk, up=True); _send(0x11, up=True); time.sleep(0.3)

def get_clip():
    win32clipboard.OpenClipboard()
    try:
        return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
    except Exception:
        return None
    finally:
        win32clipboard.CloseClipboard()

def set_clip(s):
    win32clipboard.OpenClipboard(); win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, s); win32clipboard.CloseClipboard()

def shot(path, box=(0, 0, 1920, 1080)):
    ImageGrab.grab(all_screens=True).crop(box).save(path)

def goto(cmd):
    click(360, 85)          # 명령창
    ctrl(0x41)              # Ctrl+A (기존 글자 선택)
    type_text(cmd); key(0x0D)

def wheel(x, y, notches):
    win32api.SetCursorPos((x, y)); time.sleep(0.2)
    win32api.mouse_event(0x0800, 0, 0, int(120 * notches), 0); time.sleep(0.5)

def enter():
    win32api.keybd_event(0x0D, 0x1c, 0, 0); time.sleep(0.05); win32api.keybd_event(0x0D, 0x1c, 2, 0)

def vk_type(s, delay=0.003):
    """가상키 타이핑. 붙여넣기(Ctrl+V/Shift+Insert)는 터미널 명령창에서 안 되므로 지연을 최소화해 사실상 한 번에 입력
    (2026-10-15 실측: 지연 3ms로 'SDSK 1339513698' 15글자 누락 없이 입력, 약 1초)."""
    for c in s:
        v = win32api.VkKeyScan(c); vk = v & 0xff; sc = U.MapVirtualKeyW(vk, 0); sh = (v >> 8) & 1
        if sh: win32api.keybd_event(0x10, 0, 0, 0)
        win32api.keybd_event(vk, sc, 0, 0); win32api.keybd_event(vk, sc, 2, 0)
        if sh: win32api.keybd_event(0x10, 0, 2, 0)
        time.sleep(delay)

def _exe_name(hwnd):
    import os, win32process, win32api
    try:
        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
        ph = win32api.OpenProcess(0x1000, False, pid)
        buf = ctypes.create_unicode_buffer(520); sz = ctypes.c_ulong(520)
        ctypes.windll.kernel32.QueryFullProcessImageNameW(int(ph), 0, buf, ctypes.byref(sz))
        return os.path.basename(buf.value).lower()
    except Exception:
        return "?"

def ensure_terminal():
    """Bloomberg 터미널 창을 복원해 맨 앞으로 가져오고, 실제로 앞에 왔는지 확인한다. 실패하면 예외.
    (2026-10-12 사고: 터미널이 최소화되고 SAP가 앞에 있는 채로 click/타이핑이 실행돼 입력이 SAP로 갈 뻔함.
    모든 터미널 조작은 이 함수가 성공한 뒤에만 한다.)"""
    import win32gui
    cands = []
    def f(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == "wdmm-Win32Window" and _exe_name(h) == "bplus64.exe":
            title = win32gui.GetWindowText(h)
            if title and title != "wdmm-Win32Window" and not title.startswith(("IB -", "Launchpad")):
                cands.append(h)
    win32gui.EnumWindows(f, None)
    if not cands:
        raise RuntimeError("Bloomberg 터미널 창을 찾지 못했습니다.")
    # 터미널 본체는 가장 큰 창(작은 보조 창과 구분). 최소화(-32000)된 창은 크기를 못 재므로 마지막 순위.
    def area(h):
        l, t_, r, b = win32gui.GetWindowRect(h)
        return (r - l) * (b - t_) if l > -30000 else 0
    hwnd = max(cands, key=area)
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE); time.sleep(0.8)
    # SetForegroundWindow 제한 우회: Alt 키를 한 번 눌러 준다 (focus_guard.py와 같은 방식)
    # 사용자가 다른 창(메신저 등)을 쓰는 중이면 잠깐 뺏길 수 있어 몇 번 재시도한다. 끝내 못 가져오면 입력 없이 중단.
    for _ in range(4):
        if win32gui.GetForegroundWindow() == hwnd:
            break
        win32api.keybd_event(0x12, 0, 0, 0); win32api.keybd_event(0x12, 0, 2, 0)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(1.0)
    if win32gui.GetForegroundWindow() != hwnd:
        raise RuntimeError("터미널을 맨 앞으로 가져오지 못했습니다(다른 창이 앞에 있음). 입력하지 않고 중단합니다.")
    l, t_, r, b = win32gui.GetWindowRect(hwnd)
    if not (r - l >= 1800 and b - t_ >= 1000 and -50 <= l <= 50):
        raise RuntimeError(f"터미널 창 위치/크기가 예상(왼쪽 모니터 전체화면)과 다릅니다: {(l, t_, r, b)}")
    return hwnd

def ensure_english_ime(hwnd):
    """입력기가 한글 모드면 영문으로 바꾼다. (2026-10-13 사고: 한글 모드에서 vk_type('SDSK 번호')가
    한글 자모로 입력돼 터미널 검색창에 '간포생명보험…' 같은 자동완성이 뜸 -> 티켓 대신 엉뚱한 화면.)
    IME 변환 모드의 NATIVE 비트(1)가 켜져 있으면 한글. 한/영 키(VK_HANGUL 0x15)로 토글 후 재확인."""
    try:
        ime = ctypes.windll.imm32.ImmGetDefaultIMEWnd(hwnd)
        if not ime:
            return
        IMC_GETCONVERSIONMODE = 0x0001
        def native():
            return bool(U.SendMessageW(ime, 0x0283, IMC_GETCONVERSIONMODE, 0) & 1)   # WM_IME_CONTROL
        # 2026-10-13 실측: 한/영 키(VK_HANGUL) 토글은 변환모드 9->1로만 바뀌고 영문이 안 됨(작업표시줄 '가' 그대로).
        # IMC_SETCONVERSIONMODE=0 으로 직접 설정하면 즉시 영문('A')이 된다.
        for _ in range(2):
            if not native():
                return
            U.SendMessageW(ime, 0x0283, 0x0002, 0)       # IMC_SETCONVERSIONMODE = 0
            time.sleep(0.4)
        if native():
            raise RuntimeError("입력기를 영문으로 바꾸지 못했습니다. 입력하지 않고 중단합니다.")
    except RuntimeError:
        raise
    except Exception:
        pass

INBOX_ROW_X = 700          # 수신함 줄 클릭 x (제목 글자 위)
MAIL_SDSK_LINK = (278, 275)  # 메일 상단 '블룸버그 기능 (SDSK)' 파란 링크 (2026-10-15 3건 실측, 위치 동일)

def open_ticket_from_inbox(row_y, wait=8):
    """수신함의 한 줄(y 좌표)을 클릭해 메일을 열고(읽음 처리됨) 상단 SDSK 링크를 클릭해 티켓으로 들어간다.
    사용자가 읽음 표시로 처리 여부를 눈으로 확인할 수 있게 하기 위한 방식(2026-10-15 사용자 요청). 먼저 goto('MSG') 상태여야 한다.
    반환: (메일 화면 캡처, 티켓 화면 캡처) PIL 이미지. 메일 본문에 최신 글이 그대로 들어 있어 메일 캡처만으로도 최신 내용을 읽을 수 있다."""
    ensure_english_ime(ensure_terminal())
    click(INBOX_ROW_X, row_y); time.sleep(4)
    mail = grab_box((0, 0, 1920, 1080))
    ensure_terminal(); click(*MAIL_SDSK_LINK); time.sleep(wait)
    return mail, grab_box((0, 0, 1920, 1080))

def grab_box(box):
    return ImageGrab.grab(all_screens=True).crop(box).convert("RGB")

def goto(cmd, wait=5):
    ensure_english_ime(ensure_terminal())
    click(360, 85); time.sleep(0.3)
    ctrl(0x41); time.sleep(0.2)          # 명령창에 남은 글자와 섞이지 않게 전체 선택 후 덮어쓰기 (2026-10-13 메뉴 화면으로 빠진 사례)
    vk_type(cmd); time.sleep(0.6); enter(); time.sleep(wait)

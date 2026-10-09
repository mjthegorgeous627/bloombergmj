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

def vk_type(s):
    for c in s:
        v = win32api.VkKeyScan(c); vk = v & 0xff; sc = U.MapVirtualKeyW(vk, 0); sh = (v >> 8) & 1
        if sh: win32api.keybd_event(0x10, 0, 0, 0)
        win32api.keybd_event(vk, sc, 0, 0); time.sleep(0.02); win32api.keybd_event(vk, sc, 2, 0)
        if sh: win32api.keybd_event(0x10, 0, 2, 0)
        time.sleep(0.04)

def goto(cmd, wait=5):
    click(360, 85); time.sleep(0.3)
    vk_type(cmd); time.sleep(0.6); enter(); time.sleep(wait)

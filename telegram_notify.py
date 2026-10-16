"""텔레그램 폰 알림 - 카톡 "나에게 보내기"는 폰 푸시가 안 떠서(win_notify.py 참고) 추가한 보조 채널.

telegram_config.json({"bot_token": ..., "chat_id": ...})이 없으면 조용히 아무것도 안 한다.
알림 실패가 SAP 자동화/아침 루틴을 절대 막으면 안 되므로 모든 예외는 삼킨다.

2026-10-16 사용자 규칙:
- 알림은 한국 영업일(주말/공휴일 제외) 08:30 ~ 18:00 에만 보낸다. 그 밖의 일반 알림(오류/상태)은 보내지 않는다.
- 새 오더는 버리지 않는다. 보낼 수 없는 시간에 생긴 오더는 telegram_queue.json 대기열에 쌓아 두고,
  다음 영업일 아침 요약(morning_routine.py)이 나간 뒤 쌓인 순서대로 한 건씩 보낸다
  ("마지막에 보낸 오더 이후부터 순서대로").
- 아침 요약이 나가기 전에는(그날 첫 사이클 결과를 요약에 세어야 하므로) 오더도 대기열에 모은다.
  단 아침 루틴이 안 돌았을 때 오더가 영영 안 가는 일이 없도록 MORNING_FALLBACK 이후에는 그냥 보낸다.
"""
import json
import logging
import os
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime, time as dtime

logger = logging.getLogger(__name__)

_BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_BASE, "telegram_config.json")
STATE_FILE = os.path.join(_BASE, "telegram_queue.json")   # 오더 내용이 들어 있어 .gitignore 대상

WINDOW_START = dtime(8, 30)
WINDOW_END = dtime(18, 0)
MORNING_FALLBACK = dtime(9, 30)   # 아침 요약이 이 시각까지도 안 나가면 대기열을 그냥 풀어 보낸다

_proc_lock = threading.Lock()


def _now():
    return datetime.now()


# ── 시간대 / 영업일 ──────────────────────────────────────────────────────────

def _is_business_day(d):
    try:
        from holiday_check import skip_reason
        return skip_reason(d) is None
    except Exception:
        return d.weekday() < 5   # 공휴일 라이브러리 문제 시 주말만이라도 거른다


def in_window(now=None):
    now = now or _now()
    return _is_business_day(now.date()) and WINDOW_START <= now.time() < WINDOW_END


# ── 실제 전송 ────────────────────────────────────────────────────────────────

def _send_now(text):
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
        data = urllib.parse.urlencode({"chat_id": cfg["chat_id"], "text": str(text)[:4000]}).encode("utf-8")
        urllib.request.urlopen(
            f"https://api.telegram.org/bot{cfg['bot_token']}/sendMessage", data=data, timeout=10
        ).read()
        return True
    except FileNotFoundError:
        return False
    except Exception as e:
        logger.warning(f"[텔레그램] 전송 실패(무시): {type(e).__name__}")  # URL에 토큰이 있어 예외 상세는 남기지 않는다
        return False


def send_telegram(text, wait=False):
    """일반 알림(오류/상태). 영업시간(08:30~18:00) 밖이면 보내지 않는다.
    wait=False: 백그라운드 스레드로 보내 SAP 루프를 안 막는다. wait=True: 끝날 때까지 기다린다."""
    if not text or not os.path.exists(CONFIG_FILE) or not in_window():
        return False
    if wait:
        return _send_now(text)
    threading.Thread(target=_send_now, args=(text,), daemon=False).start()
    return True


# ── 오더 대기열 (프로세스 간 공유: SAP 루프 / 아침 루틴 / workbench) ─────────

class _FileLock:
    """다른 프로세스와 대기열 파일을 동시에 고치지 않도록 하는 간단한 잠금 파일."""
    path = STATE_FILE + ".lock"

    def __enter__(self):
        deadline = time.time() + 10
        while True:
            try:
                self.fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                return self
            except FileExistsError:
                try:   # 죽은 프로세스가 남긴 잠금은 30초 지나면 치운다
                    if time.time() - os.path.getmtime(self.path) > 30:
                        os.remove(self.path)
                        continue
                except OSError:
                    pass
                if time.time() > deadline:
                    raise TimeoutError("telegram_queue 잠금 대기 초과")
                time.sleep(0.05)

    def __exit__(self, *exc):
        try:
            os.close(self.fd)
            os.remove(self.path)
        except OSError:
            pass


def _load():
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as fh:
            st = json.load(fh)
    except Exception:
        st = {}
    st.setdefault("pending", [])
    st.setdefault("morning_done_date", "")
    return st


def _save(st):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False)
    os.replace(tmp, STATE_FILE)


def order_action(order_data_list):
    """카톡 제목과 같은 기준: 배송+회수가 같이 있으면 교체."""
    prefixes = {i.get("order_prefix", "") for i in order_data_list}
    if "배송" in prefixes and "회수" in prefixes:
        return "교체"
    return "배송" if "배송" in prefixes else "회수"


def _is_ready(st, now):
    """지금 오더를 바로 보내도 되는 상태인가 (영업시간 + 아침 요약 완료 또는 대기 한도 경과)."""
    if not in_window(now):
        return False
    return st["morning_done_date"] == now.date().isoformat() or now.time() >= MORNING_FALLBACK


def _format_entry(e, now):
    ts = datetime.fromisoformat(e["ts"])
    if ts.date() == now.date() and ts.time() >= WINDOW_START:
        return e["text"]
    return f"🕐 {ts.strftime('%m/%d %H:%M')} 수집\n{e['text']}"   # 시간 밖/밤사이에 모인 오더


def _flush(now=None):
    """대기열을 받은 순서대로 한 건씩 보낸다. 실패하면 그 건을 맨 앞에 되돌리고 멈춘다(다음 기회에 이어서)."""
    sent = 0
    while True:
        now = _now()
        with _FileLock():
            st = _load()
            if not st["pending"]:
                break
            entry = st["pending"].pop(0)
            _save(st)
        if _send_now(_format_entry(entry, now)):
            sent += 1
            time.sleep(0.3)
        else:
            with _FileLock():
                st = _load()
                st["pending"].insert(0, entry)
                _save(st)
            break
    return sent


def _send_order_impl(text, action):
    now = _now()
    with _proc_lock:
        with _FileLock():
            st = _load()
            st["pending"].append({"ts": now.isoformat(timespec="seconds"), "action": action, "text": text})
            _save(st)
            ready = _is_ready(st, now)
        if ready:
            _flush()   # 방금 넣은 건 포함, 쌓여 있던 순서 그대로


def send_order(text, action):
    """새 오더 1건. 보낼 수 있는 시간이면 바로, 아니면 대기열에 보관(백그라운드)."""
    if not text or not os.path.exists(CONFIG_FILE):
        return False
    threading.Thread(target=lambda: _safe(_send_order_impl, text, action), daemon=False).start()
    return True


def _safe(fn, *a):
    try:
        return fn(*a)
    except Exception as e:
        logger.warning(f"[텔레그램] 대기열 처리 실패(무시): {type(e).__name__}: {e}")


# ── 아침 요약 (morning_routine.py에서 호출) ──────────────────────────────────

def send_morning_summary(status_lines, ok):
    """아침 요약 1건을 보낸 뒤, 쌓인 오더를 순서대로 모두 보내고 '오늘 아침 완료'로 표시한다.
    영업시간이 아니면(예: 휴일에 수동 실행) 아무것도 안 한다."""
    if not os.path.exists(CONFIG_FILE) or not in_window():
        return False
    with _FileLock():
        pend = list(_load()["pending"])
    counts = {"배송": 0, "교체": 0, "회수": 0}
    for e in pend:
        counts[e.get("action", "")] = counts.get(e.get("action", ""), 0) + 1
    head = "☀️ 아침 루틴 정상" if ok else "⚠️ 아침 루틴 확인 필요"
    orders = (f"🆕 새 오더 {len(pend)}건 (배송 {counts['배송']} · 교체 {counts['교체']} · 회수 {counts['회수']})"
              if pend else "🆕 새 오더 없음")
    link = "📱 모바일 workbench: Chrome 원격 데스크톱으로 PC에 접속해서 확인"
    _send_now("\n".join([head] + list(status_lines) + [orders, link]))
    with _proc_lock:
        _flush()
        with _FileLock():
            st = _load()
            st["morning_done_date"] = _now().date().isoformat()
            _save(st)
        _flush()   # 위 표시 직전에 막 들어온 건이 있으면 마저 보낸다
    return True

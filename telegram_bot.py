"""텔레그램 봇 리모컨 (2026-10-16 사용자 요청: 원격 데스크톱이 불편해서 폰에서 직접 쓰는 화면이 필요).

- 폰 텔레그램 대화창이 화면이다: 메뉴/날짜 목록/오더 카드 + 인라인 버튼. S/N·메모는 "카드의 버튼을 누른 뒤 글을 보내서" 입력한다.
- PC가 텔레그램 서버에 주기적으로 묻는 방식(long polling, 밖으로 나가는 요청만) - PC에 들어오는 연결/추가 설치가 없다.
- 동작은 PC 화면과 같은 서버 주소(/api/items/status, /api/items/edit_field …)를 그대로 호출한다 -> 같은 DB, 같은 규칙.
- 보안: telegram_config.json 의 chat_id(=사용자 본인) 에서 온 메시지/버튼만 처리한다. 다른 사람은 완전히 무시.
  재시작 때 쌓여 있던 옛 버튼 누름은 실행하지 않는다(예전 카드의 Done이 뒤늦게 눌리는 사고 방지).
- workbench_app.py 가 시작할 때 백그라운드 스레드로 돈다(단일 인스턴스 뮤텍스가 중복 실행을 막는다).
"""
import html
import json
import logging
import os
import threading
import time
import urllib.error
import urllib.request

logger = logging.getLogger(__name__)

_BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(_BASE, "telegram_config.json")
LOCAL = "http://127.0.0.1:8765"
PAGE_SIZE = 10

STATUS_EMOJI = {"done": "🔵", "pending": "🔴", "cancelled": "⛔", "billed": "🟢", "": "⚪"}


def esc(s):
    return html.escape(str(s if s is not None else ""))


def action_of(item_type):
    t = str(item_type or "")
    if "회수" in t:
        return "회수"
    if t == "배송" or t.startswith("Delayed"):
        return "배송"
    return t or "배송"


def group_rows(rows):
    """화면(groupRows)과 같은 기준: 같은 날짜 안에서 orderLabel이 같은 행끼리 한 카드."""
    groups, index = [], {}
    for r in rows:
        key = (r.get("orderLabel") or "").strip()
        if key in index:
            groups[index[key]]["rows"].append(r)
        else:
            index[key] = len(groups)
            groups.append({"key": key, "rows": [r]})
    return groups


def group_status(rows):
    st = [r.get("boardStatus") or "" for r in rows]
    if any(s == "pending" for s in st):
        return "pending"
    if all(s == "cancelled" for s in st):
        return "cancelled"
    if all(s in ("done", "billed") for s in st):
        return "billed" if any(s == "billed" for s in st) else "done"
    return ""


def group_has_error(rows):
    return any(r.get("portalError") or r.get("podError") or r.get("snReloError") for r in rows)


def summary_of(rows):
    """카드 요약줄 - 카톡 문구와 같은 제품명 규칙(kakao_handler._get_product_name)."""
    first = rows[0]
    if not first.get("orderNo"):   # 메모 행
        t = " / ".join(r.get("description") or "" for r in rows if r.get("description"))
        return t or first.get("memo") or ""
    try:
        from kakao_handler import _get_product_name
    except Exception:
        def _get_product_name(d, m=""):
            return (str(d).split() or [""])[0]
    cnt = {}
    for r in rows:
        k = (_get_product_name(r.get("description"), r.get("material")), action_of(r.get("itemType")))
        cnt[k] = cnt.get(k, 0) + 1
    return "[" + " / ".join(f"{p} {n}개 {a}" for (p, a), n in cnt.items()) + "]"


def kb(rows):
    return {"inline_keyboard": rows}


def btn(text, data):
    return {"text": text, "callback_data": data}


class Bot:
    def __init__(self):
        with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
        self.token = cfg["bot_token"]
        self.chat_id = int(cfg["chat_id"])
        self.awaiting = None   # {"kind": "serial"|"memo", "itemId": int, "iso": str}

    # ── 외부 호출 (테스트에서 덮어쓴다) ─────────────────────────────────────
    def api(self, method, **params):
        params = {k: v for k, v in params.items() if v is not None}
        req = urllib.request.Request(
            f"https://api.telegram.org/bot{self.token}/{method}",
            data=json.dumps(params, ensure_ascii=False).encode("utf-8"),
            headers={"Content-Type": "application/json"},
        )
        timeout = (params.get("timeout") or 0) + 15
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            body = ""
            try:
                body = e.read().decode("utf-8", "replace")
            except Exception:
                pass
            return {"ok": False, "error_code": e.code, "description": body[:200]}

    def local(self, method, path, body=None):
        data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else (b"{}" if method == "POST" else None)
        req = urllib.request.Request(LOCAL + path, data=data, method=method, headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.load(r)

    # ── 화면 만들기 ────────────────────────────────────────────────────────
    def board(self):
        return self.local("GET", "/api/board_data")

    def find_group(self, item_id):
        for d in self.board():
            for g in group_rows(d["rows"]):
                if any(r["itemId"] == item_id for r in g["rows"]):
                    return d, g
        return None, None

    def menu_view(self):
        text = "📋 <b>Bloomberg Workbench</b>\n무엇을 할까요?"
        return text, kb([
            [btn("📅 오늘", "d:today:0"), btn("📅 날짜 선택", "ds")],
            [btn("⚠️ 오류 오더", "er"), btn("📊 상태", "st")],
            [btn("🔄 SAP 전체조회", "sap:all"), btn("▶ Start Loop", "sap:start")],
            [btn("■ 루프 정지", "sap:stop")],
        ])

    def dates_view(self):
        board = self.board()
        rows = []
        line = []
        for d in board:
            n = len(group_rows(d["rows"]))
            line.append(btn(f"{d['iso'][5:].replace('-', '/')} ({n})", f"d:{d['iso']}:0"))
            if len(line) == 3:
                rows.append(line); line = []
        if line:
            rows.append(line)
        rows.append([btn("🏠 메뉴", "m")])
        return "📅 <b>날짜를 고르세요</b>", kb(rows)

    def list_view(self, iso, page):
        board = self.board()
        if iso == "today":
            from datetime import date
            iso = date.today().isoformat()
        day = next((d for d in board if d["iso"] == iso), None)
        if not day:
            return f"📅 {esc(iso)}\n표시할 오더가 없습니다.", kb([[btn("📅 날짜 선택", "ds"), btn("🏠 메뉴", "m")]])
        groups = group_rows(day["rows"])
        pages = max(1, (len(groups) + PAGE_SIZE - 1) // PAGE_SIZE)
        page = max(0, min(page, pages - 1))
        chunk = groups[page * PAGE_SIZE:(page + 1) * PAGE_SIZE]
        lines = [f"📅 <b>{esc(day.get('label') or iso)}</b> · {len(groups)}건" + (f" (페이지 {page + 1}/{pages})" if pages > 1 else "")]
        nums = []
        for i, g in enumerate(chunk, 1):
            rs = g["rows"]
            em = STATUS_EMOJI.get(group_status(rs), "⚪") + ("⚠️" if group_has_error(rs) else "")
            cust = rs[0].get("customer") or ""
            lines.append(f"{i}. {em} <b>{esc(g['key'] or '(번호없음)')}</b>\n    {esc(summary_of(rs))}" + (f" · {esc(cust)}" if cust else ""))
            nums.append(btn(str(i), f"c:{rs[0]['itemId']}"))
        keyboard = [nums[j:j + 5] for j in range(0, len(nums), 5)]
        nav = []
        if page > 0:
            nav.append(btn("◀ 이전", f"d:{iso}:{page - 1}"))
        if page < pages - 1:
            nav.append(btn("다음 ▶", f"d:{iso}:{page + 1}"))
        keyboard.append(nav + [btn("🔄", f"d:{iso}:{page}"), btn("📅 날짜", "ds"), btn("🏠", "m")])
        return "\n".join(lines), kb(keyboard)

    def errors_view(self):
        board = self.board()
        lines = ["⚠️ <b>오류 오더</b>"]
        keyboard, n = [], 0
        for d in board:
            for g in group_rows(d["rows"]):
                if group_has_error(g["rows"]):
                    n += 1
                    rs = g["rows"]
                    lines.append(f"{n}. {esc(d['iso'][5:])} <b>{esc(g['key'])}</b> {esc(summary_of(rs))}")
                    keyboard.append([btn(f"{n}. {g['key']}"[:40], f"c:{rs[0]['itemId']}")])
        if not n:
            lines.append("오류가 있는 오더가 없습니다. 👍")
        keyboard.append([btn("🏠 메뉴", "m")])
        return "\n".join(lines), kb(keyboard[:25])

    def card_view(self, item_id, note=""):
        d, g = self.find_group(item_id)
        if not g:
            return "⚠️ 항목을 찾을 수 없습니다 (이미 마감/삭제됐을 수 있습니다).", kb([[btn("🏠 메뉴", "m")]])
        rs = g["rows"]; f = rs[0]
        st = group_status(rs)
        lines = []
        if note:
            lines.append(note)
        lines.append(f"{STATUS_EMOJI.get(st, '⚪')} <b>{esc(g['key'] or '(번호없음)')}</b>  <i>{esc(d['iso'][5:])}</i>")
        lines.append(f"<b>{esc(summary_of(rs))}</b>")
        for label, key in (("고객", "customer"), ("전화", "phone"), ("주소", "address"), ("코드", "deliveryNo")):
            if f.get(key):
                lines.append(f"{label}: {esc(str(f[key]).replace(chr(10), ' / '))}")
        if f.get("memo"):
            lines.append(f"📝 {esc(f['memo'])}")
        lines.append("")
        keyboard = [[btn("✅ Done", f"s:{item_id}:D"), btn("🔴 Ready", f"s:{item_id}:R")],
                    [btn("⛔ Cancel", f"s:{item_id}:C"), btn("⚪ 초기화", f"s:{item_id}:X")]]
        for i, r in enumerate(rs, 1):
            sn = r.get("serial") or ""
            mark = {"done": "🔵", "pending": "🔴", "cancelled": "⛔", "billed": "🟢"}.get(r.get("boardStatus") or "", "⚪")
            lines.append(f"{mark} {i}) {esc(r.get('itemType') or '')} · {esc(r.get('description') or '(품목)')} · M/N {esc(r.get('material') or '-')}")
            if f.get("orderNo"):
                lines.append(f"    S/N: {esc(sn) if sn else '❗ 없음'}")
            if len(keyboard) < 6:
                keyboard.append([btn(f"✏️ S/N {i}" if f.get("orderNo") else f"✏️ 내용 {i}", f"sn:{r['itemId']}")] )
        for key, label in (("portalErrorMsg", "Portal"), ("podErrorMsg", "POD"), ("snReloErrorMsg", "S/N")):
            msgs = {r.get(key) for r in rs if r.get(key)}
            for m in msgs:
                lines.append(f"⚠️ {label} 오류: {esc(m)}")
        # S/N 버튼들을 한 줄에 모은다
        sn_btns = [b[0] for b in keyboard[2:]]
        keyboard = keyboard[:2] + [sn_btns[i:i + 3] for i in range(0, len(sn_btns), 3)]
        keyboard.append([btn("📝 메모 수정", f"mm:{item_id}")])
        keyboard.append([btn("🔄", f"c:{item_id}"), btn(f"← {d['iso'][5:].replace('-', '/')} 목록", f"d:{d['iso']}:0"), btn("🏠", "m")])
        return "\n".join(lines), kb(keyboard)

    def status_view(self):
        try:
            s = self.local("GET", "/api/sap/status")
            loop = "🟢 실행 중" if s.get("loopRunning") else "🔴 중지됨"
        except Exception:
            loop = "❓ 확인 실패"
        board = self.board()
        from datetime import date
        today = next((d for d in board if d["iso"] == date.today().isoformat()), None)
        n_today = len(group_rows(today["rows"])) if today else 0
        n_err = sum(1 for d in board for g in group_rows(d["rows"]) if group_has_error(g["rows"]))
        return (f"📊 <b>상태</b>\nSAP 루프: {loop}\n오늘 오더: {n_today}건\n오류 오더: {n_err}건",
                kb([[btn("🔄", "st"), btn("🏠 메뉴", "m")]]))

    # ── 동작 ───────────────────────────────────────────────────────────────
    def set_status(self, item_id, code):
        d, g = self.find_group(item_id)
        if not g:
            return None, "항목을 찾을 수 없습니다"
        rs = g["rows"]
        if code == "D":
            groups = {"done": [r["itemId"] for r in rs]}; label = "Done(파랑)"
        elif code == "R":   # 화면 Ready와 같은 규칙: 회수는 대기(빨강), 나머지는 완료
            groups = {"pending": [r["itemId"] for r in rs if r.get("itemType") == "회수"],
                      "done": [r["itemId"] for r in rs if r.get("itemType") != "회수"]}; label = "Ready(회수 대기 · 나머지 완료)"
        elif code == "C":
            groups = {"cancelled": [r["itemId"] for r in rs]}; label = "Cancel(회색)"
        else:
            groups = {"": [r["itemId"] for r in rs]}; label = "초기화(흰색)"
        for status, ids in groups.items():
            if ids:
                res = self.local("POST", "/api/items/status", {"itemIds": ids, "status": status})
                if res.get("error"):
                    return None, res["error"]
        return label, None

    def send(self, text, markup=None):
        return self.api("sendMessage", chat_id=self.chat_id, text=text, parse_mode="HTML", reply_markup=markup, disable_web_page_preview=True)

    def edit(self, msg_id, text, markup=None):
        r = self.api("editMessageText", chat_id=self.chat_id, message_id=msg_id, text=text, parse_mode="HTML", reply_markup=markup, disable_web_page_preview=True)
        if not r.get("ok") and "not modified" not in str(r.get("description", "")):
            self.send(text, markup)   # 수정이 안 되면(너무 오래된 메시지 등) 새로 보낸다
        return r

    def show(self, msg_id, view):
        text, markup = view
        if msg_id:
            self.edit(msg_id, text, markup)
        else:
            self.send(text, markup)

    def on_callback(self, cq):
        data = cq.get("data") or ""
        msg_id = (cq.get("message") or {}).get("message_id")
        answer = ""
        p = data.split(":")
        try:
            if data == "m":
                self.awaiting = None
                self.show(msg_id, self.menu_view())
            elif p[0] == "d":
                self.show(msg_id, self.list_view(p[1], int(p[2]) if len(p) > 2 else 0))
            elif data == "ds":
                self.show(msg_id, self.dates_view())
            elif data == "er":
                self.show(msg_id, self.errors_view())
            elif data == "st":
                self.show(msg_id, self.status_view())
            elif p[0] == "c":
                self.awaiting = None
                self.show(msg_id, self.card_view(int(p[1])))
            elif p[0] == "s":
                label, err = self.set_status(int(p[1]), p[2])
                if err:
                    answer = "실패: " + err
                    self.show(msg_id, self.card_view(int(p[1]), f"❌ 실패: {esc(err)}"))
                else:
                    answer = label + " 처리됨"
                    self.show(msg_id, self.card_view(int(p[1]), f"✅ {esc(label)} 처리됨"))
            elif p[0] in ("sn", "mm"):
                kind = "serial" if p[0] == "sn" else "memo"
                item_id = int(p[1])
                d, g = self.find_group(item_id)
                if not g:
                    answer = "항목을 찾을 수 없습니다"
                else:
                    self.awaiting = {"kind": kind, "itemId": item_id, "msgId": msg_id}
                    cur = next((r for r in g["rows"] if r["itemId"] == item_id), g["rows"][0])
                    what = "S/N" if kind == "serial" else "메모"
                    now = (cur.get("serial") if kind == "serial" else g["rows"][0].get("memo")) or "(비어 있음)"
                    self.send(f"✏️ {what}을(를) 이 메시지에 <b>답장</b>하듯 보내 주세요.\n현재: <code>{esc(now)}</code>\n지우려면 <code>-</code> 한 글자, 취소는 /cancel",
                              {"force_reply": True, "input_field_placeholder": what + " 입력"})
            elif p[0] == "sap":
                names = {"all": "SAP 전체조회", "start": "SAP 루프 시작", "stop": "SAP 루프 정지"}
                self.show(msg_id, (f"❓ <b>{names[p[1]]}</b>을(를) 실행할까요?",
                                   kb([[btn("✅ 실행", f"sapy:{p[1]}"), btn("취소", "m")]])))
            elif p[0] == "sapy":
                path = {"all": "/api/sap/run_all", "start": "/api/sap/start_loop", "stop": "/api/sap/stop_loop"}[p[1]]
                res = self.local("POST", path)
                msg = res.get("message") or ("시작됨" if res.get("started") else str(res))
                answer = msg[:150]
                self.show(msg_id, (f"✅ {esc(msg)}", kb([[btn("📊 상태", "st"), btn("🏠 메뉴", "m")]])))
        except Exception as e:
            logger.warning(f"[텔레그램 봇] 버튼 처리 실패({data}): {type(e).__name__}: {e}")
            answer = "오류: " + type(e).__name__
        self.api("answerCallbackQuery", callback_query_id=cq["id"], text=answer[:190] if answer else None)

    def on_message(self, m):
        text = (m.get("text") or "").strip()
        if text in ("/start", "/menu", "메뉴"):
            self.awaiting = None
            self.send(*self.menu_view())
        elif text in ("/today", "오늘"):
            self.send(*self.list_view("today", 0))
        elif text == "/errors":
            self.send(*self.errors_view())
        elif text == "/status":
            self.send(*self.status_view())
        elif text == "/cancel":
            self.awaiting = None
            self.send("취소했습니다.", kb([[btn("🏠 메뉴", "m")]]))
        elif self.awaiting:
            aw, self.awaiting = self.awaiting, None
            value = "" if text == "-" else text
            try:
                if aw["kind"] == "serial":
                    res = self.local("POST", "/api/items/edit_field", {"itemId": aw["itemId"], "field": "serial", "value": value})
                    what = "S/N"
                else:
                    d, g = self.find_group(aw["itemId"])
                    ids = [r["itemId"] for r in g["rows"]] if g else [aw["itemId"]]
                    res = self.local("POST", "/api/orders/edit_field", {"itemIds": ids, "field": "memo", "value": value})
                    what = "메모"
                if res.get("error"):
                    raise RuntimeError(res["error"])
                self.send(*self.card_view(aw["itemId"], f"✅ {what} 저장됨" + ("(비움)" if not value else "")))
            except Exception as e:
                self.send(f"❌ 저장 실패: {esc(e)}", kb([[btn("🏠 메뉴", "m")]]))
        else:
            self.send("메뉴에서 선택하세요.", self.menu_view()[1])

    def handle_update(self, u):
        cq = u.get("callback_query")
        msg = u.get("message")
        sender = ((cq or msg or {}).get("from") or {}).get("id")
        chat = ((cq or {}).get("message") or msg or {}).get("chat", {}).get("id")
        if sender != self.chat_id or chat != self.chat_id:   # 사용자 본인 대화방만 - 나머지는 아무 응답 없이 무시
            logger.warning("[텔레그램 봇] 허용되지 않은 발신자 무시")
            return
        if cq:
            self.on_callback(cq)
        elif msg:
            self.on_message(msg)

    def run_forever(self):
        offset = None
        # 재시작 전에 쌓인 업데이트(옛 버튼 누름 등)는 실행하지 않고 건너뛴다
        try:
            r = self.api("getUpdates", offset=-1, timeout=0)
            if r.get("ok") and r.get("result"):
                offset = r["result"][-1]["update_id"] + 1
        except Exception:
            pass
        logger.info("[텔레그램 봇] 시작")
        fails = 0
        while True:
            try:
                r = self.api("getUpdates", offset=offset, timeout=25, allowed_updates=["message", "callback_query"])
                if not r.get("ok"):
                    if r.get("error_code") == 409:
                        logger.warning("[텔레그램 봇] 다른 곳에서 같은 봇을 사용 중(409) - 30초 후 재시도")
                        time.sleep(30)
                    else:
                        time.sleep(min(60, 5 * (fails + 1))); fails += 1
                    continue
                fails = 0
                for u in r.get("result", []):
                    offset = u["update_id"] + 1
                    try:
                        self.handle_update(u)
                    except Exception:
                        logger.warning("[텔레그램 봇] 업데이트 처리 실패", exc_info=True)
            except Exception as e:
                fails += 1
                logger.warning(f"[텔레그램 봇] 폴링 실패({type(e).__name__}) - 재시도")   # 예외 상세엔 토큰 주소가 들어갈 수 있어 남기지 않는다
                time.sleep(min(60, 5 * fails))


def start_background():
    """workbench_app.py 가 서버 시작 때 호출. 설정 파일이 없으면 아무것도 안 한다."""
    if not os.path.exists(CONFIG_FILE):
        return False
    bot = Bot()
    threading.Thread(target=bot.run_forever, daemon=True, name="telegram-bot").start()
    return True

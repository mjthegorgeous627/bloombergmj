"""티켓 저장소 (2026-10-15): 티켓 하나당 기록 한 줄. 카드(처리할 것/보관함)는 이 기록으로 그린다.

- 새 글이 올라오면 이미 저장된 글은 그대로 두고(= Summary), 저장에 없던 글만 '안 본 글'(= NEW)이 된다.
- 처리([확인]/[예, 업데이트]/[건너뛰기])하면 안 본 글이 모두 본 글이 되고 티켓은 보관함(state=handled)으로 간다.
- 6개월(183일) 동안 활동 없는 티켓은 저장할 때 자동 삭제한다.
- workbench.db와 분리된 파일(tickets/runs/ticket_store.db)이라 배송장 데이터에 영향이 없다.
"""
import json
import os
import re
import sqlite3
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNS = os.path.join(BASE_DIR, "tickets", "runs")
DB = os.path.join(RUNS, "ticket_store.db")
KEEP_DAYS = 183


def _now():
    return datetime.now().isoformat(timespec="seconds")


def _con():
    os.makedirs(RUNS, exist_ok=True)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.execute("""CREATE TABLE IF NOT EXISTS tickets(
        key TEXT PRIMARY KEY, kind TEXT, title TEXT, srep TEXT, closed INTEGER DEFAULT 0,
        posts TEXT, todo TEXT, board TEXT, state TEXT, history TEXT,
        first_seen TEXT, updated_at TEXT, last_report TEXT)""")
    con.execute("DROP TABLE IF EXISTS watch")           # 95건 목록 방식은 폐기(확인 필요 = 카드로 남기는 방식)
    cols = [r[1] for r in con.execute("PRAGMA table_info(tickets)")]
    if "last_checked" not in cols:
        con.execute("ALTER TABLE tickets ADD COLUMN last_checked TEXT")
    if "due" not in cols:
        con.execute("ALTER TABLE tickets ADD COLUMN due TEXT")           # 내가 처리/확인해야 하는 날(보관함 정렬 기준)
    return con


def _j(v, default):
    try:
        return json.loads(v) if v else default
    except Exception:
        return default


def _row(r):
    d = dict(r)
    d["posts"] = _j(d.get("posts"), [])
    d["board"] = _j(d.get("board"), None)
    d["history"] = _j(d.get("history"), [])
    d["due_auto"] = todo_date(d.get("todo"))
    d["due_eff"] = d.get("due") or d["due_auto"] or ""
    d["tag"] = d["kind"] if d["kind"] in ("Mail", "List") else ("Closed" if d.get("closed") else ("New" if not d["history"] else "Updated"))
    return d


def todo_date(todo):
    """To do 문장 속 첫 날짜(2026-10-31 / 10/31 / 10월 31일)를 YYYY-MM-DD로. 없으면 None. 연도 없으면 올해(6개월 넘게 지난 날짜면 내년)."""
    t = str(todo or "")
    m = re.search(r"(?<!\d)(20\d\d)-(\d{1,2})-(\d{1,2})(?!\d)", t)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.search(r"(?<![\d/])(\d{1,2})\s*[/월]\s*(\d{1,2})\s*일?(?![\d:/])", t)
        if not m:
            return None
        mo, d = int(m.group(1)), int(m.group(2))
        y = datetime.now().year
    try:
        dt = datetime(y, mo, d)
    except ValueError:
        return None
    if not m or not re.search(r"20\d\d-", m.group(0)):
        if dt < datetime.now() - timedelta(days=180):
            dt = datetime(y + 1, mo, d)
    return dt.strftime("%Y-%m-%d")


def _post_key(p):
    n = p.get("n")
    return ("n", int(n)) if str(n or "").isdigit() else ("t", (p.get("who") or "") + "|" + (p.get("text") or "")[:60])


def _clean_key(item):
    k = str(item.get("key") or "").strip()
    if item.get("kind") == "Mail":
        return "mail:" + re.sub(r"\s+", " ", k.lower())[:120]
    return re.sub(r"\D", "", k) or k


def extract_json(text):
    """Claude 출력에서 JSON 객체를 꺼낸다(```json 블록 또는 첫 { ~ 마지막 })."""
    m = re.search(r"```json\s*(\{.*\})\s*```", text, re.S)
    raw = m.group(1) if m else text[text.find("{"): text.rfind("}") + 1]
    return json.loads(raw)


def ingest(data, report_name=""):
    """Claude가 낸 {"tickets":[...]}를 저장소에 합친다. 반환: {"new":n, "updated":n, "unchanged":n}"""
    stats = {"new": 0, "updated": 0, "unchanged": 0}
    with _con() as con:
        for it in data.get("tickets") or []:
            key = _clean_key(it)
            if not key:
                continue
            row = con.execute("SELECT * FROM tickets WHERE key=?", (key,)).fetchone()
            old = _row(row) if row else None
            posts = list(old["posts"]) if old else []
            have = {_post_key(p) for p in posts}
            added = 0
            for src in it.get("posts") or []:
                p = {"n": src.get("n"), "who": src.get("who") or "", "text": src.get("text") or "", "when": src.get("when") or ""}
                k = _post_key(p)
                if k in have:
                    continue                       # 이미 저장된 글(본 글이든 안 본 글이든)은 그대로 둔다
                # Claude가 new:false로 표시한 글은 '배경 글'(처음 보는 티켓의 앞선 흐름 등)이라 본 글로 저장한다
                p["seen"] = (src.get("new") is False)
                posts.append(p)
                have.add(k)
                if not p["seen"]:
                    added += 1
            posts.sort(key=lambda p: (0, int(p["n"])) if str(p.get("n") or "").isdigit() else (1, 0))
            board = it.get("board") if isinstance(it.get("board"), dict) else None
            if old is not None and not added and old["state"] in ("handled", "watching"):
                stats["unchanged"] += 1            # 새 글이 없으면 처리한/확인 필요 티켓은 내용은 그대로 둔다
                if old["state"] == "watching":
                    if it.get("closed") and not old.get("closed"):       # 확인하던 티켓이 해결됨 -> 사용자가 보게 처리할 것으로
                        con.execute("UPDATE tickets SET closed=1, state='pending', todo=?, updated_at=? WHERE key=?",
                                    (it.get("todo") or old["todo"], _now(), key))
                    else:                                                 # 다시 읽었으니 마지막 확인일만 갱신
                        con.execute("UPDATE tickets SET last_checked=? WHERE key=?", (_now(), key))
                continue
            state = "pending" if (old is None or added) else old["state"]
            stats["new" if old is None else "updated"] += 1
            con.execute("""INSERT INTO tickets(key,kind,title,srep,closed,posts,todo,board,state,history,first_seen,updated_at,last_report,last_checked)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(key) DO UPDATE SET kind=excluded.kind,title=excluded.title,srep=excluded.srep,closed=excluded.closed,
                posts=excluded.posts,todo=excluded.todo,board=excluded.board,state=excluded.state,updated_at=excluded.updated_at,last_report=excluded.last_report,last_checked=excluded.last_checked""",
                (key, it.get("kind") or ("Mail" if key.startswith("mail:") else "SDSK"), it.get("title") or (old or {}).get("title") or key,
                 it.get("srep") or (old or {}).get("srep") or "", 1 if it.get("closed") else 0,
                 json.dumps(posts, ensure_ascii=False), it.get("todo") or "", json.dumps(board, ensure_ascii=False) if board else None,
                 state, json.dumps((old or {}).get("history") or [], ensure_ascii=False),
                 (old or {}).get("first_seen") or _now(), _now(), report_name, _now()))
        for it in data.get("tickets") or []:                     # Claude가 준 due(할 일 날짜)는 사용자가 아직 안 정했을 때만 채운다
            d0 = str(it.get("due") or "")
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d0):
                con.execute("UPDATE tickets SET due=? WHERE key=? AND (due IS NULL OR due='')", (d0, _clean_key(it)))
        con.execute("DELETE FROM tickets WHERE updated_at < ?", ((datetime.now() - timedelta(days=KEEP_DAYS)).isoformat(timespec="seconds"),))
    return stats


def pending():
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE state='pending' ORDER BY updated_at DESC").fetchall()   # 처리할 것: 새로 읽은/업데이트된 티켓이 맨 위(열자마자 확인)
    return [_row(r) for r in rows]


def archive(q="", limit=200):
    like = f"%{q}%"
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE state='handled' AND (key LIKE ? OR title LIKE ? OR posts LIKE ?) "
                           "ORDER BY updated_at DESC LIMIT ?", (like, like, like, int(limit))).fetchall()
        total = con.execute("SELECT COUNT(*) FROM tickets WHERE state='handled'").fetchone()[0]
    return [_row(r) for r in rows], total


def handle(key, result, to=None, due=None):
    """처리 버튼([확인]/[예, 업데이트]/[건너뛰기]): 안 본 글을 본 글로 바꾸고,
    to='watching'(보관함 탭: 해결될 때까지 다시 확인) / 'done'(완료함 탭). to가 없으면 열린 티켓은 보관함, 닫힌 티켓/메일은 완료함. 마지막 확인일을 찍는다."""
    with _con() as con:
        r = con.execute("SELECT * FROM tickets WHERE key=?", (key,)).fetchone()
        if not r:
            raise ValueError("티켓 기록을 찾을 수 없습니다.")
        d = _row(r)
        for p in d["posts"]:
            p["seen"] = True
        d["history"].append({"result": (result or "확인")[:20], "at": _now()})
        state = {"watching": "watching", "done": "handled"}.get(to) or ("handled" if (d.get("closed") or d["kind"] == "Mail") else "watching")
        con.execute("UPDATE tickets SET posts=?, state=?, history=?, updated_at=?, last_checked=? WHERE key=?",
                    (json.dumps(d["posts"], ensure_ascii=False), state, json.dumps(d["history"], ensure_ascii=False), _now(), _now(), key))
        if due is not None:
            con.execute("UPDATE tickets SET due=? WHERE key=?", (due if re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(due)) else None, key))
    return {"handled": True, "state": state}


def watching():
    """확인 필요: 해결되기 전까지 중간중간 확인해야 하는 열린 티켓. 오래 확인 안 한 순서."""
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE state='watching' ORDER BY last_checked, updated_at").fetchall()
    return sorted([_row(r) for r in rows], key=lambda d: (d["due_eff"] == "", d["due_eff"]))


def set_due(key, due):
    """보관함/처리할 것 카드의 '할 일 날짜' 저장(빈 값이면 날짜 없음)."""
    due = str(due or "")
    if due and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", due):
        raise ValueError("날짜 형식이 올바르지 않습니다.")
    with _con() as con:
        con.execute("UPDATE tickets SET due=? WHERE key=?", (due or None, key))
    return {"due": due or None}


def recheck(key, result="다시 확인"):
    """[확인함]: 업데이트가 없어도 다시 확인했다는 날짜 도장."""
    with _con() as con:
        r = con.execute("SELECT * FROM tickets WHERE key=?", (key,)).fetchone()
        if not r:
            raise ValueError("티켓 기록을 찾을 수 없습니다.")
        h = _row(r)["history"]; h.append({"result": (result or "다시 확인")[:20], "at": _now()})
        con.execute("UPDATE tickets SET last_checked=?, history=? WHERE key=?", (_now(), json.dumps(h, ensure_ascii=False), key))
    return {"at": _now()}


def resolve(key):
    """[해결됨]: 확인 필요에서 빼서 보관함으로."""
    with _con() as con:
        r = con.execute("SELECT * FROM tickets WHERE key=?", (key,)).fetchone()
        if not r:
            raise ValueError("티켓 기록을 찾을 수 없습니다.")
        h = _row(r)["history"]; h.append({"result": "Done", "at": _now()})
        con.execute("UPDATE tickets SET state='handled', history=?, updated_at=? WHERE key=?", (json.dumps(h, ensure_ascii=False), _now(), key))
    return {"resolved": True}


def known_summary(limit=60):
    """러너가 Claude에게 넘기는 '이미 아는 티켓' 표: 키, 제목, 마지막 글 번호/시각, 상태."""
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE kind!='Mail' ORDER BY updated_at DESC LIMIT ?", (limit,)).fetchall()
    out = []
    for r in rows:
        d = _row(r)
        nums = [int(p["n"]) for p in d["posts"] if str(p.get("n") or "").isdigit()]
        last = max(d["posts"], key=lambda p: int(p["n"]) if str(p.get("n") or "").isdigit() else -1, default=None)
        out.append(f"SDSK {d['key']} | 마지막 글 #{max(nums) if nums else '?'} ({(last or {}).get('when', '')}) | { {'handled': '보관(해결)', 'watching': '확인 필요(내가 확인 중)'}.get(d['state'], '처리 안 함') } | {d['title'][:50]}")
    return out


def force_pending(keys):
    """사용자가 번호로 직접 지정한 티켓: 이미 처리/보관한 것이어도 다시 '처리할 것'으로 올려서 내용을 보게 한다."""
    n = 0
    with _con() as con:
        for k in keys or []:
            k = re.sub(r"\D", "", str(k))
            r = con.execute("UPDATE tickets SET state='pending', updated_at=? WHERE key=? AND state!='pending'", (_now(), k))
            n += r.rowcount
    return n

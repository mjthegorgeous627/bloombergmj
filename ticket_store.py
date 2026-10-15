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
    d["tag"] = "Mail" if d["kind"] == "Mail" else ("Closed" if d.get("closed") else ("New" if not d["history"] else "Updated"))
    return d


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
            if old is not None and not added and old["state"] == "handled":
                stats["unchanged"] += 1            # 새 글이 없으면 처리한 티켓은 건드리지 않는다
                continue
            state = "pending" if (old is None or added) else old["state"]
            stats["new" if old is None else "updated"] += 1
            con.execute("""INSERT INTO tickets(key,kind,title,srep,closed,posts,todo,board,state,history,first_seen,updated_at,last_report)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(key) DO UPDATE SET kind=excluded.kind,title=excluded.title,srep=excluded.srep,closed=excluded.closed,
                posts=excluded.posts,todo=excluded.todo,board=excluded.board,state=excluded.state,updated_at=excluded.updated_at,last_report=excluded.last_report""",
                (key, it.get("kind") or ("Mail" if key.startswith("mail:") else "SDSK"), it.get("title") or (old or {}).get("title") or key,
                 it.get("srep") or (old or {}).get("srep") or "", 1 if it.get("closed") else 0,
                 json.dumps(posts, ensure_ascii=False), it.get("todo") or "", json.dumps(board, ensure_ascii=False) if board else None,
                 state, json.dumps((old or {}).get("history") or [], ensure_ascii=False),
                 (old or {}).get("first_seen") or _now(), _now(), report_name))
        con.execute("DELETE FROM tickets WHERE updated_at < ?", ((datetime.now() - timedelta(days=KEEP_DAYS)).isoformat(timespec="seconds"),))
    return stats


def pending():
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE state='pending' ORDER BY updated_at DESC").fetchall()
    return [_row(r) for r in rows]


def archive(q="", limit=200):
    like = f"%{q}%"
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE state='handled' AND (key LIKE ? OR title LIKE ? OR posts LIKE ?) "
                           "ORDER BY updated_at DESC LIMIT ?", (like, like, like, int(limit))).fetchall()
        total = con.execute("SELECT COUNT(*) FROM tickets WHERE state='handled'").fetchone()[0]
    return [_row(r) for r in rows], total


def handle(key, result):
    """처리: 안 본 글을 모두 본 글로, 보관함으로, 이력에 한 줄 추가."""
    with _con() as con:
        r = con.execute("SELECT * FROM tickets WHERE key=?", (key,)).fetchone()
        if not r:
            raise ValueError("티켓 기록을 찾을 수 없습니다.")
        d = _row(r)
        for p in d["posts"]:
            p["seen"] = True
        d["history"].append({"result": (result or "확인")[:20], "at": _now()})
        con.execute("UPDATE tickets SET posts=?, state='handled', history=?, board=NULL, updated_at=? WHERE key=?",
                    (json.dumps(d["posts"], ensure_ascii=False), json.dumps(d["history"], ensure_ascii=False), _now(), key))
    return {"handled": True}


def known_summary(limit=60):
    """러너가 Claude에게 넘기는 '이미 아는 티켓' 표: 키, 제목, 마지막 글 번호/시각, 상태."""
    with _con() as con:
        rows = con.execute("SELECT * FROM tickets WHERE kind!='Mail' ORDER BY updated_at DESC LIMIT ?", (limit,)).fetchall()
    out = []
    for r in rows:
        d = _row(r)
        nums = [int(p["n"]) for p in d["posts"] if str(p.get("n") or "").isdigit()]
        last = max(d["posts"], key=lambda p: int(p["n"]) if str(p.get("n") or "").isdigit() else -1, default=None)
        out.append(f"SDSK {d['key']} | 마지막 글 #{max(nums) if nums else '?'} ({(last or {}).get('when', '')}) | {'처리함' if d['state'] == 'handled' else '처리 안 함'} | {d['title'][:50]}")
    return out

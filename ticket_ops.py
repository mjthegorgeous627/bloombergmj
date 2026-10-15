"""workbench "티켓 확인" 버튼용 서버 쪽 함수 (2026-10-15). 실제 작업은 tickets/ticket_check_run.py가 숨김 프로세스로 한다.
카드 데이터는 ticket_store.py(티켓당 한 줄)에 있고, 여기서는 상태/처리/적용/열기를 맡는다."""
import json
import os
import subprocess
import sys

import ticket_store

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNS = os.path.join(BASE_DIR, "tickets", "runs")


def ticket_check_start():
    from terminal_gate import lock_active
    info = lock_active()
    if info:
        return {"started": False, "message": f"이미 터미널 작업이 진행 중입니다({info.get('owner')})."}
    exe = sys.executable
    # pythonw로 떠 있는 서버에서도 콘솔 창이 뜨지 않게
    proc = subprocess.Popen([exe, "-X", "utf8", os.path.join("tickets", "ticket_check_run.py")], cwd=BASE_DIR,
                            creationflags=subprocess.CREATE_NO_WINDOW,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
    return {"started": True, "pid": proc.pid}


def ticket_check_status():
    try:
        with open(os.path.join(RUNS, "status.json"), "r", encoding="utf-8") as fh:
            st = json.load(fh)
    except Exception:
        st = {"state": "idle", "message": ""}
    # 러너가 죽어 상태가 running/waiting으로 남은 경우: 터미널 잠금이 없으면 멈춘 것으로 표시
    if st.get("state") in ("waiting", "running"):
        from terminal_gate import lock_active
        if not lock_active():
            st = {"state": "error", "message": "티켓 확인이 중간에 멈췄습니다(프로세스 종료).", "updated": st.get("updated")}
    return st


def ticket_pending():
    return {"cards": ticket_store.pending()}


def ticket_archive_list(q=""):
    items, total = ticket_store.archive(q)
    return {"items": items, "total": total}


def ticket_handle(key, result):
    return ticket_store.handle(str(key or ""), result)


def ticket_apply(p):
    """제안 1건을 보드에 반영한다(사용자가 [적용]을 눌렀을 때만). 기존 items.py 함수를 그대로 써서 Ctrl+Z가 된다."""
    import items
    action = (p.get("action") or "").strip()
    if action == "append_memo":
        oid = int(p.get("row_id") or 0)
        add = (p.get("memo_append") or "").strip()
        if not oid or not add:
            raise ValueError("row_id 또는 memo_append가 없습니다.")
        with items.connect() as con:
            o = con.execute("SELECT memo FROM orders WHERE id=?", (oid,)).fetchone()
            it = con.execute("SELECT id FROM order_items WHERE order_id=? AND deleted_at IS NULL ORDER BY id LIMIT 1", (oid,)).fetchone()
        if not o or not it:
            raise ValueError(f"행 {oid}을(를) 보드에서 찾을 수 없습니다.")
        old = (o["memo"] or "").rstrip()
        if add in old:
            return {"applied": False, "message": "이미 같은 줄이 Memo에 있습니다."}
        items.update_order_field([it["id"]], "memo", (old + chr(10) if old else "") + add)
        return {"applied": True, "message": f"행 {oid} Memo에 한 줄 추가했습니다."}
    if action == "set_date":
        oid = int(p.get("row_id") or 0)
        with items.connect() as con:
            ids = [r["id"] for r in con.execute("SELECT id FROM order_items WHERE order_id=? AND deleted_at IS NULL", (oid,)).fetchall()]
        if not ids:
            raise ValueError(f"행 {oid}을(를) 보드에서 찾을 수 없습니다.")
        items.move_items_date(ids, p.get("date") or "")
        return {"applied": True, "message": f"행 {oid} 날짜를 {p.get('date')}(으)로 옮겼습니다."}
    if action == "new_row":
        n = p.get("new_row") or {}
        label = (n.get("order_no") or "").strip() or ("SDSK " + str(p.get("ticket") or ""))
        title = (n.get("item") or "").strip() or label
        cells = [title, n.get("mn") or "", n.get("sn") or "", n.get("customer") or "", n.get("phone") or "",
                 n.get("address") or "", n.get("memo") or ""]
        r = items.add_manual_note(cells, n.get("date") or None, n.get("type") or "배송", label, 7)
        return {"applied": True, "message": f"새 행을 만들었습니다 ({label})."}
    raise ValueError("적용할 수 없는 제안입니다.")


def ticket_open(number):
    """카드의 SDSK 번호 클릭 -> 터미널에서 그 티켓 화면을 연다(숨김 프로세스, SAP 루프와 충돌하지 않게 잠금 사용)."""
    import re
    from terminal_gate import lock_active
    num = re.sub(r"\D", "", str(number or ""))
    if len(num) < 6:
        raise ValueError("티켓 번호가 올바르지 않습니다.")
    info = lock_active()
    if info:
        return {"started": False, "message": f"터미널이 사용 중입니다({info.get('owner')}). 끝난 뒤 다시 눌러 주세요."}
    subprocess.Popen([sys.executable, "-X", "utf8", os.path.join("tickets", "ticket_open_run.py"), num], cwd=BASE_DIR,
                     creationflags=subprocess.CREATE_NO_WINDOW,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL)
    return {"started": True, "message": f"터미널에서 SDSK {num} 티켓을 엽니다."}

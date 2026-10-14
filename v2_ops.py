"""workbench ver.2(/v2) 전용 서버 기능 - 전부 "추가"만 한다(기존 v1 API/컬럼은 그대로).

- 빠른 메모(side_memos): 오른쪽 열의 임시 메모. 고정 메모(notepad_pages)와는 별개.
- 오더번호 수정(rename_order_no): v1은 서버가 일부러 막아 둔 기능이라, v2 전용 경로에서만
  중복/빈값 방어 규칙을 걸고 허용한다(실행취소 가능).
- OBD/티켓 바로가기: 새 프로세스로 띄우는 순수 화면 이동(수집/저장 없음).
"""
from __future__ import annotations

import json
import os
import subprocess
import sys

from db import BASE_DIR, _log_action_batch, add_event, connect, ensure_db, now_text


# ---------------- 빠른 메모 ----------------
def list_quick_notes():
    ensure_db()
    with connect() as con:
        rows = con.execute(
            "SELECT id, body, pinned, done, created_at FROM side_memos ORDER BY id DESC"
        ).fetchall()
    return [dict(r) for r in rows]


def add_quick_note(body):
    body = (body or "").strip()
    if not body:
        return {"added": False}
    ensure_db()
    with connect() as con:
        cur = con.execute(
            "INSERT INTO side_memos(body, pinned, done, created_at) VALUES (?, 0, 0, ?)",
            (body, now_text()),
        )
    return {"added": True, "id": cur.lastrowid}


def update_quick_note(note_id, fields):
    ensure_db()
    sets, vals = [], []
    for k in ("body", "pinned", "done"):
        if k in fields:
            sets.append(f"{k}=?")
            vals.append(fields[k] if k == "body" else int(bool(fields[k])))
    if not sets:
        return {"updated": False}
    with connect() as con:
        con.execute(f"UPDATE side_memos SET {', '.join(sets)} WHERE id=?", [*vals, int(note_id)])
    return {"updated": True}


def delete_quick_note(note_id):
    ensure_db()
    with connect() as con:
        con.execute("DELETE FROM side_memos WHERE id=?", (int(note_id),))
    return {"deleted": True}


# ---------------- 오더번호 수정 ----------------
def rename_order_no(item_ids, new_no):
    """화면에서 오더번호 숫자 부분을 고친 것을 저장. 거부 조건:
    - 빈 값 / 숫자 이외 글자(오더번호는 SAP 키라 숫자만)
    - 같은 (order_no, delivery_no) 조합이 이미 다른 오더에 있음(UNIQUE 충돌 + 자동화가 엉뚱한 행을 집게 됨)
    undo(edit_order) 지원 - 기존 실행취소(Ctrl+Z)로 되돌아간다."""
    new_no = str(new_no or "").strip()
    if not new_no.isdigit():
        return {"updated": 0, "error": "오더번호는 숫자만 입력할 수 있습니다."}
    item_ids = [int(i) for i in (item_ids or []) if str(i).strip()]
    if not item_ids:
        return {"updated": 0, "error": "대상 행이 없습니다."}
    ensure_db()
    with connect() as con:
        ph = ",".join("?" for _ in item_ids)
        order_ids = sorted({
            r["order_id"] for r in con.execute(
                f"SELECT DISTINCT order_id FROM order_items WHERE id IN ({ph})", item_ids
            ).fetchall()
        })
        if len(order_ids) != 1:
            return {"updated": 0, "error": "병합된 여러 오더는 한 번에 번호를 바꿀 수 없습니다. 병합을 먼저 해제하세요."}
        oid = order_ids[0]
        cur = con.execute("SELECT id, order_no, delivery_no FROM orders WHERE id=?", (oid,)).fetchone()
        if not cur:
            return {"updated": 0, "error": "오더를 찾지 못했습니다."}
        if (cur["order_no"] or "") == new_no:
            return {"updated": 0}
        dup = con.execute(
            "SELECT id FROM orders WHERE order_no=? AND COALESCE(delivery_no,'')=? AND id<>?",
            (new_no, cur["delivery_no"] or "", oid),
        ).fetchone()
        if dup:
            return {"updated": 0, "error": f"이미 같은 오더번호({new_no})가 보드에 있습니다."}
        con.execute("UPDATE orders SET order_no=? WHERE id=?", (new_no, oid))
        _log_action_batch(
            con, "edit_order", "오더 order_no 수정",
            [(oid, json.dumps({"field": "order_no", "value": cur["order_no"]}))],
        )
        add_event(con, oid, f"order_no 수정 (workbench v2): {cur['order_no']} -> {new_no}")
    return {"updated": 1}


# ---------------- 바로가기 ----------------
def open_obd(delivery_no):
    """OBD(배송문서)를 새 SAP 세션 VL02N으로 연다(open_session.py obd). 세션 0~3은 건드리지 않음."""
    delivery_no = "".join(ch for ch in str(delivery_no or "") if ch.isdigit())
    if not delivery_no:
        return {"started": False, "message": "OBD 번호가 없습니다."}
    proc = subprocess.Popen(
        [sys.executable, "open_session.py", "obd", delivery_no], cwd=str(BASE_DIR),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def open_ticket(code):
    """SDSK/ORD 코드를 터미널 화면으로 연다(tickets/term_ops.goto - 기존 티켓 자동화와 같은 경로)."""
    code = " ".join(str(code or "").split())
    if not code:
        return {"started": False, "message": "코드가 없습니다."}
    proc = subprocess.Popen(
        [sys.executable, "-c", "import sys; sys.path.insert(0, 'tickets'); import term_ops; term_ops.goto(sys.argv[1], wait=8)", code],
        cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


# ---------------- 퀵 기사님 안내문 (Word) ----------------
LOGO_PATH = BASE_DIR / "assets" / "bloomberg_logo.png"


def quick_logo_bytes():
    return LOGO_PATH.read_bytes()


def _has_hangul(s):
    return any("\uac00" <= ch <= "\ud7a3" for ch in s)


def quick_suggest(company_en):
    """과거 이력(history.db)에서 같은 회사의 한글 주소 줄을 찾아 후보로 돌려준다(최대 3개).
    회사명 한글 표기는 이력에 없는 경우가 많아 후보에 넣지 않는다 - 팝업에서 직접 입력."""
    import history_search
    company_en = " ".join(str(company_en or "").split())
    if len(company_en) < 4:
        return {"suggestions": []}
    res = history_search.search(company_en)
    rows = res["results"] if isinstance(res, dict) else res
    out, seen = [], set()
    for r in rows:
        cells = r.get("cells") or []
        addr = cells[6] if len(cells) > 6 else ""
        lines = [ln.strip() for ln in str(addr).splitlines() if _has_hangul(ln) and "cust#" not in ln]
        # '도착후 연락' 같은 메모성 줄은 주소가 아니라서 제외(도로명/동/층 숫자가 있는 줄만)
        lines = [ln for ln in lines if any(ch.isdigit() for ch in ln) and "연락" not in ln]
        if not lines:
            continue
        text = " ".join(lines)
        if text in seen:
            continue
        seen.add(text)
        out.append({"text": text, "date": r.get("sheetDate") or ""})
        if len(out) >= 3:
            break
    return {"suggestions": out}


# ---------------- 서명 용지 (POD/POC) ----------------
def print_sign_form(kind, order):
    """POD(ZOR 배송)/POC(ZRE 회수) 서명 용지를 채워서 기본 프린터로 인쇄한다(sign_forms.py를 별도 프로세스로).
    교체(ZRX)·내부배송(ZINT/ZINP 등)은 대상이 아니라 서버에서도 한 번 더 거른다."""
    import tempfile
    kind = str(kind or "").lower()
    prefix = str((order or {}).get("prefix") or "").upper()
    if kind not in ("pod", "poc"):
        return {"started": False, "message": "종류는 pod 또는 poc 입니다."}
    if (kind == "pod" and prefix != "ZOR") or (kind == "poc" and prefix != "ZRE"):
        return {"started": False, "message": f"{kind.upper()} 서명 용지는 {'ZOR' if kind == 'pod' else 'ZRE'} 오더만 인쇄합니다 ({prefix or '?'} 제외)."}
    if not (order.get("items") or []):
        return {"started": False, "message": "품목이 없습니다."}
    fd, path = tempfile.mkstemp(prefix="signjob_", suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump({"mode": "print", "kind": kind, "order": order}, fh, ensure_ascii=False)
    proc = subprocess.Popen(
        [sys.executable, "sign_forms.py", path], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}

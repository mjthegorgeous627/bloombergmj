"""Order/item CRUD (list/get/delete/restore/move/merge/highlight/manual
notes) and board data assembly (per-date board rows, holidays, custom
types, notes pages) for the workbench dashboard.

Split out of workbench_app.py on 2026-09-16 - see db.py's module docstring
and BLOOMBERG_HANDOFF.md. Pure code motion.
"""

from __future__ import annotations

import json
import re
import time
from datetime import date, timedelta
from pathlib import Path

from db import (
    DB_PATH,
    DELETED_ITEM_RETENTION_DAYS,
    MERGE_FIELDS,
    _log_action_batch,
    add_event,
    connect,
    ensure_db,
    now_text,
)
from holiday_check import kr_holiday_labels
from order_domain import resolve_item_type

# Used by _date_label() below - kept local since it's a pure display helper
# with no other overlap with the SAP/Portal automation modules.
_WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

def order_payload(row, items):
    data = dict(row)
    data["items"] = [dict(item) for item in items]
    data["serial_count"] = sum(1 for item in items if item["serial"])
    data["material_only"] = bool(items) and all(item["mode"] == "material_only" for item in items)
    data["item_summary"] = ", ".join(
        f"{item['material'] or 'no-mat'} x{item['qty']}" + (f" / {item['serial']}" if item["serial"] else "")
        for item in items
    )
    return data


def list_orders(query="", status=""):
    ensure_db()
    where = []
    params = []
    if query:
        like = f"%{query}%"
        where.append("(order_no LIKE ? OR delivery_no LIKE ? OR customer LIKE ? OR memo LIKE ?)")
        params.extend([like, like, like, like])
    if status:
        where.append("status=?")
        params.append(status)
    sql = "SELECT * FROM orders"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY CASE status WHEN 'new' THEN 0 WHEN 'working' THEN 1 WHEN 'posted' THEN 2 WHEN 'printed' THEN 3 ELSE 4 END, updated_at DESC, id DESC LIMIT 300"
    with connect() as con:
        rows = con.execute(sql, params).fetchall()
        result = []
        for row in rows:
            items = con.execute(
                "SELECT * FROM order_items WHERE order_id=? AND deleted_at IS NULL ORDER BY line_no", (row["id"],)
            ).fetchall()
            result.append(order_payload(row, items))
        return result


def get_order(order_id):
    with connect() as con:
        row = con.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        if not row:
            return None
        items = con.execute(
            "SELECT * FROM order_items WHERE order_id=? AND deleted_at IS NULL ORDER BY line_no", (order_id,)
        ).fetchall()
        events = con.execute("SELECT * FROM events WHERE order_id=? ORDER BY id DESC LIMIT 80", (order_id,)).fetchall()
        data = order_payload(row, items)
        data["events"] = [dict(e) for e in events]
        return data


def delete_items(item_ids):
    """Soft-delete: marks order_items rows with deleted_at instead of removing
    them, so a mistaken delete stays recoverable in the DB for
    DELETED_ITEM_RETENTION_DAYS before purge_deleted_items() actually drops it
    (also undoable via Ctrl+Z / "실행취소" for a short while - see
    _log_action_batch / undo_last_action). The parent order is never touched
    here - once every one of its items is soft-deleted it just stops appearing
    on the board (board_orders_for_date only counts live items)."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids:
        return {"deleted": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT id, order_id, deleted_at FROM order_items WHERE id IN ({placeholders})", item_ids
        ).fetchall()
        order_ids = {r["order_id"] for r in rows}
        con.execute(
            f"UPDATE order_items SET deleted_at=? WHERE id IN ({placeholders})", [now_text(), *item_ids]
        )
        _log_action_batch(con, "delete", f"{len(item_ids)}건 삭제", [(r["id"], r["deleted_at"]) for r in rows])
        for order_id in order_ids:
            add_event(
                con, order_id,
                f"{len(item_ids)}건 중 일부 항목 삭제 (workbench, {DELETED_ITEM_RETENTION_DAYS}일간 복구 가능)",
            )
    return {"deleted": len(item_ids)}


def list_deleted_items():
    """The actual browsable "trash" the DELETED_ITEM_RETENTION_DAYS window
    promises - Ctrl+Z only reaches the last UNDO_HISTORY_LIMIT actions, so a
    delete from a few days ago (already outside that window) previously had
    no user-facing way back short of a direct DB query. Anything still
    returned here is by definition not yet hard-purged, since purge_deleted_
    items() runs on every ensure_db() call."""
    ensure_db()
    with connect() as con:
        rows = con.execute(
            """
            SELECT i.*, o.order_no, o.order_type, o.delivery_no, o.customer, o.phone, o.source_date
            FROM order_items i JOIN orders o ON o.id = i.order_id
            WHERE i.deleted_at IS NOT NULL
            ORDER BY i.deleted_at DESC
            """
        ).fetchall()
        return [dict(r) for r in rows]


def restore_items(item_ids):
    """Un-does a soft-delete directly, for items outside Ctrl+Z's reach.
    Reuses the "delete" action type for its own undo-log entry (same
    deleted_at column, same restore-by-setting-it-back semantics either
    direction) - undoing a restore correctly re-deletes the item."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids:
        return {"restored": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT id, order_id, deleted_at FROM order_items WHERE id IN ({placeholders})", item_ids
        ).fetchall()
        order_ids = {r["order_id"] for r in rows}
        con.execute(f"UPDATE order_items SET deleted_at=NULL WHERE id IN ({placeholders})", item_ids)
        _log_action_batch(con, "delete", f"{len(item_ids)}건 복구", [(r["id"], r["deleted_at"]) for r in rows])
        for order_id in order_ids:
            add_event(con, order_id, f"{len(item_ids)}건 항목 복구 (workbench)")
    return {"restored": len(item_ids)}


def move_items_date(item_ids, iso):
    """Set a per-item date override. Items of the same order are free to end up
    on different dates this way (see board_orders_for_date) - moving one item
    never touches its sibling items or the parent order's own source_date."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids or not re.match(r"^\d{4}-\d{2}-\d{2}$", iso or ""):
        return {"moved": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT id, order_id, item_date FROM order_items WHERE id IN ({placeholders})", item_ids
        ).fetchall()
        order_ids = {r["order_id"] for r in rows}
        # Also clear board_pos: this path (checkbox + date picker, or dropping
        # into empty space) doesn't say WHERE in the target date to land, so
        # dragging in a stale board_pos number left over from the item's
        # previous date would place it at a semi-random spot instead of the
        # predictable "end of the list" a NULL gives it. The precise-position
        # drag path (onRowDrop) immediately re-sets board_pos for the whole
        # target date right after this call, overriding the NULL anyway.
        con.execute(f"UPDATE order_items SET item_date=?, board_pos=NULL WHERE id IN ({placeholders})", [iso, *item_ids])
        _log_action_batch(con, "move_date", f"{len(item_ids)}건 {iso}(으)로 이동", [(r["id"], r["item_date"]) for r in rows])
        for order_id in order_ids:
            add_event(con, order_id, f"{iso}(으)로 항목 이동 (workbench)")
    return {"moved": len(item_ids)}


def set_items_board_status(item_ids, status):
    """Per-item Done/Ready marker for the workbench board - independent of
    orders.status, which tracks the separate SAP/Portal automation pipeline."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids:
        return {"updated": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT id, board_status, order_id FROM order_items WHERE id IN ({placeholders})", item_ids
        ).fetchall()
        con.execute(
            f"UPDATE order_items SET board_status=?, sn_relo_error='' WHERE id IN ({placeholders})",
            [status, *item_ids],
        )
        _log_action_batch(
            con, "status", f"{len(item_ids)}건 상태 변경 → {status or '(초기화)'}",
            [(r["id"], r["board_status"]) for r in rows],
        )
        # Done/Ready/Cancel/초기화 버튼은 사용자가 그 행을 직접 확인/처리했다는
        # 신호이므로, run_pod_update()가 재시도 "시작" 시점에 pod_error를 지우는
        # 것과 같은 이유로 여기서도 지운다(sn_relo_error도 위에서 함께 지움).
        # 안 지우면 .row-error의 CSS background(orange, 3719행)가
        # .row-done/.row-pending/.row-cancelled와 같은 specificity에서 소스
        # 순서상 나중이라 항상 이겨 - board_status를 아무리 바꿔도 행이 계속
        # 주황으로 보임. 실사고 2026-08-27: ZOR 7847766 - POD 자동처리가 화면
        # 인식 실패로 에러난 뒤 사용자가 포털에서 수동으로 완료 처리하고 Done을
        # 눌러도 계속 주황이었음.
        # pod_error/portal_error는 order_items가 아니라 orders 테이블에 있고
        # 이미 오더 단위로 모든 항목에 표시되므로(품목별 구분 없음), 지우는
        # 단위도 동일하게 오더 단위로 맞춘다.
        order_ids = sorted({r["order_id"] for r in rows if r["order_id"]})
        if order_ids:
            oph = ",".join("?" for _ in order_ids)
            con.execute(f"UPDATE orders SET pod_error='', portal_error='' WHERE id IN ({oph})", order_ids)
    return {"updated": len(item_ids)}


def add_manual_note(cells, target_date=None, item_type=None, order_label=None, cell_count=None):
    """Create one free-standing memo row on the board, not tied to any real
    SAP/Excel order - 2026-08-20 user request ("행 생성은 오더 가져올 때만
    되는데, 메모용으로 한 줄씩 추가하고 싶다"). Every other row on the board
    comes from load_today_from_excel()/import-rows-json, both of which key
    off a real SAP order number; this is the one path that makes a row out
    of nothing but typed text.

    order_no is ALWAYS left blank (never user-settable) - it's the real
    matching key SAP resync/POD/Portal automation key off (see orderLabel's
    onblur handler and update_order_field()'s own comment), so a manual row
    must never carry one a script could mistake for a real order. order_type
    is what actually renders as the visible Order# label when order_no is
    blank (orderLabel = f"{order_type} {order_no}".strip()) - it defaults to
    the literal 'memo' but is now user-settable via `order_label` (2026-08-27
    user request: e.g. "SDSK1333608359"-style non-SAP shipment codes need to
    go in that same column, not just the literal word "memo"). delivery_no is
    stored as SQL NULL (not ''), not shown to the client - orders has
    UNIQUE(order_no, delivery_no), and with order_no always '' here, plain ''
    for every memo row would collide on the second one; SQL NULL is never
    equal to itself under UNIQUE so any number of blank-order_no memo rows
    can coexist. material is left blank when the caller doesn't fill that
    slot: run_portal()/zrec_handler's prepare step/register_serial all raise
    on a blank material rather than silently doing something to a real
    delivery, so this row fails safely if someone ever checks it and hits a
    Portal/ZREC button by mistake.

    cells: list of up to 7 free-text strings, positionally mapped to
    제목(description)/M-N(material)/S-N(serial)/Customer/Phone/Address/Memo
    - 2026-08-25 change (user request). A memo row isn't a real order, so
    forcing everything into just the 제목 field (the old single `text`
    argument) undersold what it could hold; splitting it into 7 rigidly
    *labeled* fields up front was rejected as needless complexity for
    something meant to stay a free note. addManualRow()'s modal instead asks
    how many of these 7 slots to use and only renders that many boxes -
    cells[0] (제목) is the one still required (same as the old `text`).
    item_type is no longer restricted to a hardcoded 배송/회수 pair - it now
    accepts whatever the client's Type dropdown offers (allTypeOptions() on
    the JS side, including any session-added custom types), matching every
    other row's Type field.

    cell_count: how many of the 7 cells addManualRow()'s modal was set to
    (1-7, see manual_cell_count column) - stored verbatim so the board can
    later reproduce the exact same cell boundaries the user chose, instead of
    re-inferring them from which cells happen to be non-blank (a cell left
    blank because it was optional is still a cell, not something to silently
    merge away - 2026-08-27 user report, "Index Event" 제목만 채우고 나머지
    6칸을 비운 채 7칸으로 만들었는데 1칸으로만 보임). Falls back to
    len(cells) (i.e. no merging - one cell per non-default slot) if not
    given, which is only reachable via a raw API call bypassing the modal."""
    ensure_db()
    cells = [str(c or "").strip() for c in (cells or [])][:7]
    cells += [""] * (7 - len(cells))
    description, material, serial, customer, phone, address, memo = cells
    if not description:
        raise RuntimeError("제목이 비어 있습니다.")
    target_date = (target_date or now_text()[:10]).strip()
    item_type = (item_type or "").strip() or "배송"
    order_label = (order_label or "").strip() or "memo"
    try:
        cell_count = int(cell_count)
    except (TypeError, ValueError):
        cell_count = 7
    cell_count = max(1, min(7, cell_count))
    with connect() as con:
        cur = con.execute(
            """
            INSERT INTO orders(order_no, delivery_no, order_type, customer, phone, address, memo,
                status, source_sheet, source_date, created_at, updated_at, manual_cell_count)
            VALUES ('', NULL, ?, ?, ?, ?, ?, 'new', 'manual-memo', ?, ?, ?, ?)
            """,
            (order_label, customer, phone, address, memo, target_date, now_text(), now_text(), cell_count),
        )
        order_id = cur.lastrowid
        con.execute(
            """
            INSERT INTO order_items(order_id, line_no, description, material, serial, qty, mode, item_type, item_date, board_status)
            VALUES (?, 1, ?, ?, ?, 1, 'material_only', ?, ?, '')
            """,
            (order_id, description, material, serial, item_type, target_date),
        )
        add_event(con, order_id, f"수동 메모 행 추가 ({item_type}): {description}")
    return {"orderId": order_id}


ITEM_EDIT_FIELDS = ("description", "material", "serial", "mode", "item_type")
# item_type added 2026-08-26 - the board's Type dropdown (배송/회수/Bloomberg/
# Delayed/회수 Delayed/custom, see allTypeOptions() on the client) changed
# only the in-memory ROWS array (row.itemType) and was never actually sent
# to the server at all - looked fine until the next full reload, at which
# point resolve_item_type()'s own bug (see its docstring) additionally
# stomped anything but a literal 배송/회수 back to a guessed default. Both
# needed fixing together: this makes the dropdown's change actually
# persist (and, via manual_item_edits below, survive the next SAP
# auto-resync), and resolve_item_type() now trusts whatever got saved here.
# mode 추가 (2026-08-18) - 원래 여기 있던 실측: ZINP 551059036 "COPYHOLDER
# FELLOWES BOOKLIFT"(qty=1) 건은 _scan_sheet_groups()의 import-시점 mode
# 휴리스틱(qty>1일 때만 "serial 없음"으로 추정)이 놓쳐서 "1. Serial 등록 &
# QR 인쇄"가 "Serial Number가 비어 있습니다" 에러로 실패했던 게 시작.
# 처음엔 board에 수동 토글 배지를 만들었는데 사용자 피드백으로 제거함 -
# 대신 run_portal()이 버튼 클릭 시점의 item["serial"] 값 자체(비어있으면
# 그걸로 충분)로 no_serial을 판정하도록 고쳐서, 사람이 아무것도 안 해도
# "serial을 안 적었으면 시리얼 없는 품목"이라는 사용자의 실제 워크플로우
# 그대로 처리됨 - run_portal()의 no_serial 판정 주석 참고. mode 필드는
# API로는 여전히 편집 가능하게 남겨뒀지만(다른 필드들과 일관되게) 이 결정
# 경로에선 더 이상 안 쓰임. material/qty는 no_serial이어도 항상 그대로
# 검증/입력됨(portal_register_serial.py의 _validate_page()) - serial
# 입력만 건너뛴다.
# order_type added 2026-08-06 - orderLabel was one of the last board cells
# still screen-only (see board_row()'s own comment). order_no itself is
# deliberately NOT in this list and never will be through this path - it's
# the real matching key SAP resync (import_sap_rows/load_today_from_excel),
# POD delivery resolution (_resolve_pod_order), and Portal automation's
# Excel lookups all key off, so silently renaming it in place risks
# orphaning those lookups or colliding with a different order that happens
# to already use the new number. The client only ever sends the order_type
# portion of an edited orderLabel for exactly this reason - see its own
# onblur handler. deliveryNo (OBD/ORD/SDSK lines) has its own dedicated
# update_delivery_codes() instead of going through this generic path, since
# it's really 2 underlying columns (delivery_no + extra_codes) at once.
ORDER_EDIT_FIELDS = ("customer", "phone", "address", "memo", "order_type")


def update_item_field(item_id, field, value):
    """Persists an inline edit to one item's description/material/serial cell.
    Previously these cells were screen-only - editing them updated the
    in-browser ROWS array and nothing else, so a typed-in serial number
    (exactly the kind Portal automation needs) silently vanished on the next
    page reload or workbench restart. Undo-able like every other action here
    (prev_value carries {field, value} as JSON since the touched column
    varies per call, same pattern merge/unmerge already use)."""
    if field not in ITEM_EDIT_FIELDS:
        raise ValueError(f"invalid field: {field}")
    item_id = int(item_id)
    with connect() as con:
        row = con.execute(f"SELECT id, order_id, {field} FROM order_items WHERE id=?", (item_id,)).fetchone()
        if not row:
            return {"updated": False}
        con.execute(f"UPDATE order_items SET {field}=? WHERE id=?", (value, item_id))
        # Mark this field as human-edited so the next SAP auto-sync reimport
        # (load_today_from_excel(), ~every 10 min) doesn't silently overwrite
        # it back from Excel - see manual_item_edits' own comment in ensure_db().
        con.execute(
            "INSERT OR IGNORE INTO manual_item_edits(item_id, field) VALUES (?, ?)",
            (item_id, field),
        )
        _log_action_batch(
            con, "edit_item", f"항목 {field} 수정",
            [(item_id, json.dumps({"field": field, "value": row[field]}))],
        )
        add_event(con, row["order_id"], f"{field} 수정 (workbench): {value}")
    return {"updated": True}


def update_order_field(item_ids, field, value):
    """Persists an inline edit to a shared customer/phone/address/memo cell -
    same screen-only gap as update_item_field(), for the order-level fields.
    Resolves the checked/edited item ids down to their distinct order ids
    (a rendered block can span multiple real orders once merged - see
    merge_orders()) and updates all of them to the new value, matching what
    the client already shows (the shared cell's rowSpan covers every row in
    the merged block)."""
    if field not in ORDER_EDIT_FIELDS:
        raise ValueError(f"invalid field: {field}")
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids:
        return {"updated": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        order_ids = sorted({
            r["order_id"] for r in con.execute(
                f"SELECT DISTINCT order_id FROM order_items WHERE id IN ({placeholders})", item_ids
            ).fetchall()
        })
        if not order_ids:
            return {"updated": 0}
        order_placeholders = ",".join("?" for _ in order_ids)
        prior = con.execute(f"SELECT id, {field} FROM orders WHERE id IN ({order_placeholders})", order_ids).fetchall()
        con.execute(f"UPDATE orders SET {field}=? WHERE id IN ({order_placeholders})", [value, *order_ids])
        _log_action_batch(
            con, "edit_order", f"오더 {field} 수정",
            [(r["id"], json.dumps({"field": field, "value": r[field]})) for r in prior],
        )
        for order_id in order_ids:
            add_event(con, order_id, f"{field} 수정 (workbench): {value}")
    return {"updated": len(order_ids)}


def update_delivery_codes(item_ids, text):
    """Persists an edit to the Order# cell's secondary code line(s) - OBD/
    ORD/SDSK, one per line (see board_row()'s combined `deliveryNo`). Splits
    the edited multi-line text back into the 2 real columns underneath:
    whichever line starts with "OBD" (label stripped) becomes delivery_no -
    the real matching key POD resolution/the orders table's own
    UNIQUE(order_no, delivery_no) depend on, so it must stay a single clean
    value - everything else (ORD/SDSK/etc) is joined back into extra_codes,
    pure free text with no matching-key role. Same resolve-to-distinct-
    order-ids pattern update_order_field() uses."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    if not item_ids:
        return {"updated": 0}
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    obd_lines = [ln for ln in lines if re.match(r"^OBD\b", ln, re.I)]
    other_lines = [ln for ln in lines if not re.match(r"^OBD\b", ln, re.I)]
    new_delivery_no = re.sub(r"^\s*OBD\s*[:#-]?\s*", "", obd_lines[0], flags=re.I).strip() if obd_lines else ""
    new_extra_codes = "\n".join(other_lines + obd_lines[1:])  # a stray 2nd "OBD ..." line just rides along as text
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        order_ids = sorted({
            r["order_id"] for r in con.execute(
                f"SELECT DISTINCT order_id FROM order_items WHERE id IN ({placeholders})", item_ids
            ).fetchall()
        })
        if not order_ids:
            return {"updated": 0}
        order_placeholders = ",".join("?" for _ in order_ids)
        prior = con.execute(
            f"SELECT id, delivery_no, extra_codes FROM orders WHERE id IN ({order_placeholders})", order_ids
        ).fetchall()
        con.execute(
            f"UPDATE orders SET delivery_no=?, extra_codes=? WHERE id IN ({order_placeholders})",
            [new_delivery_no, new_extra_codes, *order_ids],
        )
        _log_action_batch(
            con, "edit_delivery_codes", "오더 배송코드 수정",
            [(r["id"], json.dumps({
                "delivery_no": r["delivery_no"], "extra_codes": r["extra_codes"],
            })) for r in prior],
        )
        for order_id in order_ids:
            add_event(con, order_id, f"배송코드 수정 (workbench): {new_delivery_no} / {new_extra_codes}")
    return {"updated": len(order_ids)}


HIGHLIGHT_FIELDS = (
    "type", "order", "description", "material", "serial",
    "customer", "phone", "address", "memo",
)
HIGHLIGHT_BG_VALUES = {"yellow", "pink"}
HIGHLIGHT_TEXT_VALUES = {"red"}


def set_cell_highlight(item_id, field, style):
    """Sets (or, if `style` is empty, clears) one cell's manual highlight -
    background color / text color / bold, whole-cell only (not per-character;
    see [[project-workbench-sap-migration]] round 15 for why). `style` is the
    FULL desired {"bg":..,"color":..,"bold":..} state, not a delta - the
    client always sends the merged result so this can stay a plain replace,
    same as update_item_field/update_order_field. Undo-able like every other
    action here."""
    if field not in HIGHLIGHT_FIELDS:
        raise ValueError(f"invalid field: {field}")
    item_id = int(item_id)
    style = style or {}
    bg = style.get("bg") or None
    color = style.get("color") or None
    bold = bool(style.get("bold"))
    if bg is not None and bg not in HIGHLIGHT_BG_VALUES:
        raise ValueError(f"invalid bg: {bg}")
    if color is not None and color not in HIGHLIGHT_TEXT_VALUES:
        raise ValueError(f"invalid color: {color}")
    clean = {}
    if bg:
        clean["bg"] = bg
    if color:
        clean["color"] = color
    if bold:
        clean["bold"] = True
    new_json = json.dumps(clean) if clean else None
    with connect() as con:
        if not con.execute("SELECT 1 FROM order_items WHERE id=?", (item_id,)).fetchone():
            return {"updated": False}
        prior = con.execute(
            "SELECT style_json FROM cell_highlights WHERE item_id=? AND field=?", (item_id, field)
        ).fetchone()
        prior_json = prior["style_json"] if prior else None
        if new_json:
            con.execute(
                "INSERT INTO cell_highlights(item_id, field, style_json) VALUES(?,?,?) "
                "ON CONFLICT(item_id, field) DO UPDATE SET style_json=excluded.style_json",
                (item_id, field, new_json),
            )
        else:
            con.execute("DELETE FROM cell_highlights WHERE item_id=? AND field=?", (item_id, field))
        _log_action_batch(
            con, "edit_highlight", f"{field} 강조 표시 변경",
            [(item_id, json.dumps({"field": field, "style_json": prior_json}))],
        )
    return {"updated": True}


def reorder_items(item_ids_in_order):
    """Persist a manual row order within one date section. `item_ids_in_order`
    is the FULL ordered list of item ids currently visible in that section
    (not just the one that got dragged) - the client always resends its whole
    per-date order after any drag, so board_pos can just be a plain dense
    0..N-1 sequence instead of juggling fractional/gap positions server-side."""
    item_ids = [int(i) for i in item_ids_in_order if str(i).strip()]
    if not item_ids:
        return {"reordered": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT id, board_pos FROM order_items WHERE id IN ({placeholders})", item_ids
        ).fetchall()
        prev_by_id = {r["id"]: r["board_pos"] for r in rows}
        con.executemany(
            "UPDATE order_items SET board_pos=? WHERE id=?",
            [(pos, item_id) for pos, item_id in enumerate(item_ids)],
        )
        _log_action_batch(
            con, "reorder", f"{len(item_ids)}건 순서 변경",
            [(item_id, prev_by_id.get(item_id)) for item_id in item_ids],
        )
    return {"reordered": len(item_ids)}


def merge_orders(item_ids, fields):
    """Manual "Excel merge cells", independently per shared field, across
    DIFFERENT orders (e.g. three separate ZOR order numbers that are, in real
    life, one customer - or sometimes just one shared address under
    different contact names, hence per-field instead of one all-or-nothing
    merge_group). Resolves the checked rows down to their distinct order ids
    and stamps only the requested `fields` columns with a fresh group id. The
    board only actually renders the merged rowSpan when those orders' rows
    also end up adjacent (see reorder_items()) - same constraint real Excel
    has: you can only merge a contiguous cell range."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    fields = [f for f in (fields or []) if f in MERGE_FIELDS]
    if not item_ids or not fields:
        return {"merged": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        order_ids = sorted({
            r["order_id"] for r in con.execute(
                f"SELECT DISTINCT order_id FROM order_items WHERE id IN ({placeholders})", item_ids
            ).fetchall()
        })
        if len(order_ids) < 2:
            return {"merged": 0, "message": "서로 다른 오더에 속한 행을 2건 이상 선택하세요"}
        order_placeholders = ",".join("?" for _ in order_ids)
        columns = [f"merge_group_{f}" for f in fields]
        prior = con.execute(
            f"SELECT id, {', '.join(columns)} FROM orders WHERE id IN ({order_placeholders})", order_ids
        ).fetchall()
        # Per field, reuse a group id one of the SELECTED orders already has
        # (if any) instead of always minting a fresh one. Without this,
        # merging A+B and then later merging B+C (the natural way to build
        # up a 3-way merge one pair at a time) would stamp B with a brand
        # new id and silently orphan A - A's old id would then match no one,
        # so it'd quietly fall back out of the merged block. Picking the
        # smallest existing id among the current selection keeps repeated
        # merges of overlapping selections growing one connected group
        # instead of leaving pieces behind.
        new_ids_by_column = {
            col: (min((r[col] for r in prior if r[col]), default=None) or f"m{int(time.time() * 1000)}")
            for col in columns
        }
        set_clause = ", ".join(f"{c}=?" for c in columns)
        con.execute(
            f"UPDATE orders SET {set_clause} WHERE id IN ({order_placeholders})",
            [new_ids_by_column[c] for c in columns] + order_ids,
        )
        _log_action_batch(
            con, "merge", f"오더 {len(order_ids)}건 {'/'.join(fields)} 병합",
            [(r["id"], json.dumps({c: r[c] for c in columns})) for r in prior],
        )
        for order_id in order_ids:
            add_event(con, order_id, f"{'/'.join(fields)} 칸 병합 (workbench, {', '.join(new_ids_by_column.values())})")
    return {"merged": len(order_ids), "groupIds": new_ids_by_column, "fields": fields}


def unmerge_orders(item_ids, fields=None):
    """Clears merge_group_<field> for the distinct orders touched by the
    checked rows - undoes merge_orders(). `fields` omitted/empty means "clear
    all 4" (the common full-unmerge case); pass a subset to unmerge just one
    or two fields while leaving the rest merged."""
    item_ids = [int(i) for i in item_ids if str(i).strip()]
    fields = [f for f in (fields or MERGE_FIELDS) if f in MERGE_FIELDS]
    if not item_ids or not fields:
        return {"unmerged": 0}
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        order_ids = sorted({
            r["order_id"] for r in con.execute(
                f"SELECT DISTINCT order_id FROM order_items WHERE id IN ({placeholders})", item_ids
            ).fetchall()
        })
        if not order_ids:
            return {"unmerged": 0}
        order_placeholders = ",".join("?" for _ in order_ids)
        columns = [f"merge_group_{f}" for f in fields]
        prior = con.execute(
            f"SELECT id, {', '.join(columns)} FROM orders WHERE id IN ({order_placeholders})", order_ids
        ).fetchall()
        set_clause = ", ".join(f"{c}=NULL" for c in columns)
        con.execute(f"UPDATE orders SET {set_clause} WHERE id IN ({order_placeholders})", order_ids)
        _log_action_batch(
            con, "unmerge", f"오더 {len(order_ids)}건 {'/'.join(fields)} 병합 해제",
            [(r["id"], json.dumps({c: r[c] for c in columns})) for r in prior],
        )
        for order_id in order_ids:
            add_event(con, order_id, f"{'/'.join(fields)} 칸 병합 해제 (workbench)")
    return {"unmerged": len(order_ids)}


# ── SAP controls ──────────────────────────────────────────────────────────
#
# These spawn the exact same scripts launcher.py (0minjung.bat) uses
# (startup.py / main.py / order.py / open_session.py) - both front-ends drive
# the same 4 shared SAP GUI sessions via the machine-wide SAP GUI Scripting
# COM object (see sap_handler.get_sap_gui_auto), so it does not matter which
# UI started a given process; two of them touching those sessions at once is
# what's dangerous. Portal automation hit exactly this failure mode once for
# real (see portal_lock.py's 2026-05-28 incident note, a shared Chrome tab
# instead of a shared SAP session) - _running_sap_loop_pids() exists so we
# never repeat that with SAP.



def board_orders_for_date(target_date_iso):
    """
    Items are grouped onto a date by their *effective* date - item_date if set,
    otherwise the parent order's source_date - not by order. So one order's
    items can legitimately land on different dates (e.g. a ZRX's 배송 leg ships
    today, its 회수 leg is picked up a few days later after being moved on the
    board). The order's shared fields (customer/phone/address/memo) still come
    along with whichever of its items land on this date.
    """
    ensure_db()
    with connect() as con:
        orders_by_id = {row["id"]: row for row in con.execute("SELECT * FROM orders").fetchall()}
        item_rows = con.execute(
            """
            SELECT i.* FROM order_items i
            JOIN orders o ON o.id = i.order_id
            WHERE i.deleted_at IS NULL
              AND COALESCE(NULLIF(i.item_date, ''), o.source_date) = ?
            ORDER BY (i.board_pos IS NULL), i.board_pos, i.order_id, i.line_no
            """,
            (target_date_iso,),
        ).fetchall()
        grouped = {}
        for item in item_rows:
            grouped.setdefault(item["order_id"], []).append(item)
        result = []
        for order_id, items in grouped.items():
            order_row = orders_by_id.get(order_id)
            if order_row:
                result.append(order_payload(order_row, items))
        return result


def board_row(order, item, idx, total, highlights_by_item=None):
    """
    One flat row of data per item line, sent to the client as JSON. The
    client groups rows that share (orderLabel, customer, phone, address,
    memo) into one visually-merged block (rowspan) - this can span *more*
    than one `orders` DB row, e.g. a delivery/pickup pair imported as two
    separate rows for the same real-world order. Kept as flat data (not
    server-built HTML) so the client can recompute grouping after a move,
    edit, or delete without needing the server round-trip.
    """
    order_label = f"{order.get('order_type') or ''} {order.get('order_no') or ''}".strip()
    delivery_no = order.get("delivery_no") or ""
    # OBD (delivery_no, the real matching key) and ORD/SDSK (extra_codes,
    # free text - see its own comment in ensure_db()) render as one combined
    # multi-line "Order#" cell but live in two separate columns underneath -
    # see update_delivery_codes() for how an edit gets split back apart.
    code_lines = ([f"OBD {delivery_no}"] if delivery_no else []) + [
        ln for ln in (order.get("extra_codes") or "").splitlines() if ln.strip()
    ]
    return {
        "rowKey": f"{order['id']}-{idx}",
        "orderId": order["id"],
        "itemId": item["id"],
        "orderLabel": order_label,
        # Real order_no, sent separately from the combined orderLabel string
        # so the client can tell whether an edit to orderLabel touched the
        # number portion - see orderLabel's onblur handler and
        # update_order_field()'s own comment for why order_no itself is
        # never persisted through that edit path.
        "orderNo": order.get("order_no") or "",
        # Secondary order-code line(s) shown under the main label, e.g. "OBD 92149447" -
        # the Bloomberg-side delivery code (Outbound Delivery), as distinct from the SAP
        # order code (ZOR/ZRE/ZRX/ZINP/ZINT/...) in orderLabel above. Free-text editable
        # on the client, one code per line, so more lines (ORD/SDSK/etc) can be added.
        "deliveryNo": "\n".join(code_lines),
        "itemType": resolve_item_type(order, item, idx, total),
        "boardStatus": item.get("board_status") or "",
        # Portal(Serial/QR)/POD 자동화 실패 표시 (todo b, 2026-08-18) - order
        # 단위 값을 이 order에 속한 모든 item row에 동일하게 붙임. 클라이언트는
        # 이걸로 행 색칠(row-error) + 라벨 옆 ⚠ 배지(클릭 시 사유 팝업)를 그림.
        "portalError": bool(order.get("portal_error")),
        "portalErrorMsg": order.get("portal_error") or "",
        "podError": bool(order.get("pod_error")),
        "podErrorMsg": order.get("pod_error") or "",
        # S/N 오더 반영(paper relo 판별) 실패 표시 (2026-08-19) - portalError/
        # podError와 같은 시각 패턴(row-error 주황 + ⚠ 배지)이지만 이건 항목
        # 단위라 order가 아니라 item에서 직접 읽음 - 회수 항목의 S/N은
        # 항목별로 다르기 때문. sn_relo_one() 참고.
        "snReloError": bool(item.get("sn_relo_error")),
        "snReloErrorMsg": item.get("sn_relo_error") or "",
        # 메모 행 압축 렌더링용 (2026-08-27) - addManualRow() 모달에서 고른
        # 칸 수 그대로, buildGroupRows가 내용 유무 대신 이 값으로 칸 경계를
        # 재현함. 실제 오더에는 항상 NULL.
        "manualCellCount": order.get("manual_cell_count"),
        "mergeGroups": {f: order.get(f"merge_group_{f}") or "" for f in MERGE_FIELDS},
        "description": item.get("description") or "",
        "material": item.get("material") or "",
        "serial": item.get("serial") or "",
        "mode": item.get("mode") or "serial",
        "customer": order.get("customer") or "",
        "phone": order.get("phone") or "",
        "address": order.get("address") or "",
        "memo": order.get("memo") or "",
        # Manual per-cell highlight (see cell_highlights/set_cell_highlight) -
        # {field: {"bg":..,"color":..,"bold":..}}, only for fields that
        # actually have one set. Always keyed to THIS item's id, whether or
        # not this particular row ends up being the one that renders a given
        # field's <td> (order-level fields only render on a group's first
        # row - see buildGroupRows on the client).
        "highlights": (highlights_by_item or {}).get(item["id"], {}),
    }


def all_board_dates():
    """
    Every distinct effective date (item_date if set, else the parent order's
    source_date) that has at least one LIVE item - no forced "today" anchor.
    A date's banner disappears once it's fully cleared out (e.g. via "마감"
    at end of day) instead of sticking around empty; a fresh item landing
    under today (import/SAP sync) makes today's section reappear on its own.
    A single Excel sheet can hold several embedded date blocks weeks apart
    (see row_date_header) - the board is not just "today vs tomorrow".
    """
    ensure_db()
    with connect() as con:
        rows = con.execute(
            """
            SELECT DISTINCT COALESCE(NULLIF(i.item_date, ''), o.source_date) AS eff_date
            FROM order_items i
            JOIN orders o ON o.id = i.order_id
            WHERE i.deleted_at IS NULL
              AND COALESCE(NULLIF(i.item_date, ''), o.source_date) IS NOT NULL
              AND COALESCE(NULLIF(i.item_date, ''), o.source_date) != ''
            """
        ).fetchall()
    return sorted({r["eff_date"] for r in rows})


def _date_label(d):
    weekday = _WEEKDAY_KR[d.weekday()]
    today = date.today()
    if d == today:
        suffix = " (Today)"
    elif d == today + timedelta(days=1):
        suffix = " (Tomorrow)"
    else:
        suffix = ""
    return f"{d.strftime('%Y. %m. %d')} {weekday}요일{suffix}"


_DB_WAL_PATH = Path(str(DB_PATH) + "-wal")


def board_version():
    """Cheap change-detector for the client's fast poll loop (2026-09-01,
    사용자 지적: 20초는 너무 느리다/바뀌면 바로 반영되어야 한다).

    Tried SQLite's `PRAGMA data_version` first, but a quick throwaway-copy
    test showed it does NOT reliably reflect a write made by a different
    connection once you open a brand-new connection per call (as connect()
    does here) - it can keep returning a stale value. Falls back instead to
    the same trusted pattern already used for the online dashboard sync
    watcher (board_sync.py, 2026-08-25: 폴링→이벤트 감시로 전환, mtime 기반):
    the newest mtime across workbench.db and its -wal sidecar. Every commit
    in WAL mode appends to the -wal file and bumps its mtime, so this catches
    writes from ANY process (SAP auto-loop, ZREC/Portal handlers, main.py,
    ...) without needing each write path to separately announce itself -
    and it's just two stat() calls, no DB connection at all.
    """
    m = DB_PATH.stat().st_mtime
    if _DB_WAL_PATH.exists():
        m = max(m, _DB_WAL_PATH.stat().st_mtime)
    return m


def board_data():
    with connect() as con:
        highlight_rows = con.execute("SELECT item_id, field, style_json FROM cell_highlights").fetchall()
    highlights_by_item = {}
    for r in highlight_rows:
        highlights_by_item.setdefault(r["item_id"], {})[r["field"]] = json.loads(r["style_json"]) if r["style_json"] else {}

    data = []
    for iso in all_board_dates():
        orders = board_orders_for_date(iso)
        rows = []
        for order in orders:
            items = order.get("items") or []
            n = len(items)
            rows.extend(board_row(order, item, idx, n, highlights_by_item) for idx, item in enumerate(items))
        data.append({"iso": iso, "label": _date_label(date.fromisoformat(iso)), "rows": rows})
    return data


def nonworking_days_payload():
    """workbench 날짜 배너를 빨간색으로 표시할 정보 (2026-09-13). 주말은 굳이
    서버가 계산해서 안 보내고 클라이언트가 그날의 Date.getDay()로 직접 판단
    (JS가 이미 알고 있는 iso 문자열로 충분) - 여기선 클라이언트가 계산할 수
    없는 두 가지만 내려준다: (1) 한국 공휴일 이름(holiday_check.kr_holiday_labels,
    설날/추석 등 음력 계산 포함), (2) 사용자가 수동으로 지정한 날짜 목록. 연도
    범위는 작년~내년+1까지만 - board는 그 밖의 먼 미래/과거 날짜를 사실상
    보여줄 일이 없다."""
    this_year = date.today().year
    holidays_map = kr_holiday_labels(range(this_year - 1, this_year + 3))
    ensure_db()
    with connect() as con:
        rows = con.execute("SELECT date FROM manual_holidays").fetchall()
    return {"holidays": holidays_map, "manual": [r["date"] for r in rows]}


def set_manual_holiday(iso, on):
    """워크벤치 날짜 배너의 수동 빨간날 지정 토글. iso 형식이 아니면(오타 등)
    조용히 무시 - 잘못된 값이 테이블에 쌓이는 것보단 안전하다."""
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", iso or ""):
        return {"ok": False, "error": "invalid date"}
    ensure_db()
    with connect() as con:
        if on:
            con.execute(
                "INSERT OR IGNORE INTO manual_holidays(date, created_at) VALUES (?, ?)",
                (iso, now_text()),
            )
        else:
            con.execute("DELETE FROM manual_holidays WHERE date=?", (iso,))
    return {"ok": True}


def list_custom_types():
    ensure_db()
    with connect() as con:
        rows = con.execute("SELECT name FROM custom_types ORDER BY created_at").fetchall()
    return [r["name"] for r in rows]


def add_custom_type(name):
    name = (name or "").strip()
    if not name:
        return {"added": False}
    ensure_db()
    with connect() as con:
        con.execute(
            "INSERT OR IGNORE INTO custom_types(name, created_at) VALUES (?, ?)",
            (name, now_text()),
        )
    return {"added": True}


def list_pages():
    ensure_db()
    with connect() as con:
        rows = con.execute(
            "SELECT id, title, body, sort_pos, created_at, updated_at "
            "FROM notepad_pages ORDER BY sort_pos, id"
        ).fetchall()
    return [dict(r) for r in rows]


def add_page(title=""):
    title = (title or "").strip() or "새 메모장"
    ensure_db()
    now = now_text()
    with connect() as con:
        max_pos = con.execute("SELECT COALESCE(MAX(sort_pos), -1) FROM notepad_pages").fetchone()[0]
        cur = con.execute(
            "INSERT INTO notepad_pages(title, body, sort_pos, created_at, updated_at) "
            "VALUES (?, '', ?, ?, ?)",
            (title, max_pos + 1, now, now),
        )
        page_id = cur.lastrowid
    return {"added": True, "id": page_id, "title": title}


def update_page(page_id, field, value):
    """field='body' is saved verbatim (no strip) so a note that intentionally
    ends with blank lines keeps them - only 'title' gets trimmed, and an
    empty trimmed title is rejected rather than silently blanking the tab."""
    if field not in ("title", "body"):
        return {"updated": False}
    if field == "title":
        value = (value or "").strip()
        if not value:
            return {"updated": False}
    ensure_db()
    with connect() as con:
        con.execute(
            f"UPDATE notepad_pages SET {field}=?, updated_at=? WHERE id=?",
            (value, now_text(), page_id),
        )
    return {"updated": True}


def delete_page(page_id):
    ensure_db()
    with connect() as con:
        con.execute("DELETE FROM notepad_pages WHERE id=?", (page_id,))
    return {"deleted": True}



"""Excel <-> workbench sync: importing SAP/portal scan rows into the DB,
and exporting/mirroring finalized workbench state back out to Excel.

Split out of workbench_app.py on 2026-09-16 - see db.py's module docstring
and BLOOMBERG_HANDOFF.md. Pure code motion.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import openpyxl

import excel_handler
from config import EXCEL_PATH
from db import (
    BASE_DIR,
    _pop_prior_item_state,
    add_event,
    connect,
    digits,
    ensure_db,
    extract_qty,
    logger,
    now_text,
    parse_item_type,
    parse_order_text,
    row_date_header,
    text_value,
    today_sheet_names,
)

def _scan_sheet_groups(ws, sheet_name, row_start=1, row_end=None):
    """Parse one Excel sheet (or an explicit [row_start, row_end] slice of
    it) into order/item groups, same shape load_today_from_excel() used to
    build inline. Shared by load_today_from_excel() (whole "today" sheet(s),
    the old automatic path) and import_excel_range() (an explicit sheet +
    row range, the manual emergency-only tool - see both callers' own
    docstrings for why this is now manual-only, not part of the auto-sync
    loop). Date-header tracking (row_date_header()) still scans from row 1
    even when row_start is later, so a scoped range still gets the right
    date if its own rows don't repeat the header - callers that want a
    literal "just these rows, whatever date they claim" should pass
    row_start=1 if that matters."""
    row_end = row_end if row_end is not None else ws.max_row
    current = None
    current_date_iso = date.today().isoformat()  # until the first date-block header is seen
    groups = []
    for row in range(1, row_end + 1):
        header_date = row_date_header(ws, row)
        if header_date:
            current_date_iso = header_date
            current = None  # next order line starts a fresh group under this date
            continue

        # Order/customer identity tracking always runs from row 1, regardless
        # of row_start - a scoped range (import_excel_range()) commonly asks
        # for just the item rows of a block (e.g. rows 42-43) whose own
        # order-number/customer row sits a row or two earlier (row 41) and
        # wouldn't otherwise repeat there. Only the actual item line below is
        # gated by row_start, so "rows 42-43" correctly still attaches to the
        # right order/customer instead of silently importing nothing (a real
        # bug hit in production - see project_sap_watchdog_fix memory).
        order_text = text_value(ws.cell(row, 1).value)
        desc = text_value(ws.cell(row, 2).value)
        material = digits(ws.cell(row, 3).value)
        serial = text_value(ws.cell(row, 4).value)
        customer = text_value(ws.cell(row, 5).value)
        phone = text_value(ws.cell(row, 6).value)
        address = text_value(ws.cell(row, 7).value)
        memo = text_value(ws.cell(row, 8).value)
        item_type = parse_item_type(order_text)

        if order_text:
            order_type, order_no, delivery_no = parse_order_text(order_text)
            if order_no:
                current = {
                    "order_no": order_no,
                    "delivery_no": delivery_no,
                    "order_type": order_type,
                    "source_sheet": sheet_name,
                    "source_date": current_date_iso,
                    "source_row_start": row,
                    "source_row_end": row,
                    "customer": customer,
                    "phone": phone,
                    "address": address,
                    "memo": memo,
                    "items": [],
                }
                groups.append(current)
        if not current:
            continue
        if not any([desc, material, serial, customer, phone, address, memo]):
            continue
        if order_text:
            _, _, row_delivery = parse_order_text(order_text)
            if row_delivery:
                current["delivery_no"] = row_delivery
        current["source_row_end"] = row
        for key, value in [("customer", customer), ("phone", phone), ("address", address), ("memo", memo)]:
            if value and not current.get(key):
                current[key] = value
        is_header_item = (desc.strip().lower() == "item" and serial.strip().upper() in {"S/N", "SN", "SERIAL"})
        is_header_material = text_value(ws.cell(row, 3).value).strip().upper() in {"M/N", "MN", "MATERIAL"}
        if is_header_item or is_header_material:
            continue
        if row < row_start:
            continue  # identity/context above is still tracked; only the item itself is skipped
        if material or desc or serial:
            qty = extract_qty(desc)
            # Default is serial. Material-only is a rare exception; qty>1 with no S/N is the practical signal.
            mode = "material_only" if (not serial and material and qty > 1) else "serial"
            current["items"].append(
                {
                    "description": desc,
                    "material": material,
                    "serial": serial,
                    "qty": qty,
                    "mode": mode,
                    "item_type": item_type,
                }
            )
    return groups


def _ingest_groups(con, groups, additive=False):
    """Upsert already-parsed order/item groups (from _scan_sheet_groups())
    into workbench.db. Default (additive=False, load_today_from_excel()'s
    whole-sheet reimport) deletes and reinserts every matched order's items
    wholesale, which is why manual_item_edits/item_date/board_status/
    deleted_at/cell_highlights all need their own explicit carry-over logic
    below - see each's own comment. That wholesale replace is only safe
    when the scan actually covers an order's *complete* row range (true for
    a whole-sheet scan), and this used to be reused as-is by
    import_excel_range() too.

    Real bug found 2026-08-07: a scoped import_excel_range() scan (e.g.
    "just rows 27-28") only ever sees a *partial slice* of a sheet - if
    those rows happen to belong to an order that already exists in
    workbench.db (a real case: ZOR 7825897, order id 57, carried forward
    from its original 8/4 collection into 8/7's sheet - the user asked for
    just 2 of its rows), the old code still ran the same wholesale
    DELETE-all-items-then-reinsert-only-what-this-scan-saw path, silently
    destroying every item of that order NOT in the requested 2-row window -
    with no soft-delete, no undo-log entry (unlike every other
    data-mutating action in this file), so it wasn't even Ctrl+Z-recoverable
    afterward. additive=True (used by import_excel_range() - see its own
    docstring) fixes this: for an existing order, never deletes anything -
    only adds items from the scan that don't already match an existing
    (material, serial) item, same "never touch a field once written"
    guarantee import_sap_rows() already gives the routine SAP path.
    Returns the number of order-groups upserted."""
    imported = 0
    merged_groups = {}
    for group in groups:
        if not group["items"]:
            continue
        key = (group.get("order_no") or "", group.get("delivery_no") or "")
        if key not in merged_groups:
            merged_groups[key] = group
            continue
        target = merged_groups[key]
        target["items"].extend(group["items"])
        target["source_row_end"] = max(target.get("source_row_end") or 0, group.get("source_row_end") or 0)
        for field in ("customer", "phone", "address", "memo", "order_type", "source_date"):
            if group.get(field) and not target.get(field):
                target[field] = group[field]

    for group in merged_groups.values():
        group_date = group.get("source_date") or date.today().isoformat()
        existing = con.execute(
            "SELECT id, status FROM orders WHERE order_no=? AND COALESCE(delivery_no,'')=COALESCE(?, '')",
            (group["order_no"], group.get("delivery_no") or ""),
        ).fetchone()
        prior_by_key = {}
        prior_by_line = {}
        manual_edits_by_item = {}
        if existing and additive:
            # Scoped/partial scan (import_excel_range()) - this group may
            # only be a slice of the order's real full item list, so unlike
            # the branch below, never delete anything. Only bump
            # updated_at; leave order_type/source_sheet/source_date etc.
            # alone too (a partial scan's own row numbers/date don't
            # describe the order as a whole and shouldn't overwrite
            # whatever it already has).
            order_id = existing["id"]
            # Counter, not a set: several existing items can legitimately
            # share one (material, serial="") key (e.g. serial-less
            # material-only lines) - each incoming item should only match
            # off ONE such existing item, same one-to-one consumption
            # _pop_prior_item_state() already does for the non-additive
            # path below. Otherwise N>1 genuinely-new items sharing a key
            # with each other (not with anything pre-existing) would
            # wrongly dedupe against each other and under-recover.
            # Real bug found 2026-08-07 (order 67072686, sheet 8-10 rows
            # 19-20): this used to count ALL items regardless of deleted_at,
            # so an item a user had soft-deleted (still sitting in the 7-day
            # trash, e.g. from an earlier "체크된 행 삭제") permanently blocked
            # ever re-adding it via a scoped range import - the whole point
            # of the retry was to bring that exact item back, and it silently
            # matched against the deleted row and got skipped again. A
            # soft-deleted item is hidden, not "still on the board" - only
            # live (non-deleted) items should count as already-present.
            existing_counts = Counter()
            for prow in con.execute(
                "SELECT material, serial FROM order_items WHERE order_id=? AND deleted_at IS NULL", (order_id,)
            ).fetchall():
                existing_counts[(prow["material"] or "", prow["serial"] or "")] += 1
            next_line = con.execute(
                "SELECT COALESCE(MAX(line_no), 0) AS m FROM order_items WHERE order_id=?", (order_id,)
            ).fetchone()["m"]
            new_items = []
            for item in group["items"]:
                key = (item.get("material") or "", item.get("serial") or "")
                if existing_counts.get(key, 0) > 0:
                    existing_counts[key] -= 1
                    continue  # already on the board - additive-only, never rewritten
                new_items.append(item)
            for item in new_items:
                next_line += 1
                con.execute(
                    """
                    INSERT INTO order_items(order_id, line_no, description, material, serial, qty, mode, item_type)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        order_id, next_line, item["description"], item["material"], item["serial"],
                        item["qty"], item["mode"], item.get("item_type", ""),
                    ),
                )
            con.execute("UPDATE orders SET updated_at=? WHERE id=?", (now_text(), order_id))
            add_event(con, order_id, f"Excel 범위 불러오기({group['source_sheet']} {group['source_row_start']}-{group['source_row_end']}): 항목 {len(new_items)}개 추가")
            imported += 1
            continue
        if existing:
            order_id = existing["id"]
            # Real bug fixed 2026-08-03: this used to also overwrite
            # customer/phone/address/memo with whatever Excel
            # currently has, on EVERY reimport (the SAP auto-sync
            # loop runs this roughly every ~10 min for every order
            # still present in Excel) - silently reverting any
            # manual edit made via update_order_field() (the inline
            # cell editing added 2026-07-31) within one sync cycle,
            # with no visible error - it just looked like "my edit
            # didn't save" whenever the user happened to notice
            # later. item_date/board_status/deleted_at (per-item)
            # were already preserved across reimport; these four
            # order-level fields now get the same treatment - only
            # order_type and the source-sheet bookkeeping fields
            # (which aren't user-editable in workbench) still sync
            # from Excel every time.
            con.execute(
                """
                UPDATE orders SET order_type=?,
                    source_sheet=?, source_row_start=?, source_row_end=?, source_date=?, updated_at=?
                WHERE id=?
                """,
                (
                    group.get("order_type"), group.get("source_sheet"),
                    group.get("source_row_start"), group.get("source_row_end"), group_date,
                    now_text(), order_id,
                ),
            )
            prior_rows = con.execute(
                "SELECT id, line_no, description, material, serial, item_type, item_date, board_status, deleted_at "
                "FROM order_items WHERE order_id=?",
                (order_id,),
            ).fetchall()
            prior_ids = [prow["id"] for prow in prior_rows]
            # Which (item_id, field) pairs a human has directly
            # edited on the board - those fields must survive this
            # reimport untouched instead of taking Excel's fresh
            # value (see manual_item_edits' comment in ensure_db()
            # and its use just below in the insert loop).
            if prior_ids:
                placeholders = ",".join("?" for _ in prior_ids)
                for erow in con.execute(
                    f"SELECT item_id, field FROM manual_item_edits WHERE item_id IN ({placeholders})",
                    prior_ids,
                ).fetchall():
                    manual_edits_by_item.setdefault(erow["item_id"], set()).add(erow["field"])
            for prow in prior_rows:
                pkey = (prow["material"] or "", prow["serial"] or "")
                prior_by_key.setdefault(pkey, []).append(prow)
                prior_by_line[prow["line_no"]] = prow
            con.execute("DELETE FROM order_items WHERE order_id=?", (order_id,))
        else:
            cur = con.execute(
                """
                INSERT INTO orders(order_no, delivery_no, order_type, customer, phone, address, memo,
                    status, source_sheet, source_row_start, source_row_end, source_date, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'new', ?, ?, ?, ?, ?, ?)
                """,
                (
                    group.get("order_no"), group.get("delivery_no"), group.get("order_type"),
                    group.get("customer"), group.get("phone"), group.get("address"), group.get("memo"),
                    group.get("source_sheet"), group.get("source_row_start"), group.get("source_row_end"),
                    group_date, now_text(), now_text(),
                ),
            )
            order_id = cur.lastrowid
        for idx, item in enumerate(group["items"], start=1):
            prior = _pop_prior_item_state(prior_by_key, prior_by_line, item, idx)
            edited_fields = manual_edits_by_item.get(prior["id"], set()) if prior is not None else set()
            # A manually-edited field keeps its prior (human-typed)
            # value instead of taking Excel's fresh one - see
            # manual_item_edits' comment in ensure_db(). Fields not
            # in edited_fields still sync from Excel normally (e.g.
            # a serial SAP collects a few cycles after the item was
            # first created blank must still come through).
            description = prior["description"] if "description" in edited_fields else item["description"]
            material = prior["material"] if "material" in edited_fields else item["material"]
            serial = prior["serial"] if "serial" in edited_fields else item["serial"]
            # item_type added to this carry-over 2026-08-26, same reasoning as
            # description/material/serial above - a Type dropdown value the
            # user set by hand (now persisted via ITEM_EDIT_FIELDS, see its
            # own comment) shouldn't get silently overwritten by the next
            # emergency Excel reimport either.
            item_type = prior["item_type"] if prior is not None and "item_type" in edited_fields else item.get("item_type", "")
            cur = con.execute(
                """
                INSERT INTO order_items(order_id, line_no, description, material, serial, qty, mode, item_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    order_id, idx, description, material, serial,
                    item["qty"], item["mode"], item_type,
                ),
            )
            if prior is not None:
                if prior["item_date"] or prior["board_status"] or prior["deleted_at"]:
                    con.execute(
                        "UPDATE order_items SET item_date=?, board_status=?, deleted_at=? WHERE id=?",
                        (prior["item_date"], prior["board_status"], prior["deleted_at"], cur.lastrowid),
                    )
                # cell_highlights is keyed by item_id, which changes
                # on every reimport (order_items gets DELETEd and
                # reinserted above, not updated in place) - without
                # this, a highlight set on an item silently vanishes
                # the next time its order gets reimported. Re-point
                # it at the new id instead, same "carry state across
                # the delete+reinsert" treatment item_date/board_status/
                # deleted_at already get via prior above.
                con.execute(
                    "UPDATE cell_highlights SET item_id=? WHERE item_id=?",
                    (cur.lastrowid, prior["id"]),
                )
                if edited_fields:
                    con.execute(
                        "UPDATE manual_item_edits SET item_id=? WHERE item_id=?",
                        (cur.lastrowid, prior["id"]),
                    )
        add_event(con, order_id, f"Imported from Excel sheet {group['source_sheet']} rows {group['source_row_start']}-{group['source_row_end']}")
        imported += 1
    return imported


def load_today_from_excel():
    """Manual/emergency tool only - as of 2026-08-06, nothing calls this
    automatically anymore. SAP's own scripts (main.py/manual_order_handler.py)
    push freshly-collected orders straight into workbench.db via
    import_sap_rows() the moment they write them to Excel, instead of relying
    on this function to re-scan and re-parse the whole sheet ~every 10 min
    (that automatic Excel->Workbench round-trip was the actual source of the
    "workbench randomly goes down / edits vanish" reports - see
    project_sap_watchdog_fix memory - since every reimport had to re-derive
    structured data from Excel cell text and needed ever-growing per-field
    "don't overwrite this" carve-outs). Kept for real recovery scenarios
    (workbench.db lost/corrupted, or Excel was hand-edited and workbench
    needs to catch up) - use import_excel_range() instead when you only want
    a specific sheet/row range rather than all of today's sheet(s)."""
    ensure_db()
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    sheets = today_sheet_names(wb.sheetnames)
    if not sheets:
        wb.close()
        raise RuntimeError(f"today sheet not found in {EXCEL_PATH}")
    try:
        with connect() as con:
            imported = 0
            for sheet_name in sheets:
                groups = _scan_sheet_groups(wb[sheet_name], sheet_name)
                imported += _ingest_groups(con, groups)
    finally:
        wb.close()
    return {"imported": imported, "sheets": sheets}


def import_excel_range(sheet_name, row_start=1, row_end=None):
    """Manual, scoped emergency import - "이 시트의 이 행부터 이 행까지만
    가져오기", a one-shot, explicitly-targeted alternative to
    load_today_from_excel()'s whole-sheet scan. Meant for the rare case
    where something needs to be pulled back in from Excel by hand (e.g. a
    row someone typed directly into Excel, or recovering after a data
    problem) - NOT for routine syncing, which is now import_sap_rows()'s job
    (see load_today_from_excel()'s own docstring for why routine automatic
    Excel->Workbench reimport was removed)."""
    ensure_db()
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise RuntimeError(f"sheet not found: {sheet_name}")
        ws = wb[sheet_name]
        row_start = max(1, int(row_start or 1))
        row_end = int(row_end) if row_end else ws.max_row
        groups = _scan_sheet_groups(ws, sheet_name, row_start=row_start, row_end=row_end)
        with connect() as con:
            # additive=True: this is a scoped/partial slice of the sheet by
            # design (see this function's own docstring) - an order caught
            # in that slice may have more items elsewhere that this scan
            # never saw, so existing orders must only ever gain items here,
            # never lose them (see _ingest_groups()'s docstring for the
            # real 2026-08-07 incident this fixes).
            imported = _ingest_groups(con, groups, additive=True)
    finally:
        wb.close()
    return {"imported": imported, "sheet": sheet_name, "rowStart": row_start, "rowEnd": row_end}


def import_sap_rows(rows):
    """The real, routine SAP->workbench path (replaces the old automatic
    load_today_from_excel() reimport, see its own docstring). Takes the exact
    same structured row shape excel_handler.write_orders_to_excel() consumes -
    one order's items, each a dict with order_prefix/order_type/order_num/
    obd/extra_orders/description/quantity/material/serial_number/customer/
    phone/company/street/street2/memo (see sap_handler.build_excel_rows(),
    zrma_handler.build_excel_rows_zrma(), vl10g_handler's row builder) -
    called by main.py/manual_order_handler.py right after they write the
    same rows to Excel, so workbench.db reflects a freshly-collected order
    without ever re-parsing Excel cell text for it.

    Deliberately additive-only, unlike the Excel-reimport path: finds-or-
    creates the orders row, then appends new order_items rows starting after
    whatever line_no already exists - it never deletes/rewrites existing
    items. Each SAP handler only calls this once per genuinely new order it
    just detected (order_tracker's processed-order tracking prevents the
    same order being processed twice), so there's no reimport-driven need for
    manual_item_edits-style field preservation here - a field just never gets
    overwritten because the row that already exists is never touched again."""
    if not rows:
        return {"imported": 0}
    first = rows[0]
    order_no = str(first.get("order_num") or "").strip()
    if not order_no:
        return {"imported": 0, "error": "missing order_num"}
    delivery_no = str(first.get("obd") or "").strip()
    order_type = first.get("order_type") or ""
    today_iso = date.today().isoformat()

    address_parts = []
    if first.get("company"):
        address_parts.append(first["company"])
    street, street2 = first.get("street") or "", first.get("street2") or ""
    if street and street2:
        address_parts.append(f"{street},\n{street2}")
    elif street:
        address_parts.append(street)
    elif street2:
        address_parts.append(street2)
    if first.get("cust_no"):
        address_parts.append(f"(cust# {first['cust_no']})")
    new_address = "\n".join(address_parts)

    ensure_db()
    with connect() as con:
        existing = con.execute(
            "SELECT id FROM orders WHERE order_no=? AND COALESCE(delivery_no,'')=COALESCE(?, '')",
            (order_no, delivery_no or ""),
        ).fetchone()
        if existing:
            order_id = existing["id"]
            # Real bug found 2026-08-07 (order 7837832, ZOR): the order's
            # very first collection attempt failed (ZZ block, RPC
            # disconnect) and got saved as a placeholder shell with
            # customer/phone blank and a "[Delivery Block: ...]" memo. A
            # later retry collected the real customer/phone/address/memo
            # fine, but this branch used to only bump updated_at - the
            # additive-only "never touch an existing order field again"
            # rule (meant to protect a human's manual edit) was also
            # silently locking in that first failed attempt's blank/
            # placeholder fields forever. Fix: still never overwrite a
            # field that already holds real data (human-edited or a prior
            # good collection), but DO backfill a field that's still
            # genuinely blank - there's nothing to protect there.
            cur_row = con.execute(
                "SELECT customer, phone, address, memo FROM orders WHERE id=?", (order_id,)
            ).fetchone()
            backfill = {
                "customer": first.get("customer") or "",
                "phone": first.get("phone") or "",
                "address": new_address,
                "memo": first.get("memo") or "",
            }
            set_clauses, params = [], []
            for field, new_val in backfill.items():
                if new_val and not (cur_row[field] or "").strip():
                    set_clauses.append(f"{field}=?")
                    params.append(new_val)
            set_clauses.append("updated_at=?")
            params.append(now_text())
            params.append(order_id)
            con.execute(f"UPDATE orders SET {', '.join(set_clauses)} WHERE id=?", params)
            max_line = con.execute(
                "SELECT COALESCE(MAX(line_no), 0) AS m FROM order_items WHERE order_id=?", (order_id,)
            ).fetchone()["m"]
        else:
            extra_codes = "\n".join(first.get("extra_orders") or [])
            cur = con.execute(
                """
                INSERT INTO orders(order_no, delivery_no, order_type, customer, phone, address, memo,
                    status, source_sheet, source_row_start, source_row_end, source_date, created_at, updated_at,
                    extra_codes)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'new', ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    order_no, delivery_no, order_type, first.get("customer") or "", first.get("phone") or "",
                    new_address, first.get("memo") or "",
                    "SAP-direct", None, None, today_iso, now_text(), now_text(),
                    extra_codes,
                ),
            )
            order_id = cur.lastrowid
            max_line = 0

        for offset, item in enumerate(rows, start=1):
            serial = item.get("serial_number") or ""
            material = str(item.get("material") or "")
            try:
                qty = max(1, int(float(item.get("quantity") or 1)))
            except (ValueError, TypeError):
                qty = 1
            mode = "material_only" if (not serial and material and qty > 1) else "serial"
            item_type = item.get("order_prefix") or first.get("order_prefix") or ""
            con.execute(
                """
                INSERT INTO order_items(order_id, line_no, description, material, serial, qty, mode, item_type)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (order_id, max_line + offset, item.get("description") or "", material, serial, qty, mode, item_type),
            )
        add_event(con, order_id, f"SAP 자동수집 반영 ({len(rows)}행)")
    return {"imported": len(rows), "orderId": order_id}


def export_dates_to_excel(dates_payload):
    """마감(확정)된 데이터를 배송장 파일(EXCEL_PATH, "2026 배송장.xlsx") 하나에만 쓴다.

    2026-10-08 통합: 예전엔 별도 "마감 파일"(FINALIZE_EXPORT_PATH)에 openpyxl로 먼저
    쓰고 공유파일에 COM으로 한 번 더 반영했는데, 두 파일이 달라지는 사고가 반복돼
    공유파일 하나로 합쳤다. Dashboard / 이력 검색 / 퀵 금액 정리도 이 파일을 읽는다.

    {iso날짜: rows}를 날짜별 "M-D" 시트로 쓰고(없으면 날짜 순서 위치에 새로 만들고, 있으면
    맨 끝에 이어붙임 - excel_handler.write_finalized_sheet_com), 쓰기가 실패하면 예외를
    그대로 올려 호출부(/api/export)가 마감 실패로 처리한다 - 이 경우 workbench.db 항목은
    삭제되지 않는다. 성공한 뒤에만 Dashboard 갱신과 이력 색인을 백그라운드로 건다.
    쓰기 직전에 직전 상태를 `.before_last_마감` 사본 1개로 덮어써 둔다(한 단계 되돌리기용)."""
    excel_path = Path(EXCEL_PATH)
    if not excel_path.exists():
        raise FileNotFoundError(f"배송장 파일을 찾을 수 없음: {excel_path}")

    backup_path = excel_path.with_name(f"{excel_path.stem}.before_last_마감{excel_path.suffix}")
    shutil.copy2(excel_path, backup_path)

    excel_handler.mirror_finalized_rows_to_excel_path(dates_payload)

    result = {}
    for iso, rows in dates_payload.items():
        d = date.fromisoformat(iso)
        result[iso] = {"sheet": f"{d.month}-{d.day}", "rows": len(rows)}
    logger.info(f"배송장 마감 반영 완료: {list(dates_payload.keys())}")
    _trigger_dashboard_refresh(list(dates_payload.keys()))
    _trigger_history_reindex(excel_path)
    return {"exported": result, "backup": str(backup_path), "archive": {"ok": True}}


def _trigger_history_reindex(excel_path):
    """마감 직후 방금 저장된 마감 파일을 이력 검색 색인(history.db)에 반영한다
    (2026-10-01). 백그라운드 스레드라 마감 자체를 막지 않고, 실패해도 마감은
    정상 완료된 상태다."""
    try:
        import history_search

        history_search.reindex_async(files=[Path(excel_path)])
    except Exception:
        logger.warning("이력 검색 색인 갱신 실행 실패 (마감 자체는 정상 완료됨)", exc_info=True)


def _trigger_dashboard_refresh(iso_dates):
    """Fire-and-forget: launch vendor_dashboard.py in the background to
    refresh the local Vendor Activities Dashboard copy/paste sheet for the
    dates just 마감-ed. No console window (same pattern as every other
    background launch in this app - see startup.py/main.py above), and any
    failure to even start it is logged and swallowed rather than raised,
    since the 마감 export above has already succeeded by this point."""
    if not iso_dates:
        return
    try:
        subprocess.Popen(
            [sys.executable, "vendor_dashboard.py", "--dates", ",".join(iso_dates)],
            cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        logger.warning("Dashboard 자동 갱신 실행 실패 (마감 자체는 정상 완료됨)", exc_info=True)



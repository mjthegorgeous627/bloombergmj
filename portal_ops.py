"""Portal (TES-AMM/ShipERP) automation triggers: login, packing/ship,
label printing, and POD preview/update.

Split out of workbench_app.py on 2026-09-16 - see db.py's module docstring
and BLOOMBERG_HANDOFF.md. Pure code motion.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading

from db import BASE_DIR, add_event, connect, ensure_db, now_text
from items import set_items_board_status
from order_domain import resolve_item_type
from sap_ops import _run_hidden_with_toast

def portal_login_manual():
    """"Portal 로그인 (수동)" - just opens the automation Chrome profile
    (right debug port, right user-data-dir) to the SSO start page and stops
    there; no autofill, no auto-clicking Next, no BUIT wait-loop (see
    portal_login.py --open-only). For logging in by hand while the automated
    login flow (run as step 0 of every "Serial 등록 & QR 인쇄") is being
    sorted out - once this window sits on the real logged-in portal page,
    that step just detects the existing session and passes straight through,
    so a manual login here works exactly as well as an automated one. Leave
    the Chrome window open and reuse it - no need to reopen for each order
    (see portal_register_serial.py's own delivery-page navigation).
    2026-08-25: portal_login.py itself has no logging module (pure print()),
    so unlike morning_routine.py this one has no file-based fallback - the
    toast from _run_hidden_with_toast() is the only record of what it
    printed if it fails."""
    pid = _run_hidden_with_toast(
        [sys.executable, "portal_login.py", "--open-only"], "Portal 로그인 (수동)",
    )
    return {"started": True, "pid": pid}


def print_latest_zpl():
    """"2. 최근 ZPL 수동 인쇄" - was a pure UI stub ("다음 단계에서 연결
    예정") until 2026-08-07. Mirrors launcher.py's own button exactly:
    spawns print_zpl_file.py with no arguments, which auto-selects the most
    recently modified .ZPL/.zpl file in the Windows Downloads folder (the
    QR/label just downloaded from Bloomberg Portal) and sends it straight to
    the Zebra printer (ZDesigner GK420t) - see print_zpl_file.py's own
    find_latest_zpl()/send_raw(). Also handles a PDF label the same way
    (parses it via zebra_handler and rebuilds a ZPL label) - see
    print_zpl_file.py's is_pdf() branch. 2026-08-25: used to run in its own
    CREATE_NEW_CONSOLE window so the script's own print()'d success/failure
    messages (including "연결된 프린터 목록" when the printer name doesn't
    match) were visible immediately; now windowless via
    _run_hidden_with_toast(), same rationale as portal_login_manual() -
    print_zpl_file.py also has no logging module of its own, so the toast is
    the only place that "연결된 프린터 목록" text can still surface on
    failure."""
    pid = _run_hidden_with_toast([sys.executable, "print_zpl_file.py"], "ZPL 인쇄")
    return {"started": True, "pid": pid}


def _portal_pids_for_order(order_no):
    """Live processes actually running portal_ship_and_print.py for this
    order number right now - mirrors _running_sap_loop_pids()'s approach
    (a live process-list check, not just trusting a DB flag). Needed because
    'working' alone isn't reliable: if workbench itself gets restarted while
    a run is genuinely in progress, that run's watcher thread dies with the
    old process and never flips the order back to 'printed'/'error' - the
    status is then stuck at 'working' forever with nothing actually running,
    permanently blocking every future retry unless this is checked."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'portal_ship_and_print\\.py' -and "
        f"$_.CommandLine -match '{re.escape(str(order_no))}' }} | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        return []
    return [int(p.strip()) for p in out.splitlines() if p.strip().isdigit()]


def run_portal(order_id, skip_qr=False):
    """Kicks off portal_ship_and_print.py for one order. Guards against
    starting a second run for the same order while one is already
    'working' (started but not yet finished, AND a process for it is
    actually still alive - see _portal_pids_for_order()) or already
    'printed' (finished) - a double-click or an impatient second click
    before the first run's toast even lands would otherwise fire a second
    full Serial registration + Packing Post + QR print for an order already
    mid-flight or already done. portal_lock.py's shared-Chrome-tab lock
    (added after the real 2026-05-28 incident) stops the two runs from
    colliding mid-navigation, but would still let a genuine SECOND full run
    complete - this catches it before that even starts.

    skip_qr: passes --skip-qr through to portal_ship_and_print.py (stop
    after Serial registration + Packing Post; don't download/print QR
    labels). Used by the "4. 스페어 배송처리" flow (runSpareOneshot() in the
    dashboard JS, always skips QR), where a spare unit doesn't need a
    freshly printed label. The regular "1. Serial 등록 & QR 인쇄" button
    never passes this, so its behavior is unchanged."""
    ensure_db()
    with connect() as con:
        row = con.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        items = con.execute(
            "SELECT * FROM order_items WHERE order_id=? AND deleted_at IS NULL ORDER BY line_no", (order_id,)
        ).fetchall()
        if not row or not items:
            raise RuntimeError("order/items not found")
        if row["status"] == "printed":
            return {
                "started": False,
                "message": "이 오더는 이미 완료입니다 (중복 실행 방지). 다시 실행하려면 먼저 상태를 확인하세요.",
            }
        if row["status"] == "working":
            if _portal_pids_for_order(row["order_no"]):
                return {
                    "started": False,
                    "message": "이 오더는 이미 처리 중입니다 (중복 실행 방지). 다시 실행하려면 먼저 상태를 확인하세요.",
                }
            add_event(con, order_id, "이전 Portal 실행이 중단된 것으로 보임 (해당 프로세스 없음) - 재실행 허용")
        # Bloomberg Portal only ever has an outbound delivery (OBD) record for
        # the order's 배송(delivery) leg - build_excel_rows_zrma() in
        # zrma_handler.py deliberately leaves obd blank for 회수(collection)
        # items, since a pickup/return unit was never shipped through this
        # delivery in the first place. Sending a 회수 item's serial through
        # portal_register_serial.py means asking the delivery's page to find a
        # row/quantity slot that doesn't exist for it - it will genuinely find
        # nothing to match, not fail transiently. Real incident (2026-08-12,
        # order 67074842/ZRX): 배송 leg's serial registered correctly and the
        # delivery completed, then the loop moved on to the 회수 leg's serial
        # and crashed with "material ... not found" because that delivery
        # only ever had 1 unit's worth of row to begin with. Filtering here
        # keeps 회수 items out of what's sent to the portal entirely - their
        # serial is already captured via SAP/Excel/kakao at collection time,
        # they have no Portal-side action left to do.
        order_dict = dict(row)
        item_dicts = [dict(i) for i in items]
        total_items = len(item_dicts)
        material_items = []
        material_item_index = {}  # material -> index into material_items, for same-material merge below
        skipped_return_items = []
        for idx, item in enumerate(item_dicts):
            leg = resolve_item_type(order_dict, item, idx, total_items)
            entry = {
                "material": item["material"],
                "qty": str(item["qty"] or 1),
                "serials": [item["serial"]] if item["serial"] else [],
                # 2026-08-18: serial 칸이 비어있으면 그 자체로 "시리얼 없는
                # 품목"으로 취급 (import 시점 mode 휴리스미틱 - qty>1 신호 -
                # 에 의존하지 않음, ZINP 551059036 qty=1 케이스가 놓쳤던
                # 자리). 사용자 워크플로우: Portal 화면에 시리얼 입력칸
                # 자체가 없는 품목은 workbench에도 애초에 serial을 안 적음.
                # run_zrec_prepare()의 non_serial 판정과 같은 패턴.
                "no_serial": item["mode"] == "material_only" or not item["serial"],
                "description": item["description"] or "",
            }
            if leg == "회수":
                skipped_return_items.append(entry)
                continue
            # SAP combines multiple units of the same material into ONE
            # outbound delivery line, so Portal shows a single Pick Quantity
            # row per material - not one row per unit. Real incident
            # (2026-08-18, order 67842970/ZOR, 일반키보드 x2): each row was
            # sent to portal_register_serial.py as its own qty=1 entry, so
            # the first entry filled Pick Quantity=1 (should've been 2) and
            # the second entry then looked for a *second* DOM row for that
            # material - which never existed on the combined-line page - and
            # errored. Merging same-material rows here into one entry with
            # summed qty and a combined serials list matches what the Portal
            # page actually expects: enter total qty once, then add each
            # serial in turn (Serial 입력 → + → Serial 입력 → +...).
            existing_idx = material_item_index.get(entry["material"])
            if existing_idx is not None and entry["material"]:
                existing = material_items[existing_idx]
                existing["qty"] = str(int(existing["qty"] or 1) + int(entry["qty"] or 1))
                existing["serials"].extend(entry["serials"])
                existing["no_serial"] = existing["no_serial"] and entry["no_serial"]
                continue
            if entry["material"]:
                material_item_index[entry["material"]] = len(material_items)
            material_items.append(entry)

        if not material_items:
            add_event(
                con,
                order_id,
                "Portal 등록 생략: 회수(collection) 전용 오더 - Bloomberg Portal에 등록할 배송 항목 없음",
            )
            return {
                "started": False,
                "message": "이 오더는 회수(수거) 전용이라 Bloomberg Portal에 등록할 배송 항목이 없습니다.",
            }
        if skipped_return_items:
            add_event(
                con,
                order_id,
                f"Portal 등록에서 회수(collection) 품목 {len(skipped_return_items)}건 제외 "
                "(해당 품목은 Portal 배송 등록 대상이 아님)",
            )

        material_only = all(i["no_serial"] for i in material_items)
        total_qty = sum(int(i["qty"] or 1) for i in material_items)
        first = material_items[0]
        args = [
            sys.executable,
            str(BASE_DIR / "portal_ship_and_print.py"),
            "--order", row["order_no"],
            "--delivery", row["delivery_no"] or "",
            "--material", first["material"] or "",
            "--qty", str(total_qty or 1),
            "--items-json", json.dumps(material_items, ensure_ascii=False),
        ]
        if not material_only:
            first_serial = next((i["serials"][0] for i in material_items if i["serials"]), "")
            if first_serial:
                args.extend(["--serial", first_serial])
        else:
            args.append("--material-only")
        if skip_qr:
            args.append("--skip-qr")
        delivery_no_was_blank = not (row["delivery_no"] or "").strip()
        # portal_error is cleared the instant a fresh run STARTS, not only on
        # success - so retrying an errored order clears its orange highlight
        # right away instead of leaving it lit until the retry also finishes.
        con.execute(
            "UPDATE orders SET status='working', portal_error='', updated_at=? WHERE id=?",
            (now_text(), order_id),
        )
        add_event(con, order_id, "Portal automation started: " + " ".join(args[1:]))

    # Output is captured (not sent to a visible console) specifically so a
    # failure reason is recoverable at all from this path: launched via
    # CREATE_NEW_CONSOLE alone, a crashing script's traceback would print to
    # a console window that closes the instant the process exits (nothing
    # keeps it open), so the *only* trace of why it failed was automation.log
    # - and only for whatever the script bothered to log via logger.error()
    # before dying, not any uncaught exception. Capturing here means the
    # actual failure text reaches the board (events / run_portal's caller)
    # regardless of whether that log call exists.
    # PYTHONIOENCODING forces the child's stdout to actual UTF-8 bytes -
    # without it, a redirected (non-console) stdout on Windows falls back to
    # the system codepage (e.g. cp949), and portal_ship_and_print.py never
    # reconfigures its own stdout encoding the way main.py/order.py do, so
    # decoding that pipe as "utf-8" here would silently mangle every Korean
    # message into mojibake.
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.Popen(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        creationflags=subprocess.CREATE_NO_WINDOW,
    )

    def watcher():
        output = proc.stdout.read()
        code = proc.wait()
        with connect() as con:
            if code == 0:
                con.execute("UPDATE orders SET status='printed', updated_at=? WHERE id=?", (now_text(), order_id))
                add_event(con, order_id, "Portal automation completed")
                # Real gap hit live (2026-08-12, order 67075418/ZRX): when the
                # ZRMA order was collected, its outbound delivery didn't exist
                # in SAP/VL06O yet, so obd_map had nothing for it and
                # orders.delivery_no was left blank. portal_ship_and_print.py
                # still found the real delivery (by then created) via its own
                # live SAP lookup and ran successfully - but that discovery
                # never made it back into workbench.db, so the later POD
                # (배송완료) button failed with "Delivery#/OBD가 없습니다"
                # (_pod_args() requires a non-blank delivery_no). Backfilling
                # here, only when the column started out blank, closes that
                # gap without risking overwriting an already-correct value.
                if delivery_no_was_blank:
                    match = re.search(r"^\s*Delivery:\s*(\S+)", output, re.MULTILINE)
                    found_delivery = (match.group(1) if match else "").strip()
                    if found_delivery:
                        con.execute(
                            "UPDATE orders SET delivery_no=?, updated_at=? WHERE id=? AND COALESCE(delivery_no,'')=''",
                            (found_delivery, now_text(), order_id),
                        )
                        add_event(con, order_id, f"Delivery# 자동 보정: (비어있음) → {found_delivery}")
            else:
                # First line is a clean one-sentence summary (for a toast to
                # show at a glance); the rest is the fuller tail for anyone
                # digging further. Prefers the script's own "[실패] ..." line
                # (portal_ship_and_print.py prints this for exactly this
                # purpose) over just "last line printed", which is often a
                # SAP-log noise line rather than the actual reason.
                lines = [ln for ln in output.strip().splitlines() if ln.strip()]
                summary = next((ln for ln in reversed(lines) if ln.startswith("[실패]")), None) or (lines[-1] if lines else "(출력 없음)")
                detail = "\n".join(lines[-15:])
                con.execute(
                    "UPDATE orders SET status='error', portal_error=?, updated_at=? WHERE id=?",
                    (f"{summary}\n{detail}", now_text(), order_id),
                )
                add_event(con, order_id, f"Portal automation failed: exit {code} - {summary}\n{detail}", "error")

    threading.Thread(target=watcher, daemon=True).start()
    return {"started": True, "pid": proc.pid}


POD_PENDING_FILE = BASE_DIR / "pod_pending.json"


def _pod_pids_for_delivery(delivery_no):
    """Live-process check for portal_update_pod.py runs against this
    delivery#, same reasoning/technique as _portal_pids_for_order() - a
    process-list check instead of a DB flag, so nothing goes stale if
    workbench itself restarts mid-run. Deliberately NOT gated through
    orders.status like run_portal() is: that column is the Serial/QR
    pipeline's own new/working/posted/printed/error state machine, and POD
    (delivered-marking) is a separate, later workflow step - writing to it
    here would corrupt run_portal()'s own duplicate-run guard."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'portal_update_pod\\.py' -and "
        f"$_.CommandLine -match '{re.escape(str(delivery_no))}' }} | "
        "Select-Object -ExpandProperty ProcessId"
    )
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL, text=True, timeout=20,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        return []
    return [int(p.strip()) for p in out.splitlines() if p.strip().isdigit()]


def _pod_items_data(con, item_ids):
    """description/serial/qty/item_type for exactly the given item ids - this
    is the workbench.db data passed to portal_update_pod.py's --items-json
    bypass (see prepare_values()'s items_override) instead of letting it
    re-derive everything from Excel. Deliberately not filtered by a single
    order_id (see _resolve_pod_order() for why: a 배송/회수 pair can be two
    DIFFERENT `orders` DB rows that only share a rendered label on screen)."""
    if not item_ids:
        return []
    placeholders = ",".join("?" for _ in item_ids)
    items = con.execute(f"SELECT * FROM order_items WHERE id IN ({placeholders})", item_ids).fetchall()
    return [
        {
            "description": i["description"] or "",
            "serial": i["serial"] or "",
            "qty": i["qty"] or 1,
            "item_type": i["item_type"] or "",
        }
        for i in items
    ]


def _resolve_pod_order(con, order_ids):
    """Real bug hit live (2026-08-03): a single board block that LOOKS like
    one order (same rendered "ZRX 12345" label, one shared checkbox-select-
    all) can actually be TWO separate `orders` DB rows - Excel sometimes
    imports a ZRX's 배송 and 회수 legs as two distinct order records that
    groupRows() only unifies visually by matching orderLabel text (same
    mechanism the merge feature and the stale-duplicate-order issue from
    earlier in this project both already hit). Checking both rows and firing
    POD per raw `orders.id` therefore split into two incomplete runs - one
    order (with a real delivery_no) got "delivered" with no pickup info, the
    other (delivery_no empty, since only one Bloomberg Portal delivery
    record actually exists for the whole ZRX movement) failed outright.
    Fix: resolve the WHOLE checked label-group to the single `orders` row
    that actually carries a delivery_no - that's the one real Portal
    delivery both legs are filed against. Raises if the group spans more
    than one distinct order_no (a real mixed-order selection, not this
    pattern) or more than one DIFFERENT non-empty delivery_no (genuine
    ambiguity - not something to silently guess)."""
    if not order_ids:
        raise RuntimeError("order not found")
    placeholders = ",".join("?" for _ in order_ids)
    rows = con.execute(f"SELECT * FROM orders WHERE id IN ({placeholders})", order_ids).fetchall()
    if not rows:
        raise RuntimeError("order not found")
    order_nos = {r["order_no"] for r in rows}
    if len(order_nos) > 1:
        raise RuntimeError("체크된 행이 서로 다른 오더번호에 걸쳐 있습니다: " + ", ".join(sorted(order_nos)))
    with_delivery = [r for r in rows if r["delivery_no"]]
    delivery_nos = {r["delivery_no"] for r in with_delivery}
    if len(delivery_nos) > 1:
        raise RuntimeError("체크된 행들의 Delivery#가 서로 다릅니다: " + ", ".join(sorted(delivery_nos)) + " - 확인이 필요합니다.")
    primary = with_delivery[0] if with_delivery else rows[0]
    customer = next((r["customer"] for r in rows if r["customer"]), "")
    return primary, customer


def _pod_args(order_row, customer, items, signed_by="", dt="", remarks="", pickup_done=False):
    if not order_row["delivery_no"]:
        raise RuntimeError("이 오더에는 Delivery#/OBD가 없습니다 (workbench에 값이 비어 있음)")
    if not items:
        raise RuntimeError("POD 처리할 항목이 없습니다 (체크된 행을 확인하세요)")
    args = [
        sys.executable, str(BASE_DIR / "portal_update_pod.py"),
        "--delivery", order_row["delivery_no"],
        "--order", order_row["order_no"],
        "--items-json", json.dumps(items, ensure_ascii=False),
    ]
    resolved_signed_by = signed_by or (customer or "")
    if resolved_signed_by:
        args.extend(["--signed-by", resolved_signed_by])
    if dt:
        args.extend(["--datetime", dt])
    if remarks:
        args.extend(["--remarks", remarks])
    if pickup_done:
        args.append("--pickup-done")
    return args


def pod_preview(order_ids, signed_by="", dt="", remarks="", pickup_done=False, item_ids=None):
    """POD (Proof of Delivery / "mark as delivered") preview - ported from
    launcher.py's "POD 확인/저장 선택" button. Runs portal_update_pod.py
    --preview, which only does an item lookup (workbench.db via
    --items-json, not Excel - see prepare_values()) and writes
    pod_pending.json; it never touches Chrome/Portal at all, so this can run
    synchronously (fast) instead of needing the spawn-and-poll pattern
    run_portal()/run_pod_update() use for the actual browser-touching step.
    `order_ids` is every distinct orders.id among the checked rows (see
    _resolve_pod_order() for why there can be more than one)."""
    ensure_db()
    with connect() as con:
        primary, customer = _resolve_pod_order(con, order_ids)
        items = _pod_items_data(con, item_ids or [])
        args = _pod_args(primary, customer, items, signed_by, dt, remarks, pickup_done) + ["--preview"]

    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if result.returncode != 0:
        lines = [ln for ln in (result.stdout or "").strip().splitlines() if ln.strip()]
        summary = next((ln for ln in reversed(lines) if ln.startswith("[실패]")), None) or (lines[-1] if lines else "미리보기 실패")
        raise RuntimeError(summary)
    if not POD_PENDING_FILE.exists():
        raise RuntimeError("pod_pending.json이 생성되지 않았습니다.")
    payload = json.loads(POD_PENDING_FILE.read_text(encoding="utf-8"))
    payload["podOrderId"] = primary["id"]
    return payload


def run_pod_update(order_ids, signed_by="", dt="", remarks="", pickup_done=False, item_ids=None):
    """The real "click Update POD" step - ported from launcher.py's "POD
    최종 저장" button. Sources order#/delivery#/items directly from
    workbench.db (already resolved at import/edit time) rather than
    re-deriving from Excel by order number alone - portal_update_pod.py's own
    Excel lookup only searches TODAY's dated sheet and would silently miss
    (or misclassify) a leg that's on a different day, and also can't tell
    which specific rows the user actually checked. Passing --items-json
    (built from item_ids, the exact checked rows) sidesteps both: the right
    delivery# is always used, and the remarks text reflects exactly what was
    checked, not whatever Excel happens to contain. `order_ids` is every
    distinct orders.id among the checked rows - see _resolve_pod_order()."""
    ensure_db()
    resolved_item_ids = item_ids or []
    with connect() as con:
        primary, customer = _resolve_pod_order(con, order_ids)
        if _pod_pids_for_delivery(primary["delivery_no"]):
            return {"started": False, "message": "이 오더의 POD 처리가 이미 실행 중입니다 (중복 실행 방지)."}
        items = _pod_items_data(con, resolved_item_ids)
        args = _pod_args(primary, customer, items, signed_by, dt, remarks, pickup_done) + ["--update"]
        for oid in order_ids:
            add_event(con, oid, "POD 처리 시작: " + " ".join(args[2:]))
        # pod_error 초기화 - run_portal()의 portal_error와 같은 이유: 재시도가
        # "시작"되는 순간 orange 표시가 사라지도록, 성공을 기다리지 않고 지금 지움.
        placeholders = ",".join("?" for _ in order_ids)
        con.execute(f"UPDATE orders SET pod_error='' WHERE id IN ({placeholders})", order_ids)

    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.Popen(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        creationflags=subprocess.CREATE_NO_WINDOW,
    )

    def watcher():
        output = proc.stdout.read()
        code = proc.wait()
        pod_success = (code == 0)
        with connect() as con:
            if pod_success:
                for oid in order_ids:
                    add_event(con, oid, "POD 처리 완료 (Update POD 클릭됨)")
            else:
                lines = [ln for ln in output.strip().splitlines() if ln.strip()]
                summary = next((ln for ln in reversed(lines) if ln.startswith("[실패]")), None) or (lines[-1] if lines else "(출력 없음)")
                detail = "\n".join(lines[-15:])
                for oid in order_ids:
                    add_event(con, oid, f"POD 처리 실패: exit {code} - {summary}\n{detail}", "error")
                placeholders = ",".join("?" for _ in order_ids)
                con.execute(
                    f"UPDATE orders SET pod_error=? WHERE id IN ({placeholders})",
                    [f"{summary}\n{detail}", *order_ids],
                )
        # 포털에서 POD가 정상 처리된 것까지 확인됐을 때(에러 배너 없이 Update
        # POD 클릭 성공, code==0 - portal_update_pod.py가 클릭 후
        # _collect_errors()로 실제 확인함)만 여기 온다. 사용자 요청,
        # 2026-08-14: 배송 행은 Done(파랑), 회수 행은 Ready to process(빨강)로
        # 자동 색칠 - processDone()/processPickupReady() 버튼을 수동으로 누르는
        # 것과 동일한 board_status 값. item_type을 직접 신뢰하고(회수
        # 아이템은 원래 OBD가 비어 있어 POD 자체가 안 걸리므로 여기 들어오는
        # 건 사실상 항상 배송이지만, 혹시 모를 회수 항목도 안전하게 구분 -
        # ZRX의 fallback 반반 나누기 휴리스틱은 절대 안 씀, run_portal()의
        # 2026-08-12 사고와 같은 이유로 item_type 원본 값만 신뢰).
        # set_items_board_status()가 자기 자신의 connect()를 새로 여므로, 위
        # add_event용 connect()의 트랜잭션이 커밋되어 닫힌 뒤에(with 블록
        # 밖에서) 실행 - 같은 DB 파일에 쓰기 커넥션 두 개를 동시에 열어두면
        # SQLite가 잠길 수 있어서 순서를 분리함.
        if pod_success and resolved_item_ids:
            with connect() as con:
                placeholders = ",".join("?" for _ in resolved_item_ids)
                type_rows = con.execute(
                    f"SELECT id, item_type FROM order_items WHERE id IN ({placeholders})",
                    resolved_item_ids,
                ).fetchall()
            pickup_ids = [r["id"] for r in type_rows if (r["item_type"] or "").strip() == "회수"]
            delivery_ids = [r["id"] for r in type_rows if (r["item_type"] or "").strip() != "회수"]
            if delivery_ids:
                set_items_board_status(delivery_ids, "done")
            if pickup_ids:
                set_items_board_status(pickup_ids, "pending")

    threading.Thread(target=watcher, daemon=True).start()
    return {"started": True, "pid": proc.pid, "podOrderId": primary["id"]}


def pid_running(pid):
    try:
        out = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command",
             f"Get-Process -Id {int(pid)} -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id"],
            stderr=subprocess.DEVNULL, text=True, timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        return False
    return bool(out.strip())


# ── Dashboard board rendering ────────────────────────────────────────────
#
# STRUCTURE-ONLY PASS: SAP/Portal/Process/Excel control panel + one
# continuous table (sticky header, dates stacked downward as you scroll).
# The board itself is rendered entirely client-side from a JSON row list
# (__BOARD_DATA__ below) rather than server-built HTML strings, because
# merging same-order rows, inline text editing, and delete all need to
# recompute the table's grouping/rowspans after every action - far easier
# to redo from one in-memory row list (JS `ROWS`) than to patch rowspans
# into hand-built HTML piecemeal. Button handlers and moves are still
# client-side placeholders for now (see JS `toast(...)` calls) - wiring
# them to real actions (portal automation, Excel export, persisting a
# date move / edit / delete) is the next step, deliberately kept separate
# from this layout pass.

_WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]


# infer_item_type()/resolve_item_type() moved to order_domain.py (2026-09-08,
# see its own module docstring) so quick_courier.py can reuse the same
# 배송/회수 판정 without importing this whole http.server app.



"""Local operations workbench for SAP/Bloomberg delivery automation.

This app moves day-to-day work state out of the shipping Excel file and into a
small SQLite database. Excel is treated as an import source, not the runtime
source of truth for portal execution.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from config import EXCEL_PATH

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "workbench.db"
LOG_FILE = BASE_DIR / "automation.log"

ORDER_RE = re.compile(r"\b(ZOR|ZRX|SDSK|ZINP|ZINT|ZRE|ORD)\s*[:#-]?\s*(\d{7,})\b", re.I)
ANY_NUM_RE = re.compile(r"\b\d{7,}\b")
OBD_RE = re.compile(r"\bOBD\s*[:#-]?\s*(\d{7,})\b", re.I)
QTY_RE = re.compile(r"\bQty\s*[:#-]?\s*(\d+)\b|\bx\s*(\d+)\b", re.I)

# A single day-sheet (e.g. "7-27") actually contains several date blocks stacked
# inside it: a lone "2026년 7월 27일 월요일"-style row in column G, then an
# "Order # / item / M/N / S/N / customer / phone / ADDRESS / memo" header row,
# then that date's orders - repeated for each date. DATE_HEADER_RE finds those
# block boundaries; TYPE_PREFIX_RE reads the real 배송/회수 marker that sits at
# the front of column A on every item row.
DATE_HEADER_RE = re.compile(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")
TYPE_PREFIX_RE = re.compile(r"^\s*(배송|회수)\b")


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_sheet_names(sheetnames, today=None):
    today = today or date.today()
    candidates = [f"{today.month}-{today.day}", f"{today.month}.{today.day}", f"{today.month}_{today.day}"]
    return [name for name in candidates if name in sheetnames]


def digits(value) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def text_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def extract_qty(description) -> int:
    match = QTY_RE.search(str(description or ""))
    if not match:
        return 1
    try:
        return max(1, int(match.group(1) or match.group(2)))
    except Exception:
        return 1


def parse_order_text(value):
    text = str(value or "")
    matches = ORDER_RE.findall(text)
    order_type = matches[0][0].upper() if matches else ""
    order_no = matches[0][1] if matches else ""
    if not order_no:
        nums = ANY_NUM_RE.findall(text)
        order_no = nums[0] if nums else ""
    obd_match = OBD_RE.search(text)
    delivery_no = obd_match.group(1) if obd_match else ""
    return order_type, order_no, delivery_no


def parse_item_type(order_text):
    """배송/회수 from the front of column A on an item row, e.g. '배송 ZRX 67059008 | ...'."""
    m = TYPE_PREFIX_RE.match(str(order_text or ""))
    return m.group(1) if m else ""


def row_date_header(ws, row):
    """
    If `row` is one of the sheet's embedded date-block banners (column G holds
    only a "2026년 7월 27일 월요일"-style string, every other column is blank),
    return its ISO date. Otherwise None. Checked narrowly (all other columns
    empty) so a memo like "8/22로 연기됨" inside a normal item row is never
    mistaken for a new date block.
    """
    g_val = text_value(ws.cell(row, 7).value)
    m = DATE_HEADER_RE.match(g_val)
    if not m:
        return None
    for col in (1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 13, 14):
        if text_value(ws.cell(row, col).value):
            return None
    year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return None


def ensure_db():
    with sqlite3.connect(DB_PATH) as con:
        con.execute("PRAGMA journal_mode=WAL")
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS orders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_no TEXT NOT NULL,
                delivery_no TEXT,
                order_type TEXT,
                customer TEXT,
                phone TEXT,
                address TEXT,
                memo TEXT,
                status TEXT NOT NULL DEFAULT 'new',
                source_sheet TEXT,
                source_row_start INTEGER,
                source_row_end INTEGER,
                source_date TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(order_no, delivery_no)
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS order_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                line_no INTEGER NOT NULL,
                description TEXT,
                material TEXT,
                serial TEXT,
                qty INTEGER NOT NULL DEFAULT 1,
                mode TEXT NOT NULL DEFAULT 'serial',
                item_type TEXT NOT NULL DEFAULT '',
                FOREIGN KEY(order_id) REFERENCES orders(id) ON DELETE CASCADE
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER,
                level TEXT NOT NULL DEFAULT 'info',
                message TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        # Migration for DBs created before item_type existed.
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN item_type TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass


def connect():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    return con


def add_event(con, order_id, message, level="info"):
    con.execute(
        "INSERT INTO events(order_id, level, message, created_at) VALUES (?, ?, ?, ?)",
        (order_id, level, message, now_text()),
    )


def load_today_from_excel():
    ensure_db()
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    imported = 0
    sheets = today_sheet_names(wb.sheetnames)
    if not sheets:
        wb.close()
        raise RuntimeError(f"today sheet not found in {EXCEL_PATH}")

    try:
        with connect() as con:
            for sheet_name in sheets:
                ws = wb[sheet_name]
                current = None
                current_date_iso = date.today().isoformat()  # until the first date-block header is seen
                groups = []
                for row in range(1, ws.max_row + 1):
                    header_date = row_date_header(ws, row)
                    if header_date:
                        current_date_iso = header_date
                        current = None  # next order line starts a fresh group under this date
                        continue

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
                    if existing:
                        order_id = existing["id"]
                        con.execute(
                            """
                            UPDATE orders SET order_type=?, customer=?, phone=?, address=?, memo=?,
                                source_sheet=?, source_row_start=?, source_row_end=?, source_date=?, updated_at=?
                            WHERE id=?
                            """,
                            (
                                group.get("order_type"), group.get("customer"), group.get("phone"),
                                group.get("address"), group.get("memo"), group.get("source_sheet"),
                                group.get("source_row_start"), group.get("source_row_end"), group_date,
                                now_text(), order_id,
                            ),
                        )
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
                        con.execute(
                            """
                            INSERT INTO order_items(order_id, line_no, description, material, serial, qty, mode, item_type)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (
                                order_id, idx, item["description"], item["material"], item["serial"],
                                item["qty"], item["mode"], item.get("item_type", ""),
                            ),
                        )
                    add_event(con, order_id, f"Imported from Excel sheet {group['source_sheet']} rows {group['source_row_start']}-{group['source_row_end']}")
                    imported += 1
    finally:
        wb.close()
    return {"imported": imported, "sheets": sheets}


EXPORT_HEADER = ["Order #", "item", "M/N", "S/N", "customer", "phone", "ADDRESS", "memo"]

_KOREAN_WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]

# Column widths and per-column cell styling, matched against an existing
# hand-formatted sheet (verified against "7-27": header fill is the theme
# accent2 color tinted +0.8, or "FBE3D6" computed out to RGB; data-row fill
# is the plain blue "00B0F0" Excel uses for "done" rows). Column A carries
# its own item_type/order_label/delivery_no line on every row (never blank
# for continuation - the original sheets always repeat it), while columns
# E-H (customer/phone/address/memo) are merged down the whole customer
# group, same as the original sheets build via SAP collection.
_EXPORT_COL_WIDTHS = {1: 23.375, 2: 40.0, 3: 13.75, 4: 14.5, 5: 16.625, 6: 19.875, 7: 47.75, 8: 25.125}
_EXPORT_HEADER_FILL = PatternFill(fill_type="solid", fgColor="FBE3D6")
_EXPORT_DATA_FILL = PatternFill(fill_type="solid", fgColor="00B0F0")
_EXPORT_THIN_BORDER = Border(*(Side(style="thin"),) * 4)
_EXPORT_COL_ALIGN = {
    1: Alignment(vertical="top", wrap_text=True),
    2: Alignment(horizontal="left", vertical="center", wrap_text=True),
    3: Alignment(horizontal="center", vertical="center"),
    4: Alignment(horizontal="center", vertical="center"),
    5: Alignment(horizontal="center", vertical="center"),
    6: Alignment(horizontal="center", vertical="center", wrap_text=True),
    7: Alignment(horizontal="center", vertical="center", wrap_text=True),
    8: Alignment(horizontal="center", vertical="top", wrap_text=True),
}


def _unique_sheet_name(wb, base_name):
    """
    Never overwrites an existing sheet - if `base_name` (e.g. "7-29") is
    already taken, appends " (2)", " (3)", ... same convention Excel itself
    uses for "copy of this sheet", until a free name is found.
    """
    if base_name not in wb.sheetnames:
        return base_name
    n = 2
    while f"{base_name} ({n})" in wb.sheetnames:
        n += 1
    return f"{base_name} ({n})"


def _order_cell_text(row):
    order_cell = str(row.get("itemType") or "").strip()
    label = str(row.get("orderLabel") or "").strip()
    if label:
        order_cell = f"{order_cell} {label}".strip()
    codes = str(row.get("deliveryNo") or "").strip()
    if codes:
        order_cell = f"{order_cell}\n{codes}" if order_cell else codes
    return order_cell


def _write_export_sheet(wb, sheet_name, iso, rows):
    d = date.fromisoformat(iso)
    ws = wb.create_sheet(sheet_name)

    for col, width in _EXPORT_COL_WIDTHS.items():
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width

    # Row 1: date banner (bold, only the ADDRESS column carries the text,
    # same position as the original sheets).
    date_cell = ws.cell(row=1, column=7)
    date_cell.value = f"{d.year}년 {d.month}월 {d.day}일 {_KOREAN_WEEKDAYS[d.weekday()]}요일"
    date_cell.font = Font(bold=True)
    date_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Row 2: column headers.
    for col, label in enumerate(EXPORT_HEADER, start=1):
        cell = ws.cell(row=2, column=col)
        cell.value = label
        cell.fill = _EXPORT_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _EXPORT_THIN_BORDER

    # Data rows, grouped by isGroupFirst so E-H (customer/phone/address/memo)
    # can be merged down the whole group like the original sheets do.
    row_cursor = 3
    i = 0
    n = len(rows)
    while i < n:
        j = i + 1
        while j < n and not rows[j].get("isGroupFirst"):
            j += 1
        group = rows[i:j]
        group_start = row_cursor

        for offset, row in enumerate(group):
            r = row_cursor + offset
            values = {
                1: _order_cell_text(row),
                2: row.get("description") or "",
                3: row.get("material") or "",
                4: row.get("serial") or "",
                5: row.get("customer") or "" if offset == 0 else "",
                6: row.get("phone") or "" if offset == 0 else "",
                7: row.get("address") or "" if offset == 0 else "",
                8: row.get("memo") or "" if offset == 0 else "",
            }
            for col in range(1, 9):
                cell = ws.cell(row=r, column=col)
                cell.value = values[col]
                cell.fill = _EXPORT_DATA_FILL
                cell.alignment = _EXPORT_COL_ALIGN[col]
                cell.border = _EXPORT_THIN_BORDER
                if col == 8:
                    cell.font = Font(bold=True)

        group_end = row_cursor + len(group) - 1
        if group_end > group_start:
            for col in (5, 6, 7, 8):
                ws.merge_cells(start_row=group_start, start_column=col, end_row=group_end, end_column=col)

        row_cursor = group_end + 1
        i = j

    return ws, row_cursor - 1


def export_dates_to_excel(dates_payload):
    """
    Writes the board's current (client-side) state for the given dates into
    the live shipping workbook - one NEW sheet per date, named the same
    "{month}-{day}" way `today_sheet_names()` looks for them (e.g. "7-29"),
    or "7-29 (2)", "7-29 (3)", ... if that name is already taken - existing
    sheets are never deleted or overwritten (the board's data still comes
    from the client's in-memory `ROWS`, not workbench.db, since nothing this
    session persists there yet). Formatting (column widths, header fill,
    blue "done" fill, borders, merged customer/phone/address/memo per
    group) matches the hand-built sheets SAP collection produces.

    Takes a full backup of the workbook before writing - this is the one
    action in this app that touches the real, live Excel file (everything
    else here is screen-only), so a mistake here is not casually undoable
    the way an in-browser action is.
    """
    excel_path = Path(EXCEL_PATH)
    if not excel_path.exists():
        raise RuntimeError(f"Excel file not found: {excel_path}")

    backup_path = excel_path.with_name(
        f"{excel_path.stem}.backup_before_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}{excel_path.suffix}"
    )
    shutil.copy2(excel_path, backup_path)

    wb = openpyxl.load_workbook(excel_path)
    try:
        result = {}
        for iso, rows in dates_payload.items():
            d = date.fromisoformat(iso)
            sheet_name = _unique_sheet_name(wb, f"{d.month}-{d.day}")
            _write_export_sheet(wb, sheet_name, iso, rows)
            result[iso] = {"sheet": sheet_name, "rows": len(rows)}
        wb.save(excel_path)
        return {"exported": result, "backup": str(backup_path)}
    finally:
        wb.close()


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
            items = con.execute("SELECT * FROM order_items WHERE order_id=? ORDER BY line_no", (row["id"],)).fetchall()
            result.append(order_payload(row, items))
        return result


def get_order(order_id):
    with connect() as con:
        row = con.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        if not row:
            return None
        items = con.execute("SELECT * FROM order_items WHERE order_id=? ORDER BY line_no", (order_id,)).fetchall()
        events = con.execute("SELECT * FROM events WHERE order_id=? ORDER BY id DESC LIMIT 80", (order_id,)).fetchall()
        data = order_payload(row, items)
        data["events"] = [dict(e) for e in events]
        return data


def run_portal(order_id):
    ensure_db()
    with connect() as con:
        row = con.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        items = con.execute("SELECT * FROM order_items WHERE order_id=? ORDER BY line_no", (order_id,)).fetchall()
        if not row or not items:
            raise RuntimeError("order/items not found")
        material_items = []
        for item in items:
            material_items.append(
                {
                    "material": item["material"],
                    "qty": str(item["qty"] or 1),
                    "serials": [item["serial"]] if item["serial"] else [],
                    "no_serial": item["mode"] == "material_only",
                    "description": item["description"] or "",
                }
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
        con.execute("UPDATE orders SET status='working', updated_at=? WHERE id=?", (now_text(), order_id))
        add_event(con, order_id, "Portal automation started: " + " ".join(args[1:]))

    proc = subprocess.Popen(args, cwd=str(BASE_DIR), creationflags=0x00000010 if os.name == "nt" else 0)

    def watcher():
        code = proc.wait()
        with connect() as con:
            if code == 0:
                con.execute("UPDATE orders SET status='printed', updated_at=? WHERE id=?", (now_text(), order_id))
                add_event(con, order_id, "Portal automation completed")
            else:
                con.execute("UPDATE orders SET status='error', updated_at=? WHERE id=?", (now_text(), order_id))
                add_event(con, order_id, f"Portal automation failed: exit {code}", "error")

    threading.Thread(target=watcher, daemon=True).start()
    return {"started": True, "pid": proc.pid}


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


def infer_item_type(order_type, idx, total):
    """
    배송/회수 fallback heuristic, used only for legacy rows imported before
    load_today_from_excel started capturing the real per-row marker from
    column A (item_type). ZRE-only orders are pure pickups, everything else
    defaults to delivery, and ZRX (exchange - genuinely carries both) is
    split in half just so the layout has something to show. Do not trust
    ZRX's split for real portal actions.
    """
    order_type = (order_type or "").upper()
    if order_type == "ZRE":
        return "회수"
    if order_type == "ZRX" and total > 1:
        return "회수" if idx >= (total + 1) // 2 else "배송"
    return "배송"


def resolve_item_type(order, item, idx, total):
    real = str(item.get("item_type") or "").strip()
    if real in ("배송", "회수"):
        return real
    return infer_item_type(order.get("order_type"), idx, total)


def board_orders_for_date(target_date_iso):
    ensure_db()
    with connect() as con:
        rows = con.execute(
            "SELECT * FROM orders WHERE source_date=? ORDER BY id", (target_date_iso,)
        ).fetchall()
        result = []
        for row in rows:
            items = con.execute(
                "SELECT * FROM order_items WHERE order_id=? ORDER BY line_no", (row["id"],)
            ).fetchall()
            result.append(order_payload(row, items))
        return [o for o in result if o["items"]]


def board_row(order, item, idx, total):
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
    return {
        "rowKey": f"{order['id']}-{idx}",
        "orderId": order["id"],
        "orderLabel": order_label,
        # Secondary order-code line(s) shown under the main label, e.g. "OBD 92149447" -
        # the Bloomberg-side delivery code (Outbound Delivery), as distinct from the SAP
        # order code (ZOR/ZRE/ZRX/ZINP/ZINT/...) in orderLabel above. Free-text editable
        # on the client, one code per line, so more lines (ORD/SDSK/etc) can be added.
        "deliveryNo": f"OBD {delivery_no}" if delivery_no else "",
        "itemType": resolve_item_type(order, item, idx, total),
        "description": item.get("description") or "",
        "material": item.get("material") or "",
        "serial": item.get("serial") or "",
        "customer": order.get("customer") or "",
        "phone": order.get("phone") or "",
        "address": order.get("address") or "",
        "memo": order.get("memo") or "",
    }


def all_board_dates():
    """
    Every distinct source_date that has at least one order with items, plus
    today (so today's section always renders as an anchor even if empty).
    A single Excel sheet can hold several embedded date blocks weeks apart
    (see row_date_header) - the board is not just "today vs tomorrow".
    """
    ensure_db()
    with connect() as con:
        rows = con.execute(
            """
            SELECT DISTINCT o.source_date FROM orders o
            JOIN order_items i ON i.order_id = o.id
            WHERE o.source_date IS NOT NULL AND o.source_date != ''
            """
        ).fetchall()
    dates = {r["source_date"] for r in rows}
    dates.add(date.today().isoformat())
    return sorted(dates)


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


def board_data():
    data = []
    for iso in all_board_dates():
        orders = board_orders_for_date(iso)
        rows = []
        for order in orders:
            items = order.get("items") or []
            n = len(items)
            rows.extend(board_row(order, item, idx, n) for idx, item in enumerate(items))
        data.append({"iso": iso, "label": _date_label(date.fromisoformat(iso)), "rows": rows})
    return data


def render_dashboard_page():
    data_json = json.dumps(board_data(), ensure_ascii=False)
    return PAGE_TEMPLATE.replace("__BOARD_DATA__", data_json)


PAGE_TEMPLATE = r'''
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bloomberg Dashboard Seoul</title>
<style>
*{box-sizing:border-box}
:root{
  font-family:-apple-system,"Segoe UI","Malgun Gothic",Arial,sans-serif;
  color:#1e293b;background:#f1f4f9;
  --line:#e2e8f0; --line-strong:#cbd5e1;
}
body{margin:0}

.topbar{height:54px;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;
  display:flex;align-items:center;justify-content:space-between;padding:0 22px;
  box-shadow:0 1px 3px rgba(0,0,0,.15)}
.topbar h1{font-size:17px;margin:0;font-weight:600;letter-spacing:.2px}
#clock{font-size:13px;color:#cbd5e1;font-variant-numeric:tabular-nums}

.controls{display:flex;gap:12px;padding:14px 16px 0}
.ctrl-col{flex:1;background:#fff;border:1px solid var(--line);border-radius:10px;
  padding:12px;box-shadow:0 1px 2px rgba(15,23,42,.04)}
.ctrl-col h3{margin:0 0 9px;font-size:12.5px;font-weight:700;color:#64748b;
  text-transform:uppercase;letter-spacing:.5px}

.btn{height:32px;border:1px solid #d1d5db;background:#fff;border-radius:7px;padding:0 12px;
  cursor:pointer;font-size:12.5px;color:#334155;transition:background .12s,box-shadow .12s,transform .05s}
.btn:hover{background:#f8fafc;box-shadow:0 1px 2px rgba(15,23,42,.06)}
.btn:active{transform:translateY(1px)}
.btn.block{display:block;width:100%;text-align:left;height:auto;padding:9px 12px;
  margin-bottom:7px;line-height:1.45;font-weight:500}
.btn.primary{background:#2563eb;color:#fff;border-color:#2563eb}
.btn.primary:hover{background:#1d4ed8}
.btn.warn{background:#fff7ed;border-color:#fdba74;color:#9a3412}
.btn.warn:hover{background:#ffedd5}
.btn .hint{font-size:10.5px;font-weight:400;opacity:.8;display:block;margin-top:1px}
.ctrl-inline{display:flex;gap:6px;margin-bottom:7px}
.ctrl-inline .btn{flex:1.3;text-align:left}
.ctrl-inline input{flex:1;border:1px solid var(--line-strong);border-radius:7px;padding:0 9px;
  font-size:12.5px;color:#1e293b}
.ctrl-inline input:focus{outline:2px solid #93c5fd;outline-offset:-1px}

.board-wrap{margin:16px;background:#fff;border:1px solid var(--line);border-radius:10px;
  box-shadow:0 1px 2px rgba(15,23,42,.04)}
table.board-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12.5px}
.board-table thead th{position:sticky;top:0;z-index:5;background:#e8720c;color:#fff;
  text-align:center;padding:9px 10px;font-weight:600;letter-spacing:.2px;
  box-shadow:0 1px 0 rgba(0,0,0,.08)}
.board-table thead th:first-child{border-top-left-radius:0}

.date-banner-row td{background:linear-gradient(90deg,#1e3a5f,#25507e);color:#fff;
  font-weight:600;font-size:13.5px;padding:9px 12px;letter-spacing:.2px}
.date-banner-inner{display:flex;align-items:center;justify-content:space-between}
.date-banner-right{display:flex;align-items:center;gap:10px}
.date-export-check{width:15px;height:15px;cursor:pointer}
.date-toggle{background:transparent;border:none;color:#fff;font-size:13px;cursor:pointer;
  padding:2px 10px;border-radius:5px;line-height:1.4}
.date-toggle:hover{background:rgba(255,255,255,.18)}
tbody.date-group.collapsed .item-row,tbody.date-group.collapsed .empty-row{display:none}

.item-row td{padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:middle;
  background:#fff;transition:background .12s}
.item-row:hover td{background:#f8fafc}
.item-row[draggable="true"]{cursor:grab}
.item-row.dragging{opacity:.35}
.item-row.row-done td{background:#eff6ff}
.item-row.row-done:hover td{background:#e6f0ff}
.item-row.row-pending td{background:#fef2f2}
.item-row.row-pending:hover td{background:#fde8e8}

.type-cell{text-align:center}
.type-select{font-weight:700;border:1px solid transparent;background:transparent;
  border-radius:5px;padding:2px 4px;cursor:pointer;font-size:12.5px}
.type-select:hover{border-color:var(--line-strong)}
.type-delivery{color:#2563eb}
.type-pickup{color:#dc2626}
.type-other{color:#9333ea}
.ord-cell{font-weight:600}
.ord-cell .muted{display:block;font-weight:400;color:#94a3b8;font-size:11px;margin-top:3px;white-space:pre-wrap}
.cust-cell{color:#334155}
.addr-cell{white-space:pre-wrap}
.addr-cell::first-line{font-weight:700;color:#1e293b}
.check-cell{text-align:center}
.check-cell input{width:15px;height:15px;cursor:pointer}
.handle-cell{text-align:center;cursor:grab}
.drag-handle{display:inline-block;color:#cbd5e1;font-size:15px;cursor:grab;user-select:none}
.drag-handle:hover{color:#94a3b8}

[contenteditable="true"]{outline:none;border-radius:5px;padding:2px 4px;margin:-2px -4px;
  cursor:text;min-height:1.2em}
[contenteditable="true"]:hover{background:#f8fafc}
[contenteditable="true"]:focus{background:#eff6ff;box-shadow:0 0 0 2px #93c5fd}

.board-table thead th:nth-child(1){width:3%}
.board-table thead th:nth-child(2){width:5%}
.board-table thead th:nth-child(3){width:9%}
.board-table thead th:nth-child(4){width:15%}
.board-table thead th:nth-child(5){width:6%}
.board-table thead th:nth-child(6){width:7%}
.board-table thead th:nth-child(7){width:7%}
.board-table thead th:nth-child(8){width:6%}
.board-table thead th:nth-child(9){width:16%}
.board-table thead th:nth-child(10){width:8%}
.board-table thead th:nth-child(11){width:4%}
/* Column width/alignment on body rows is keyed by explicit class, NOT
   nth-child - a merged group's non-first rows omit the rowspan'd Order#/
   Customer/Phone/Address/Memo <td>s entirely (they're covered by the first
   row's rowspan), which shifts every later cell's nth-child index left of
   its true column. nth-child rules here silently applied the WRONG
   column's width/alignment to Item/M-N/S-N/check on every row after a
   group's first - this is what caused the "first row looks right, every
   row after it looks off" misalignment. Classes don't shift with cell
   count, so this class is what's now considered authoritative. */
.item-row td.handle-cell{width:3%}
.item-row td.type-cell{width:5%}
.item-row td.ord-cell{width:9%}
.item-row td.item-cell{width:15%}
.item-row td.mn-cell{width:6%;text-align:center}
.item-row td.sn-cell{width:7%;text-align:center}
.item-row td.customer-cell{width:7%;text-align:center}
.item-row td.phone-cell{width:6%;text-align:center}
.item-row td.addr-cell{width:16%;text-align:center}
.item-row td.memo-cell{width:8%;text-align:center}
.item-row td.check-cell{width:4%}

.empty-row td{padding:22px 10px;text-align:center;color:#94a3b8;font-size:12.5px}
.date-group.drag-over .date-banner-row td{background:linear-gradient(90deg,#2563eb,#3b82f6)}

.toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%) translateY(20px);
  background:#0f172a;color:#fff;padding:10px 18px;border-radius:8px;opacity:0;
  transition:.25s;font-size:13px;z-index:50;max-width:80vw;text-align:center;
  box-shadow:0 6px 20px rgba(0,0,0,.25)}
.toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
</style>
</head>
<body>
<div class="topbar"><h1>Bloomberg Dashboard Seoul</h1><div id="clock"></div></div>

<div class="controls">
  <div class="ctrl-col">
    <h3>SAP</h3>
    <button class="btn block" onclick="runSap('start')">1. SAP Start &amp; Loop</button>
    <button class="btn block" onclick="runSap('all')">2. SAP 전체 조회</button>
    <div class="ctrl-inline">
      <button class="btn" onclick="runOrderNoAction()">3. 오더번호 반영</button>
      <input id="orderNoInput" placeholder="오더번호">
    </div>
    <div class="ctrl-inline">
      <button class="btn" onclick="runNewSessionAction()">4. New SAP session open</button>
      <input id="tcodeInput" placeholder="T-code (선택)">
    </div>
  </div>
  <div class="ctrl-col">
    <h3>Bloomberg Portal</h3>
    <button class="btn block primary" onclick="runPortalBatch()">1. Serial 등록 &amp; QR 인쇄<span class="hint">아래 체크된 행 대상</span></button>
    <button class="btn block" onclick="runLatestZpl()">2. 최근 ZPL 수동 인쇄</button>
  </div>
  <div class="ctrl-col">
    <h3>Process</h3>
    <button class="btn block primary" onclick="processDone()">Done<span class="hint">체크된 행 전체 완료 (파랑)</span></button>
    <button class="btn block warn" onclick="processPickupReady()">Ready to process<span class="hint">체크된 회수만 대기 (빨강), 나머지 완료</span></button>
  </div>
  <div class="ctrl-col">
    <h3>Excel</h3>
    <div class="ctrl-inline">
      <input type="date" id="moveDateInput" onchange="moveCheckedToDate(this.value); this.value='';">
    </div>
    <div class="hint" style="font-size:10.5px;color:#94a3b8;margin:-3px 0 9px">체크된 행을 이 날짜로 이동 (목록에 없는 날짜도 새로 생김)</div>
    <button class="btn block primary" onclick="exportChecked()">1. 선택 날짜 엑셀로 내보내기<span class="hint">날짜 줄의 체크박스로 선택 - 기존 시트는 그대로 두고 새 시트(예: 7-29, 이미 있으면 7-29 (2))에 기록</span></button>
    <button class="btn block warn" onclick="deleteChecked()">2. 체크된 행 삭제<span class="hint">화면에서만 제거 - DB는 그대로</span></button>
  </div>
</div>

<div class="board-wrap">
  <table class="board-table" id="boardTable">
    <thead><tr>
      <th></th><th>Type</th><th>Order #</th><th>Item</th><th>M/N</th><th>S/N</th>
      <th>Customer</th><th>Phone</th><th>Address</th><th>Memo</th><th>check</th>
    </tr></thead>
  </table>
</div>

<script>
function tick(){document.getElementById('clock').textContent=new Date().toLocaleString('ko-KR',{weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'})}
setInterval(tick,1000);tick();

function toast(msg){
  const t=document.createElement('div');
  t.className='toast';
  t.textContent=msg;
  document.body.appendChild(t);
  requestAnimationFrame(()=>t.classList.add('show'));
  setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},1900);
}

// -- Control panel: structure-only placeholders, wired to real actions next --
function runSap(kind){toast('SAP '+kind+' - 다음 단계에서 연결 예정')}
function runOrderNoAction(){const v=document.getElementById('orderNoInput').value.trim();toast('오더번호 반영: '+(v||'(입력 없음)')+' - 다음 단계에서 연결 예정')}
function runNewSessionAction(){const v=document.getElementById('tcodeInput').value.trim();toast('새 SAP 세션 열기: '+(v||'(tcode 없음)')+' - 다음 단계에서 연결 예정')}
function runPortalBatch(){
  const keys=[...document.querySelectorAll('.row-check:checked')];
  if(!keys.length){toast('체크된 행이 없습니다');return}
  toast(keys.length+'건 선택됨 - Portal 자동화 연결은 다음 단계');
}
function runLatestZpl(){toast('최근 ZPL 수동 인쇄 - 다음 단계에서 연결 예정')}

// ── Board data model ─────────────────────────────────────────────────────
// ROWS is the single in-memory source of truth for the whole board (one
// entry per item line). Every action (move, edit, delete, type change,
// Done/Ready) mutates ROWS then calls renderAll() to rebuild the table
// from scratch - simplest way to keep same-order merging (rowspan) correct
// after arbitrary edits, rather than patching DOM rowspans incrementally.
let ROWS = [];
const DATE_LABELS = {};
const COLLAPSED_DATES = new Set();
const EXPORT_DATES = new Set();

// Type dropdown options. BASE_TYPES is fixed; CUSTOM_TYPES accumulates
// whatever the user adds via the dropdown's "+ 추가" entry (session-only,
// like everything else here) - shared across every row's dropdown, not
// just the one it was added from.
const BASE_TYPES = ['배송','회수','Bloomberg','Delayed','회수 Delayed'];
let CUSTOM_TYPES = [];
function allTypeOptions(){ return [...BASE_TYPES, ...CUSTOM_TYPES]; }
function typeClassFor(v){
  if(v==='회수' || v==='회수 Delayed') return 'type-pickup';
  if(v==='배송') return 'type-delivery';
  return 'type-other';
}

(function loadInitialData(){
  const data = __BOARD_DATA__;
  data.forEach(d=>{
    DATE_LABELS[d.iso] = d.label;
    d.rows.forEach(r=>ROWS.push({...r, date: d.iso, uiStatus: ''}));
  });
})();

const WEEKDAY_KR = ['월','화','수','목','금','토','일'];
function formatDateLabel(iso){
  const d=new Date(iso+'T00:00:00');
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  const todayIso=new Date().toLocaleDateString('sv-SE');
  const t=new Date(); t.setDate(t.getDate()+1);
  const tomorrowIso=t.toLocaleDateString('sv-SE');
  let suffix='';
  if(iso===todayIso) suffix=' (Today)';
  else if(iso===tomorrowIso) suffix=' (Tomorrow)';
  return `${y}. ${m}. ${day} ${WEEKDAY_KR[(d.getDay()+6)%7]}요일${suffix}`;
}
function labelFor(iso){
  if(!DATE_LABELS[iso]) DATE_LABELS[iso]=formatDateLabel(iso);
  return DATE_LABELS[iso];
}
function allDatesSorted(){
  const set=new Set(Object.keys(DATE_LABELS));
  ROWS.forEach(r=>set.add(r.date));
  set.add(new Date().toLocaleDateString('sv-SE'));
  return [...set].sort();
}

// -- Same-order merge: rows sharing the same order label (SAP order type +
// number) are grouped into one visual block (rowspan on the shared
// columns) - this is what lets two separate order records for the same
// real-world order (e.g. a delivery row and a pickup row imported from
// Excel under the same order number, like ZRX 67066585's delivery/pickup
// pair) collapse into a single block, even when one of those records is
// missing customer/phone/address/memo (Excel only carries those on the
// delivery leg; the pickup leg's row inherits them via the first-non-blank
// fallback in sharedTd() below rather than requiring an exact match). --
function mergeKey(row){
  return (row.orderLabel||'').trim();
}
function groupRows(rows){
  const groups=[]; const indexByKey=new Map();
  rows.forEach(row=>{
    const key=mergeKey(row);
    if(indexByKey.has(key)) groups[indexByKey.get(key)].rows.push(row);
    else { indexByKey.set(key, groups.length); groups.push({key, rows:[row]}); }
  });
  return groups;
}
// first non-blank value for a shared field across a merged group - same
// backfill rule the display uses (see sharedTd in buildGroupRows), reused
// by the Excel export so a blank pickup-leg row exports with the same
// customer/phone/address/memo the board shows for it, not a blank cell.
function resolvedField(group, field){
  return group.rows.map(r=>r[field]).find(v=>v && String(v).trim()) || '';
}

function emptyRowTr(){
  const td=document.createElement('td');
  td.colSpan=11;
  td.textContent='여기로 행을 드래그하거나, 체크 후 상단 Excel 칸의 날짜로 이동할 수 있습니다 (화면상으로만 - 저장 연결은 다음 단계)';
  const tr=document.createElement('tr');
  tr.className='empty-row';
  tr.appendChild(td);
  return tr;
}

function makeEditableTd(value, className, onCommit){
  const td=document.createElement('td');
  if(className) td.className=className;
  td.contentEditable='true';
  td.spellcheck=false;
  td.textContent=value||'';
  td.addEventListener('blur', ()=>{ onCommit(td.textContent.trim()); renderAll(); });
  td.addEventListener('keydown', e=>{ if(e.key==='Enter'){ e.preventDefault(); td.blur(); } });
  return td;
}

function buildGroupRows(group){
  const trs=[]; const n=group.rows.length;
  group.rows.forEach((row, i)=>{
    const tr=document.createElement('tr');
    tr.className='item-row'+(row.uiStatus==='done'?' row-done':row.uiStatus==='pending'?' row-pending':'');
    tr.dataset.rowKey=row.rowKey;

    // drag handle - a dedicated grip so dragging doesn't fight with
    // clicking/selecting text inside the row's editable cells or dropdown
    const handleTd=document.createElement('td');
    handleTd.className='handle-cell';
    const grip=document.createElement('span');
    grip.className='drag-handle';
    grip.textContent='⠿';
    grip.draggable=true;
    grip.addEventListener('dragstart', e=>{e.dataTransfer.setData('text/plain', row.rowKey); tr.classList.add('dragging')});
    grip.addEventListener('dragend', ()=>tr.classList.remove('dragging'));
    handleTd.appendChild(grip);
    tr.appendChild(handleTd);

    // type - dropdown (배송/회수/Bloomberg/Delayed/회수 Delayed/... plus any
    // custom type added via "+ 추가"), per item (a merged block can still
    // mix a delivery line and a pickup line for the same order)
    const typeTd=document.createElement('td');
    typeTd.className='type-cell';
    const sel=document.createElement('select');
    const applySelClass=()=>{sel.className='type-select '+typeClassFor(sel.value)};
    allTypeOptions().forEach(v=>{
      const opt=document.createElement('option'); opt.value=v; opt.textContent=v;
      if(v===row.itemType) opt.selected=true;
      sel.appendChild(opt);
    });
    if(row.itemType && !allTypeOptions().includes(row.itemType)){
      const cur=document.createElement('option'); cur.value=row.itemType; cur.textContent=row.itemType;
      cur.selected=true; sel.insertBefore(cur, sel.firstChild);
    }
    const addOpt=document.createElement('option');
    addOpt.value='__add__'; addOpt.textContent='+ 추가';
    sel.appendChild(addOpt);
    applySelClass();
    sel.addEventListener('change', ()=>{
      if(sel.value==='__add__'){
        const newType=(window.prompt('새 Type 이름을 입력하세요 (모든 행의 드롭다운에 추가됩니다)')||'').trim();
        if(newType && !allTypeOptions().includes(newType)) CUSTOM_TYPES.push(newType);
        if(newType) row.itemType=newType;
        renderAll();
        return;
      }
      row.itemType=sel.value;
      applySelClass();
    });
    typeTd.appendChild(sel);
    tr.appendChild(typeTd);

    if(i===0){
      const ordTd=document.createElement('td');
      ordTd.className='ord-cell';
      ordTd.rowSpan=n;

      // main line: the SAP order code (ZOR/ZRE/ZRX/ZINP/ZINT/...) + order number - bold, black
      const main=document.createElement('div');
      main.contentEditable='true'; main.spellcheck=false; main.textContent=row.orderLabel;
      main.addEventListener('blur', ()=>{
        const v=main.textContent.trim();
        group.rows.forEach(r=>r.orderLabel=v);
        renderAll();
      });
      main.addEventListener('keydown', e=>{ if(e.key==='Enter'){ e.preventDefault(); main.blur(); } });
      ordTd.appendChild(main);

      // secondary order-code line(s): Bloomberg-side codes (OBD/ORD/SDSK/...), one per
      // line, same small muted style, same free-text edit as everything else - not just
      // a bare number anymore, and more lines can be typed in directly if more than one
      // secondary code applies to this order.
      const codeLines=[...new Set(
        group.rows.flatMap(r=>String(r.deliveryNo||'').split('\n')).map(s=>s.trim()).filter(Boolean)
      )];
      const extra=document.createElement('div');
      extra.className='muted';
      extra.contentEditable='true'; extra.spellcheck=false;
      extra.textContent=codeLines.join('\n');
      extra.addEventListener('blur', ()=>{
        const v=extra.innerText.split('\n').map(s=>s.trim()).filter(Boolean).join('\n');
        group.rows.forEach(r=>r.deliveryNo=v);
        renderAll();
      });
      ordTd.appendChild(extra);

      tr.appendChild(ordTd);
    }

    tr.appendChild(makeEditableTd(row.description, 'item-cell', v=>row.description=v));
    tr.appendChild(makeEditableTd(row.material, 'mn-cell', v=>row.material=v));
    tr.appendChild(makeEditableTd(row.serial, 'sn-cell', v=>row.serial=v));

    if(i===0){
      // first non-blank value across the group, not just group.rows[0] - a
      // merged pickup leg with no customer/phone/address of its own (e.g.
      // ZRX 67066585's 회수 record) should still show/inherit the delivery
      // leg's values rather than rendering blank.
      const sharedTd=(field, cellClass)=>makeEditableTd(
        resolvedField(group, field), 'cust-cell '+cellClass, v=>group.rows.forEach(r=>r[field]=v)
      );
      const custTd=sharedTd('customer','customer-cell'); custTd.rowSpan=n; tr.appendChild(custTd);
      const phoneTd=sharedTd('phone','phone-cell'); phoneTd.rowSpan=n; tr.appendChild(phoneTd);
      const addrTd=sharedTd('address','addr-cell'); addrTd.rowSpan=n; tr.appendChild(addrTd);
      const memoTd=sharedTd('memo','memo-cell'); memoTd.rowSpan=n; tr.appendChild(memoTd);
    }

    const checkTd=document.createElement('td');
    checkTd.className='check-cell';
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.className='row-check'; cb.dataset.rowKey=row.rowKey;
    checkTd.appendChild(cb);
    tr.appendChild(checkTd);

    trs.push(tr);
  });
  return trs;
}

function renderAll(){
  const table=document.getElementById('boardTable');
  table.querySelectorAll('tbody.date-group').forEach(tb=>tb.remove());
  allDatesSorted().forEach(iso=>{
    const tbody=document.createElement('tbody');
    tbody.className='date-group';
    tbody.dataset.date=iso;
    tbody.addEventListener('drop', onDrop);
    tbody.addEventListener('dragover', onDragOver);

    const collapsed=COLLAPSED_DATES.has(iso);
    if(collapsed) tbody.classList.add('collapsed');

    const bannerTd=document.createElement('td');
    bannerTd.colSpan=11;
    const bannerInner=document.createElement('div');
    bannerInner.className='date-banner-inner';
    const labelSpan=document.createElement('span');
    labelSpan.textContent=labelFor(iso);

    const rightGroup=document.createElement('div');
    rightGroup.className='date-banner-right';
    const exportCb=document.createElement('input');
    exportCb.type='checkbox';
    exportCb.className='date-export-check';
    exportCb.title='엑셀로 내보낼 날짜로 선택';
    exportCb.checked=EXPORT_DATES.has(iso);
    exportCb.addEventListener('change', ()=>{
      if(exportCb.checked) EXPORT_DATES.add(iso); else EXPORT_DATES.delete(iso);
    });
    const toggleBtn=document.createElement('button');
    toggleBtn.type='button';
    toggleBtn.className='date-toggle';
    toggleBtn.textContent=collapsed ? '▸' : '▾';
    toggleBtn.title=collapsed ? '펼치기' : '접기';
    toggleBtn.addEventListener('click', ()=>{
      if(COLLAPSED_DATES.has(iso)) COLLAPSED_DATES.delete(iso); else COLLAPSED_DATES.add(iso);
      renderAll();
    });
    rightGroup.appendChild(exportCb);
    rightGroup.appendChild(toggleBtn);

    bannerInner.appendChild(labelSpan);
    bannerInner.appendChild(rightGroup);
    bannerTd.appendChild(bannerInner);
    const bannerTr=document.createElement('tr');
    bannerTr.className='date-banner-row';
    bannerTr.appendChild(bannerTd);
    tbody.appendChild(bannerTr);

    const rows=ROWS.filter(r=>r.date===iso);
    if(!rows.length){
      tbody.appendChild(emptyRowTr());
    } else {
      groupRows(rows).forEach(g=>buildGroupRows(g).forEach(tr=>tbody.appendChild(tr)));
    }
    table.appendChild(tbody);
  });
}
renderAll();

// -- Process: acts on whichever rows are checked, same checkboxes Portal uses --
function processDone(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='done'});
  renderAll();
  toast(keys.length+'건 완료 처리(파랑) - 저장은 다음 단계');
}
function processPickupReady(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  keys.forEach(k=>{
    const row=ROWS.find(r=>r.rowKey===k);
    if(!row)return;
    row.uiStatus = row.itemType==='회수' ? 'pending' : 'done';
  });
  renderAll();
  toast(keys.length+'건 처리: 회수는 대기(빨강), 나머지는 완료(파랑) - 저장은 다음 단계');
}

// -- Export: the one action in this app that actually writes a real file
// (the live shipping Excel workbook) instead of just changing what's on
// screen. Builds the current board state (post move/edit/merge/delete) for
// each date checked via the per-date checkbox next to its ▾/▸ toggle, and
// sends it to the server to write one sheet per date. --
function buildExportRows(iso){
  const rows=ROWS.filter(r=>r.date===iso);
  const out=[];
  groupRows(rows).forEach(group=>{
    const customer=resolvedField(group,'customer');
    const phone=resolvedField(group,'phone');
    const address=resolvedField(group,'address');
    const memo=resolvedField(group,'memo');
    group.rows.forEach((row, i)=>{
      out.push({
        isGroupFirst: i===0,
        itemType: row.itemType,
        orderLabel: row.orderLabel,
        deliveryNo: row.deliveryNo,
        description: row.description,
        material: row.material,
        serial: row.serial,
        customer, phone, address, memo,
      });
    });
  });
  return out;
}

async function exportChecked(){
  const isos=[...EXPORT_DATES];
  if(!isos.length){toast('내보낼 날짜를 먼저 체크하세요 (날짜 줄 오른쪽 체크박스)');return}
  const payload={};
  isos.forEach(iso=>{payload[iso]=buildExportRows(iso)});
  toast('엑셀로 내보내는 중...');
  try{
    const res=await fetch('/api/export', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({dates:payload})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'내보내기 실패')}
    const summary=Object.entries(data.exported).map(([iso,info])=>`${labelFor(iso)}: ${info.sheet}시트 ${info.rows}행`).join(' / ');
    toast('엑셀 저장 완료 - '+summary);
  }catch(err){
    toast('내보내기 실패: '+err.message);
  }
}

// -- Delete: removes checked rows from the board. Screen-only for now,
// same as every other action here (moves, Done/Ready) - nothing in this
// dashboard writes back to workbench.db yet. --
function deleteChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  ROWS = ROWS.filter(r=>!keys.includes(r.rowKey));
  renderAll();
  toast(keys.length+'건 삭제 (화면에서만 제거 - DB는 그대로)');
}

// -- Drag-and-drop between date sections (nearby dates) --
function onDragOver(e){e.preventDefault();e.currentTarget.classList.add('drag-over')}
document.addEventListener('dragleave',e=>{
  const grp=e.target.closest && e.target.closest('.date-group');
  if(grp && !grp.contains(e.relatedTarget)) grp.classList.remove('drag-over');
});
document.addEventListener('dragend',e=>{
  document.querySelectorAll('.date-group.drag-over').forEach(g=>g.classList.remove('drag-over'));
});
function onDrop(e){
  e.preventDefault();
  const tbody=e.currentTarget;
  tbody.classList.remove('drag-over');
  const key=e.dataTransfer.getData('text/plain');
  const row=ROWS.find(r=>r.rowKey===key);
  if(!row)return;
  const iso=tbody.dataset.date;
  if(row.date===iso)return;
  row.date=iso;
  renderAll();
  toast('다른 날짜로 이동 (화면상으로만 - 저장은 다음 단계)');
}

// -- Excel-panel date picker: check rows below, pick a date up top, they all
// jump there. Handles dates that don't have a section yet (e.g. board only
// shows 7/28 and 7/30 but the user wants 7/29) by creating a fresh date
// section in the right sorted position (renderAll() rebuilds every date in
// sorted order every time, so a brand-new iso just needs to exist on a
// row's `date` field to get its own section). --
function moveCheckedToDate(iso){
  if(!iso)return;
  if(!/^\d{4}-\d{2}-\d{2}$/.test(iso) || isNaN(new Date(iso+'T00:00:00').getTime())){
    toast('날짜 형식이 올바르지 않습니다. 달력에서 다시 선택해주세요');
    return;
  }
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  let moved=0;
  keys.forEach(k=>{
    const row=ROWS.find(r=>r.rowKey===k);
    if(row && row.date!==iso){row.date=iso; moved++}
  });
  renderAll();
  if(moved) toast(moved+'건을 '+labelFor(iso)+'(으)로 이동 (화면상으로만 - 저장은 다음 단계)');
  else toast('이미 해당 날짜에 있습니다');
}
</script>
</body>
</html>
'''


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        with LOG_FILE.open("a", encoding="utf-8") as fh:
            fh.write(now_text() + " [Workbench] " + (fmt % args) + "\n")

    def send_json(self, data, status=200):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0") or 0)
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self):
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/":
                payload = render_dashboard_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/api/orders":
                qs = parse_qs(parsed.query)
                self.send_json({"orders": list_orders(qs.get("q", [""])[0], qs.get("status", [""])[0])})
                return
            m = re.match(r"^/api/orders/(\d+)$", parsed.path)
            if m:
                data = get_order(int(m.group(1)))
                if not data:
                    self.send_json({"error": "not found"}, 404)
                else:
                    self.send_json(data)
                return
            self.send_json({"error": "not found"}, 404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)

    def do_POST(self):
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/api/import/today":
                self.send_json(load_today_from_excel())
                return
            if parsed.path == "/api/export":
                body = self.read_json()
                self.send_json(export_dates_to_excel(body.get("dates") or {}))
                return
            m = re.match(r"^/api/orders/(\d+)/run_portal$", parsed.path)
            if m:
                self.send_json(run_portal(int(m.group(1))))
                return
            m = re.match(r"^/api/orders/(\d+)/status$", parsed.path)
            if m:
                body = self.read_json()
                status = body.get("status", "new")
                with connect() as con:
                    con.execute("UPDATE orders SET status=?, updated_at=? WHERE id=?", (status, now_text(), int(m.group(1))))
                    add_event(con, int(m.group(1)), f"Status changed to {status}")
                self.send_json({"ok": True})
                return
            self.send_json({"error": "not found"}, 404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)


def serve(host, port, open_browser=True):
    ensure_db()
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"Workbench running: {url}", flush=True)
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


def main():
    parser = argparse.ArgumentParser(description="MJSuh local operations workbench")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-open", action="store_true")
    parser.add_argument("--import-today", action="store_true")
    args = parser.parse_args()
    ensure_db()
    if args.import_today:
        print(json.dumps(load_today_from_excel(), ensure_ascii=False, indent=2))
        return
    serve(args.host, args.port, open_browser=not args.no_open)


if __name__ == "__main__":
    main()

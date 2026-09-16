"""Local operations workbench for SAP/Bloomberg delivery automation.

This app moves day-to-day work state out of the shipping Excel file and into a
small SQLite database. Excel is treated as an import source, not the runtime
source of truth for portal execution.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from collections import Counter
from datetime import date, datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import openpyxl
import win32api
import win32event
import winerror
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import excel_handler
from config import EXCEL_PATH, FINALIZE_EXPORT_PATH, REFRESH_INTERVAL_MINUTES
from holiday_check import kr_holiday_labels, skip_reason
from order_domain import resolve_item_type

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
DB_PATH = BASE_DIR / "workbench.db"
LOG_FILE = BASE_DIR / "automation.log"
# Workbench's own HTTP access log used to be appended straight into
# automation.log (see the old Handler.log_message) - harmless on its own,
# but it meant automation.log's mtime got bumped by every dashboard click/
# poll regardless of whether the actual SAP loop (main.py/startup.py) was
# alive or hung. That silently defeated any "has automation.log gone stale"
# hang-detection (see _sap_loop_watchdog_thread() below) while a workbench
# tab happened to be open and in use - split into its own file so
# automation.log's mtime reflects only the SAP loop's own activity again.
WORKBENCH_ACCESS_LOG = BASE_DIR / "workbench_access.log"

# 2026-08-24: workbench_app.py had ZERO application-level logging - it ran in
# a bare CREATE_NEW_CONSOLE window with nothing but print(), so every past
# crash's real exception vanished the moment that window closed (confirmed
# by grepping the whole file: no `import logging` anywhere). Every prior
# "fix" for a workbench crash was therefore a guess at symptoms, not the
# actual traceback - which is almost certainly why the same class of problem
# kept recurring after each fix. WORKBENCH_ERROR_LOG is the missing net:
# sys.excepthook covers the main thread, threading.excepthook covers
# background threads (the SAP watchdog thread, per-request handler threads,
# the per-order 'watcher' threads started from Popen callbacks - any of
# which could die silently before this and leave e.g. an order stuck at
# 'working' forever with no record of why, see _pod_pids_for_delivery()'s
# docstring for a real instance of exactly that symptom).
WORKBENCH_ERROR_LOG = BASE_DIR / "workbench_error.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(WORKBENCH_ERROR_LOG, encoding="utf-8")],
)
logger = logging.getLogger("workbench")


def _log_unhandled(exc_type, exc_value, exc_tb):
    if exc_type is KeyboardInterrupt:
        return
    logger.error(
        "미처리 예외로 스레드/프로세스가 죽음:\n"
        + "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    )


sys.excepthook = _log_unhandled


def _thread_excepthook(args):
    _log_unhandled(args.exc_type, args.exc_value, args.exc_traceback)


threading.excepthook = _thread_excepthook

# How long a soft-deleted item stays recoverable before purge_deleted_items()
# actually drops it.
DELETED_ITEM_RETENTION_DAYS = 7

# How many recent workbench actions (delete/move/status) can be undone via
# Ctrl+Z / the "실행취소" button. Older batches are pruned after each action.
UNDO_HISTORY_LIMIT = 20

# The 4 "shared" order-level fields that can each independently be merged
# ("Excel merge cells") across different orders - see merge_orders(). Order
# matters here only in that it's reused as the display/prompt order.
MERGE_FIELDS = ("customer", "phone", "address", "memo")

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
        # Undo history: one action_batches row per workbench action (delete /
        # move_date / status), one action_log row per item it touched, holding
        # that item's column value from just before the action so it can be
        # restored. Capped to UNDO_HISTORY_LIMIT batches (see _log_action_batch).
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS action_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                action_type TEXT NOT NULL,
                summary TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS action_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id INTEGER NOT NULL,
                item_id INTEGER NOT NULL,
                prev_value TEXT,
                FOREIGN KEY(batch_id) REFERENCES action_batches(id) ON DELETE CASCADE
            )
            """
        )
        # Migration for DBs created before item_type existed.
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN item_type TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        # item_date: per-item date override. NULL means "use the parent order's
        # source_date" (the default for everything imported so far) - only set
        # once a single item is moved to a different day than its order, e.g. a
        # ZRX's 배송 leg ships today but its 회수 leg isn't picked up for a few
        # more days, so the two items of one order end up on different board dates.
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN item_date TEXT")
        except sqlite3.OperationalError:
            pass
        # board_status: workbench UI's own Done/Ready/Cancel marker
        # (blue/red/gray - 'done'/'pending'/'cancelled'), per item - unrelated
        # to orders.status, which is the separate SAP/Portal automation
        # pipeline stage (new/working/posted/printed). Cancel doesn't delete
        # the row, just marks it (gray + strikethrough) - added 2026-08-07.
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN board_status TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        # deleted_at: soft-delete marker. NULL = live. Set instead of actually
        # removing the row so "2. 체크된 행 삭제" is recoverable for
        # DELETED_ITEM_RETENTION_DAYS - see purge_deleted_items().
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN deleted_at TEXT")
        except sqlite3.OperationalError:
            pass
        # board_pos: manual on-screen ordering within a date section (the drag
        # handle's "move to a different row, same date" case). NULL = never
        # manually reordered, falls back to natural (order_id, line_no) order.
        # A reorder always rewrites board_pos for every item currently in that
        # date section (see reorder_items()) - a dense 0..N-1 sequence, no gap
        # bookkeeping needed.
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN board_pos INTEGER")
        except sqlite3.OperationalError:
            pass
        # merge_group_<field>: manual "Excel merge cells" grouping across
        # DIFFERENT orders that are, in real life, the same customer under
        # separate SAP order numbers - one independent group-id column per
        # shared field (not a single merge_group for all 4) because a real
        # case can share e.g. just the address while the customer name on
        # file differs (roommates, a company's several contacts, etc.) - see
        # merge_orders()/unmerge_orders(). NULL = that field isn't merged.
        # Lives on `orders` (not order_items) because customer/phone/address/
        # memo are themselves order-level fields.
        for _field in MERGE_FIELDS:
            try:
                con.execute(f"ALTER TABLE orders ADD COLUMN merge_group_{_field} TEXT")
            except sqlite3.OperationalError:
                pass
        # extra_codes: ORD/SDSK line(s) (2026-08-06) - kept separate from
        # delivery_no (the real OBD matching key POD resolution/UNIQUE(order_no,
        # delivery_no) depend on) so it can hold free multi-line text without
        # risking that key. Purely a rendering concern: board_row() joins this
        # with delivery_no into one combined "deliveryNo" string for the
        # client, so on screen/in Excel it all still looks like one Order#
        # cell with OBD/ORD/SDSK each on their own line - never a separate
        # visible column. See update_delivery_codes().
        try:
            con.execute("ALTER TABLE orders ADD COLUMN extra_codes TEXT")
        except sqlite3.OperationalError:
            pass
        # portal_error/pod_error: last-failure summary for each half of the
        # Bloomberg Portal automation (todo b, 2026-08-18) - orders.status
        # already tracks Serial/QR success/failure ('printed'/'error') but
        # POD deliberately never touches status (see run_pod_update()'s own
        # comment: writing there would corrupt run_portal()'s duplicate-run
        # guard), so POD needed its own field. Both are empty string = no
        # active failure; board_row() surfaces them as the board's "row-error"
        # orange highlight + ⚠ badge (see runSpareOneshot/run_portal/
        # run_pod_update watchers for where each gets set/cleared). Cleared
        # the instant a fresh run of that same step starts, not just on
        # success - so the moment someone retries, the orange goes away
        # instead of lingering until the retry finishes too.
        try:
            con.execute("ALTER TABLE orders ADD COLUMN portal_error TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        try:
            con.execute("ALTER TABLE orders ADD COLUMN pod_error TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        # manual_cell_count: how many of the 7 제목/M-N/S-N/Customer/Phone/
        # Address/Memo cells addManualRow()'s modal was set to when a memo row
        # was created (2026-08-27) - NULL for every non-memo row. Needed
        # because the board's compact memo-row rendering used to infer cell
        # boundaries purely from which cells had content (see buildGroupRows'
        # "!row.orderNo && n===1" branch) - which silently collapsed any cell
        # the user deliberately left blank (an optional cell IS still a cell)
        # back into its neighbor, e.g. picking 7 cells but only typing a
        # title rendered as just 1 wide cell instead of 7 separate ones.
        # Storing the actual chosen count lets rendering reproduce the exact
        # same boundaries addManualRow() showed, regardless of which cells
        # ended up empty.
        try:
            con.execute("ALTER TABLE orders ADD COLUMN manual_cell_count INTEGER")
        except sqlite3.OperationalError:
            pass
        # sn_relo_error: 회수 오더에 S/N이 없어서(X) 실제 회수된 S/N을
        # 워크벤치에 기록한 뒤 그 값을 오더(VA02)에 반영 시도했다가 실패한
        # 경우(firm/cust 불일치 등 - paper relo 필요)의 마지막 실패 사유
        # (2026-08-19). portal_error/pod_error와 같은 패턴이지만 오더 단위가
        # 아니라 order_items 단위 - 회수 항목의 S/N은 애초에 항목별로 다르기
        # 때문. 빈 문자열 = 실패 없음. 재시도를 시작하는 순간 즉시 비워짐
        # (portal_error와 같은 이유 - 재시도 중엔 orange가 안 남아있게).
        try:
            con.execute("ALTER TABLE order_items ADD COLUMN sn_relo_error TEXT NOT NULL DEFAULT ''")
        except sqlite3.OperationalError:
            pass
        # cell_highlights: manual visual emphasis (background color / text
        # color / bold) on any one rendered board cell - Type/Order#/Item/
        # M-N/S-N/Customer/Phone/Address/Memo. Keyed by (item_id, field), where
        # item_id is always the id of whichever order_items row actually owns
        # the rendered <td> - for the per-item fields (type/description/
        # material/serial) that's just that item; for the order-level fields
        # (order/customer/phone/address/memo), which render once per rowSpan
        # block, it's the first item of whichever group the block's <td>
        # physically lives on (same anchor the client already uses to build
        # that block). style_json is a compact {"bg":"yellow"|"pink",
        # "color":"red","bold":true} blob (only the keys actually set) - NULL/
        # missing row = no highlight, same "absence means default" pattern as
        # every other optional per-cell override in this file.
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS cell_highlights (
                item_id INTEGER NOT NULL,
                field TEXT NOT NULL,
                style_json TEXT,
                PRIMARY KEY(item_id, field)
            )
            """
        )
        # manual_item_edits: marks which per-item field(s) (description/
        # material/serial) a human has directly edited on the board via
        # update_item_field(). Real bug this closes: load_today_from_excel()
        # deletes and reinserts every order's items on every reimport (the
        # SAP auto-sync loop does this ~every 10 min for every order still
        # live in Excel) and, unlike item_date/board_status/deleted_at/
        # cell_highlights/the order-level customer/phone/address/memo fields
        # (all separately fixed to survive reimport, see the 2026-08 rounds
        # in project memory), description/material/serial were still always
        # overwritten from Excel every single time - so a typed-in serial
        # (exactly what Portal automation needs) could silently revert within
        # one sync cycle even after update_item_field() "saved" it. Presence
        # of an (item_id, field) row here means "don't let the next reimport
        # touch this field" - see its use in load_today_from_excel() below.
        # Deliberately does NOT get cleared by undo_last_action() (reverting
        # a value doesn't un-mark it) - a documented, accepted simplification,
        # same spirit as other small tradeoffs already noted in this file.
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS manual_item_edits (
                item_id INTEGER NOT NULL,
                field TEXT NOT NULL,
                PRIMARY KEY(item_id, field)
            )
            """
        )
        # custom_types (2026-08-26): the Type dropdown's "+ 추가" custom-type
        # option used to only push the new name into a JS-only CUSTOM_TYPES
        # array (allTypeOptions() on the client) - never sent to the server.
        # A row that was actually GIVEN a custom type still displayed fine
        # after reload (resolve_item_type()/item_type persistence, fixed the
        # same day - see their own comments) since the client synthesizes a
        # selected <option> for whatever's on the row even if it's not in the
        # known list. But the *list itself* reset to empty on every reload,
        # so the custom type silently stopped being offered as a choice for
        # any OTHER row, and had to be retyped via "+ 추가" again. This table
        # is that missing persistence - just the set of names ever added.
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS custom_types (
                name TEXT PRIMARY KEY,
                created_at TEXT
            )
            """
        )
        # notepad_pages (2026-08-31 사용자 요청, 같은 날 카드+카테고리 모델에서
        # 탭+큰 textarea 모델로 재설계 - render_notes_page()의 주석 참고): 오더/
        # 날짜와 무관한 잡메모(고객센터 번호, SAP 계정, 배송기한 규칙 등)를
        # /notes 페이지에서 탭별로 자유롭게 적는 곳. 한 tab = 한 row, body는
        # 그 tab의 전체 텍스트를 통째로 담음(항목 단위로 쪼개지 않음) - 오더
        # 보드의 undo/하이라이트/날짜 그룹과는 전혀 무관하게 독립 동작.
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS notepad_pages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL DEFAULT '',
                body TEXT NOT NULL DEFAULT '',
                sort_pos INTEGER NOT NULL DEFAULT 0,
                created_at TEXT,
                updated_at TEXT
            )
            """
        )
        # manual_holidays (2026-09-13 사용자 요청): 날짜 배너를 수동으로 빨간색
        # (휴무일)으로 지정한 날짜들. 주말/한국 공휴일(holiday_check.py)은 이미
        # 자동으로 빨갛게 표시되는데, 그 외에 회사 자체 휴무일(창립기념일 등)이나
        # 라이브러리가 못 잡는 임시공휴일(선거일 등)을 화면에서 직접 표시해두기
        # 위한 수동 오버레이 - SAP 자동화 스케줄(holiday_check.skip_reason())과는
        # 완전히 별개다: 이 테이블은 workbench 화면 표시용일 뿐, main.py/
        # morning_routine.py/watchdog.py의 자동 기동 판단에는 전혀 영향 없음.
        # 존재 자체가 "빨간색"이라는 뜻이라 컬럼은 날짜 하나뿐 - "이 날짜를
        # 다시 평일로" 되돌리는 것은 그냥 행을 지우는 것(delete_manual_holiday).
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS manual_holidays (
                date TEXT PRIMARY KEY,
                created_at TEXT
            )
            """
        )
    purge_deleted_items()


def purge_deleted_items():
    """Hard-drop items that have been soft-deleted for longer than
    DELETED_ITEM_RETENTION_DAYS, then remove any order left with zero item
    rows at all (soft-deleted-but-not-yet-purged items still count as rows,
    so an order isn't cleaned up until its items actually age out)."""
    cutoff = (datetime.now() - timedelta(days=DELETED_ITEM_RETENTION_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
    with connect() as con:
        con.execute("DELETE FROM order_items WHERE deleted_at IS NOT NULL AND deleted_at <= ?", (cutoff,))
        con.execute("DELETE FROM orders WHERE id NOT IN (SELECT DISTINCT order_id FROM order_items)")
        # Safety net for cell_highlights left pointing at an item_id that no
        # longer exists - load_today_from_excel()'s reimport re-points these
        # at the freshly-reinserted item when it can match one, but a line
        # that's genuinely gone (removed from Excel, not just re-imported)
        # has nothing to re-point to and would otherwise sit here forever.
        con.execute("DELETE FROM cell_highlights WHERE item_id NOT IN (SELECT id FROM order_items)")
        con.execute("DELETE FROM manual_item_edits WHERE item_id NOT IN (SELECT id FROM order_items)")


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


def _log_action_batch(con, action_type, summary, item_prev_pairs):
    """Record one undoable batch: item_prev_pairs is [(item_id, prev_value), ...]
    holding the touched column's value from just before this action, so
    undo_last_action() can put it back. Must be called on the same connection
    as (and before/after, doesn't matter which side of) the actual UPDATE."""
    cur = con.execute(
        "INSERT INTO action_batches(action_type, summary, created_at) VALUES (?, ?, ?)",
        (action_type, summary, now_text()),
    )
    batch_id = cur.lastrowid
    con.executemany(
        "INSERT INTO action_log(batch_id, item_id, prev_value) VALUES (?, ?, ?)",
        [(batch_id, item_id, prev_value) for item_id, prev_value in item_prev_pairs],
    )
    con.execute(
        "DELETE FROM action_batches WHERE id NOT IN "
        "(SELECT id FROM action_batches ORDER BY id DESC LIMIT ?)",
        (UNDO_HISTORY_LIMIT,),
    )


def undo_last_action():
    with connect() as con:
        batch = con.execute("SELECT * FROM action_batches ORDER BY id DESC LIMIT 1").fetchone()
        if not batch:
            return {"undone": False, "message": "되돌릴 작업이 없습니다"}
        log_rows = con.execute("SELECT * FROM action_log WHERE batch_id=?", (batch["id"],)).fetchall()
        # merge/unmerge touch one or more `orders.merge_group_<field>` columns
        # at once, and edit_delivery_codes touches delivery_no+extra_codes
        # together (item_id column here actually holds an order id for all
        # three action types) - prev_value is a JSON {column: prior_value}
        # blob since the exact set of touched columns varies per action,
        # unlike every other action type below which always touches one
        # fixed order_items column keyed by a real item id.
        if batch["action_type"] in ("merge", "unmerge", "edit_delivery_codes"):
            for row in log_rows:
                prior = json.loads(row["prev_value"]) if row["prev_value"] else {}
                for col, val in prior.items():
                    con.execute(f"UPDATE orders SET {col}=? WHERE id=?", (val, row["item_id"]))
        elif batch["action_type"] == "edit_item":
            for row in log_rows:
                prior = json.loads(row["prev_value"])
                con.execute(f"UPDATE order_items SET {prior['field']}=? WHERE id=?", (prior["value"], row["item_id"]))
        elif batch["action_type"] == "edit_order":
            for row in log_rows:
                prior = json.loads(row["prev_value"])
                con.execute(f"UPDATE orders SET {prior['field']}=? WHERE id=?", (prior["value"], row["item_id"]))
        elif batch["action_type"] == "edit_highlight":
            for row in log_rows:
                prior = json.loads(row["prev_value"])
                field, style_json = prior["field"], prior["style_json"]
                if style_json:
                    con.execute(
                        "INSERT INTO cell_highlights(item_id, field, style_json) VALUES(?,?,?) "
                        "ON CONFLICT(item_id, field) DO UPDATE SET style_json=excluded.style_json",
                        (row["item_id"], field, style_json),
                    )
                else:
                    con.execute("DELETE FROM cell_highlights WHERE item_id=? AND field=?", (row["item_id"], field))
        else:
            column = {
                "delete": "deleted_at", "move_date": "item_date", "status": "board_status", "reorder": "board_pos",
            }[batch["action_type"]]
            for row in log_rows:
                con.execute(f"UPDATE order_items SET {column}=? WHERE id=?", (row["prev_value"], row["item_id"]))
        con.execute("DELETE FROM action_batches WHERE id=?", (batch["id"],))
        return {"undone": True, "summary": batch["summary"], "affected": len(log_rows)}


def _pop_prior_item_state(prior_by_key, prior_by_line, item, line_no):
    """Match a freshly-parsed Excel item to its pre-reimport row so
    item_date/board_status/deleted_at survive a re-import. Prefers
    (material, serial) identity since that's stable even if row order
    shifts; falls back to line_no when material+serial can't identify it
    (e.g. material-only lines with no serial)."""
    key = (item.get("material") or "", item.get("serial") or "")
    bucket = prior_by_key.get(key)
    if bucket:
        return bucket.pop(0)
    return prior_by_line.get(line_no)


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
# Manual cell_highlights carried into the export: column 1 (Order #) also
# carries the Type text (see _order_cell_text below, there's no separate
# Type column in this sheet layout - matches the hand-built originals), so a
# "type" highlight is checked as a fallback for column 1 if "order" isn't
# set. Every other column maps to exactly one highlight field.
_EXPORT_HIGHLIGHT_FIELDS_BY_COL = {
    1: ("order", "type"), 2: ("description",), 3: ("material",), 4: ("serial",),
    5: ("customer",), 6: ("phone",), 7: ("address",), 8: ("memo",),
}
_EXPORT_HIGHLIGHT_FILL = {
    "yellow": PatternFill(fill_type="solid", fgColor="FFFF00"),
    "pink": PatternFill(fill_type="solid", fgColor="FFC0CB"),
}
_EXPORT_HIGHLIGHT_FONT_COLOR = {"red": "FFFF0000"}


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
    """Writes `rows` into `sheet_name`. If that sheet doesn't exist yet,
    creates it fresh (date banner + headers + data, as before). If it
    already exists - a prior 마감 for this exact date already made it -
    appends `rows` starting right after its last used row instead: no new
    "(2)"-suffixed sheet, no rewriting the banner/header above.

    2026-08-27 (user request): an order whose 마감 slips past its own day
    (e.g. a POD/SAP issue only gets resolved a few days later) should stay on
    the workbench board, unchecked, while everything else for that date gets
    마감-ed normally; once it's later resolved and checked off, 마감-ing it
    should land it at the bottom of the date it always belonged to - not
    spawn a duplicate sheet for that date. Since export_dates_to_excel() now
    always targets the exact "{month}-{day}" sheet name (no more
    _unique_sheet_name() collision-avoidance), this is the only place that
    needs to branch on "does it already exist"."""
    d = date.fromisoformat(iso)
    existing = sheet_name in wb.sheetnames
    ws = wb[sheet_name] if existing else wb.create_sheet(sheet_name)

    if existing:
        row_cursor = max(ws.max_row + 1, 3)
    else:
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
        row_cursor = 3

    # Data rows. Each of E-H (customer/phone/address/memo) merges down its
    # OWN contiguous run independently, per row["mergeFirst"][field] (built
    # client-side by buildExportRows/computeFieldSpans) - a real case can
    # share e.g. just the address across rows with different customer names,
    # so this can no longer be one shared group boundary for all 4 columns.
    # This merge run is scoped only to the rows being written THIS call - an
    # appended block never merges upward into whatever's already above it.
    n = len(rows)
    row_numbers = [row_cursor + i for i in range(n)]

    for i, row in enumerate(rows):
        r = row_numbers[i]
        mf = row.get("mergeFirst") or {}
        values = {
            1: _order_cell_text(row),
            2: row.get("description") or "",
            3: row.get("material") or "",
            4: row.get("serial") or "",
            5: (row.get("customer") or "") if mf.get("customer", True) else "",
            6: (row.get("phone") or "") if mf.get("phone", True) else "",
            7: (row.get("address") or "") if mf.get("address", True) else "",
            8: (row.get("memo") or "") if mf.get("memo", True) else "",
        }
        highlights = row.get("highlights") or {}
        for col in range(1, 9):
            cell = ws.cell(row=r, column=col)
            cell.value = values[col]
            style = next(
                (highlights[f] for f in _EXPORT_HIGHLIGHT_FIELDS_BY_COL[col] if highlights.get(f)), None
            )
            cell.fill = _EXPORT_HIGHLIGHT_FILL.get((style or {}).get("bg"), _EXPORT_DATA_FILL)
            font_kwargs = {}
            if col == 8:
                font_kwargs["bold"] = True
            if style and style.get("bold"):
                font_kwargs["bold"] = True
            if style and style.get("color") in _EXPORT_HIGHLIGHT_FONT_COLOR:
                font_kwargs["color"] = _EXPORT_HIGHLIGHT_FONT_COLOR[style["color"]]
            if font_kwargs:
                cell.font = Font(**font_kwargs)
            cell.alignment = _EXPORT_COL_ALIGN[col]
            cell.border = _EXPORT_THIN_BORDER

    for field, col in zip(("customer", "phone", "address", "memo"), (5, 6, 7, 8)):
        i = 0
        while i < n:
            j = i + 1
            while j < n and not (rows[j].get("mergeFirst") or {}).get(field, True):
                j += 1
            if j - i > 1:
                ws.merge_cells(start_row=row_numbers[i], start_column=col, end_row=row_numbers[j - 1], end_column=col)
            i = j

    row_cursor = (row_numbers[-1] + 1) if n else row_cursor
    return ws, row_cursor - 1


def export_dates_to_excel(dates_payload):
    """
    Writes the board's current (client-side) state for the given dates into
    the "마감" export workbook (FINALIZE_EXPORT_PATH - a dedicated file,
    separate from the live daily-collection workbook EXCEL_PATH, per user
    request 2026-08-14) - one sheet per date, named the same "{month}-{day}"
    way `today_sheet_names()` looks for them (e.g. "7-29"). Existing sheets
    are never deleted or overwritten: a date exported for the first time
    gets a fresh sheet, but 마감-ing that SAME date again later (e.g. one
    order's 마감 slipped a few days past everything else - 2026-08-27 user
    request) appends the newly-checked rows to the bottom of that same
    "7-29" sheet instead of spawning "7-29 (2)" - see _write_export_sheet().
    The board's data still comes from the client's in-memory `ROWS`, not
    workbench.db, since nothing this session persists there yet. Formatting
    (column widths, header fill, blue "done" fill, borders, merged
    customer/phone/address/memo per group) matches the hand-built sheets SAP
    collection produces.

    Takes a full backup of the export workbook before writing - this is the
    one action in this app that touches a real, persistent Excel file
    (everything else here is screen-only), so a mistake here is not casually
    undoable the way an in-browser action is.

    Once the save below succeeds, also kicks off vendor_dashboard.py in the
    background for every exported date (2026-08-25: 마감 IS what freezes a
    day's colors/rows into this same file, so it's the natural point to keep
    the local Vendor Activities Dashboard copy/paste sheet in sync - no
    separate manual step to remember). That refresh is fire-and-forget: its
    own failure must never make 마감 itself look like it failed.
    """
    excel_path = Path(FINALIZE_EXPORT_PATH)
    if not excel_path.exists():
        # openpyxl은 시트가 0개인 워크북 저장을 거부하므로 기본 빈 "Sheet"는
        # 그대로 둔다 - 실제 마감 시트들과 나란히 있어도 무해함.
        excel_path.parent.mkdir(parents=True, exist_ok=True)
        openpyxl.Workbook().save(excel_path)

    # 마감을 누를 때마다 새 타임스탬프 파일을 만들면 이 파일 자체가 이미
    # "마감 전용 백업 파일"이라 backup-of-a-backup이 무한정 쌓인다 (사용자
    # 지적, 2026-08-14). 그래서 백업은 이름 고정 1개만 두고 매번 덮어쓴다 -
    # "직전 저장 시점으로 한 단계만 되돌릴 수 있는" 안전장치는 유지하되
    # 파일이 쌓이지는 않음.
    backup_path = excel_path.with_name(f"{excel_path.stem}.before_last_export{excel_path.suffix}")
    shutil.copy2(excel_path, backup_path)

    wb = openpyxl.load_workbook(excel_path)
    try:
        result = {}
        for iso, rows in dates_payload.items():
            d = date.fromisoformat(iso)
            sheet_name = f"{d.month}-{d.day}"
            _write_export_sheet(wb, sheet_name, iso, rows)
            result[iso] = {"sheet": sheet_name, "rows": len(rows)}
        wb.save(excel_path)
        _trigger_dashboard_refresh(list(dates_payload.keys()))
        _mirror_finalize_to_shared_excel(dates_payload)
        return {"exported": result, "backup": str(backup_path)}
    finally:
        wb.close()


def _mirror_finalize_to_shared_excel(dates_payload):
    """마감(확정)된 데이터를 공유파일(EXCEL_PATH)에도 그대로 반영한다
    (2026-09-01 사용자 결정).

    배경: 공유파일은 다른 사람들과 공유하는 파일이라 항상 깔끔해야 하는데,
    지금까지는 매일 SAP 자동수집이 "아직 확정 전"인 진행중 상태까지 직접
    써서 지저분해졌었다. 그 자동쓰기는 이제 껐고(main.py/manual_order_
    handler.py), 대신 workbench에서 실제로 "완료"로 확정 처리된 마감 데이터만
    이 함수로 공유파일에 반영한다 - 공유파일이 "1년치 확정 기록 아카이브"
    역할을 하게 됨. 실시간 상태는 board_sync.py가 20분마다 자동으로 채우는
    REALTIME_SYNC_PATH가 대신한다.

    구조 문제와 해결: 공유파일의 기존 날짜별 시트("M-D")는 이 마감 파일과
    달리 한 시트 안에 여러 날짜 배너가 쌓이는 옛 구조(board_sync.py 참고)라,
    "한 시트=한 날짜" 가정으로 그대로 이어붙이면 데이터가 시트 맨 끝(전혀
    다른 미래 날짜 배너 밑)에 엉뚱하게 붙는다. 그런 시트(배너가 2개 이상)를
    만나면 건드리지 않고 "{sheet_name} (마감)"이라는 구분되는 이름으로 대신
    만든다(excel_handler.mirror_finalized_rows_to_excel_path 안에서 처리) -
    앞으로 새로 마감되는 날짜는 이 옛 구조 시트가 없을 것이므로(이 전환
    이후로는 공유파일에 새 날짜 시트가 자동생성 안 됨) 자연스럽게 깨끗한
    "한 시트=한 날짜" 구조로 자리잡는다.

    2026-09-02 openpyxl→COM 전환: 처음엔 openpyxl(_write_export_sheet 재사용
    시도)로 짰다가 실제로 두 번 다 PermissionError로 실패했다 - 공유파일은
    사람이 실제로 Excel에 열어놓고 보는 경우가 흔한데, openpyxl의 일반 파일
    저장은 그렇게 열려있는 파일에 못 쓴다. write_orders_to_excel()이 이미
    검증해온 xlwings/COM 경로(excel_handler.get_workbook())로 옮겼다 - 열려
    있는 세션에 그대로 붙어서 실시간 반영된다.

    실패해도 마감 자체(위에서 이미 저장 완료)를 막지 않는다 - _trigger_
    dashboard_refresh()와 같은 이유로 로그만 남기고 넘어간다."""
    excel_path = Path(EXCEL_PATH)
    if not excel_path.exists():
        logger.warning(f"공유파일 없음, 마감 반영 스킵: {excel_path}")
        return

    backup_path = excel_path.with_name(f"{excel_path.stem}.before_last_마감{excel_path.suffix}")
    try:
        shutil.copy2(excel_path, backup_path)
    except Exception:
        logger.error("공유파일 백업 실패 - 반영은 계속 시도", exc_info=True)

    try:
        excel_handler.mirror_finalized_rows_to_excel_path(dates_payload)
        logger.info(f"공유파일 마감 반영 완료: {list(dates_payload.keys())}")
    except Exception:
        logger.error("공유파일(EXCEL_PATH) 마감 반영 실패 (마감 자체는 정상 완료됨)", exc_info=True)


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


def sync_delivery_excel():
    """workbench.db의 현재 상태를 기존 배송장 엑셀(EXCEL_PATH - 예전부터 써온
    실시간 배송장, board_sync.py 상단 docstring 참고)의 오늘 시트에 반영한다.
    board_sync.sync()를 그 자리에서 바로 호출(서브프로세스 아님 - 순수
    파이썬+xlwings 호출이라 콘솔 창 걱정이 없고, 결과 건수를 그대로 돌려줄
    수 있어 마감 버튼처럼 toast에 몇 건 반영됐는지 보여줄 수 있음).

    2026-08-28 사용자 요청: workbench가 진짜 메인이 될 때까지, 예전부터
    써온 배송장 엑셀도 손으로 안 맞춰도 되게 하기 위한 임시 다리 -
    workbench가 충분히 안정화되면 이 버튼째로 치우면 됨. 이미 있는 오더는
    건드리지 않고 누락된 것만 추가하므로 여러 번 눌러도 안전."""
    import board_sync
    return board_sync.sync()


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


def _running_sap_loop_pids():
    """PIDs of any python process currently running main.py or startup.py as
    the SAP loop, regardless of whether launcher.py or this dashboard started
    it. Mirrors watchdog.py's own detection query so both agree on what
    counts as "a loop is already running" - a live process-list check instead
    of a lock file, so there's nothing to go stale if a process ever dies
    without cleaning up after itself."""
    ps_cmd = (
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "($_.CommandLine -match 'main\\.py' -or $_.CommandLine -match 'startup\\.py') } | "
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


# ---------------------------------------------------------------------------
# In-process SAP loop watchdog - ported from the old standalone watchdog.py.
#
# watchdog.py ran as its own separate console process, started manually via
# launcher.py's "워치독 시작" button, and auto-restarted startup.py whenever
# automation.log went 30+ minutes without a new line (main.py/startup.py's
# SAP GUI Scripting/COM session hanging or dying silently - a real, recurring
# failure mode: confirmed live in automation.log, e.g. 2026-08-05 09:02-09:56,
# a 54-minute silent hang - "VL06O 진입 중..." then nothing until an RPC
# failure - that only recovered because the user happened to notice on the
# workbench dashboard and clicked "SAP Start & Loop" by hand). watchdog.py
# itself stopped running sometime after 2026-07-24 (its own log has exactly
# one line, from that day, and no live python.exe process is running it) -
# nothing since then has started/monitored/restarted it, so every SAP hang
# since has needed a manual restart, which reads as "SAP randomly goes down"
# from the workbench side even though the underlying SAP GUI flakiness
# itself isn't new (the same class of COM error shows up as far back as
# March). Since workbench is the thing actually kept running/open all day
# now, the same restart logic lives here instead, tied to workbench_app.py's
# own process lifetime - a daemon thread started once from serve() - rather
# than a separate console window nobody's watching. watchdog.py is left in
# place and still usable by hand; this doesn't replace it, it just means the
# safety net no longer depends on remembering to click that button.
# ---------------------------------------------------------------------------

WATCHDOG_STALE_MINUTES = REFRESH_INTERVAL_MINUTES + 10   # 30분 무갱신 → 죽은 것으로 판단
WATCHDOG_CHECK_INTERVAL_SEC = 60
WATCHDOG_RESTART_COOLDOWN_SEC = 8 * 60                    # 재시작 직후 재판단 유예
WATCHDOG_LONG_COOLDOWN_SEC = 30 * 60                      # 반복 실패 시 확대 유예
WATCHDOG_FLAP_WINDOW_SEC = 60 * 60
WATCHDOG_FLAP_THRESHOLD = 3
WATCHDOG_BUSINESS_HOUR_START = 8
WATCHDOG_BUSINESS_HOUR_END = 20
WATCHDOG_CONSEC_FAILURE_CYCLES = 2   # 연속 전세션 연결 실패 사이클 수 → mtime staleness와 무관하게 재시작
STOP_AUTOMATION_FLAG = BASE_DIR / "stop_automation.flag"


def _watchdog_log_stale_seconds():
    if not LOG_FILE.exists():
        return None
    return time.time() - os.path.getmtime(LOG_FILE)


# 2026-08-21 실측: startup.py 프로세스가 캐시된(죽은) SAP GUI Scripting COM
# 엔진을 붙잡은 채로 계속 도는 상태(NWBC가 사라진 뒤에도 매 사이클 4세션
# 전부 "연결 실패" 에러만 남기고 "이번 실행 완료"까지 찍음)에서는 로그
# 파일이 매 사이클(기본 20분)마다 계속 갱신되므로 위 mtime 기반 staleness
# 판정이 절대 걸리지 않는다 - 워치독이 "살아있다"고 오판해 하루 종일
# 재시작을 안 함. 그래서 mtime과 별개로, 로그 "내용"을 봐서 최근 사이클들이
# 연속으로 전부 실패였는지도 판정한다.
def _watchdog_recent_cycle_texts(n):
    """automation.log에서 최근 n개 사이클(각 '실행 시작:' 구분)의 텍스트를
    오래된 순으로 반환. 파일이 커질 수 있으니 끝에서 일정 크기만 읽는다."""
    try:
        with LOG_FILE.open("rb") as fh:
            fh.seek(0, os.SEEK_END)
            size = fh.tell()
            fh.seek(max(0, size - 200_000))
            data = fh.read()
        text = data.decode("utf-8", errors="replace")
    except Exception:
        return []
    parts = text.split("실행 시작:")
    cycles = ["실행 시작:" + p for p in parts[1:]]
    return cycles[-n:] if cycles else []


def _watchdog_cycle_all_sessions_failed(cycle_text):
    """한 사이클 안에서 자동 루프가 실제로 도는 세션이 전부 '연결 실패'로
    끝났는지 판단. 2026-08-26부터 자동 루프는 세션0/1(VL06O/VL10G)만 돌고
    세션2/3(ZRMA RLKR/Q2)은 main.py의 AUTO_SCAN_ZRMA=False로 빠졌으므로
    (그 세션들 몫의 '연결 실패:' 로그 자체가 안 생김), 임계값도 4가 아니라
    2 - AUTO_SCAN_ZRMA를 다시 켜면 여기도 4로 되돌려야 한다."""
    return cycle_text.count("연결 실패:") >= 2


def _watchdog_consecutive_total_failures():
    """가장 최근에 완료된("이번 실행 완료" 있는) 사이클들부터 거꾸로 훑어서,
    연속으로 전세션 실패인 사이클이 몇 개인지 센다(중간에 하나라도 성공이
    섞이면 거기서 멈춤). 아직 진행 중인 마지막 사이클은 제외."""
    cycles = _watchdog_recent_cycle_texts(WATCHDOG_CONSEC_FAILURE_CYCLES + 1)
    completed = [c for c in cycles if "이번 실행 완료" in c]
    count = 0
    for c in reversed(completed):
        if _watchdog_cycle_all_sessions_failed(c):
            count += 1
        else:
            break
    return count


def _watchdog_append_log(message):
    try:
        with LOG_FILE.open("a", encoding="utf-8") as fh:
            fh.write(f"{now_text()} [WARNING] [Workbench-Watchdog] {message}\n")
    except Exception:
        pass


def _watchdog_restart_loop(reason):
    pids = _running_sap_loop_pids()
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        except Exception:
            pass
    if pids:
        time.sleep(3)
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    try:
        proc = subprocess.Popen(
            [sys.executable, "startup.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
        )
        _watchdog_append_log(f"{reason} → startup.py 재시작 (PID {proc.pid})")
    except Exception as exc:
        _watchdog_append_log(f"startup.py 재시작 실패: {exc}")


def _sap_loop_watchdog_thread():
    """Runs forever in a daemon thread started once from serve(). Same
    detect-and-restart rules as the old watchdog.py (see module docstring
    above) - respects a manual stop_automation.flag, only acts during
    business hours, and backs off if restarting keeps not fixing it."""
    restart_history = []
    cooldown_until = 0.0
    while True:
        time.sleep(WATCHDOG_CHECK_INTERVAL_SEC)
        now = time.time()

        if STOP_AUTOMATION_FLAG.exists():
            continue
        # 2026-09-13: watchdog.py의 동일 지점과 같은 이유(holiday_check.py
        # 참고) - STOP_AUTOMATION_FLAG가 무슨 이유로든 없어진 채로 주말/공휴일
        # 업무시간을 맞아도 여기서 한 번 더 걸러 무인 재기동을 막는다.
        if skip_reason():
            continue
        hour = datetime.now().hour
        if not (WATCHDOG_BUSINESS_HOUR_START <= hour < WATCHDOG_BUSINESS_HOUR_END):
            continue
        if now < cooldown_until:
            continue

        stale = _watchdog_log_stale_seconds()
        is_stale = stale is not None and stale > WATCHDOG_STALE_MINUTES * 60
        consec_fail = _watchdog_consecutive_total_failures()
        is_flatlined = consec_fail >= WATCHDOG_CONSEC_FAILURE_CYCLES
        if not (is_stale or is_flatlined):
            continue

        if is_stale:
            reason = f"automation.log {WATCHDOG_STALE_MINUTES}분 이상 무갱신 감지"
        else:
            reason = f"연속 {consec_fail}회 전세션 SAP 연결 실패 감지"

        restart_history = [t for t in restart_history if now - t < WATCHDOG_FLAP_WINDOW_SEC]
        _watchdog_restart_loop(reason)
        restart_history.append(now)

        if len(restart_history) >= WATCHDOG_FLAP_THRESHOLD:
            cooldown_until = now + WATCHDOG_LONG_COOLDOWN_SEC
        else:
            cooldown_until = now + WATCHDOG_RESTART_COOLDOWN_SEC


def sap_status():
    pids = _running_sap_loop_pids()
    return {"loopRunning": bool(pids), "pids": pids}


def sap_start_loop():
    """"1. SAP Start & Loop" - refuses to start a second loop (from either
    front-end) on top of one that's already running, instead of silently
    launching a process that will fight the existing one over the same 4 SAP
    sessions.

    2026-08-25: startup.py used to run in its own CREATE_NEW_CONSOLE window,
    and closing that window was the only way to stop the loop (the message
    below used to say "그 창을 먼저 종료하세요"). Now that every spawned
    script runs windowless (see sap_stop_loop()), a start here also clears
    STOP_AUTOMATION_FLAG - otherwise a loop stopped via the "중지" button
    would come back up but the watchdog thread would stay silently disabled
    forever (it skips entirely whenever that flag file exists, see
    _sap_loop_watchdog_thread())."""
    pids = _running_sap_loop_pids()
    if pids:
        return {
            "started": False,
            "message": f"이미 SAP 자동루프가 실행 중입니다 (PID {', '.join(map(str, pids))}). "
                       "새로 시작하려면 먼저 중지하세요.",
        }
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    proc = subprocess.Popen(
        [sys.executable, "startup.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def sap_stop_loop():
    """"SAP 자동루프 중지" (2026-08-25) - startup.py/main.py가 더 이상
    CREATE_NEW_CONSOLE 창으로 안 뜨면서, 그 창을 닫는 게 유일한 중지 수단이던
    게 없어져서 새로 추가한 버튼. STOP_AUTOMATION_FLAG를 남겨(워치독이 곧바로
    되살리지 못하게) 지금 떠있는 루프 프로세스를 taskkill로 종료한다. 다시
    시작하려면 sap_start_loop()/"1. SAP Start & Loop"가 이 플래그를 지운다."""
    pids = _running_sap_loop_pids()
    STOP_AUTOMATION_FLAG.write_text("", encoding="utf-8")
    if not pids:
        return {"stopped": False, "message": "실행 중인 SAP 자동루프가 없습니다."}
    killed = []
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            killed.append(pid)
        except Exception:
            pass
    return {"stopped": True, "message": f"SAP 자동루프 중지됨 (PID {', '.join(map(str, killed))})"}


def workbench_force_restart():
    """2026-08-28 신설: 트레이 아이콘의 "강제 재시작"과 완전히 동일한 동작을
    workbench 페이지 안에서도 쓸 수 있게 함(사용자가 "트레이에 있는 걸 여기서도
    바꿀 수 있냐" 요청). 지금 이 요청을 처리하는 프로세스 자신을 죽여야 하므로,
    죽이는 동작은 이 프로세스가 살아있는 채로 직접 하지 않고 - 응답을 먼저
    보내야 하고, taskkill이 자기 자신을 겨냥하면 응답 전송이 끊길 수 있음 -
    별도의 분리된(detached) PowerShell 한 줄짜리 헬퍼를 띄워서 0.8초 뒤에
    실행하게 위임한다. 그 헬퍼가 workbench_app.py 프로세스를 전부 taskkill한
    뒤(단일 인스턴스 뮤텍스는 프로세스 종료 시 OS가 자동으로 풀어줌) 하나만
    새로 띄운다."""
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    pythonw = str(pythonw) if pythonw.exists() else sys.executable
    ps_cmd = (
        "Start-Sleep -Milliseconds 800; "
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.Name -in @('python.exe','pythonw.exe') -and "
        "$_.CommandLine -match 'workbench_app\\.py' } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }; "
        "Start-Sleep -Milliseconds 500; "
        f"Start-Process -FilePath '{pythonw}' -ArgumentList 'workbench_app.py' -WorkingDirectory '{BASE_DIR}'"
    )
    try:
        subprocess.Popen(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        logger.warning("Workbench 강제 재시작 요청 (페이지에서 트리거)")
        return {"restarting": True, "message": "1~2초 후 재시작됩니다. 잠시 후 이 페이지를 새로고침하세요."}
    except Exception as exc:
        logger.exception("Workbench 강제 재시작 트리거 실패")
        return {"restarting": False, "message": f"재시작 트리거 실패: {exc}"}


def sap_run_all():
    """"2. SAP 전체 조회" - if a loop is already running somewhere, just signal
    it via the same run_now.flag file launcher.py's "즉시 조회" button uses
    (no new process touching the shared sessions). If nothing is running,
    spawn bare `main.py` (no --once) - it does an immediate full pass through
    all 4 sessions first thing inside run_loop(), then keeps looping every
    REFRESH_INTERVAL_MINUTES on its own, same as "1. SAP Start & Loop" minus
    the SAP-launch/session-setup steps (this assumes the 4 sessions already
    exist, same assumption --once made)."""
    pids = _running_sap_loop_pids()
    if pids:
        (BASE_DIR / "run_now.flag").write_text("", encoding="utf-8")
        return {"mode": "signal", "message": "실행 중인 자동루프에 즉시 조회 신호를 보냈습니다."}
    if STOP_AUTOMATION_FLAG.exists():
        try:
            STOP_AUTOMATION_FLAG.unlink()
        except Exception:
            pass
    proc = subprocess.Popen(
        [sys.executable, "main.py"], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {
        "mode": "spawned_loop", "pid": proc.pid,
        "message": "실행 중인 루프가 없어 전체 조회 후 자동 루프를 새로 시작했습니다.",
    }


SESSION_LABELS = {0: "VL06O", 1: "VL10G", 2: "ZRMA RLKR", 3: "ZRMA Q2"}


def sap_run_session(idx):
    """세션 하나만 즉시 조회 - launcher.py(0minjung.bat)의 "특정 창만 조회"
    버튼과 완전히 같은 메커니즘(run_now_N.flag)을 워크벤치 대시보드에도
    노출한다(2026-08-26, 사용자가 launcher.py 대신 워크벤치를 쓴다고 요청).
    now.py도 이 파일을 씀 - 셋 다 같은 트리거를 공유. 세션2/3(ZRMA RLKR/Q2)은
    main.py의 AUTO_SCAN_ZRMA=False로 20분 자동루프에서 빠졌으니, 그 두 개를
    확인하고 싶을 때 이 버튼이 사실상 유일한 수단이다.

    run_loop()가 이 플래그를 최대 5초 주기로만 확인하므로(sleep(5) 루프 안),
    자동루프 자체가 안 돌고 있으면 플래그를 써도 아무도 안 지켜봐서 조용히
    무시된다 - sap_run_all()의 "루프 없으면 새로 띄움" 폴백과 달리 여기선
    그렇게 안 한다(세션 하나만 보자고 새 main.py 루프를 통째로 띄우는 건
    과함) - 대신 루프가 없다고 먼저 알려준다."""
    if idx not in SESSION_LABELS:
        return {"started": False, "message": "세션 번호는 0~3이어야 합니다."}
    if not _running_sap_loop_pids():
        return {
            "started": False,
            "message": "실행 중인 SAP 자동루프가 없습니다. 먼저 '1. SAP Start & Loop'를 누르세요.",
        }
    (BASE_DIR / f"run_now_{idx}.flag").write_text("", encoding="utf-8")
    return {"started": True, "message": f"세션{idx} ({SESSION_LABELS[idx]}) 단독 조회 요청됨 (최대 5초 내 시작)"}


def sap_run_order(order_no):
    """"3. 오더번호 반영" - opens its own dedicated SAP session (see
    manual_order_handler._open_new_sap_session), so it's safe to run alongside
    a loop in general. The one exception (ZRX orders briefly reading VL06O
    session 0 for its OBD) is a narrow, low-frequency race that launcher.py's
    equivalent button doesn't guard against either - not worth extra friction
    here for a window this small."""
    order_no = (order_no or "").strip()
    if not order_no:
        return {"started": False, "message": "오더번호를 입력하세요."}
    proc = subprocess.Popen(
        [sys.executable, "order.py", order_no], cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def _zrec_item_rows(item_ids):
    """체크된 item_ids에 대한 order_items+orders 조인 행. ZREC 관련 두
    함수(zrec_lookup_batch/zrec_commit_one)와 sn_relo_one()이 공유."""
    if not item_ids:
        return []
    ensure_db()
    with connect() as con:
        placeholders = ",".join("?" for _ in item_ids)
        rows = con.execute(
            f"SELECT i.*, o.order_no AS o_order_no, o.order_type AS o_order_type, o.id AS o_id "
            f"FROM order_items i JOIN orders o ON o.id = i.order_id "
            f"WHERE i.id IN ({placeholders})",
            item_ids,
        ).fetchall()
    by_id = {r["id"]: dict(r) for r in rows}
    # 원래 요청한 순서(item_ids)대로 반환 - IN절은 순서를 보장하지 않음
    return [by_id[i] for i in item_ids if i in by_id]


def sn_relo_one(item_id):
    """회수 오더 수집 시 SAP상 S/N이 없어서('X') 워크벤치에 X로 찍혔던
    항목 하나를 대상으로 한다 - 사용자가 실제로 회수한 S/N을 이미
    item.serial에 수동 입력해둔 상태라고 가정하고, 그 값을 VA02로 오더를
    열어 Technical Objects > Serial Numbers에 입력·저장 시도한다
    (2026-08-19, 사용자 요청). sn_relo_handler.py를 서브프로세스로 실행 -
    _zrec_run_receive()와 같은 '서브프로세스 하나 = SAP 세션 하나 열고
    쓰고 닫기' 패턴 그대로.

    저장 성공(status='saved')은 그 S/N의 firm/cust가 오더와 일치한다는
    뜻 - 이 항목은 이제 일반 회수 항목과 똑같은 상태이니, 이어서 기존
    "5. ZREC 처리" 버튼으로 처리하면 된다(자동으로 이어붙이지 않음 - 이
    함수는 오더 반영 성공/실패 판정까지만 책임진다). 실패(status='error')는
    firm/cust 불일치로 추정 - paper relo 프로세스가 필요하다는 뜻이라
    sn_relo_error에 사유를 남겨 대시보드에 표시한다. 'skipped'(오더에
    이미 - 워크벤치가 모르던 - 다른 S/N이 들어있던 경우)도 에러는 아니지만
    워크벤치 기록과 다를 수 있어 확인이 필요하므로 마찬가지로 표시한다."""
    items = _zrec_item_rows([item_id])
    if not items:
        raise RuntimeError("item not found")
    item = items[0]

    serial = (item["serial"] or "").strip()
    order_no = (item["o_order_no"] or "").strip()
    order_label = f"{item['o_order_type'] or ''} {order_no}".strip()
    if not serial or serial.upper() == "X":
        raise RuntimeError("실제 회수된 S/N을 먼저 워크벤치에 입력해야 합니다 (현재 S/N 칸이 비어있거나 X)")
    if not order_no:
        raise RuntimeError("오더번호가 없어 처리할 수 없습니다")

    # portal_error와 같은 이유 - 재시도를 시작하는 순간 즉시 비워서, 이번
    # 시도가 끝나기 전까지는 화면에 지난 실패가 안 남아있게 한다.
    with connect() as con:
        con.execute("UPDATE order_items SET sn_relo_error='' WHERE id=?", (item_id,))

    args = [sys.executable, str(BASE_DIR / "sn_relo_handler.py"), "fill",
            "--order", order_no, "--serial", serial]
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=45, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-15:]) if lines else "(출력 없음)"
        with connect() as con:
            con.execute("UPDATE order_items SET sn_relo_error=? WHERE id=?", (f"실행 실패: {detail}", item_id))
        raise RuntimeError(f"S/N 오더 반영 실패 (결과 파싱 불가): {detail}")

    payload = json.loads(json_line[len("RESULT_JSON:"):])
    payload["itemId"] = item_id
    payload["orderLabel"] = order_label

    with connect() as con:
        if payload.get("status") == "saved":
            con.execute("UPDATE order_items SET sn_relo_error='' WHERE id=?", (item_id,))
            add_event(con, item["o_id"], f"S/N '{serial}' 오더 반영 성공: {payload.get('message', '')}")
        else:
            con.execute(
                "UPDATE order_items SET sn_relo_error=? WHERE id=?",
                (payload.get("message") or "알 수 없는 실패", item_id),
            )
            add_event(con, item["o_id"], f"S/N '{serial}' 오더 반영 실패({payload.get('status')}): {payload.get('message', '')}")

    return payload


def zrec_lookup_batch(item_ids):
    """"ZREC 준비" 1단계 (todo c, 2026-08-19) - 체크된 회수 항목들의
    시리얼을 전부 모아서 ZIH08에 **한 번에** 조회한다. 사용자의 실제 하루
    업무 방식("오늘 회수된 모든 시리얼을 한번에 넣고 확인한 뒤, ZREC는
    하나씩") 그대로 - 이전 버전(zrec_handler.py의 prepare 서브커맨드)은
    항목마다 ZIH08을 매번 새로 열어 조회했는데, 그러면 ZIH08 세션을 계속
    재사용/재조회하게 되어 느리고 실제 업무 방식과도 다름. 여기서는
    조회만 하고 ZREC 화면은 전혀 건드리지 않는다 - 실제로 채우고 접수하는
    건 zrec_commit_one()이 항목별로 담당(2026-08-27부터 채우기+클릭이
    한 호출로 통합됨)."""
    items = _zrec_item_rows(item_ids)
    serial_items = [
        i for i in items
        if i["mode"] != "material_only" and (i["serial"] or "").strip()
    ]
    serials = [i["serial"].strip() for i in serial_items]

    ih08_by_serial = {}
    if serials:
        args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "lookup", "--serials", *serials]
        child_env = os.environ.copy()
        child_env["PYTHONIOENCODING"] = "utf-8"
        result = subprocess.run(
            args, cwd=str(BASE_DIR), env=child_env,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
            timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        output = result.stdout or ""
        json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
        if not json_line:
            lines = [ln for ln in output.strip().splitlines() if ln.strip()]
            detail = "\n".join(lines[-10:]) if lines else "(출력 없음)"
            raise RuntimeError(f"ZIH08 조회 실패: {detail}")
        ih08_by_serial = json.loads(json_line[len("RESULT_JSON:"):])

    plans = []
    for item in items:
        serial = (item["serial"] or "").strip()
        material = (item["material"] or "").strip()
        qty = item["qty"] or 1
        non_serial = (item["mode"] == "material_only") or not serial
        order_no = item["o_order_no"] or ""
        order_label = f"{item['o_order_type'] or ''} {order_no}".strip()
        warnings = []
        ih08_info = {}

        if non_serial:
            if not material:
                warnings.append("Material 정보가 없어 ZREC 준비를 할 수 없습니다.")
            filled_order = order_no
            warnings.append("Non-serialized 품목이라 ZIH08 대조를 건너뛰었습니다 - Reference Document를 직접 확인하세요.")
        else:
            ih08_info = ih08_by_serial.get(serial, {})
            matched_order = ih08_info.get("matched_order", "")
            if matched_order:
                filled_order = matched_order
                if order_no and matched_order != order_no:
                    warnings.append(
                        f"workbench 오더번호({order_no})와 ZIH08 매칭 오더(KDAUF={matched_order})가 다릅니다 - "
                        "ZIH08 매칭값을 사용합니다. 반드시 직접 확인하세요."
                    )
            else:
                filled_order = order_no
                warnings.append("ZIH08에서 매칭된 Sales Order(KDAUF)를 찾지 못했습니다 - workbench 오더번호로 대체했습니다. 반드시 직접 확인하세요.")
            if ih08_info.get("plant") == "6507" and ih08_info.get("location") == "0052":
                warnings.append("ZIH08 조회 결과 이 시리얼은 이미 Plant 6507/Location 0052로 접수되어 있습니다 - 중복 접수 주의.")

        plans.append({
            "itemId": item["id"],
            "orderLabel": order_label,
            "serial": serial,
            "material": material,
            "qty": qty,
            "description": item.get("description") or "",
            "nonSerial": non_serial,
            "filledOrder": filled_order,
            "ih08": ih08_info,
            "warnings": warnings,
        })

    return {"plans": plans}


def _zrec_run_receive(item_id, order_no, commit, timeout):
    """zrec_handler.py의 receive 서브커맨드를 서브프로세스로 실행 -
    commit=False면 화면만 채우고 멈춤(dry-run, 수동 디버깅용으로만 남겨둠 -
    workbench는 2026-08-27부터 항상 commit=True로만 호출), commit=True면
    채우기부터 검증, Receive Equipment 클릭까지 한 번에 실행한다.
    zrec_commit_one()의 실행부."""
    items = _zrec_item_rows([item_id])
    if not items:
        raise RuntimeError("item not found")
    item = items[0]

    serial = (item["serial"] or "").strip()
    material = (item["material"] or "").strip()
    qty = item["qty"] or 1
    non_serial = (item["mode"] == "material_only") or not serial
    if non_serial and not material:
        raise RuntimeError("Material 정보가 없어 ZREC 처리를 할 수 없습니다 (Non-serialized인데 Material도 없음)")
    if not non_serial and not serial:
        raise RuntimeError("Serial 정보가 없어 ZREC 처리를 할 수 없습니다")
    order_no = (order_no or item["o_order_no"] or "").strip()
    if not order_no:
        raise RuntimeError("오더번호가 없어 ZREC 처리를 할 수 없습니다")

    args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "receive", "--order", order_no]
    if non_serial:
        args += ["--material", material, "--qty", str(qty), "--non-serial"]
    else:
        args += ["--serial", serial]
    if commit:
        args.append("--commit")

    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=timeout, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-15:]) if lines else "(출력 없음)"
        action = "Receive Equipment 클릭" if commit else "ZREC 채우기"
        raise RuntimeError(f"{action} 실패 (결과 파싱 불가): {detail}")

    payload = json.loads(json_line[len("RESULT_JSON:"):])
    payload["itemId"] = item_id
    payload["orderLabel"] = f"{item['o_order_type'] or ''} {item['o_order_no'] or ''}".strip()
    with connect() as con:
        verb = "ZREC 실제 접수(Receive Equipment 클릭)" if commit else "ZREC 준비"
        add_event(con, item["o_id"], f"{verb} 완료: " + " ".join(args[2:]))
    return payload


def zrec_commit_one(item_id, order_no):
    """"ZREC 준비" 2단계 (todo c, 2026-08-19 - 사용자 확정: 확인 팝업에서
    "확인"을 누르면 사람이 SAP에서 직접 누르는 대신 자동으로 Receive
    Equipment를 클릭) - ZREC 화면에 Plant/Reference Document/Serial(또는
    Non-serialized Material+수량)을 채우고, 화면값을 재확인(검증)한 뒤 곧바로
    Receive Equipment를 클릭해 성공/실패 팝업까지 한 번에 처리한다.

    2026-08-27 이전에는 이 채우기를 별도의 zrec_fill_one() 호출(dry-run)로
    먼저 한 번 하고, 이 함수가 세션을 리셋해서 "안전하게" 똑같은 값을
    다시 채운 뒤 클릭했다. 하지만 배치 확인창(runZrecPrepare 2단계)이
    ZIH08 조회 결과만으로 이미 사람 승인을 받은 뒤 이 함수가 항목마다
    중간 팝업 없이 곧바로 호출되므로, fill_one이 채운 화면을 사람이 보고
    판단하는 순간이 실제로는 없었다 - 세션 리셋+동일 값 재입력은 안전
    효과 없이 SAP 왕복만 두 배로 내는 낭비였다(사용자 지적, 2026-08-27).
    지금은 이 함수 하나가 채우기부터 클릭까지 한 세션에서 끝낸다 -
    안전장치는 그대로 유지: _validate_zrec_fields()가 클릭 직전 화면값을
    재확인하고(엉뚱한 값으로 접수 방지), 클릭 후에는 아래처럼 상태바로
    1차 판정한다.

    **주의**: zrec_handler.py의 성공/실패 판정(_popup_text 문구 매칭)은
    실사용에서 매번 오판(실제 성공을 실패로 표시)하는 게 확인됐다
    (2026-08-19). 상태바 MessageType 우선 판정으로 고쳤지만 그 판정
    자체도 참고용일 뿐이니 결과의 success 값을 곧이곧대로 믿지 말 것 -
    실제 done 처리 여부는 이어지는 zrec_verify_batch()의 ZIH08 재조회
    (Plant/Location 6507/0052 확인)로 최종 결정된다."""
    return _zrec_run_receive(item_id, order_no, commit=True, timeout=30)


def zrec_verify_batch(item_ids):
    """"ZREC 준비" 4단계(완료 확인) - 방금 Receive Equipment를 누른
    항목들의 시리얼을 ZIH08로 재조회해서 Plant/Location이 6507/0052로
    들어왔는지 확인한다 (todo c, 2026-08-19: 사용자 요청 - "zih08 조회후
    6507 0052 들어왔는지 조회 확인"). Non-serialized 항목은 ZIH08이
    시리얼 기반이라 확인 대상에서 제외 - Non-serialized는 ZIH08/시리얼
    개념이 없어 애초에 이 확인 방법이 적용 안 됨."""
    items = _zrec_item_rows(item_ids)
    serial_items = [i for i in items if i["mode"] != "material_only" and (i["serial"] or "").strip()]
    serials = [i["serial"].strip() for i in serial_items]
    if not serials:
        return {"results": []}

    args = [sys.executable, str(BASE_DIR / "zrec_handler.py"), "verify", "--serials", *serials]
    child_env = os.environ.copy()
    child_env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(
        args, cwd=str(BASE_DIR), env=child_env,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
        timeout=60, creationflags=subprocess.CREATE_NO_WINDOW,
    )
    output = result.stdout or ""
    json_line = next((ln for ln in output.splitlines() if ln.startswith("RESULT_JSON:")), None)
    if not json_line:
        lines = [ln for ln in output.strip().splitlines() if ln.strip()]
        detail = "\n".join(lines[-10:]) if lines else "(출력 없음)"
        raise RuntimeError(f"ZREC 완료 확인 실패: {detail}")
    by_serial = json.loads(json_line[len("RESULT_JSON:"):])

    results = []
    for item in serial_items:
        serial = item["serial"].strip()
        row = by_serial.get(serial, {})
        results.append({
            "itemId": item["id"],
            "orderLabel": f"{item['o_order_type'] or ''} {item['o_order_no'] or ''}".strip(),
            "serial": serial,
            "receivedOk": bool(row.get("received_ok")),
            "plant": row.get("plant", ""),
            "location": row.get("location", ""),
        })
    return {"results": results}


def sap_open_session(tcode):
    """"4. New SAP session open" - just asks SAP to spawn a new window; never
    touches sessions 0-3's screens, safe to run anytime including mid-loop
    (open_session.py already anticipated this - session lookups go through
    the SessionNumber map, not raw index, precisely so an extra session
    appearing mid-run doesn't shift anything)."""
    args = [sys.executable, "open_session.py"]
    tcode = (tcode or "").strip()
    if tcode:
        args.append(tcode)
    proc = subprocess.Popen(args, cwd=str(BASE_DIR), creationflags=subprocess.CREATE_NO_WINDOW)
    return {"started": True, "pid": proc.pid}


def sap_open_order(order_no):
    """"오더 바로가기" - 새 SAP 세션을 열어 오더번호 앞자리로 VA02(6*)/VA03
    (그 외)을 자동판별해서 그 오더 화면으로 바로 들어간다(open_session.py의
    open_order()). "3. 오더번호 반영"(sap_run_order, order.py)과 달리 항목
    추출/workbench 반영/카카오 전송을 전혀 안 하는 순수 화면 바로가기라
    아무 때나(자동루프 중에도) 안전하게 눌러도 된다 - sap_open_session()과
    같은 이유로 세션 0~3을 건드리지 않음."""
    order_no = (order_no or "").strip()
    if not order_no:
        return {"started": False, "message": "오더번호를 입력하세요."}
    proc = subprocess.Popen(
        [sys.executable, "open_session.py", "order", order_no], cwd=str(BASE_DIR),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    return {"started": True, "pid": proc.pid}


def _run_hidden_with_toast(args, label, cwd=None):
    """2026-08-25: replaces the CREATE_NEW_CONSOLE pattern several one-shot
    scripts used purely so their own print()'d success/failure text was
    visible immediately (morning_routine.py / portal_login.py / print_zpl_file.py
    - see each caller's docstring below for why). Runs windowless, captures
    stdout+stderr instead of letting them go to a console that no longer
    exists, and reports the outcome as a Windows toast once the process
    exits - a background watcher thread, same shape as run_pod_update()'s."""
    proc = subprocess.Popen(
        args, cwd=str(cwd or BASE_DIR),
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
        encoding="utf-8", errors="replace", creationflags=subprocess.CREATE_NO_WINDOW,
    )

    def watcher():
        output = proc.stdout.read() if proc.stdout else ""
        code = proc.wait()
        try:
            from win_notify import send_windows_notification
            if code == 0:
                send_windows_notification(f"{label} 완료", "", duration="short")
            else:
                lines = [ln.strip() for ln in output.strip().splitlines() if ln.strip()]
                tail = lines[-1] if lines else f"exit {code}"
                send_windows_notification(f"{label} 실패", tail, duration="long")
        except Exception:
            logger.warning(f"{label} 결과 토스트 실패(무시)", exc_info=True)

    threading.Thread(target=watcher, daemon=True).start()
    return proc.pid


def sap_run_morning_routine():
    """"0. 아침 루틴 강제 실행" - morning_routine.py를 지금 바로 수동으로
    돌린다. 매일 08:30 Windows 작업 스케줄러(SAP_Morning_Routine 작업)가
    자동으로 실행해주지만, 혹시 그게 안 돌았을 때를 대비한 수동 대체
    버튼(사용자 요청, 2026-08-14). morning_routine.py 자체가 이미 켜져
    있는 건 건너뛰는 로직이라 SAP 루프/workbench가 이미 떠 있어도 안전.
    2026-08-25: 예전엔 콘솔을 띄워서 뭐가 실행됐는지 눈으로 바로 확인했지만,
    이제 창 없이 돌리는 대신 끝나면 토스트로 결과를 알린다(morning_routine.py
    자체 로그는 morning_routine.log에 그대로 남음)."""
    pid = _run_hidden_with_toast([sys.executable, "morning_routine.py"], "아침 루틴")
    return {"started": True, "pid": pid}


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


def render_notes_page():
    """2026-08-31 사용자 요청: 오더/날짜와 무관한 잡메모(링크, 전화번호, 계정정보
    등)를 관리할 곳이 없어서 신설. 처음엔 항목별 카드+카테고리 UI로 만들었는데,
    실제로 옮기려는 내용(고객센터 번호/SAP 계정/배송기한 규칙 등 엑셀에 박스로
    적어둔 15개+ 잡다한 블록)이 "카테고리 하나 + 한 줄" 같은 정형 항목이 아니라
    사용자가 힘들다고 함 - 항목 단위 add/카테고리 입력 자체가 안 맞는 모델이었음.
    그래서 카드 모델을 버리고 탭(페이지) 여러 개 + 탭마다 큰 textarea 하나인
    메모장 방식으로 교체 - 엑셀 내용을 그대로 복사/붙여넣기만 하면 되게."""
    pages_json = json.dumps(list_pages(), ensure_ascii=False)
    return NOTES_PAGE_TEMPLATE.replace("__PAGES__", pages_json)


def render_mobile_page():
    """2026-08-26 사용자 요청: 회사 밖에서 폰으로 workbench를 볼 때(원격
    데스크톱 화면을 확대해서 보는 게 아니라) 터치에 맞는 화면이 필요해서 추가.
    데스크톱용 PAGE_TEMPLATE(복잡한 표/컨트롤 패널 전부)을 반응형으로 늘리는
    대신 완전히 별도의 가벼운 페이지로 분리했다 - 기존 데스크톱 화면을 건드릴
    위험 없이, SAP 루프 상태 확인 + 세션별 즉시 조회 버튼 + 오늘/이후 오더를
    카드로 보는 것까지만 딱 필요한 만큼만 담는다. 실제 SAP 화면 스크린샷은
    안 보여준다(사용자 확인: workbench가 긁어온 오더 데이터면 충분) - 이미
    board_data()가 보여주는 게 SAP에서 방금 긁어온 그 데이터 그대로다."""
    data_json = json.dumps(board_data(), ensure_ascii=False)
    return MOBILE_PAGE_TEMPLATE.replace("__BOARD_DATA__", data_json)


def render_dashboard_page():
    data_json = json.dumps(board_data(), ensure_ascii=False)
    custom_types_json = json.dumps(list_custom_types(), ensure_ascii=False)
    nonworking_json = json.dumps(nonworking_days_payload(), ensure_ascii=False)
    return (
        PAGE_TEMPLATE
        .replace("__BOARD_DATA__", data_json)
        .replace("__CUSTOM_TYPES__", custom_types_json)
        .replace("__NONWORKING_DAYS__", nonworking_json)
        .replace("__DELETED_ITEM_RETENTION_DAYS__", str(DELETED_ITEM_RETENTION_DAYS))
    )


MOBILE_PAGE_TEMPLATE = (TEMPLATES_DIR / "mobile.html").read_text(encoding="utf-8")


NOTES_PAGE_TEMPLATE = (TEMPLATES_DIR / "notes.html").read_text(encoding="utf-8")



PAGE_TEMPLATE = (TEMPLATES_DIR / "dashboard.html").read_text(encoding="utf-8")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        with WORKBENCH_ACCESS_LOG.open("a", encoding="utf-8") as fh:
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
            if parsed.path == "/mobile":
                payload = render_mobile_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/notes":
                payload = render_notes_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/api/pages":
                self.send_json({"pages": list_pages()})
                return
            if parsed.path == "/api/orders":
                qs = parse_qs(parsed.query)
                self.send_json({"orders": list_orders(qs.get("q", [""])[0], qs.get("status", [""])[0])})
                return
            if parsed.path == "/api/board_data":
                # 새 오더가 SAP 수집으로 workbench.db에 들어와도 이미 열려있는
                # 브라우저 탭은 새로고침 전엔 모른다 - 클라이언트가 이 엔드포인트를
                # 주기적으로 폴링해서 새 itemId만 감지해 추가한다(사용자 요청,
                # 2026-08-14). board_data()는 page-load 때 __BOARD_DATA__로 쓰는
                # 것과 완전히 같은 함수라 모양이 항상 일치함.
                self.send_json(board_data())
                return
            if parsed.path == "/api/board_version":
                # board_data()의 가벼운 버전 - 클라이언트가 이걸 자주(1~2초)
                # 찔러보고 바뀐 걸 감지했을 때만 진짜 board_data()를 부른다.
                # 2026-09-01, see board_version()'s own docstring.
                self.send_json({"version": board_version()})
                return
            if parsed.path == "/api/sap/status":
                self.send_json(sap_status())
                return
            if parsed.path == "/api/items/deleted":
                self.send_json({"items": list_deleted_items()})
                return
            if parsed.path == "/api/nonworking_days":
                self.send_json(nonworking_days_payload())
                return
            m = re.match(r"^/api/orders/(\d+)$", parsed.path)
            if m:
                data = get_order(int(m.group(1)))
                if not data:
                    self.send_json({"error": "not found"}, 404)
                else:
                    self.send_json(data)
                return
            m = re.match(r"^/api/process/(\d+)/running$", parsed.path)
            if m:
                self.send_json({"running": pid_running(int(m.group(1)))})
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
            if parsed.path == "/api/import/range":
                body = self.read_json()
                self.send_json(import_excel_range(
                    body.get("sheet") or "", body.get("rowStart") or 1, body.get("rowEnd") or None,
                ))
                return
            if parsed.path == "/api/export":
                body = self.read_json()
                self.send_json(export_dates_to_excel(body.get("dates") or {}))
                return
            if parsed.path == "/api/sync_delivery_excel":
                self.send_json(sync_delivery_excel())
                return
            if parsed.path == "/api/items/delete":
                body = self.read_json()
                self.send_json(delete_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/items/restore":
                body = self.read_json()
                self.send_json(restore_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/nonworking_days/set":
                body = self.read_json()
                self.send_json(set_manual_holiday(body.get("date") or "", bool(body.get("on"))))
                return
            if parsed.path == "/api/items/move_date":
                body = self.read_json()
                self.send_json(move_items_date(body.get("itemIds") or [], body.get("date") or ""))
                return
            if parsed.path == "/api/items/status":
                body = self.read_json()
                self.send_json(set_items_board_status(body.get("itemIds") or [], body.get("status") or ""))
                return
            if parsed.path == "/api/orders/add_manual":
                body = self.read_json()
                self.send_json(add_manual_note(
                    body.get("cells") or [], body.get("date") or "", body.get("itemType") or "",
                    body.get("orderLabel") or "", body.get("cellCount"),
                ))
                return
            if parsed.path == "/api/items/edit_field":
                body = self.read_json()
                self.send_json(update_item_field(body.get("itemId"), body.get("field") or "", body.get("value") or ""))
                return
            if parsed.path == "/api/orders/edit_field":
                body = self.read_json()
                self.send_json(update_order_field(body.get("itemIds") or [], body.get("field") or "", body.get("value") or ""))
                return
            if parsed.path == "/api/orders/edit_delivery_codes":
                body = self.read_json()
                self.send_json(update_delivery_codes(body.get("itemIds") or [], body.get("value") or ""))
                return
            if parsed.path == "/api/items/highlight":
                body = self.read_json()
                self.send_json(set_cell_highlight(body.get("itemId"), body.get("field") or "", body.get("style") or {}))
                return
            if parsed.path == "/api/items/reorder":
                body = self.read_json()
                self.send_json(reorder_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/items/merge":
                body = self.read_json()
                self.send_json(merge_orders(body.get("itemIds") or [], body.get("fields") or []))
                return
            if parsed.path == "/api/items/unmerge":
                body = self.read_json()
                self.send_json(unmerge_orders(body.get("itemIds") or [], body.get("fields") or []))
                return
            if parsed.path == "/api/undo":
                self.send_json(undo_last_action())
                return
            if parsed.path == "/api/sap/start_loop":
                self.send_json(sap_start_loop())
                return
            if parsed.path == "/api/sap/stop_loop":
                self.send_json(sap_stop_loop())
                return
            if parsed.path == "/api/workbench/restart":
                self.send_json(workbench_force_restart())
                return
            if parsed.path == "/api/custom_types/add":
                body = self.read_json()
                self.send_json(add_custom_type(body.get("name") or ""))
                return
            if parsed.path == "/api/pages/add":
                body = self.read_json()
                self.send_json(add_page(body.get("title") or ""))
                return
            if parsed.path == "/api/pages/edit":
                body = self.read_json()
                self.send_json(update_page(body.get("id"), body.get("field") or "", body.get("value") if body.get("value") is not None else ""))
                return
            if parsed.path == "/api/pages/delete":
                body = self.read_json()
                self.send_json(delete_page(body.get("id")))
                return
            if parsed.path == "/api/sap/morning_routine":
                self.send_json(sap_run_morning_routine())
                return
            if parsed.path == "/api/sap/run_all":
                self.send_json(sap_run_all())
                return
            if parsed.path == "/api/sap/run_session":
                body = self.read_json()
                self.send_json(sap_run_session(int(body.get("session", -1))))
                return
            if parsed.path == "/api/sap/run_order":
                body = self.read_json()
                self.send_json(sap_run_order(body.get("orderNo") or ""))
                return
            if parsed.path == "/api/sap/open_session":
                body = self.read_json()
                self.send_json(sap_open_session(body.get("tcode") or ""))
                return
            if parsed.path == "/api/sap/open_order":
                body = self.read_json()
                self.send_json(sap_open_order(body.get("orderNo") or ""))
                return
            if parsed.path == "/api/portal/login_manual":
                self.send_json(portal_login_manual())
                return
            if parsed.path == "/api/portal/print_latest_zpl":
                self.send_json(print_latest_zpl())
                return
            m = re.match(r"^/api/orders/(\d+)/run_portal$", parsed.path)
            if m:
                body = self.read_json()
                self.send_json(run_portal(int(m.group(1)), skip_qr=bool(body.get("skipQr"))))
                return
            if parsed.path == "/api/sn_relo/fill_one":
                body = self.read_json()
                self.send_json(sn_relo_one(body.get("itemId")))
                return
            if parsed.path == "/api/zrec/lookup_batch":
                body = self.read_json()
                self.send_json(zrec_lookup_batch([int(i) for i in (body.get("itemIds") or [])]))
                return
            if parsed.path == "/api/zrec/commit_one":
                body = self.read_json()
                self.send_json(zrec_commit_one(body.get("itemId"), body.get("orderNo") or ""))
                return
            if parsed.path == "/api/zrec/verify_batch":
                body = self.read_json()
                self.send_json(zrec_verify_batch([int(i) for i in (body.get("itemIds") or [])]))
                return
            if parsed.path == "/api/pod/preview":
                body = self.read_json()
                self.send_json(pod_preview(
                    [int(i) for i in (body.get("orderIds") or [])],
                    body.get("signedBy") or "", body.get("datetime") or "",
                    body.get("remarks") or "", bool(body.get("pickupDone")),
                    body.get("itemIds") or [],
                ))
                return
            if parsed.path == "/api/pod/update":
                body = self.read_json()
                self.send_json(run_pod_update(
                    [int(i) for i in (body.get("orderIds") or [])],
                    body.get("signedBy") or "", body.get("datetime") or "",
                    body.get("remarks") or "", bool(body.get("pickupDone")),
                    body.get("itemIds") or [],
                ))
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


# 2026-08-28: workbench_app.py 서버 자체는(SAP 루프와 달리, sap_start_loop()
# 참고) 여태 중복 실행을 전혀 막지 않았다 - workbench.bat을 두 번 실행하면
# 그냥 두 서버가 동시에 뜬다. 실측으로 확인한 결과 Python HTTPServer의
# allow_reuse_address=True 기본값 + Windows의 느슨한 SO_REUSEADDR 처리
# 때문에 두 번째 프로세스도 같은 포트(8765)에 조용히 bind까지 성공해서
# "에러도 없이 어느 쪽이 응답하는지 알 수 없는" 상태가 된다. 실제로
# 사용자가 "workbench가 안 보인다"며 workbench.bat을 손으로 한 번 더
# 실행해서 이 상태가 실제로 재현됨(자세한 경위는 세션 기록 참고).
# 포트/프로세스 목록 확인은 그 자체로 확인-후-실행 사이 경합(race)이
# 남으므로, 원자적인 Win32 named mutex로 "이 프로세스가 유일한 서버인지"를
# 판정한다.
_SINGLE_INSTANCE_MUTEX_NAME = "Global\\MJSuh_Workbench_App_SingleInstance"
_single_instance_mutex_handle = None  # GC/커널 핸들 해제로 락이 풀리지 않도록 계속 참조 유지


def _acquire_single_instance_lock():
    """뮤텍스 획득에 성공하면 True(이 프로세스가 유일한 서버), 이미 다른
    workbench_app.py 서버가 뮤텍스를 쥐고 있으면 False."""
    global _single_instance_mutex_handle
    handle = win32event.CreateMutex(None, False, _SINGLE_INSTANCE_MUTEX_NAME)
    already_running = (win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS)
    if already_running:
        return False
    _single_instance_mutex_handle = handle  # 참조 유지 (프로세스 종료 시 자동 해제됨)
    return True


def serve(host, port, open_browser=True):
    url = f"http://{host}:{port}/"
    if not _acquire_single_instance_lock():
        # 이미 다른 workbench_app.py 서버가 떠 있음 - 새 서버를 띄우는 대신
        # 기존 서버의 브라우저 탭만 열어준다(수동 실행/워치독/아침 루틴
        # 어느 경로로 여기 왔든 동일하게 안전한 동작).
        logger.info(f"workbench_app.py 이미 실행 중인 인스턴스가 있음 - 새 서버 없이 브라우저만 엶: {url}")
        if open_browser:
            webbrowser.open(url)
        return
    # 위 뮤텍스가 최종 방어선이지만, 혹시 뮤텍스가 뚫리는 경우에도(예: 다른
    # 사용자 세션) 포트가 이미 쓰이고 있으면 조용히 이중 bind되지 않고
    # 시끄럽게(예외로) 실패하도록 재사용 허용을 꺼둔다.
    ThreadingHTTPServer.allow_reuse_address = False
    ensure_db()
    threading.Thread(target=_sap_loop_watchdog_thread, daemon=True).start()
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Workbench running: {url}", flush=True)
    logger.info(f"Workbench 서버 시작: {url} (PID {os.getpid()})")
    # 2026-08-25: workbench.bat/워치독 재시작 모두 pythonw(콘솔 창 없음)로
    # 바뀌면서, "떠 있나"를 눈으로 확인할 창이 사라졌다 - 그 대신 매 시작마다
    # 토스트 한 번씩 띄워서 최소한의 확인 수단을 남겨둔다. 죽었다가 워치독이
    # 재시작하는 비정상 케이스는 workbench_watchdog_check.py가 재시작 '직전'에
    # 별도의 더 눈에 띄는 토스트를 먼저 띄우므로, 이건 항상 뒤이어 오는 평범한
    # "정상 기동" 확인용 - 실패해도(윈도우 알림 API 자체 오류 등) 서버 구동을
    # 막으면 안 되므로 예외는 무조건 삼킨다.
    try:
        from win_notify import send_windows_notification
        send_windows_notification("Workbench 시작됨", url, duration="short")
    except Exception:
        logger.warning("시작 토스트 알림 실패(무시)", exc_info=True)
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except Exception:
        logger.exception("serve_forever()가 처리되지 않은 예외로 종료됨 - 이게 workbench가 죽은 실제 원인")
        raise
    finally:
        # serve_forever()는 정상 상황에서 절대 반환되지 않는다 - 여기 도달했다는
        # 것 자체가 곧 workbench 프로세스 종료라는 뜻이므로 항상 남긴다.
        logger.warning("Workbench 서버 종료(serve_forever 반환/예외)")


def main():
    parser = argparse.ArgumentParser(description="MJSuh local operations workbench")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-open", action="store_true")
    # Manual/emergency-only from here on - see load_today_from_excel()'s own
    # docstring for why. The routine SAP->workbench path is --import-rows-json.
    parser.add_argument("--import-today", action="store_true")
    parser.add_argument(
        "--import-excel-range", nargs=3, metavar=("SHEET", "START_ROW", "END_ROW"),
        help="한 시트의 특정 행 범위만 1회성으로 불러오기 (비상용) - 예: --import-excel-range 8-6 10 25",
    )
    # The routine path: main.py/manual_order_handler.py call this right after
    # writing a freshly-collected order to Excel, passing the exact same rows
    # as a JSON file (avoids a huge argv string) - see import_sap_rows().
    parser.add_argument("--import-rows-json", metavar="PATH")
    args = parser.parse_args()
    ensure_db()
    if args.import_today:
        print(json.dumps(load_today_from_excel(), ensure_ascii=False, indent=2))
        return
    if args.import_excel_range:
        sheet, start_row, end_row = args.import_excel_range
        print(json.dumps(import_excel_range(sheet, start_row, end_row), ensure_ascii=False, indent=2))
        return
    if args.import_rows_json:
        with open(args.import_rows_json, "r", encoding="utf-8") as fh:
            rows = json.load(fh)
        print(json.dumps(import_sap_rows(rows), ensure_ascii=False, indent=2))
        return
    serve(args.host, args.port, open_browser=not args.no_open)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        logger.exception("workbench_app.py 최상위에서 처리되지 않은 예외로 프로세스 종료")
        raise

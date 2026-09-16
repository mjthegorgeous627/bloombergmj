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


MOBILE_PAGE_TEMPLATE = r'''
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<title>Bloomberg Dashboard (Mobile)</title>
<style>
:root{--bg:#f1f4f9;--card:#fff;--line:#e2e8f0;--text:#0f172a;--sub:#64748b;--accent:#2563eb}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;background:var(--bg);color:var(--text);font-family:-apple-system,"Malgun Gothic",sans-serif;
  font-size:15px;padding-bottom:24px}
.topbar{position:sticky;top:0;z-index:10;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;
  padding:10px 14px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 6px rgba(0,0,0,.15)}
.topbar h1{font-size:15px;margin:0;font-weight:700}
#clock{font-size:11px;color:#cbd5e1}
#loopStatus{font-size:12px;display:flex;align-items:center;gap:5px;margin-top:2px}
#loopStatus .dot{width:8px;height:8px;border-radius:50%;background:#64748b;display:inline-block}
#loopStatus.running .dot{background:#22c55e;box-shadow:0 0 5px #22c55e}
#loopStatus.running{color:#86efac}
#loopStatus.stopped .dot{background:#f87171}
#loopStatus.stopped{color:#fca5a5}

.panel{background:var(--card);margin:10px;border-radius:12px;padding:10px;box-shadow:0 1px 3px rgba(15,23,42,.08)}
.btn-row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.btn-row.cols4{grid-template-columns:repeat(2,1fr)}
.btn{border:1px solid #d1d5db;background:#fff;border-radius:10px;padding:12px 8px;font-size:14px;
  font-weight:600;color:#334155;text-align:center}
.btn:active{background:#f1f5f9;transform:translateY(1px)}
.btn.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
.btn.warn{background:#dc2626;color:#fff;border-color:#dc2626}
.btn.ghost{background:#f8fafc}
.section-title{font-size:11px;font-weight:700;color:var(--sub);text-transform:uppercase;
  letter-spacing:.4px;margin:4px 2px 8px}

#filterBox{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:15px;margin:0 10px 6px;
  width:calc(100% - 20px)}
.date-chips{display:flex;gap:6px;overflow-x:auto;padding:0 10px 8px;-webkit-overflow-scrolling:touch}
.date-chip{flex:0 0 auto;padding:6px 12px;border-radius:999px;background:#e2e8f0;color:#334155;
  font-size:12.5px;font-weight:600;white-space:nowrap}
.date-chip.today{background:#2563eb;color:#fff}

.date-heading{margin:16px 10px 6px;font-size:13px;font-weight:700;color:var(--sub)}
.order-card{background:var(--card);margin:0 10px 10px;border-radius:12px;padding:12px;
  box-shadow:0 1px 3px rgba(15,23,42,.08);border-left:4px solid #cbd5e1}
.order-card.st-done{border-left-color:#2563eb;opacity:.7}
.order-card.st-cancelled{border-left-color:#9ca3af;opacity:.55}
.order-card.st-error{border-left-color:#dc2626}
.oc-head{display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:6px}
.oc-label{font-weight:700;font-size:14.5px}
.oc-badges{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end}
.badge{font-size:10.5px;font-weight:700;padding:2px 7px;border-radius:999px;white-space:nowrap}
.badge.type{background:#e0e7ff;color:#3730a3}
.badge.err{background:#fee2e2;color:#b91c1c}
.badge.done{background:#dbeafe;color:#1e40af}
.badge.cancelled{background:#e5e7eb;color:#4b5563;text-decoration:line-through}
.oc-code{font-size:11.5px;color:var(--sub);white-space:pre-line;margin-bottom:6px}
.oc-items{margin:6px 0;padding:8px 10px;background:#f8fafc;border-radius:8px}
.oc-item{font-size:13px;line-height:1.5}
.oc-item + .oc-item{margin-top:4px;padding-top:4px;border-top:1px dashed var(--line)}
.oc-mn{color:var(--sub)}
.oc-row{display:flex;gap:6px;font-size:13.5px;margin-top:4px}
.oc-row .k{color:var(--sub);flex:0 0 auto;width:40px}
.oc-row .v{flex:1}
.oc-row a{color:var(--accent);text-decoration:none;font-weight:600}
.empty{text-align:center;color:var(--sub);padding:30px 10px;font-size:13.5px}

.toast{position:fixed;left:50%;bottom:20px;transform:translate(-50%,20px);background:#0f172a;color:#fff;
  padding:10px 16px;border-radius:10px;font-size:13.5px;opacity:0;transition:opacity .2s,transform .2s;
  z-index:50;max-width:88vw;text-align:center}
.toast.show{opacity:1;transform:translate(-50%,0)}
</style>
</head>
<body>
<div class="topbar">
  <div><h1>Bloomberg Dashboard</h1><div id="loopStatus"><span class="dot"></span><span class="txt">확인 중...</span></div></div>
  <div id="clock"></div>
</div>

<div class="panel">
  <div class="section-title">SAP 자동루프</div>
  <div class="btn-row">
    <button class="btn primary" onclick="startLoop()">Start &amp; Loop</button>
    <button class="btn warn" onclick="stopLoop()">중지</button>
  </div>
</div>

<div class="panel">
  <div class="section-title">세션별 즉시 조회 (루프 실행 중이어야 동작)</div>
  <div class="btn-row cols4">
    <button class="btn ghost" onclick="runSession(0)">VL06O</button>
    <button class="btn ghost" onclick="runSession(1)">VL10G</button>
    <button class="btn ghost" onclick="runSession(2)">ZRMA RLKR</button>
    <button class="btn ghost" onclick="runSession(3)">ZRMA Q2</button>
  </div>
</div>

<div class="panel">
  <div class="btn-row" style="grid-template-columns:1fr">
    <button class="btn primary" onclick="reloadBoard()">↻ 오더 목록 새로고침</button>
  </div>
</div>

<input id="filterBox" placeholder="오더번호/고객명/주소/메모 검색...">
<div class="date-chips" id="dateChips"></div>
<div id="boardArea"></div>

<script>
let BOARD = __BOARD_DATA__;

function tick(){document.getElementById('clock').textContent=new Date().toLocaleString('ko-KR',{hour:'2-digit',minute:'2-digit',second:'2-digit'})}
setInterval(tick,1000);tick();

function toast(msg){
  const t=document.createElement('div');
  t.className='toast';
  t.textContent=msg;
  document.body.appendChild(t);
  requestAnimationFrame(()=>t.classList.add('show'));
  setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},2200);
}

async function updateLoopStatus(){
  const el=document.getElementById('loopStatus');
  try{
    const res=await fetch('/api/sap/status');
    const data=await res.json();
    el.classList.remove('running','stopped');
    if(data.loopRunning){
      el.classList.add('running');
      el.querySelector('.txt').textContent='루프 실행 중 (PID '+data.pids.join(', ')+')';
    }else{
      el.classList.add('stopped');
      el.querySelector('.txt').textContent='루프 중지됨';
    }
  }catch(err){
    el.classList.remove('running','stopped');
    el.querySelector('.txt').textContent='상태 확인 실패';
  }
}
updateLoopStatus();
setInterval(updateLoopStatus,8000);

async function startLoop(){
  try{
    const res=await fetch('/api/sap/start_loop',{method:'POST'});
    const data=await res.json();
    toast(data.started ? '시작됨 (PID '+data.pid+')' : data.message);
    updateLoopStatus();
  }catch(err){ toast('시작 실패: '+err.message); }
}
async function stopLoop(){
  if(!confirm('SAP 자동루프를 중지할까요?')) return;
  try{
    const res=await fetch('/api/sap/stop_loop',{method:'POST'});
    const data=await res.json();
    toast(data.message);
    updateLoopStatus();
  }catch(err){ toast('중지 실패: '+err.message); }
}
async function runSession(idx){
  try{
    const res=await fetch('/api/sap/run_session',{
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({session:idx})
    });
    const data=await res.json();
    toast(data.message);
  }catch(err){ toast('조회 실패: '+err.message); }
}

async function reloadBoard(){
  try{
    const res=await fetch('/api/board_data');
    BOARD=await res.json();
    renderBoard();
    toast('오더 목록 새로고침됨');
  }catch(err){ toast('새로고침 실패: '+err.message); }
}

function statusClass(row){
  if(row.boardStatus==='done') return 'st-done';
  if(row.boardStatus==='cancelled') return 'st-cancelled';
  if(row.portalError||row.podError||row.snReloError) return 'st-error';
  return '';
}

function orderCard(rows){
  const first=rows[0];
  const div=document.createElement('div');
  div.className='order-card '+statusClass(first);

  const head=document.createElement('div'); head.className='oc-head';
  const label=document.createElement('div'); label.className='oc-label'; label.textContent=first.orderLabel||'(번호없음)';
  const badges=document.createElement('div'); badges.className='oc-badges';
  if(first.itemType){ const b=document.createElement('span'); b.className='badge type'; b.textContent=first.itemType; badges.appendChild(b); }
  if(first.boardStatus==='done'){ const b=document.createElement('span'); b.className='badge done'; b.textContent='완료'; badges.appendChild(b); }
  if(first.boardStatus==='cancelled'){ const b=document.createElement('span'); b.className='badge cancelled'; b.textContent='취소'; badges.appendChild(b); }
  if(first.portalError){ const b=document.createElement('span'); b.className='badge err'; b.textContent='Portal 오류'; badges.appendChild(b); }
  if(first.podError){ const b=document.createElement('span'); b.className='badge err'; b.textContent='POD 오류'; badges.appendChild(b); }
  if(first.snReloError){ const b=document.createElement('span'); b.className='badge err'; b.textContent='S/N 오류'; badges.appendChild(b); }
  head.appendChild(label); head.appendChild(badges);
  div.appendChild(head);

  if(first.deliveryNo){ const c=document.createElement('div'); c.className='oc-code'; c.textContent=first.deliveryNo; div.appendChild(c); }

  const items=document.createElement('div'); items.className='oc-items';
  rows.forEach(r=>{
    const it=document.createElement('div'); it.className='oc-item';
    const parts=[r.description,r.material,r.serial].filter(Boolean);
    it.textContent = parts.length ? parts.join(' · ') : '(품목 정보 없음)';
    items.appendChild(it);
  });
  div.appendChild(items);

  const infoRows=[
    ['고객', first.customer],
    ['주소', first.address],
    ['메모', first.memo],
  ];
  infoRows.forEach(([k,v])=>{
    if(!v) return;
    const r=document.createElement('div'); r.className='oc-row';
    const kk=document.createElement('div'); kk.className='k'; kk.textContent=k;
    const vv=document.createElement('div'); vv.className='v'; vv.textContent=v;
    r.appendChild(kk); r.appendChild(vv);
    div.appendChild(r);
  });
  if(first.phone){
    const r=document.createElement('div'); r.className='oc-row';
    const kk=document.createElement('div'); kk.className='k'; kk.textContent='전화';
    const vv=document.createElement('div'); vv.className='v';
    const a=document.createElement('a'); a.href='tel:'+first.phone.replace(/[^0-9+]/g,''); a.textContent=first.phone;
    vv.appendChild(a);
    r.appendChild(kk); r.appendChild(vv);
    div.appendChild(r);
  }
  return div;
}

function matchesFilter(row, q){
  if(!q) return true;
  const hay=[row.orderLabel,row.orderNo,row.customer,row.address,row.memo,row.description,row.material,row.serial]
    .filter(Boolean).join(' ').toLowerCase();
  return hay.includes(q);
}

function renderBoard(){
  const area=document.getElementById('boardArea');
  const chips=document.getElementById('dateChips');
  area.innerHTML=''; chips.innerHTML='';
  const q=(document.getElementById('filterBox').value||'').trim().toLowerCase();

  let anyRendered=false;
  BOARD.forEach(dateGroup=>{
    const rows=dateGroup.rows.filter(r=>matchesFilter(r,q));
    const chip=document.createElement('a');
    chip.className='date-chip'+(dateGroup.label.includes('(Today)')?' today':'');
    chip.textContent=dateGroup.label.replace(/\s*\(Today\)|\s*\(Tomorrow\)/,m=>m.includes('Today')?' 오늘':' 내일');
    chip.href='#d-'+dateGroup.iso;
    chips.appendChild(chip);
    if(!rows.length) return;
    anyRendered=true;

    const heading=document.createElement('div');
    heading.className='date-heading'; heading.id='d-'+dateGroup.iso;
    heading.textContent=dateGroup.label+' · '+rows.length+'건';
    area.appendChild(heading);

    const byOrder=new Map();
    rows.forEach(r=>{
      if(!byOrder.has(r.orderId)) byOrder.set(r.orderId,[]);
      byOrder.get(r.orderId).push(r);
    });
    byOrder.forEach(orderRows=>area.appendChild(orderCard(orderRows)));
  });

  if(!anyRendered){
    const e=document.createElement('div'); e.className='empty';
    e.textContent = q ? '검색 결과가 없습니다.' : '표시할 오더가 없습니다.';
    area.appendChild(e);
  }
}

document.getElementById('filterBox').addEventListener('input', renderBoard);
renderBoard();
</script>
</body>
</html>
'''


NOTES_PAGE_TEMPLATE = r'''
<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bloomberg Dashboard - 메모</title>
<style>
*{box-sizing:border-box}
html,body{height:100%;margin:0}
:root{font-family:-apple-system,"Segoe UI","Malgun Gothic",Arial,sans-serif;
  color:#1e293b;background:#f1f4f9;--line:#e2e8f0;--line-strong:#cbd5e1;--accent:#2563eb}
body{display:flex;flex-direction:column}

.topbar{height:38px;flex:0 0 auto;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;
  display:flex;align-items:center;justify-content:space-between;padding:0 18px;
  box-shadow:0 1px 3px rgba(0,0,0,.15)}
.topbar h1{font-size:14.5px;margin:0;font-weight:600;letter-spacing:.2px}
.topbar-right{display:flex;align-items:center;gap:10px}
#saveState{font-size:11px;color:#86efac;opacity:0;transition:opacity .3s}
#saveState.show{opacity:1}
.topbar a{color:#cbd5e1;text-decoration:none;font-size:11.5px;border:1px solid #475569;
  border-radius:5px;padding:3px 8px}
.topbar a:hover{background:rgba(255,255,255,.1)}

.tabbar{flex:0 0 auto;display:flex;align-items:flex-end;gap:2px;background:#e2e8f0;
  padding:8px 12px 0;overflow-x:auto}
.tab{display:flex;align-items:center;gap:6px;background:#dde3ec;color:#475569;
  padding:8px 14px;border-radius:8px 8px 0 0;font-size:12.5px;font-weight:600;
  cursor:pointer;white-space:nowrap;border:1px solid transparent;border-bottom:none;
  max-width:220px}
.tab:hover{background:#eef1f5}
.tab.active{background:#fff;color:#0f172a;border-color:var(--line)}
.tab .t-title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;outline:none}
.tab .t-title[contenteditable="true"]{cursor:text}
.tab .t-close{color:#94a3b8;font-size:12px;line-height:1;padding:1px 3px;border-radius:3px}
.tab .t-close:hover{background:#fee2e2;color:#dc2626}
.tab-add{flex:0 0 auto;background:none;border:none;color:#64748b;font-size:16px;
  cursor:pointer;padding:6px 10px;line-height:1}
.tab-add:hover{color:#1e293b}

.editor-wrap{flex:1;min-height:0;background:#fff;padding:0}
#editor{width:100%;height:100%;border:none;outline:none;resize:none;padding:18px 22px;
  font-family:"Cascadia Mono","Consolas","Malgun Gothic",monospace;font-size:13px;
  line-height:1.65;color:#1e293b;background:#fff}
#editor::placeholder{color:#94a3b8;font-family:inherit}

.empty-state{flex:1;display:flex;align-items:center;justify-content:center;
  flex-direction:column;gap:10px;color:#94a3b8;font-size:13px}
.empty-state button{height:32px;border:1px solid #d1d5db;background:#fff;border-radius:7px;
  padding:0 16px;cursor:pointer;font-size:13px;color:#334155}
.empty-state button:hover{background:#f8fafc}
</style>
</head>
<body>
<div class="topbar">
  <h1>Bloomberg Dashboard - 메모</h1>
  <div class="topbar-right"><span id="saveState">저장됨</span><a href="/">← 대시보드</a></div>
</div>
<div class="tabbar" id="tabbar"></div>
<div class="editor-wrap" id="editorWrap" style="display:none">
  <textarea id="editor" placeholder="여기에 자유롭게 입력하거나 엑셀 내용을 그대로 붙여넣으세요..."></textarea>
</div>
<div class="empty-state" id="emptyState" style="display:none">
  <div>아직 메모장이 없습니다.</div>
  <button onclick="createPage()">+ 첫 메모장 만들기</button>
</div>
<script>
let PAGES = __PAGES__;
let activeId = PAGES.length ? PAGES[0].id : null;
let saveTimer = null;

async function api(path, body){
  const res = await fetch(path, body ? {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)} : {});
  return res.json();
}

function currentPage(){ return PAGES.find(p=>p.id===activeId); }

function enterTitleEdit(title, p){
  title.contentEditable = 'true';
  title.focus();
  document.execCommand('selectAll', false, null);
}

function renderTabs(){
  const bar = document.getElementById('tabbar');
  bar.innerHTML = '';
  PAGES.forEach(p=>{
    const tab = document.createElement('div');
    tab.className = 'tab' + (p.id===activeId ? ' active' : '');
    tab.onclick = ()=>switchTo(p.id);

    const title = document.createElement('span');
    title.className = 't-title';
    title.textContent = p.title;
    title.dataset.pageId = p.id;
    title.onclick = e=>{ if(title.isContentEditable) e.stopPropagation(); };
    title.ondblclick = e=>{ e.stopPropagation(); enterTitleEdit(title, p); };
    title.onblur = ()=>{
      title.contentEditable = 'false';
      const val = title.textContent.trim();
      if(val && val !== p.title){
        p.title = val;
        api('/api/pages/edit', {id:p.id, field:'title', value:val});
      } else {
        title.textContent = p.title;
      }
    };
    title.onkeydown = e=>{
      if(e.key==='Enter'){ e.preventDefault(); title.blur(); }
      if(e.key==='Escape'){ title.textContent = p.title; title.blur(); }
    };
    tab.appendChild(title);

    if(p.id===activeId){
      const close = document.createElement('span');
      close.className = 't-close';
      close.textContent = '✕';
      close.title = '메모장 삭제';
      close.onclick = e=>{ e.stopPropagation(); removePage(p.id); };
      tab.appendChild(close);
    }
    bar.appendChild(tab);
  });
  const add = document.createElement('button');
  add.className = 'tab-add';
  add.textContent = '+';
  add.title = '새 메모장';
  add.onclick = ()=>createPage();
  bar.appendChild(add);
}

function renderEditor(){
  const wrap = document.getElementById('editorWrap');
  const empty = document.getElementById('emptyState');
  if(!PAGES.length){
    wrap.style.display = 'none';
    empty.style.display = 'flex';
    return;
  }
  wrap.style.display = 'block';
  empty.style.display = 'none';
  const editor = document.getElementById('editor');
  editor.value = (currentPage() || {}).body || '';
}

function switchTo(id){
  flushSave(true);
  activeId = id;
  renderTabs();
  renderEditor();
  document.getElementById('editor').focus();
}

function showSaved(){
  const el = document.getElementById('saveState');
  el.classList.add('show');
  clearTimeout(el._hideTimer);
  el._hideTimer = setTimeout(()=>el.classList.remove('show'), 1200);
}

function scheduleSave(){
  clearTimeout(saveTimer);
  saveTimer = setTimeout(()=>flushSave(false), 700);
}

function flushSave(immediate){
  clearTimeout(saveTimer);
  const page = currentPage();
  if(!page) return;
  const editor = document.getElementById('editor');
  const val = editor.value;
  if(val === page.body) return;
  page.body = val;
  api('/api/pages/edit', {id: page.id, field:'body', value: val}).then(showSaved);
}

document.getElementById('editor').addEventListener('input', scheduleSave);
window.addEventListener('beforeunload', ()=>flushSave(true));

async function createPage(){
  const res = await api('/api/pages/add', {title:'새 메모장'});
  if(res.added){
    PAGES.push({id:res.id, title:res.title, body:'', sort_pos:PAGES.length, created_at:'', updated_at:''});
    activeId = res.id;
    renderTabs();
    renderEditor();
    const titleEl = document.querySelector('.tab.active .t-title');
    if(titleEl) enterTitleEdit(titleEl, currentPage());
  }
}

function removePage(id){
  if(!confirm('이 메모장을 삭제할까요? 내용은 복구할 수 없습니다.')) return;
  PAGES = PAGES.filter(p=>p.id!==id);
  if(activeId===id) activeId = PAGES.length ? PAGES[0].id : null;
  renderTabs();
  renderEditor();
  api('/api/pages/delete', {id});
}

renderTabs();
renderEditor();
</script>
</body>
</html>
'''



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

.sticky-top{position:sticky;top:0;z-index:20;background:#f1f4f9;
  box-shadow:0 4px 10px rgba(15,23,42,.08)}

.topbar{height:38px;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;
  display:flex;align-items:center;justify-content:space-between;padding:0 18px;
  box-shadow:0 1px 3px rgba(0,0,0,.15)}
.topbar h1{font-size:14.5px;margin:0;font-weight:600;letter-spacing:.2px}
.topbar-right{display:flex;align-items:center;gap:12px}
#clock{font-size:11.5px;color:#cbd5e1;font-variant-numeric:tabular-nums}
#sapLoopWrap{position:relative}
#sapLoopStatus{font-size:11.5px;display:flex;align-items:center;gap:5px;cursor:pointer;
  padding:3px 6px;border-radius:5px;user-select:none}
#sapLoopStatus:hover{background:rgba(255,255,255,.1)}
#sapLoopStatus .dot{width:7px;height:7px;border-radius:50%;background:#64748b;display:inline-block}
#sapLoopStatus.running .dot{background:#22c55e;box-shadow:0 0 4px #22c55e}
#sapLoopStatus.running{color:#86efac}
#sapLoopStatus.stopped .dot{background:#f87171}
#sapLoopStatus.stopped{color:#fca5a5}
/* 2026-09-13 사용자 요청: 주말/공휴일(자동화가 안 도는 날)을 달력 숫자로
   빨갛게 보여주고, 그 외 날짜(임시공휴일 등)도 클릭으로 수동 지정 가능하게 -
   시계 옆 📅 버튼을 눌러서 여는 드롭다운, 위 sapLoopMenu와 같은 패턴. */
#calWrap{position:relative}
.cal-toggle-btn{background:none;border:1px solid #475569;border-radius:5px;color:#cbd5e1;
  font-size:13px;padding:3px 7px;cursor:pointer;line-height:1}
.cal-toggle-btn:hover{background:#334155}
#calPanel{display:none;position:absolute;top:100%;right:0;margin-top:4px;
  background:#1e293b;border:1px solid #334155;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.35);
  padding:10px;z-index:50;width:220px}
#calPanel.open{display:block}
.cal-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}
.cal-head button{background:none;border:none;color:#cbd5e1;font-size:13px;cursor:pointer;padding:2px 6px}
.cal-head button:hover{color:#fff}
.cal-title{color:#e2e8f0;font-size:12.5px;font-weight:600}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:2px;text-align:center}
.cal-wd{color:#64748b;font-size:10.5px;padding:2px 0}
.cal-day{color:#e2e8f0;font-size:11.5px;padding:4px 0;border-radius:5px;cursor:pointer}
.cal-day:hover{background:#334155}
.cal-day.nonworking{color:#f87171;font-weight:700}
.cal-day.today{box-shadow:inset 0 0 0 1px #60a5fa}
.cal-legend{margin-top:6px;font-size:10px;color:#94a3b8;line-height:1.4}

/* 2026-08-28: 트레이 아이콘 우클릭 메뉴와 같은 기능(시작/중지/강제 재시작)을
   페이지 안에서도 쓰게 - 시계 옆 상태 점을 눌러서 여는 드롭다운. */
#sapLoopMenu{display:none;position:absolute;top:100%;right:0;margin-top:4px;
  background:#1e293b;border:1px solid #334155;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.35);
  min-width:190px;z-index:50;overflow:hidden}
#sapLoopMenu.open{display:block}
#sapLoopMenu button{display:block;width:100%;text-align:left;background:none;border:none;
  color:#e2e8f0;font-size:12px;padding:9px 12px;cursor:pointer}
#sapLoopMenu button:hover{background:#334155}
#sapLoopMenu button.danger{color:#fca5a5}
#sapLoopMenu hr{border:none;border-top:1px solid #334155;margin:2px 0}

.controls{display:flex;align-items:flex-start;gap:8px;padding:8px 12px 0}
.ctrl-col{flex:1;background:#fff;border:1px solid var(--line);border-radius:8px;
  padding:7px 9px;box-shadow:0 1px 2px rgba(15,23,42,.04)}
.ctrl-col h3{margin:0 0 5px;font-size:11px;font-weight:700;color:#64748b;
  text-transform:uppercase;letter-spacing:.5px;cursor:pointer;user-select:none;
  display:flex;align-items:center;justify-content:space-between}
.ctrl-col-toggle{font-size:10px;color:#94a3b8}
.ctrl-col.collapsed .ctrl-col-body{display:none}
.ctrl-col.collapsed h3{margin-bottom:0}
.ctrl-divider{height:1px;background:var(--line);margin:5px 0}
/* 2026-08-25: 세로로 한 줄씩 쌓이던 짧은 버튼들을 같은 줄에 여러 개 묶어서
   컨트롤 패널 전체 높이를 줄이기 위한 그리드 래퍼 - 컨트롤 패널이 커져서
   그 아래 오더 목록이 안 보이는 문제(사용자 신고, 2026-08-25) 대응. 긴
   설명형 버튼(예: "4-2. S/N 오더 반영 시도...")은 여기 넣지 않고 기존처럼
   한 줄 전체를 씀 - 그리드에 넣으면 줄바꿈이 심해져 오히려 더 높아짐.*/
.btn-grid{display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-bottom:5px}
.btn-grid.cols4{grid-template-columns:repeat(4,1fr)}
.btn-grid .btn.block{margin-bottom:0}

.btn{height:26px;border:1px solid #d1d5db;background:#fff;border-radius:6px;padding:0 10px;
  cursor:pointer;font-size:12px;color:#334155;transition:background .12s,box-shadow .12s,transform .05s}
.btn:hover{background:#f8fafc;box-shadow:0 1px 2px rgba(15,23,42,.06)}
.btn:active{transform:translateY(1px)}
.btn.block{display:block;width:100%;text-align:left;height:auto;padding:5px 9px;
  margin-bottom:4px;line-height:1.25;font-weight:500}
.btn.primary{background:#2563eb;color:#fff;border-color:#2563eb}
.btn.primary:hover{background:#1d4ed8}
.btn.warn{background:#dc2626;border-color:#dc2626;color:#fff;font-weight:700}
.btn.warn:hover{background:#b91c1c}
.btn.finish{background:#0f172a;color:#fff;border-color:#0f172a}
.btn.finish:hover{background:#1e293b}
.btn.muted{background:#6b7280;color:#fff;border-color:#6b7280}
.btn.muted:hover{background:#4b5563}
.btn .hint{font-size:10.5px;font-weight:400;opacity:.8;display:block;margin-top:1px}
.merge-field-row{display:flex;flex-wrap:wrap;gap:2px 10px;font-size:11.5px;color:#475569;
  margin:2px 0 7px;padding:6px 8px;background:#f8fafc;border:1px solid var(--line);border-radius:7px}
.merge-field-row label{display:flex;align-items:center;gap:4px;cursor:pointer;white-space:nowrap}
.merge-field-row input{cursor:pointer}
.ctrl-inline{display:flex;gap:6px;margin-bottom:7px}
.ctrl-inline .btn{flex:1.3;text-align:left}
.ctrl-inline input{flex:1;border:1px solid var(--line-strong);border-radius:7px;padding:0 9px;
  font-size:12.5px;color:#1e293b}
.ctrl-inline input:focus{outline:2px solid #93c5fd;outline-offset:-1px}

.board-wrap{margin:16px;background:#fff;border:1px solid var(--line);border-radius:10px;
  box-shadow:0 1px 2px rgba(15,23,42,.04)}
table.board-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px}
.board-table thead th{position:sticky;top:var(--sticky-top-h,0px);z-index:5;background:#e8720c;color:#fff;
  text-align:center;padding:9px 10px;font-weight:600;letter-spacing:.2px;
  box-shadow:0 1px 0 rgba(0,0,0,.08)}
.board-table thead th:first-child{border-top-left-radius:0}

/* 2026-08-27 user request: 날짜 배너 줄(예: "2026. 08. 26 수요일")이 너무 높아서
   화면에 보이는 오더 행 수가 줄어듦 - padding/font-size를 줄여 한 줄당 높이를
   압축(약 40px→28px), 그만큼 아래 오더 목록이 더 많이 보이도록. */
.date-banner-row td{background:linear-gradient(90deg,#1e3a5f,#25507e);color:#fff;
  font-weight:600;font-size:12.5px;padding:5px 12px;letter-spacing:.2px}
.date-banner-inner{display:flex;align-items:center;justify-content:space-between}
.date-banner-right{display:flex;align-items:center;gap:10px}
.date-export-check{width:14px;height:14px;cursor:pointer}
.date-toggle{background:transparent;border:none;color:#fff;font-size:12px;cursor:pointer;
  padding:1px 8px;border-radius:5px;line-height:1.3}
.date-toggle:hover{background:rgba(255,255,255,.18)}
tbody.date-group.collapsed .item-row,tbody.date-group.collapsed .empty-row{display:none}

/* 2026-08-27 user request: 오더 라인 간격/글자 조금 줄이기 (너무 많이는 말고) -
   board-table font-size 12.5→12px과 함께, 줄 하나당 약 3-4px 정도만 압축. */
.item-row td{padding:5px 10px;border-bottom:1px solid var(--line);vertical-align:middle;
  background:#fff;transition:background .12s}
.item-row:hover td{background:#f8fafc}
.item-row[draggable="true"]{cursor:grab}
.item-row.dragging{opacity:.35}
.item-row.drop-before td{box-shadow:inset 0 2px 0 0 #2563eb}
.item-row.drop-after td{box-shadow:inset 0 -2px 0 0 #2563eb}
.item-row.row-done td{background:#bfdbfe}
.item-row.row-done:hover td{background:#93c5fd}
.item-row.row-pending td{background:#fecaca}
.item-row.row-pending:hover td{background:#fca5a5}
.item-row.row-cancelled td{background:#d1d5db;color:#6b7280;text-decoration:line-through}
.item-row.row-cancelled:hover td{background:#b0b6c0}
/* Portal(Serial/QR) 또는 POD 자동화 실패 표시 (todo b, 2026-08-18) - done(파랑)/
   pending(빨강)과 안 겹치는 주황 계열, 다음 재실행이 성공하면 서버가
   portal_error/pod_error를 비워서 자동으로 사라짐. row-pending 등과 동시에
   붙어도(둘 다 해당하는 행) CSS 선언 순서상 이게 이겨서 실패가 우선 보임. */
.item-row.row-error td{background:#fed7aa;box-shadow:inset 4px 0 0 0 #ea580c}
.item-row.row-error:hover td{background:#fdba74}
.order-error-badge{cursor:pointer;margin-left:4px;color:#ea580c;font-weight:700}

.type-cell{text-align:center}
.type-select{font-weight:700;border:1px solid transparent;background:transparent;
  border-radius:5px;padding:2px 4px;cursor:pointer;font-size:12px}
.type-select:hover{border-color:var(--line-strong)}
.type-delivery{color:#2563eb}
.type-pickup{color:#dc2626}
.type-other{color:#9333ea}
.ord-cell{font-weight:600}
.ord-cell .muted{display:block;font-weight:400;color:#94a3b8;font-size:11px;margin-top:3px;white-space:pre-wrap}
.order-select-all{float:right;width:14px;height:14px;cursor:pointer;margin-left:6px}
.cust-cell{color:#334155;white-space:pre-wrap}
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
.board-table thead th:nth-child(3){width:4%}
.board-table thead th:nth-child(4){width:9%}
.board-table thead th:nth-child(5){width:15%}
.board-table thead th:nth-child(6){width:6%}
.board-table thead th:nth-child(7){width:7%}
.board-table thead th:nth-child(8){width:7%}
.board-table thead th:nth-child(9){width:6%}
.board-table thead th:nth-child(10){width:16%}
.board-table thead th:nth-child(11){width:8%}
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
.item-row td.item-cell{width:15%;white-space:pre-wrap}
.item-row td.mn-cell{width:6%;text-align:center;white-space:pre-wrap}
.item-row td.sn-cell{width:7%;text-align:center;white-space:pre-wrap}
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

.trash-overlay{display:none;position:fixed;inset:0;background:rgba(15,23,42,.45);z-index:100;
  align-items:flex-start;justify-content:center;padding:40px 20px}
.trash-overlay.show{display:flex}
.trash-panel{background:#fff;border-radius:10px;max-width:1000px;width:100%;max-height:85vh;
  overflow-y:auto;padding:16px 20px;box-shadow:0 12px 40px rgba(0,0,0,.3)}
.trash-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.trash-header h3{font-size:15px}
.trash-list{margin-top:10px}
.trash-table{width:100%;border-collapse:collapse;font-size:12px}
.trash-table th{text-align:left;padding:6px 8px;border-bottom:2px solid var(--line-strong);
  color:#64748b;font-weight:600}
.trash-table td{padding:6px 8px;border-bottom:1px solid var(--line)}
.trash-empty{padding:20px;text-align:center;color:#94a3b8;font-size:12.5px}

/* 배송장 인쇄: hidden on screen, only shown (and only content left visible)
   inside @media print. Built fresh from scratch, plain black-on-white table -
   deliberately NOT reusing .board-table's classes/colors, so the printed
   slip stays forced-B&W regardless of the board's on-screen styling or the
   browser's own "background graphics" print setting. */
#printArea{display:none}
.print-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:9.5px;color:#000;
  -webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact}
.print-table th,.print-table td{border:1px solid #000;padding:3px 4px;text-align:left;
  vertical-align:top;white-space:pre-wrap;word-break:break-word;background:#fff;color:#000}
.print-table th{font-weight:700;text-align:center}
.print-table td:nth-child(4),.print-table td:nth-child(5){text-align:center}
.print-table tr{break-inside:avoid}
/* 배송 강조: 흑백 인쇄에서도 배경 음영은 회색 톤으로 남는다(위 print-color-adjust:exact
   가 브라우저의 "배경 그래픽" 인쇄 옵션 체크 여부와 무관하게 강제 출력되게 함).
   혹시 그마저 안 찍히는 환경 대비로 테두리 두께도 같이 바꿔 이중 안전장치를 둠 -
   회수는 기존 그대로(흰 배경/얇은 테두리), 배송만 음영+굵은 테두리로 대비. */
.print-table td.print-type-delivery{background:#c2c2c2;font-weight:700;border-width:2px}
#printArea h2{font-size:15px;margin:0 0 6px;color:#000}
#printArea h3{font-size:12px;margin:14px 0 4px;color:#000}
#printArea h3:first-of-type{margin-top:0}
@media print{
  @page{size:A4 landscape;margin:10mm}
  body>*:not(#printArea){display:none!important}
  #printArea{display:block!important}
}

.hl-menu{position:fixed;z-index:200;background:#fff;border:1px solid var(--line-strong);
  border-radius:8px;box-shadow:0 8px 24px rgba(15,23,42,.18);padding:4px;
  display:flex;flex-direction:column;min-width:130px}
.hl-menu-btn{display:block;width:100%;text-align:left;border:none;background:none;
  padding:6px 10px;font-size:12px;border-radius:5px;cursor:pointer;color:#334155}
.hl-menu-btn:hover{background:#f1f5f9}
.hl-menu-btn.hl-swatch-yellow{background:#fde047}
.hl-menu-btn.hl-swatch-yellow:hover{background:#facc15}
.hl-menu-btn.hl-swatch-pink{background:#f9a8d4}
.hl-menu-btn.hl-swatch-pink:hover{background:#f472b6}
.hl-menu-btn.hl-swatch-red{color:#dc2626;font-weight:700}
</style>
</head>
<body>
<div class="sticky-top">
<div class="topbar"><h1>Bloomberg Dashboard Seoul</h1><div class="topbar-right"><a href="/notes" style="color:#cbd5e1;text-decoration:none;font-size:11.5px;border:1px solid #475569;border-radius:5px;padding:3px 8px" title="오더/날짜와 무관한 잡메모(링크, 전화번호 등)">📌 메모</a><div id="calWrap"><button type="button" class="cal-toggle-btn" onclick="toggleCalendarPanel(event)" title="자동화 비운영일 달력 - 빨간 숫자는 주말/공휴일(또는 수동 지정)이라 그날 SAP 자동화가 돌지 않습니다. 숫자를 클릭하면 수동으로 빨간날을 지정/해제할 수 있습니다">📅</button><div id="calPanel"></div></div><div id="sapLoopWrap"><div id="sapLoopStatus" onclick="toggleSapLoopMenu(event)"><span class="dot"></span><span class="txt">SAP 루프 확인 중...</span></div><div id="sapLoopMenu">
  <button onclick="runSap('start');closeSapLoopMenu()">SAP 루프 시작</button>
  <button onclick="runSap('stop');closeSapLoopMenu()">SAP 루프 중지</button>
  <hr>
  <button class="danger" onclick="restartWorkbench();closeSapLoopMenu()">Workbench 강제 재시작</button>
</div></div><div id="clock"></div></div></div>

<div class="controls">
  <div class="ctrl-col" id="col-sap">
    <h3 onclick="toggleCtrlCol('col-sap')">SAP<span class="ctrl-col-toggle">▾</span></h3>
    <div class="ctrl-col-body">
      <!-- 2026-08-27 user request: SAP 칼럼이 다른 3개 칼럼보다 2줄만큼 더 길어서
           보드(오더 목록)가 보이는 칸이 줄어듦 - 옆 칼럼들과 높이를 맞추려고 기존
           7줄을 5줄로 압축. 3680행 주석의 경고(긴 설명형 버튼을 그리드에 넣으면
           줄바꿈이 심해져 오히려 더 길어짐)를 피하려고, 아래 두 그리드는 원래도
           짧았던 라벨은 그대로 두고 긴 라벨만 축약했다 - 잘려나간 상세 설명은
           전부 title 툴팁에 그대로 남아있음 (호버로 확인 가능). -->
      <div class="btn-grid cols4">
        <button class="btn block" onclick="runMorningRoutine()" title="0. 아침 루틴 강제 실행 - 매일 08:30 작업 스케줄러가 자동 실행하는 것과 동일 - 그게 혹시 안 돌았을 때 수동으로 대신 실행. 이미 켜져 있는 건 건너뜀">0.아침루틴</button>
        <button class="btn block" onclick="runSap('start')" title="1. SAP Start & Loop">1.Start&amp;Loop</button>
        <button class="btn block warn" onclick="runSap('stop')" title="SAP 자동루프 중지 - 콘솔 창이 없어서(2026-08-25) 더 이상 창을 닫아서 멈출 수 없음 - 대신 이 버튼으로 중지. 워치독이 되살리지 않도록 표시해두고, 다시 켜려면 왼쪽 Start&Loop를 누르면 됨">루프중지</button>
        <button class="btn block" onclick="runSap('all')" title="2. SAP 전체 조회">2.전체조회</button>
      </div>
      <div class="btn-grid cols4" title="자동루프가 이미 돌고 있어야 동작 (최대 5초 내 시작). ZRMA RLKR/Q2는 2026-08-26부터 20분 자동루프에서 빠져있어 확인하려면 여기를 눌러야 함">
        <button class="btn" onclick="runSapSession(0)">VL06O</button>
        <button class="btn" onclick="runSapSession(1)">VL10G</button>
        <button class="btn" onclick="runSapSession(2)">ZRMA&nbsp;RLKR</button>
        <button class="btn" onclick="runSapSession(3)">ZRMA Q2</button>
      </div>
      <div class="ctrl-inline">
        <button class="btn" onclick="runOrderNoAction()">3. 오더번호 반영</button>
        <input id="orderNoInput" placeholder="오더번호" onkeydown="if(event.key==='Enter'){event.preventDefault();runOrderNoAction();}">
      </div>
      <div class="ctrl-inline" title="새 SAP 세션을 열어 오더번호 앞자리로 VA02(6*)/VA03(그 외)을 자동판별해서 그 오더 화면으로 바로 들어감 - 항목수집/workbench반영/카카오전송 없이 화면만 여는 순수 바로가기(위 3번 오더번호 반영과 다름)">
        <button class="btn" onclick="runOrderJumpAction()">오더 바로가기</button>
        <input id="orderJumpInput" placeholder="오더번호" onkeydown="if(event.key==='Enter'){event.preventDefault();runOrderJumpAction();}">
      </div>
      <div class="ctrl-inline">
        <button class="btn" onclick="runNewSessionAction()">4. New SAP session open</button>
        <input id="tcodeInput" placeholder="T-code (선택)" onkeydown="if(event.key==='Enter'){event.preventDefault();runNewSessionAction();}">
      </div>
      <div class="btn-grid">
        <button class="btn block" onclick="runSnReloBatch()" title="4-2. S/N 오더 반영 시도 (X→실제 S/N, paper relo 판별) - S/N칸에 실제 회수된 S/N을 먼저 입력해두고 체크 - 오더 원본(VA02)에 그 S/N을 입력·저장 시도합니다. 저장되면 firm/cust 일치 → 이어서 5번 ZREC 처리 가능. 저장 안 되면(에러) firm/cust 불일치로 paper relo 필요 - 항목에 ⚠ 표시됩니다.">4-2. S/N 오더 반영 시도</button>
        <button class="btn block" onclick="runZrecPrepare()" title="5. ZREC 처리 (ZIH08 조회 → 표로 확인 → 자동 Receive) - 체크된(Ready to process) 회수 항목의 시리얼을 ZIH08에 한 번에 조회 → 오더번호/M-N/S-N/수량/아이템명 표를 확인창 하나로 모아 보여줌 → 확인을 누르면 항목마다 순서대로 자동으로 Receive Equipment 클릭(실제 SAP 접수! 중간 팝업 없음) → 전체 끝나면 ZIH08로 6507/0052 접수 확인까지, 확인된 항목은 파란색(완료)으로 자동 전환">5. ZREC 처리</button>
      </div>
    </div>
  </div>
  <div class="ctrl-col" id="col-portal">
    <h3 onclick="toggleCtrlCol('col-portal')">Bloomberg Portal<span class="ctrl-col-toggle">▾</span></h3>
    <div class="ctrl-col-body">
      <button class="btn block" onclick="openPortalLoginManual()" title="Chrome만 열림 - 자동입력/자동클릭 없음, 직접 로그인+B-Unit 인증. 켜둔 채로 재사용됨">Portal 로그인 (수동)</button>
      <div class="btn-grid">
        <button class="btn block primary" onclick="runPortalBatch()" title="체크된 행의 오더별로 확인창 후 하나씩 순서대로 실행 (동시 실행 안 함)">1. Serial 등록 &amp; QR 인쇄</button>
        <button class="btn block" onclick="runLatestZpl()" title="Downloads 폴더에서 가장 최근 .ZPL/.zpl(또는 PDF 라벨)을 찾아 Zebra 프린터로 바로 전송">2. 최근 ZPL 수동 인쇄</button>
      </div>
      <div class="btn-grid">
        <button class="btn block" onclick="runPodBatch()" title="체크된 행 기준으로 Portal POD(배송완료) 처리 - 배송만 체크하면 배송만, 배송+회수 둘 다 체크하면 둘 다 완료로 저장. 오더별 미리보기 확인 후 저장">3. POD (배송완료) 처리</button>
        <button class="btn block" onclick="runSpareOneshot()" title="누르면 POD 내용 확인창이 먼저 뜨고, 확인하면 Serial 등록부터 POD(배송완료) 저장까지 중간 팝업 없이 한 번에 진행 (QR 인쇄는 생략)">4. 스페어 배송처리 (serial+POD처리)</button>
      </div>
      <div class="ctrl-divider"></div>
      <div class="btn-grid cols4">
        <button class="btn block primary" onclick="processDone()" title="체크된 행 전체 완료 (파랑)">Done</button>
        <button class="btn block warn" onclick="processPickupReady()" title="체크된 회수만 대기 (빨강), 나머지 완료">Ready</button>
        <button class="btn block muted" onclick="processCancel()" title="체크된 행 전체 취소 (회색) - 행 전체에 표시, 삭제되지는 않음">Cancel</button>
        <button class="btn block" onclick="processReset()" title="체크된 행을 Done/Ready/Cancel 이전 상태(흰색)로 되돌림 - 잘못 완료 처리됐거나 아직 진행 안 된 행을 표시 없는 상태로 초기화">초기화</button>
      </div>
    </div>
  </div>
  <div class="ctrl-col" id="col-excel">
    <h3 onclick="toggleCtrlCol('col-excel')">Excel<span class="ctrl-col-toggle">▾</span></h3>
    <div class="ctrl-col-body">
      <div class="ctrl-inline">
        <input type="date" id="moveDateInput" onchange="moveCheckedToDate(this.value); this.value='';" title="체크된 행을 이 날짜로 이동 (목록에 없는 날짜도 새로 생김)">
      </div>
      <button class="btn block finish" onclick="finalizeChecked()" title="날짜 줄의 체크박스를 누르면 그 날짜의 행이 전부 기본 체크됨 - 아직 해결 안 된 행은 체크만 해제하면 이번 마감에서 빠지고 보드에 그대로 남음(며칠 뒤 그 행만 체크해서 마감하면 원래 날짜의 시트 맨 아래에 이어붙음). 엑셀 저장이 성공해야만 workbench.db에서도 삭제됨">마감 (내보내기 + 삭제)</button>
      <button class="btn block" onclick="printChecked()" title="날짜 줄의 체크박스(날짜 전체) 또는 행 체크박스(오더 단위, 한쪽 leg만 체크해도 같은 오더 전체가 인쇄됨)로 선택 - A4 가로, 흑백">배송장 인쇄</button>
      <button class="btn block" onclick="syncDeliveryExcel()" title="workbench에 있는데 아직 기존 배송장 엑셀(2026 배송장.xlsx)에 없는 오더만 골라 오늘 시트의 해당 날짜 섹션에 채워 넣음 - 이미 있는 건 건드리지 않으니 여러 번 눌러도 안전. workbench가 완전히 메인이 될 때까지 기존 배송장 엑셀도 손으로 안 맞추기 위한 임시 버튼(2026-08-28)">기존 배송장 엑셀로 동기화</button>
      <div class="ctrl-divider"></div>
      <div class="ctrl-inline" title="평상시엔 필요 없음 - SAP 수집이 workbench.db에 직접 반영되므로 자동 동기화는 더 이상 엑셀을 다시 읽지 않음. 엑셀에 직접 입력/수정한 내용을 워크벤치로 가져와야 할 때만 시트명과 행 범위를 지정해 1회성으로 사용">
        <input type="text" id="excelRangeSheet" placeholder="시트명 (예: 8-6)" size="10">
        <input type="number" id="excelRangeStart" placeholder="시작행" min="1" style="width:5em">
        <input type="number" id="excelRangeEnd" placeholder="종료행" min="1" style="width:5em">
      </div>
      <button class="btn block" onclick="importExcelRange()" title="위 시트의 지정한 행 범위만 1회성으로 workbench.db에 불러옵니다 - 비상시에만 사용">3. 엑셀에서 범위 불러오기 (비상용)</button>
    </div>
  </div>
  <div class="ctrl-col" id="col-cleanup">
    <h3 onclick="toggleCtrlCol('col-cleanup')">정리<span class="ctrl-col-toggle">▾</span></h3>
    <div class="ctrl-col-body">
      <div class="btn-grid">
        <button class="btn block" onclick="addManualRow()" title="오더 가져오기 없이, 텍스트만으로 새 행 1개를 오늘 날짜에 추가 - 실제 SAP 오더가 아니므로 Portal/ZREC 등 자동화 버튼 대상으로는 쓰지 마세요. 다른 날짜로 옮기려면 추가된 뒤 체크하고 아래 Excel 칸의 날짜 이동을 쓰세요">0. 메모 행 추가</button>
        <button class="btn block" onclick="deleteChecked()" title="workbench.db에서도 삭제됨 - 실행취소로 되돌리기 가능">1. 체크된 행 삭제</button>
      </div>
      <div class="merge-field-row">
        <label><input type="checkbox" class="merge-field-cb" value="order">Order#</label>
        <label><input type="checkbox" class="merge-field-cb" value="customer" checked>Customer</label>
        <label><input type="checkbox" class="merge-field-cb" value="phone" checked>Phone</label>
        <label><input type="checkbox" class="merge-field-cb" value="address" checked>Address</label>
        <label><input type="checkbox" class="merge-field-cb" value="memo" checked>Memo</label>
      </div>
      <div class="btn-grid">
        <button class="btn block" onclick="mergeChecked()" title="같은 날짜 안에서 서로 붙어있어야 화면에 병합되어 보임 - 떨어져 있으면 드래그(⠿)로 옆에 붙이세요">2. 체크된 행 병합</button>
        <button class="btn block" onclick="unmergeChecked()">3. 병합 해제</button>
      </div>
      <div class="ctrl-divider"></div>
      <div class="btn-grid">
        <button class="btn block" onclick="undoLastAction()" title="삭제/날짜이동/완료처리를 최근 20건까지 되돌림">실행취소 (Ctrl+Z)</button>
        <button class="btn block" onclick="openTrash()" title="실행취소 범위(최근 20건) 밖의 오래된 삭제도 7일 안이면 여기서 복구 가능">삭제된 항목 보기 (휴지통)</button>
      </div>
    </div>
  </div>
</div>
</div>

<div id="trashOverlay" class="trash-overlay">
  <div class="trash-panel">
    <div class="trash-header">
      <h3>삭제된 항목 (7일간 보관 후 완전 삭제)</h3>
      <button class="btn" onclick="closeTrash()">닫기</button>
    </div>
    <button class="btn block primary" onclick="restoreCheckedTrash()">선택 항목 복구</button>
    <div id="trashList" class="trash-list"></div>
  </div>
</div>

<!-- 2026-08-20: 메모 행 추가 모달 - 오더 가져오기 없이 텍스트만으로 board에
     새 행을 만드는 기능(사용자 요청). order#는 항상 빈 값으로 저장되고
     (add_manual_note() 참고 - 실제 오더와 안 헷갈리게), 구분/날짜는 여기서
     사용자가 직접 선택. trash-overlay와 같은 오버레이 패턴 재사용.
     2026-08-25: 내용 칸이 원래 Item(description) 하나뿐이었는데, 사용자
     요청으로 Type/check/Order#를 제외한 나머지 7칸(Item/M-N/S-N/Customer/
     Phone/Address/Memo) 전부를 자유롭게 쓸 수 있게 바꿈 - 메모 행은 실제
     오더가 아니라서 그 7칸에 각각 "무엇을 적어야 하는지" 강제할 이유가
     없다는 게 요청 취지. 다만 처음부터 7칸을 다 보여주면 오히려 복잡해
     보이므로(사용자 지적), 몇 칸을 쓸지 먼저 고르면 그 수만큼만 렌더링 -
     renderManualRowCells() 참고. -->
<div id="manualRowOverlay" class="trash-overlay">
  <div class="trash-panel" style="max-width:420px">
    <div class="trash-header">
      <h3>메모 행 추가</h3>
      <button class="btn" onclick="closeManualRowModal()">닫기</button>
    </div>
    <label style="display:block;margin-bottom:10px;font-size:13px;color:var(--muted,#555)">Order# (선택 - 비워두면 "memo"로 표시됨)
      <input type="text" id="manualRowOrderLabel" placeholder="예: SDSK1333608359" style="display:block;width:100%;margin-top:4px;padding:6px 8px;box-sizing:border-box">
    </label>
    <label style="display:block;margin-bottom:10px;font-size:13px;color:var(--muted,#555)">몇 칸을 쓸까요? (제목/M-N/S-N/Customer/Phone/Address/Memo 중 앞에서부터 - 기본값은 일반 오더와 같은 7칸 전부, 필요하면 줄여서 몇 칸만 합쳐 쓸 수 있음)
      <input type="number" id="manualRowCellCount" min="1" max="7" value="7" style="width:70px;margin-left:8px;padding:4px 6px;box-sizing:border-box" onchange="renderManualRowCells()">
    </label>
    <div id="manualRowCells" style="margin-bottom:4px"></div>
    <label style="display:block;margin:10px 0;font-size:13px;color:var(--muted,#555)">구분
      <select id="manualRowType" style="width:100%;margin-top:4px;padding:6px 8px;box-sizing:border-box"></select>
    </label>
    <label style="display:block;margin-bottom:14px;font-size:13px;color:var(--muted,#555)">날짜
      <input type="date" id="manualRowDate" style="width:100%;margin-top:4px;padding:6px 8px;box-sizing:border-box">
    </label>
    <div style="text-align:right">
      <button class="btn" onclick="closeManualRowModal()">취소</button>
      <button class="btn primary" onclick="submitManualRow()">추가</button>
    </div>
  </div>
</div>

<div class="board-wrap">
  <table class="board-table" id="boardTable">
    <thead><tr>
      <th></th><th>Type</th><th>check</th><th>Order #</th><th>Item</th><th>M/N</th><th>S/N</th>
      <th>Customer</th><th>Phone</th><th>Address</th><th>Memo</th>
    </tr></thead>
  </table>
</div>

<div id="printArea"></div>

<script>
function tick(){document.getElementById('clock').textContent=new Date().toLocaleString('ko-KR',{weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'})}
setInterval(tick,1000);tick();

// 콘솔 창이 없어진(2026-08-25) 뒤로 SAP 자동루프가 실제로 돌고 있는지 화면에서
// 확인할 방법이 없어져서 추가 - 시계 옆에 점+문구로 표시. 백엔드 /api/sap/status는
// 이미 있었는데(2026-08-25) 프론트에서 아무도 안 부르고 있었음.
async function updateLoopStatus(){
  const el=document.getElementById('sapLoopStatus');
  try{
    const res=await fetch('/api/sap/status');
    const data=await res.json();
    el.classList.remove('running','stopped');
    if(data.loopRunning){
      el.classList.add('running');
      el.querySelector('.txt').textContent='SAP 루프 실행 중 (PID '+data.pids.join(', ')+')';
    }else{
      el.classList.add('stopped');
      el.querySelector('.txt').textContent='SAP 루프 중지됨';
    }
  }catch(err){
    el.classList.remove('running','stopped');
    el.querySelector('.txt').textContent='SAP 루프 상태 확인 실패';
  }
}
updateLoopStatus();
setInterval(updateLoopStatus,5000);

function toggleCtrlCol(id){
  const col=document.getElementById(id);
  const collapsed=col.classList.toggle('collapsed');
  col.querySelector('.ctrl-col-toggle').textContent=collapsed?'▸':'▾';
  syncStickyTopHeight();
}

// Keeps the board's sticky column-header row (already position:sticky) glued
// directly beneath the sticky-top control bar instead of overlapping it.
// Re-measured whenever the control bar's own height can change (column
// collapse/expand, window resize) since that height isn't fixed.
function syncStickyTopHeight(){
  const el=document.querySelector('.sticky-top');
  if(el) document.documentElement.style.setProperty('--sticky-top-h', el.offsetHeight+'px');
}
window.addEventListener('resize', syncStickyTopHeight);
syncStickyTopHeight();

function toast(msg){
  const t=document.createElement('div');
  t.className='toast';
  t.textContent=msg;
  document.body.appendChild(t);
  requestAnimationFrame(()=>t.classList.add('show'));
  setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},1900);
}

// -- Control panel: structure-only placeholders, wired to real actions next --
async function runMorningRoutine(){
  try{
    const res=await fetch('/api/sap/morning_routine', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('아침 루틴 강제 실행됨 (PID '+data.pid+') - 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('아침 루틴 실행 실패: '+err.message);
  }
}
async function runSap(kind){
  try{
    if(kind==='start'){
      const res=await fetch('/api/sap/start_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'시작 실패');
      toast(data.started ? 'SAP 자동루프 시작됨 (PID '+data.pid+')' : data.message);
      updateLoopStatus();
    } else if(kind==='stop'){
      if(!window.confirm('SAP 자동루프를 지금 중지할까요? (워치독도 다시 켤 때까지 재시작하지 않습니다)')) return;
      const res=await fetch('/api/sap/stop_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'중지 실패');
      toast(data.message);
      updateLoopStatus();
    } else if(kind==='all'){
      const res=await fetch('/api/sap/run_all', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
      toast(data.message);
    }
  }catch(err){
    toast('SAP '+kind+' 실패: '+err.message);
  }
}
// 2026-08-28: 트레이 아이콘 우클릭 메뉴와 같은 조작을 시계 옆 상태 점에서도
// 할 수 있게 - 클릭하면 열리는 드롭다운. 문서 아무 곳이나 클릭하면 닫힘.
function toggleSapLoopMenu(ev){
  ev.stopPropagation();
  document.getElementById('sapLoopMenu').classList.toggle('open');
}
function closeSapLoopMenu(){
  document.getElementById('sapLoopMenu').classList.remove('open');
}
document.addEventListener('click', closeSapLoopMenu);

// 2026-09-13 사용자 요청: 시계 옆 📅 버튼 - 주말/한국 공휴일(자동화 비운영일)
// 을 달력 숫자로 빨갛게 보여주고, 클릭으로 수동 지정/해제도 가능하게. 색
// 판단은 이미 있는 nonworkingInfo()/MANUAL_HOLIDAYS를 그대로 재사용 -
// 서버의 /api/nonworking_days/set에 저장되는 것도 동일.
let CAL_VIEW = (() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; })();

function _calIso(y, m, day){
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function toggleCalendarPanel(ev){
  ev.stopPropagation();
  const panel = document.getElementById('calPanel');
  const opening = !panel.classList.contains('open');
  panel.classList.toggle('open');
  if(opening){
    const today = new Date();
    CAL_VIEW = { y: today.getFullYear(), m: today.getMonth() };
    renderCalendarPanel();
  }
}
function closeCalendarPanel(){
  document.getElementById('calPanel').classList.remove('open');
}
document.addEventListener('click', closeCalendarPanel);

function calShiftMonth(delta, ev){
  ev.stopPropagation();
  CAL_VIEW.m += delta;
  if(CAL_VIEW.m < 0){ CAL_VIEW.m = 11; CAL_VIEW.y -= 1; }
  if(CAL_VIEW.m > 11){ CAL_VIEW.m = 0; CAL_VIEW.y += 1; }
  renderCalendarPanel();
}

function calToggleDay(iso, ev){
  ev.stopPropagation();
  const willAdd = !MANUAL_HOLIDAYS.has(iso);
  if(willAdd) MANUAL_HOLIDAYS.add(iso); else MANUAL_HOLIDAYS.delete(iso);
  fetch('/api/nonworking_days/set', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({date: iso, on: willAdd}),
  }).catch(() => {});
  renderCalendarPanel();
}

function renderCalendarPanel(){
  const {y, m} = CAL_VIEW;
  const panel = document.getElementById('calPanel');
  const startOffset = new Date(y, m, 1).getDay();  // 0=일요일
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const now = new Date();
  const todayIso = _calIso(now.getFullYear(), now.getMonth(), now.getDate());

  let html = `<div class="cal-head">`
    + `<button type="button" onclick="calShiftMonth(-1,event)">◂</button>`
    + `<span class="cal-title">${y}. ${String(m + 1).padStart(2, '0')}</span>`
    + `<button type="button" onclick="calShiftMonth(1,event)">▸</button>`
    + `</div><div class="cal-grid">`;
  ['일', '월', '화', '수', '목', '금', '토'].forEach(w => { html += `<div class="cal-wd">${w}</div>`; });
  for(let i = 0; i < startOffset; i++) html += `<div></div>`;
  for(let day = 1; day <= daysInMonth; day++){
    const iso = _calIso(y, m, day);
    const info = nonworkingInfo(iso);
    const cls = ['cal-day'];
    if(info.non) cls.push('nonworking');
    if(iso === todayIso) cls.push('today');
    const title = info.non ? info.label : '클릭: 이 날짜를 수동으로 휴무일(빨간색) 지정';
    html += `<div class="${cls.join(' ')}" title="${title}" onclick="calToggleDay('${iso}',event)">${day}</div>`;
  }
  html += `</div><div class="cal-legend">빨간 숫자 = 그날 SAP 자동화 비운영일(주말/공휴일/수동 지정). 숫자를 클릭하면 수동으로 지정·해제할 수 있습니다.</div>`;
  panel.innerHTML = html;
}
async function restartWorkbench(){
  if(!window.confirm('Workbench를 강제로 재시작할까요?\n1~2초간 서버가 끊겼다가 새로 뜨고, 이 페이지는 자동으로 새로고침됩니다.')) return;
  try{
    const res=await fetch('/api/workbench/restart', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'재시작 트리거 실패');
    toast(data.message);
    setTimeout(()=>location.reload(), 3000);
  }catch(err){
    toast('Workbench 재시작 실패: '+err.message);
  }
}
async function runSapSession(idx){
  try{
    const res=await fetch('/api/sap/run_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({session:idx})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast(data.message);
  }catch(err){
    toast('세션 '+idx+' 조회 실패: '+err.message);
  }
}
async function runOrderNoAction(){
  const input=document.getElementById('orderNoInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/run_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    toast('오더번호 반영 시작됨: '+v+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더번호 반영 실패: '+err.message);
  }
}
async function runOrderJumpAction(){
  const input=document.getElementById('orderJumpInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/open_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    const tcode = v.startsWith('6') ? 'VA02' : 'VA03';
    toast('오더 '+v+' ('+tcode+') 여는 중 (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더 바로가기 실패: '+err.message);
  }
}
async function runNewSessionAction(){
  const input=document.getElementById('tcodeInput');
  const v=input.value.trim();
  try{
    const res=await fetch('/api/sap/open_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({tcode:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('새 SAP 세션 여는 중'+(v?' ('+v+')':'')+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('새 세션 열기 실패: '+err.message);
  }
}
async function openPortalLoginManual(){
  try{
    const res=await fetch('/api/portal/login_manual', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('Chrome 열림 (PID '+data.pid+') - 직접 로그인하세요. 이 창을 켜둔 채로 계속 재사용하면 됨');
  }catch(err){
    toast('Portal 로그인 창 열기 실패: '+err.message);
  }
}
// -- Bloomberg Portal: real automation, not a stub - runs
// portal_ship_and_print.py per distinct order among the checked rows, ONE
// AT A TIME (waits for each to fully finish via pollOrderDone() before
// starting the next), and stops on the first failure/timeout rather than
// plowing through the rest - if one order's run failed, something is likely
// systemically wrong (not logged into Portal, Chrome not running, a bad
// scrape) and the same failure would probably just repeat. Confirms the
// exact order list up front, same safety habit launcher.py's own
// messagebox.askyesno already has for this same action. --
async function pollOrderDone(orderId, timeoutMs){
  const start=Date.now();
  while(Date.now()-start < timeoutMs){
    await new Promise(r=>setTimeout(r, 3000));
    const res=await fetch('/api/orders/'+orderId);
    const data=await res.json();
    if(data.status && data.status!=='working'){
      // events[0] is the most recent event (newest-first) - on failure this
      // is run_portal()'s own add_event() call, whose first line is always
      // a clean one-sentence summary of what actually went wrong (see
      // run_portal's watcher()), not just a bare exit code.
      const detail=(data.events && data.events[0] && data.events[0].message || '').split('\n')[0];
      return {status: data.status, detail};
    }
  }
  return {status: 'timeout', detail: ''};
}
async function runPortalBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const orderIds=[...new Set(checkedRows.map(r=>r.orderId))];
  const labelFor_=id=>{ const r=checkedRows.find(r=>r.orderId===id); return r ? r.orderLabel : ('오더#'+id); };
  const summary=orderIds.map(labelFor_).join(', ');
  if(!window.confirm(
    `다음 ${orderIds.length}개 오더에 대해 "Serial 등록 & QR 인쇄"를 순서대로 실행합니다:\n\n${summary}\n\n진행할까요?`
  )) return;

  for(const orderId of orderIds){
    const label=labelFor_(orderId);
    let data;
    try{
      const res=await fetch('/api/orders/'+orderId+'/run_portal', {method:'POST'});
      data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    }catch(err){
      toast(label+' 실행 실패: '+err.message+' - 나머지 오더는 중단합니다');
      return;
    }
    if(!data.started){
      toast(label+': '+(data.message||'건너뜀'));
      continue;
    }
    toast(label+' 진행 중 (PID '+data.pid+') - 완료될 때까지 대기...');
    const result=await pollOrderDone(orderId, 360000);
    if(result.status==='printed'){
      toast(label+' 완료');
    } else if(result.status==='timeout'){
      toast(label+' 시간 초과 (6분) - 상태를 직접 확인하세요. 나머지 오더는 중단합니다');
      return;
    } else {
      toast(label+' 실패: '+(result.detail||('상태: '+result.status))+' - 나머지 오더는 중단합니다');
      return;
    }
  }
  toast('선택한 오더 처리 완료');
}
async function runLatestZpl(){
  try{
    const res=await fetch('/api/portal/print_latest_zpl', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('ZPL 인쇄 시작됨 (PID '+data.pid+') - Downloads 폴더 최신 라벨을 프린터로 전송, 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('ZPL 인쇄 실패: '+err.message);
  }
}

async function pollPidRunning(pid, timeoutMs){
  const start=Date.now();
  while(Date.now()-start < timeoutMs){
    await new Promise(r=>setTimeout(r, 3000));
    const res=await fetch('/api/process/'+pid+'/running');
    const data=await res.json();
    if(!data.running) return true;
  }
  return false;
}

// -- POD (Proof of Delivery / "mark as delivered") - ported from
// launcher.py's "3. 수동 배송 완료 처리" panel. Grouped by the board's
// rendered orderLabel ("ZRX 12345"), NOT raw orderId - real bug hit live
// (2026-08-03): a 배송/회수 pair that LOOKS like one order on screen (one
// shared label, one select-all checkbox) can actually be two separate
// `orders` DB rows (Excel sometimes imports a ZRX's two legs as distinct
// order records - groupRows() only unifies them visually by matching label
// text, same mechanism the merge feature and an earlier stale-duplicate-
// order issue both already hit). Grouping by orderId split one checked
// block into two incomplete POD runs; grouping by label and sending every
// order id in the group lets the server resolve which one actually carries
// the real Portal delivery# (_resolve_pod_order()). For each label group:
// preview (no browser touch) -> a confirm dialog showing exactly what will
// be saved -> if confirmed, the real "click Update POD" run (async, polled
// via PID liveness rather than orders.status - see run_pod_update()'s own
// comment server-side for why). Declining or failing one group's confirm
// just moves on to the next, unlike runPortalBatch() which aborts the whole
// batch - POD is a per-order human sign-off, not a pipeline where one
// failure implies something systemic is wrong. --
// POD 미리보기 조회 - 성공하면 preview 객체, 실패하면 null(+토스트).
// orderIds/itemIds/pickupDone은 podForLabel()/runSpareOneshot() 둘 다 같은
// 방식(체크된 행 그룹)으로 계산해서 넘긴다.
async function podFetchPreview(label, orderIds, itemIds, pickupDone){
  try{
    const res=await fetch('/api/pod/preview', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({orderIds, itemIds, pickupDone})
    });
    const preview=await res.json();
    if(!res.ok || preview.error) throw new Error(preview.error||'미리보기 실패');
    return preview;
  }catch(err){
    toast(label+' POD 미리보기 실패: '+err.message+' - 건너뜁니다');
    return null;
  }
}
function podConfirmText(preview, actionLabel){
  return `오더: ${preview.order}\n`+
    `Delivery#: ${preview.delivery}\n`+
    `배송시간: ${preview.delivery_datetime}\n`+
    `수신인: ${preview.signed_by}\n`+
    `배송 Serial: ${(preview.delivery_serials||[]).join(', ')||'-'}\n`+
    `회수 Serial: ${(preview.pickup_serials||[]).join(', ')||'-'}\n`+
    `Remarks: ${preview.remarks}\n\n`+
    actionLabel;
}
// 실제 POD 저장(Update POD 클릭) 실행 + 완료까지 대기 + 결과 토스트. 확인창은
// 이미 호출한 쪽(podForLabel 또는 runSpareOneshot)에서 끝난 뒤라 여기서는
// 다시 묻지 않는다.
async function podRunUpdate(label, orderIds, itemIds, pickupDone, preview){
  let data;
  try{
    const res=await fetch('/api/pod/update', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        orderIds, itemIds, pickupDone,
        signedBy: preview.signed_by, datetime: preview.delivery_datetime, remarks: preview.remarks,
      })
    });
    data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
  }catch(err){
    toast(label+' POD 저장 실패: '+err.message);
    return false;
  }
  if(!data.started){
    toast(label+': '+(data.message||'건너뜀'));
    return false;
  }
  toast(label+' POD 저장 진행 중 (PID '+data.pid+') - 완료될 때까지 대기...');
  const finished=await pollPidRunning(data.pid, 120000);
  if(!finished){
    toast(label+' POD 저장 시간 초과 (2분) - 상태를 직접 확인하세요');
    return false;
  }
  // small grace period - the watcher thread's own add_event() write can
  // land a beat after the OS process itself is gone (which is what
  // pollPidRunning just observed), so fetching events immediately can
  // occasionally race and show a stale prior event instead.
  await new Promise(r=>setTimeout(r, 500));
  try{
    const res=await fetch('/api/orders/'+data.podOrderId);
    const od=await res.json();
    const msg=(od.events && od.events[0] && od.events[0].message || '').split('\n')[0];
    toast(label+': '+(msg||'POD 처리 완료'));
  }catch(err){
    toast(label+': POD 처리 완료 (상태 확인 실패)');
  }
  return true;
}
// runPodBatch용: 미리보기 -> 확인창 -> 저장, 기존 그대로.
async function podForLabel(label, groupRows_){
  const orderIds=[...new Set(groupRows_.map(r=>r.orderId))];
  // pickupDone comes straight from which of THIS group's rows are
  // actually checked, not a separate question - 배송/회수 each have their
  // own row+checkbox now, so checking only the 배송 row means "delivery
  // only", checking both means "delivery and pickup both done", exactly
  // like checking rows already drives every other bulk action on the
  // board (Done/삭제/날짜이동/병합 etc).
  const pickupDone=groupRows_.some(r=>r.itemType==='회수' || r.itemType==='회수 Delayed');
  // Only the CHECKED rows' item ids go to the server - the remarks text
  // (Delivered/collected item counts) is built from exactly these, not
  // Excel and not every item the order happens to have, so checking just
  // the 배송 row correctly leaves the 회수 item out of "Delivered ..." too.
  const itemIds=groupRows_.map(r=>r.itemId);

  const preview=await podFetchPreview(label, orderIds, itemIds, pickupDone);
  if(!preview) return false;

  if(!window.confirm(podConfirmText(preview, '이 내용으로 POD(배송완료) 저장할까요?'))){
    toast(label+': POD 저장 건너뜀');
    return false;
  }
  return podRunUpdate(label, orderIds, itemIds, pickupDone, preview);
}
async function runPodBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const labels=[...new Set(checkedRows.map(r=>r.orderLabel))];

  for(const label of labels){
    const groupRows_=checkedRows.filter(r=>r.orderLabel===label);
    await podForLabel(label, groupRows_);
  }
  toast('POD 처리 완료');
}
// -- 스페어 배송처리 (todo a, 2026-08-18): Serial 등록 + POD(배송완료)를 한
// 번의 클릭으로 처리. 확인창은 버튼을 누른 그 순간, Serial 등록을 시작하기
// *전에* POD 미리보기 내용으로 딱 한 번만 뜬다 - 실사용해보니(2026-08-18)
// Serial 등록이 끝난 뒤에야 POD 확인창이 뜨는 구조라 결국 두 번 확인해야
// 완료됐는데, 그 확인창을 앞으로 당겨서 한 번 확인하면 Serial 등록부터
// POD 저장까지 중간에 아무 팝업 없이 끝까지 진행되도록 바꿈 (사용자 확정).
// 그래서 순서가 podForLabel()과 반대: 여기서는 미리보기→확인을 먼저 하고,
// 그 확인에서 얻은 preview 값(signed_by/datetime/remarks)을 그대로 들고
// Serial 등록 → podRunUpdate()로 이어간다 - 확인 이후에는 preview를 다시
// 조회하지 않는다(값이 그 사이 바뀔 이유가 없고, 다시 조회하면 그 결과를
// 또 확인해야 하니 원샷의 의미가 없어짐). QR 인쇄는 항상 생략(--skip-qr) -
// 스페어는 라벨 재인쇄가 보통 불필요해서 버튼 자체에서 뺐다 (사용자 확정).
async function runSpareOneshot(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const labels=[...new Set(checkedRows.map(r=>r.orderLabel))];

  for(const label of labels){
    const groupRows_=checkedRows.filter(r=>r.orderLabel===label);
    const orderIds=[...new Set(groupRows_.map(r=>r.orderId))];
    const pickupDone=groupRows_.some(r=>r.itemType==='회수' || r.itemType==='회수 Delayed');
    const itemIds=groupRows_.map(r=>r.itemId);

    const preview=await podFetchPreview(label, orderIds, itemIds, pickupDone);
    if(!preview) continue;
    if(!window.confirm(podConfirmText(preview, '이 내용으로 Serial 등록 후 POD(배송완료)까지 한 번에 진행할까요?'))){
      toast(label+': 스페어 배송처리 건너뜀');
      continue;
    }

    let portalOk=true;
    for(const orderId of orderIds){
      let data;
      try{
        const res=await fetch('/api/orders/'+orderId+'/run_portal', {
          method:'POST', headers:{'Content-Type':'application/json'},
          body:JSON.stringify({skipQr: true})
        });
        data=await res.json();
        if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
      }catch(err){
        toast(label+' Serial 등록 실행 실패: '+err.message+' - 이 오더는 건너뜁니다');
        portalOk=false; break;
      }
      if(!data.started){
        // 이미 'printed'(완료)인 경우는 실패가 아니라 그대로 POD 단계로
        // 진행 - Serial 등록만 먼저 성공해둔 뒤 원샷을 재실행하는 경우 등.
        // '처리 중'(다른 프로세스가 살아있음)인 경우만 진짜로 건너뜀.
        if((data.message||'').includes('이미 완료')){
          toast(label+': Serial 등록 이미 완료 - POD로 진행');
        } else {
          toast(label+': '+(data.message||'Serial 등록 건너뜀')+' - 이 오더는 건너뜁니다');
          portalOk=false; break;
        }
      } else {
        toast(label+' Serial 등록 진행 중 (PID '+data.pid+') - 완료 대기...');
        const result=await pollOrderDone(orderId, 360000);
        if(result.status==='printed'){
          toast(label+' Serial 등록 완료');
        } else if(result.status==='timeout'){
          toast(label+' Serial 등록 시간 초과 (6분) - 이 오더는 건너뜁니다');
          portalOk=false; break;
        } else {
          toast(label+' Serial 등록 실패: '+(result.detail||('상태: '+result.status))+' - 이 오더는 건너뜁니다');
          portalOk=false; break;
        }
      }
    }
    if(!portalOk) continue;

    await podRunUpdate(label, orderIds, itemIds, pickupDone, preview);
  }
  toast('스페어 배송처리 완료');
}

// -- S/N 오더 반영 (2026-08-19): 회수 오더 수집 시 SAP상 S/N이 없어서('X')
// 워크벤치에 X로 찍힌 항목을 위한 단계. 사용자가 S/N 칸에 실제로 회수한
// S/N을 먼저 입력해두고 체크한 뒤 이 버튼을 누르면, 그 항목들을 순서대로
// VA02 오더에 입력·저장 시도한다 - ZREC 준비(runZrecPrepare)와 같은 "체크
// → 버튼 한 번 → 항목마다 순서대로 처리" 패턴. 저장되면(firm/cust 일치)
// 그걸로 끝 - 이 항목은 이제 일반 회수 항목과 똑같은 상태이니, 이어서
// 5번 ZREC 처리 버튼을 사람이 직접 눌러 진행한다(자동 연결 안 함, 사용자
// 확정). 저장 실패(firm/cust 불일치 추정)는 paper relo가 필요하다는 뜻 -
// S/N 칸 옆에 ⚠ 배지 + 행 주황 배경으로 표시된다.
async function runSnReloBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const targets=checkedRows.filter(r=>r.serial && r.serial.trim().toUpperCase()!=='X');
  if(!targets.length){
    window.alert('체크된 항목 중 S/N 칸이 채워진(X가 아닌) 항목이 없습니다.\n실제 회수한 S/N을 먼저 S/N 칸에 입력한 뒤 다시 시도하세요.');
    return;
  }
  if(!window.confirm(
    `체크된 항목 중 S/N이 입력된 ${targets.length}개를 대상으로, 순서대로 VA02 오더를 열어 `+
    `그 S/N을 Technical Objects에 입력·저장 시도합니다.\n`+
    `저장되면 firm/cust 일치(ZREC 진행 가능), 저장 실패면 paper relo가 필요하다는 뜻입니다.\n\n진행할까요?`
  )) return;

  let okCount=0, errCount=0;
  for(const row of targets){
    let data;
    try{
      const res=await fetch('/api/sn_relo/fill_one', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({itemId: row.itemId})
      });
      data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    }catch(err){
      errCount++;
      row.snReloError=true; row.snReloErrorMsg=err.message;
      if(!window.confirm(row.orderLabel+' S/N 오더 반영 실패: '+err.message+'\n\n다음 항목으로 계속할까요?')) break;
      continue;
    }
    if(data.status==='saved'){
      okCount++;
      row.snReloError=false; row.snReloErrorMsg='';
    }else{
      errCount++;
      row.snReloError=true; row.snReloErrorMsg=data.message||'알 수 없는 실패';
    }
  }
  renderAll();
  toast(`S/N 오더 반영 완료 - 성공 ${okCount}건, 실패/확인필요 ${errCount}건`+(errCount?' (⚠ 배지 클릭으로 사유 확인)':''));
}

// -- ZREC 준비 (todo c, 2026-08-18/19, 개편 2026-08-21): 사용자의 실제
// 하루 업무 방식("오늘 회수된 모든 시리얼을 한번에 넣고 확인한 뒤, ZREC는
// 하나씩") + 2026-08-19 확정 사항("확인 누르면 Receive Equipment는
// 자동으로 클릭, 다 끝나면 ZIH08로 6507/0052 확인")을 반영한 흐름.
//   1단계: 체크된 모든 항목의 시리얼을 모아 ZIH08에 딱 한 번 조회
//          (/api/zrec/lookup_batch) - ZREC 화면은 전혀 안 건드리는
//          read-only 조회라 확인창 없이 바로 진행.
//   2단계: 조회 결과를 오더번호/M-N/S-N/수량/아이템명 한 표로 모아 확인창
//          "한 번만" 띄운다 - 원래는 항목마다 ZREC 채우기 직후 개별
//          확인창이 떴는데, 여러 건을 한 번에 처리할 때 매번 팝업이 뜨는
//          게 번거롭다는 지적(2026-08-21, 사용자) 반영. 이 확인 하나로
//          전체 배치를 승인하는 것 - "확인"을 누르면 항목마다 순서대로
//          ZREC 채우기 + Receive Equipment 자동 클릭까지 중간 팝업 없이
//          진행한다(개별 실패는 그 항목만 건너뛰고 계속). 사람이 SAP
//          ZREC 화면에 실제로 채워지는 6507/오더번호/시리얼 값을 눈으로
//          보고 승인하는 지점은 없다 - 승인은 이 확인창에서 ZIH08 조회
//          데이터를 보고 한 번에 이뤄진다(2026-08-27 확인).
//   3단계: 실제로 접수(commit)된 항목들의 시리얼을 ZIH08로 재조회해서
//          Plant/Location이 6507/0052로 들어왔는지 확인(/api/zrec/verify_batch)
//          - 확인된 항목만 board_status를 'done'(파랑)으로 자동 전환한다.
//          준비 상태(Ready to process, 빨강)는 ZIH08로 접수가 실제 확인될
//          때까지 그대로 유지 (사용자 확정, 2026-08-19: "zrec할 준비가
//          된것이 빨간색이고, zrec가 완료되면 파란색으로 처리"). 이게
//          진짜 성공/실패 판정이고, 아래 commit_one 결과 문구는 참고용일
//          뿐이니 "실패로 보인다" 같은 불확실한 문구를 화면에 띄우지
//          않는다(2026-08-19: SAP 팝업 문구 판정이 실제 성공을 실패로
//          오판하는 게 실사용에서 확인됨 - 사용자 지적, 2026-08-21: "성공
//          했는데 팝업에 실패한 것 같다는 문구가 뜬다").
async function runZrecPrepare(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));

  // 1단계: 일괄 ZIH08 조회 (read-only - 확인 없이 바로 진행)
  let plans;
  try{
    const res=await fetch('/api/zrec/lookup_batch', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({itemIds: checkedRows.map(r=>r.itemId)})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
    plans=data.plans||[];
  }catch(err){
    toast('ZIH08 일괄 조회 실패: '+err.message);
    return;
  }
  if(!plans.length){toast('ZREC 대상 항목이 없습니다');return}

  // 2단계: 전체를 한 표(줄바꿈 나열)로 모아 확인창을 딱 한 번만 띄움.
  // 항목마다 빈 줄로 섹션을 나눠 한눈에 읽히게 하고, ZIH08 경고가 있는
  // 항목은 그 아래 ⚠로 붙여서 개별 확인창 없이도 문제를 미리 볼 수 있게 함.
  const summaryLines=plans.map(p=>{
    const sn=p.nonSerial ? '(S/N 없음)' : (p.serial||'-');
    const warn=(p.warnings||[]).map(w=>'   ⚠ '+w).join('\n');
    return `${p.orderLabel}  |  M/N ${p.material||'-'}  |  S/N ${sn}  |  수량 ${p.qty}  |  ${p.description||'-'}`
      +(warn ? '\n'+warn : '');
  });
  const confirmMsg=
    `오늘 ZREC 접수할 ${plans.length}건 (오더번호 | M/N | S/N | 수량 | 아이템명):\n\n`+
    summaryLines.join('\n\n')+
    `\n\n이 내용으로 확인을 누르면 항목마다 순서대로 자동으로 Receive Equipment까지 `+
    `클릭합니다(중간 확인창 없음) - 실제로 SAP에 접수되니 내용을 꼭 확인하고 눌러주세요.\n`+
    `모두 끝나면 ZIH08로 Plant/Location(6507/0052) 접수 여부를 다시 확인합니다.`;
  if(!window.confirm(confirmMsg)) return;

  // 3단계: 항목마다 순서대로 ZREC 채우기 → 자동 Receive Equipment.
  // 위 2단계에서 이미 전체 승인을 받았으므로 항목별 확인창은 없음 -
  // 개별 실패는 토스트로만 알리고 다음 항목으로 계속 진행.
  //
  // 2026-08-27: 예전엔 여기서 fill_one(dry-run, 화면만 채움) →
  // commit_one(세션 리셋 후 똑같은 값 재입력 → 클릭) 두 번을 호출했다.
  // 그런데 fill_one이 채운 화면을 사람이 보고 판단할 시점이 이미 없다
  // (승인은 위 2단계 확인창에서 ZIH08 데이터만으로 끝남) - 재입력은
  // 안전 효과 없이 SAP 왕복만 두 배로 냈다(사용자 지적). commit_one 한
  // 번으로 채우기+검증+클릭까지 끝낸다.
  const committedItemIds=[];
  for(const plan of plans){
    try{
      const res=await fetch('/api/zrec/commit_one', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({itemId: plan.itemId, orderNo: plan.filledOrder})
      });
      const cdata=await res.json();
      if(!res.ok || cdata.error) throw new Error(cdata.error||'Receive Equipment 클릭 실패');
      // Receive Equipment 버튼은 실제로 눌렸으므로(SAP 트랜잭션 시도 완료)
      // cdata.success 값(SAP 팝업 문구 판정)과 무관하게 일단 3단계 ZIH08
      // verify_batch로 검증 대상에 넣는다 - 문구 판정은 실사용에서 실제
      // 성공을 실패로 오판하는 게 확인됐으므로(2026-08-19) 여기서는 그
      // 판정 결과를 화면에 노출하지 않고 중립적으로만 알린다. ZIH08
      // 재조회(6507/0052 확인)만이 진짜 성공/실패를 가리는 최종 판정.
      committedItemIds.push(plan.itemId);
      toast(plan.orderLabel+' Receive Equipment 클릭 완료 - ZIH08로 곧 확인합니다');
    }catch(err){
      toast(plan.orderLabel+' Receive Equipment 클릭 실패: '+err.message+' - 이 항목은 건너뜁니다');
      continue;
    }
  }

  if(!committedItemIds.length){
    toast('ZREC 처리 완료 (실제로 접수된 항목 없음)');
    return;
  }

  // 3단계: ZIH08 재조회로 Plant/Location 6507/0052 확인
  toast('ZIH08로 접수 완료 여부 확인 중...');
  try{
    const res=await fetch('/api/zrec/verify_batch', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({itemIds: committedItemIds})
    });
    const vdata=await res.json();
    if(!res.ok || vdata.error) throw new Error(vdata.error||'확인 실패');
    const results=vdata.results||[];
    const ok=results.filter(r=>r.receivedOk);
    const bad=results.filter(r=>!r.receivedOk);

    // ZREC 완료 확인된 항목은 파랑(done)으로 자동 전환 - 준비된 상태(빨강,
    // Ready to process)는 그대로 두고, 실제 접수가 ZIH08로 확인된 것만
    // 파랑으로 바뀐다 (사용자 확정, 2026-08-19: "zrec할 준비가 된것이
    // 빨간색이고, zrec가 완료되면 파란색으로 처리"). 확인 안 된 항목은
    // 빨강 그대로 남아서 다시 확인해야 함을 표시. "해당 오더를 모두"
    // 파랑으로 바꾸라는 요청이라 확인된 시리얼 자신뿐 아니라 같은
    // orderId를 공유하는 다른 행(예: 같은 오더의 다른 품목)까지 전부
    // 포함해서 done 처리한다.
    if(ok.length){
      const okItemIds=new Set(ok.map(r=>r.itemId));
      const okOrderIds=new Set(
        ok.map(r=>{ const row=ROWS.find(x=>x.itemId===r.itemId); return row ? row.orderId : null; }).filter(Boolean)
      );
      const doneRows=ROWS.filter(r=>okItemIds.has(r.itemId) || okOrderIds.has(r.orderId));
      try{
        await persistItemStatus(doneRows.map(r=>r.itemId), 'done');
        doneRows.forEach(r=>{ r.uiStatus='done'; });
        renderAll();
      }catch(err){
        toast('완료 항목 파랑 처리 저장 실패: '+err.message);
      }
    }

    let summary=`ZREC 완료 확인: ${ok.length}/${results.length}건 Plant 6507/Location 0052 확인됨(해당 오더 전체 파랑 처리).`;
    if(bad.length){
      summary+=`\n\n⚠ 아직 확인 안 된 항목(빨강 유지):\n`+bad.map(r=>`${r.orderLabel} (${r.serial}) - Plant='${r.plant}' Location='${r.location}'`).join('\n');
      window.alert(summary);
    } else {
      toast(summary);
    }
  }catch(err){
    toast('ZIH08 완료 확인 조회 실패: '+err.message);
  }
}

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
// Type/check를 제외한 나머지 7개 board 열, 왼쪽부터의 순서 - addManualRow()
// 모달과 add_manual_note()의 cells 매핑과 반드시 같은 순서여야 함. 여기,
// ROWS 선언 바로 아래에 있어야 한다: buildGroupRows()의 메모 행 렌더링이
// manualRowGroupStarts()를 통해 이 값을 참조하고, 그 렌더링은 최초
// renderAll() 호출(파일 훨씬 아래쪽) 시점에 이미 실행되므로, 이 const가 그
// 호출 지점보다 뒤에 있으면 TDZ(temporal dead zone)로 스크립트 전체가 죽는다
// - 2026-08-27 실사고: 메모 행 추가 버튼을 누르자마자 보드 전체가 하얗게
// 사라짐, 콘솔에 "Cannot access 'MANUAL_ROW_COLUMNS' before initialization".
const MANUAL_ROW_COLUMNS=['제목','M/N','S/N','Customer','Phone','Address','Memo'];

// Type dropdown options. BASE_TYPES is fixed; CUSTOM_TYPES accumulates
// whatever the user adds via the dropdown's "+ 추가" entry - shared across
// every row's dropdown, not just the one it was added from. 2026-08-26 fix:
// this used to be session-only (reset to [] on every load) even though each
// ROW's own chosen type was already being saved - the *list of known custom
// types offered in the dropdown* was the part that wasn't, so a custom type
// had to be retyped via "+ 추가" every time it was needed on a different row
// after a reload. Now seeded from the server (custom_types table via
// __CUSTOM_TYPES__, see render_dashboard_page()) and every newly-added one
// is persisted immediately - see the '__add__' branch below.
const BASE_TYPES = ['배송','회수','Bloomberg','Delayed','회수 Delayed'];
let CUSTOM_TYPES = __CUSTOM_TYPES__;
function allTypeOptions(){ return [...BASE_TYPES, ...CUSTOM_TYPES]; }

// 2026-09-13 사용자 요청: 날짜 배너를 "자동화가 안 도는 날"(주말/한국
// 공휴일) + "직접 지정한 휴무일"이면 빨간색으로. 한국 공휴일 이름은 서버
// (holiday_check.kr_holiday_labels, 음력 계산 포함)가 내려주고, 주말은
// iso 문자열만으로 클라이언트가 바로 계산할 수 있어 서버 왕복 없이 처리.
const NONWORKING = __NONWORKING_DAYS__;
const KR_HOLIDAYS = NONWORKING.holidays || {};
const MANUAL_HOLIDAYS = new Set(NONWORKING.manual || []);
function nonworkingInfo(iso){
  if(MANUAL_HOLIDAYS.has(iso)) return {non:true, label:'수동 지정 휴무일'};
  if(KR_HOLIDAYS[iso]) return {non:true, label:KR_HOLIDAYS[iso]};
  const dow = new Date(iso+'T00:00:00').getDay();
  if(dow===0 || dow===6) return {non:true, label:'주말'};
  return {non:false, label:''};
}
function typeClassFor(v){
  if(v==='회수' || v==='회수 Delayed') return 'type-pickup';
  if(v==='배송') return 'type-delivery';
  return 'type-other';
}

(function loadInitialData(){
  const data = __BOARD_DATA__;
  data.forEach(d=>{
    DATE_LABELS[d.iso] = d.label;
    d.rows.forEach(r=>ROWS.push({...r, date: d.iso, uiStatus: r.boardStatus || ''}));
  });
})();

// 새 오더 폴링 - SAP 수집이 workbench.db에 오더를 넣어도 이미 열려있는 탭은
// 새로고침 전까진 몰랐다(사용자 지적, 2026-08-14). 20초마다 서버 board_data()를
// 다시 가져와서 로컬 ROWS에 없는 itemId만 "새 오더"로 판단해 추가한다 - 기존
// 행은 절대 건드리지 않음(진행 중인 드래그/편집/선택 상태를 깨뜨리지 않기
// 위해 add-only로 제한). 예외는 사용자가 텍스트로 직접 편집하는 필드가 아니라
// 버튼/드롭다운/우클릭 메뉴로 원자적으로 바뀌는 상태값들 - 다른 탭/세션에서
// 트리거된 변화도 새로고침 없이 20초 안에 반영되도록 이 필드들만 기존 행에도
// 동기화한다: portalError/podError/snReloError(todo b, 2026-08-18 / S/N
// 반영, 2026-08-19), boardStatus(Done/Ready/Cancel/초기화 행 색칠)와
// highlights(우클릭 셀 강조), itemType(타입 드롭다운) - 2026-09-01, 사용자
// 지적("업데이트가 새로고침 없이 안 보임") 계기로 확대. memo/customer/
// address 등 자유 텍스트 필드는 여전히 건드리지 않음(다른 탭에서 타이핑
// 중인 내용을 덮어쓸 위험).
let _pollingBoard = false;
async function pollForNewOrders(){
  if(_pollingBoard) return;  // 이전 폴링이 아직 안 끝났으면 겹쳐서 돌리지 않음
  _pollingBoard = true;
  try{
    const res = await fetch('/api/board_data');
    if(!res.ok) return;
    const data = await res.json();
    const byItemId = new Map(ROWS.map(r=>[r.itemId, r]));
    let added = 0;
    let changed = false;
    data.forEach(d=>{
      d.rows.forEach(r=>{
        const existing = byItemId.get(r.itemId);
        if(existing){
          if(existing.portalError!==r.portalError || existing.portalErrorMsg!==r.portalErrorMsg
             || existing.podError!==r.podError || existing.podErrorMsg!==r.podErrorMsg
             || existing.snReloError!==r.snReloError || existing.snReloErrorMsg!==r.snReloErrorMsg){
            existing.portalError=r.portalError; existing.portalErrorMsg=r.portalErrorMsg;
            existing.podError=r.podError; existing.podErrorMsg=r.podErrorMsg;
            existing.snReloError=r.snReloError; existing.snReloErrorMsg=r.snReloErrorMsg;
            changed = true;
          }
          if(existing.boardStatus!==r.boardStatus){
            existing.boardStatus=r.boardStatus;
            existing.uiStatus=r.boardStatus || '';
            changed = true;
          }
          if(existing.itemType!==r.itemType){
            existing.itemType=r.itemType;
            changed = true;
          }
          if(JSON.stringify(existing.highlights||{})!==JSON.stringify(r.highlights||{})){
            existing.highlights=r.highlights;
            changed = true;
          }
          return;
        }
        if(!DATE_LABELS[d.iso]) DATE_LABELS[d.iso] = d.label;
        const newRow={...r, date: d.iso, uiStatus: r.boardStatus || ''};
        ROWS.push(newRow);
        byItemId.set(r.itemId, newRow);
        added++;
      });
    });
    if(added > 0){
      renderAll();
      toast('새 오더 '+added+'건 반영됨');
    } else if(changed){
      renderAll();
    }
  }catch(err){
    // 조용히 무시 - 네트워크 순간 끊김 등으로 매 폴링마다 토스트 띄우면 방해됨
  }finally{
    _pollingBoard = false;
  }
}
// 2026-09-01: 20초 폴링은 "바뀌면 바로 반영돼야지"라기엔 너무 느리다는 지적.
// 그렇다고 board_data() 전체(오더+아이템 조인 여러 개)를 1~2초마다 통째로
// 다시 불러오는 건 낭비 - 대신 /api/board_version(그냥 SQLite data_version
// pragma 하나, 사실상 공짜)을 1.5초마다 찔러보고 값이 바뀐 걸 감지했을 때만
// 진짜 pollForNewOrders()를 부른다. data_version은 이 서버의 HTTP 핸들러를
// 거치지 않고 SAP 자동루프/ZREC/Portal 등 다른 프로세스가 workbench.db에
// 직접 쓴 변경까지도 잡아내므로(각 write 경로마다 "바뀜"을 따로 통지하게
// 만드는 것보다 훨씬 안전함), 체감상 거의 즉시(최대 1.5초) 반영된다. 혹시
// 이 감지를 놓치는 경우에 대비한 belt-and-suspenders로 30초 전체 폴백도
// 유지.
let _lastBoardVersion = null;
let _checkingVersion = false;
async function checkBoardVersion(){
  if(_checkingVersion) return;
  _checkingVersion = true;
  try{
    const res = await fetch('/api/board_version');
    if(!res.ok) return;
    const {version} = await res.json();
    if(_lastBoardVersion===null){ _lastBoardVersion = version; return; }
    if(version !== _lastBoardVersion){
      _lastBoardVersion = version;
      pollForNewOrders();
    }
  }catch(err){
    // 조용히 무시
  }finally{
    _checkingVersion = false;
  }
}
setInterval(checkBoardVersion, 1500);
setInterval(pollForNewOrders, 30000);
const DELETED_ITEM_RETENTION_DAYS = __DELETED_ITEM_RETENTION_DAYS__;

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
  // Only dates that currently have rows get a section - no forced "today"
  // anchor, so finishing today's date (e.g. via "마감") makes its banner
  // disappear entirely instead of sticking around empty. DATE_LABELS is just
  // a label-text cache and intentionally excluded here.
  const set=new Set();
  ROWS.forEach(r=>set.add(r.date));
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

// -- Cross-order merge ("Excel merge cells" for customer/phone/address/memo
// across DIFFERENT order numbers that are really the same customer) - manual,
// via the "고객정보 병합" button (mergeChecked()/unmergeChecked() further
// down). Each of the 4 fields merges INDEPENDENTLY (its own group id, its
// own span) rather than one all-or-nothing block, because a real case can
// share e.g. just the address while the customer name on file differs.
// Persisted as orders.merge_group_<field>; a field only actually renders as
// one rowspan block when the merged orders' groups also end up adjacent on
// the board (same constraint real Excel has - you can only merge a
// contiguous range), which is what the drag-to-reorder-within-a-date
// feature is for. --
const MERGE_FIELDS=['customer','phone','address','memo'];
function fieldMergeKeyOf(group, field){
  // Checks EVERY row in the group, not just rows[0] - a "group" here can
  // already be a concatenation of multiple different underlying orders that
  // happen to render under the same order-number label (groupRows() collapses
  // same-labeled orders together regardless of which one is actually merged,
  // same fallback logic resolvedField() already uses for customer/phone/etc).
  // If only rows[0] were checked, a group whose FIRST underlying order isn't
  // merged would report "not merged" even when a later order folded into the
  // same block genuinely is - exactly what happened with a stale duplicate
  // "ZOR 7825897" order sitting before the real merged one.
  const row=group.rows.find(r=>r.mergeGroups && r.mergeGroups[field]);
  return row ? String(row.mergeGroups[field]).trim() : '';
}
function computeFieldSpans(groups, field){
  const spans=new Array(groups.length);
  let i=0;
  while(i<groups.length){
    const key=fieldMergeKeyOf(groups[i], field);
    let j=i+1;
    let total=groups[i].rows.length;
    const mergedRows=[...groups[i].rows];
    if(key){
      while(j<groups.length && fieldMergeKeyOf(groups[j], field)===key){
        total+=groups[j].rows.length;
        mergedRows.push(...groups[j].rows);
        j++;
      }
    }
    spans[i]={isFirst:true, span:total, mergedRows};
    for(let k=i+1;k<j;k++) spans[k]={isFirst:false, span:0, mergedRows:[]};
    i=j;
  }
  return spans;
}
function computeAllFieldSpans(groups){
  const out={};
  MERGE_FIELDS.forEach(field=>{ out[field]=computeFieldSpans(groups, field); });
  return out;
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

function makeEditableTd(value, className, onCommit, onPersist){
  const td=document.createElement('td');
  if(className) td.className=className;
  td.contentEditable='true';
  td.spellcheck=false;
  td.textContent=value||'';
  td.addEventListener('blur', ()=>{
    const v=td.textContent.trim();
    onCommit(v);
    renderAll();
    // onPersist is optional (still-screen-only fields like orderLabel/
    // deliveryNo don't pass one yet) - fires after the optimistic on-screen
    // update so typing stays instant; a failure just warns instead of
    // silently leaving the DB out of sync with what's shown.
    if(onPersist) onPersist(v).catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
  });
  td.addEventListener('keydown', e=>{
    if(e.key!=='Enter') return;
    e.preventDefault();
    if(e.altKey){
      // Alt+Enter = line break within the cell, same shortcut Excel uses -
      // insertText keeps it a plain '\n' character (matches how
      // deliveryNo/address already store multi-line text), not a nested
      // <div>/<br>, so td.textContent stays a simple string.
      document.execCommand('insertText', false, '\n');
      return;
    }
    td.blur();
  });
  return td;
}

// -- Manual cell highlight (whole-cell only, not per-character - see
// project memory round 15): right-click any highlightable <td> for a small
// menu (배경 노랑/핑크, 글자 빨강, Bold). Persists to workbench.db
// (row.highlights, keyed by {itemId, field}) and also carries into the
// Excel export (buildExportRows/_write_export_sheet) - deliberately does
// NOT show up in the 배송장 인쇄 output (print stays plain black-on-white
// per its own spec). --
const HL_BG_COLOR={yellow:'#fde047', pink:'#f9a8d4'};
const HL_TEXT_COLOR={red:'#dc2626'};
function applyHighlight(td, itemId, field, row){
  td.dataset.hlItem=itemId;
  td.dataset.hlField=field;
  const style=(row.highlights && row.highlights[field]) || null;
  td.style.background = style && HL_BG_COLOR[style.bg] ? HL_BG_COLOR[style.bg] : '';
  td.style.color = style && HL_TEXT_COLOR[style.color] ? HL_TEXT_COLOR[style.color] : '';
  td.style.fontWeight = style && style.bold ? '700' : '';
}
async function setHighlight(itemId, field, style){
  const res=await fetch('/api/items/highlight', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body:JSON.stringify({itemId, field, style})
  });
  const data=await res.json();
  if(!res.ok || data.error){ toast('강조 표시 저장 실패: '+(data.error||'')); return; }
  const row=ROWS.find(r=>r.itemId===itemId);
  if(row){
    row.highlights=row.highlights||{};
    if(Object.keys(style).length) row.highlights[field]=style; else delete row.highlights[field];
  }
  renderAll();
}
let hlMenuEl=null;
function closeHlMenu(){
  if(hlMenuEl){ hlMenuEl.remove(); hlMenuEl=null; }
  document.removeEventListener('click', closeHlMenu);
}
function openHlMenu(x, y, itemId, field, current){
  closeHlMenu();
  const menu=document.createElement('div');
  menu.className='hl-menu';
  menu.style.left=x+'px'; menu.style.top=y+'px';
  const mk=(label, cls, onClick)=>{
    const b=document.createElement('button');
    b.type='button'; b.className='hl-menu-btn'+(cls?' '+cls:''); b.textContent=label;
    b.addEventListener('click', e=>{ e.stopPropagation(); onClick(); closeHlMenu(); });
    menu.appendChild(b);
  };
  const apply=patch=>{
    const next={...current, ...patch};
    Object.keys(next).forEach(k=>{ if(!next[k]) delete next[k]; });
    setHighlight(itemId, field, next);
  };
  mk('배경 없음', '', ()=>apply({bg:undefined}));
  mk('배경 노랑', 'hl-swatch-yellow', ()=>apply({bg:'yellow'}));
  mk('배경 핑크', 'hl-swatch-pink', ()=>apply({bg:'pink'}));
  mk('글자 기본', '', ()=>apply({color:undefined}));
  mk('글자 빨강', 'hl-swatch-red', ()=>apply({color:'red'}));
  mk(current.bold ? 'Bold 끄기' : 'Bold 켜기', '', ()=>apply({bold:!current.bold}));
  mk('전체 지우기', '', ()=>setHighlight(itemId, field, {}));
  document.body.appendChild(menu);
  hlMenuEl=menu;
  setTimeout(()=>document.addEventListener('click', closeHlMenu), 0);
}
document.getElementById('boardTable').addEventListener('contextmenu', e=>{
  const td=e.target.closest('[data-hl-field]');
  if(!td) return;
  e.preventDefault();
  const itemId=parseInt(td.dataset.hlItem, 10);
  const field=td.dataset.hlField;
  const row=ROWS.find(r=>r.itemId===itemId);
  const current=(row && row.highlights && row.highlights[field]) || {};
  openHlMenu(e.clientX, e.clientY, itemId, field, current);
});

function buildGroupRows(group, allSpans, gi){
  const trs=[]; const n=group.rows.length;
  // "select this whole order" checkbox (added to ordTd below) - so merging
  // (or bulk Done/삭제/날짜이동) across orders with many items doesn't mean
  // hunting down and clicking every single item-row checkbox by hand; one
  // click selects/deselects every item this order-group owns. groupCheckboxes
  // collects each row's real .row-check as it's built; syncSelectAll keeps
  // the order-level box in check/indeterminate/unchecked sync if the user
  // still clicks individual item checkboxes directly.
  const groupCheckboxes=[];
  let selectAllCb=null;
  const syncSelectAll=()=>{
    if(!selectAllCb)return;
    const checkedCount=groupCheckboxes.filter(cb=>cb.checked).length;
    selectAllCb.checked=checkedCount>0 && checkedCount===groupCheckboxes.length;
    selectAllCb.indeterminate=checkedCount>0 && checkedCount<groupCheckboxes.length;
  };
  group.rows.forEach((row, i)=>{
    const tr=document.createElement('tr');
    tr.className='item-row'+(row.uiStatus==='done'?' row-done':row.uiStatus==='pending'?' row-pending':row.uiStatus==='cancelled'?' row-cancelled':'')
      +((row.portalError||row.podError||row.snReloError)?' row-error':'');
    tr.dataset.rowKey=row.rowKey;
    // group.key (same value mergeKey(row) would give) stamped on every row of
    // this group - onRowDragOver/onRowDrop use it to snap the insert point to
    // the GROUP's own top/bottom edge instead of a boundary that can fall
    // inside a DIFFERENT, unrelated order's multi-row block (2026-08-27 user
    // report: the drag preview line could land mid-way through a completely
    // unrelated order - different order#, different company - making it look
    // like a drop there would splice into/mix with it).
    tr.dataset.groupKey=group.key;
    // same-date reordering: dragover/drop directly on a row inserts the
    // dragged row(s) immediately before/after its GROUP (see
    // onRowDragOver/onRowDrop). stopPropagation on drop only - dragover is
    // left to bubble up to the tbody handler too, so the whole date section
    // still gets its 'drag-over' background highlight while hovering any row
    // inside it.
    tr.addEventListener('dragover', onRowDragOver);
    tr.addEventListener('drop', onRowDrop);

    // drag handle - a dedicated grip so dragging doesn't fight with
    // clicking/selecting text inside the row's editable cells or dropdown
    const handleTd=document.createElement('td');
    handleTd.className='handle-cell';
    const grip=document.createElement('span');
    grip.className='drag-handle';
    grip.textContent='⠿';
    grip.draggable=true;
    // Dragging one row now drags its WHOLE order-group (every row sharing
    // its orderLabel - an order must never be split by a drag) plus, if the
    // grabbed row is itself checked, every OTHER checked row's whole group
    // too - so checking several rows (or whole orders via the "select this
    // order" box) and dragging any one of them moves the entire selection
    // together as one block, in their existing relative order (2026-08-27
    // user request: "선택하고 드래그 드롭을 하는게 뭉텅이로 쉽게 이동
    // 가능해야하는데"). Dragging an UNchecked row still only moves that one
    // row's own group, so a plain drag never sweeps up unrelated checked
    // rows left over from something else.
    grip.addEventListener('dragstart', e=>{
      const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
      const seedKeys = (checkedKeys.length && checkedKeys.includes(row.rowKey)) ? checkedKeys : [row.rowKey];
      const seedSet=new Set(seedKeys);
      const groupKeysInvolved=new Set(ROWS.filter(r=>seedSet.has(r.rowKey)).map(r=>mergeKey(r)));
      const fullKeys=ROWS.filter(r=>groupKeysInvolved.has(mergeKey(r))).map(r=>r.rowKey);
      e.dataTransfer.setData('text/plain', fullKeys.join(','));
      const fullKeySet=new Set(fullKeys);
      document.querySelectorAll('.item-row').forEach(el=>{
        if(fullKeySet.has(el.dataset.rowKey)) el.classList.add('dragging');
      });
    });
    grip.addEventListener('dragend', ()=>{
      document.querySelectorAll('.item-row.dragging').forEach(el=>el.classList.remove('dragging'));
    });
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
        if(newType && !allTypeOptions().includes(newType)){
          CUSTOM_TYPES.push(newType);
          persistCustomType(newType)
            .catch(err=>toast('새 Type 목록 저장 실패 (이 세션에서만 보임): '+err.message));
        }
        if(newType){
          row.itemType=newType;
          persistItemField(row.itemId,'item_type',newType)
            .catch(err=>toast('Type 저장 실패 (DB 반영 안 됨): '+err.message));
        }
        renderAll();
        return;
      }
      // 2026-08-26 fix: 이전엔 여기서 row.itemType만 바꾸고 서버에 저장하는
      // 코드가 아예 없었음 - 화면에선 바뀐 것처럼 보이다가 다음 새로고침(또는
      // 워크벤치 재시작) 때 서버가 다시 내려주는 원래 값으로 조용히
      // 되돌아갔음. 배송/회수만 특별취급하던 resolve_item_type()의 별도
      // 버그와 겹쳐서 "Bloomberg로 바꿔도 자꾸 배송으로 돌아간다"는 증상으로
      // 나타남(사용자 신고, 8/26) - 그 버그도 같이 고침.
      row.itemType=sel.value;
      applySelClass();
      persistItemField(row.itemId,'item_type',sel.value)
        .catch(err=>toast('Type 저장 실패 (DB 반영 안 됨): '+err.message));
    });
    typeTd.appendChild(sel);
    applyHighlight(typeTd, row.itemId, 'type', row);
    tr.appendChild(typeTd);

    const checkTd=document.createElement('td');
    checkTd.className='check-cell';
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.className='row-check'; cb.dataset.rowKey=row.rowKey;
    cb.addEventListener('change', syncSelectAll);
    groupCheckboxes.push(cb);
    checkTd.appendChild(cb);
    tr.appendChild(checkTd);

    if(i===0){
      const ordTd=document.createElement('td');
      ordTd.className='ord-cell';
      ordTd.rowSpan=n;

      selectAllCb=document.createElement('input');
      selectAllCb.type='checkbox';
      selectAllCb.className='order-select-all';
      selectAllCb.title='이 오더의 항목 전체 선택/해제 (병합·완료·삭제·날짜이동 등에 사용)';
      selectAllCb.addEventListener('change', ()=>{
        groupCheckboxes.forEach(cb=>{cb.checked=selectAllCb.checked});
      });
      ordTd.appendChild(selectAllCb);

      // main line: the SAP order code (ZOR/ZRE/ZRX/ZINP/ZINT/...) + order number - bold, black.
      // Persisted as of 2026-08-06 - but only the TYPE portion. order_no itself is never
      // sent to the server (see update_order_field()'s own comment for why: it's the real
      // matching key SAP resync/POD/Portal automation all key off, so silently renaming it
      // in place is too risky) - if the number portion is edited, it's reverted on-screen
      // and a toast explains why, while the type-portion edit (if any) still saves normally.
      const main=document.createElement('div');
      main.contentEditable='true'; main.spellcheck=false; main.textContent=row.orderLabel;
      const orderItemIds=group.rows.map(r=>r.itemId);
      main.addEventListener('blur', ()=>{
        const v=main.textContent.trim();
        const realNo=group.rows[0].orderNo||'';
        let newType, newNo;
        if(!realNo){
          // No real SAP order number tied to this row (memo rows, and any
          // other free-text label - e.g. "SDSK1333608359") - there is
          // nothing to protect, so the whole typed text IS the label as-is.
          // 2026-08-27 bug fix: this used to always split on the first space
          // and treat everything after it as "the order number", discarding
          // it whenever it didn't match a real order_no - for a blank-
          // order_no row that meant ANY space in the text silently ate
          // everything typed after it (e.g. "SDSK 1333608359" saved as just
          // "SDSK"). Splitting only makes sense when there IS a real order_no
          // to compare the typed "number" portion against.
          newType=v;
          newNo='';
        } else {
          const parts=v.split(/\s+/);
          newType=parts[0]||'';
          newNo=parts.slice(1).join(' ');
          if(newNo && newNo!==realNo){
            toast('오더번호는 여기서 수정할 수 없습니다 - 타입만 저장됩니다');
          }
        }
        const savedLabel=(newType+' '+realNo).trim();
        group.rows.forEach(r=>r.orderLabel=savedLabel);
        renderAll();
        persistOrderField(orderItemIds, 'order_type', newType)
          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
      });
      main.addEventListener('keydown', e=>{ if(e.key==='Enter'){ e.preventDefault(); main.blur(); } });
      ordTd.appendChild(main);

      // 실패 배지 (todo b, 2026-08-18) - portalError/podError는 order 단위라
      // 그룹의 모든 행이 같은 값을 갖고 있음(row 대신 group.rows[0]에서 읽어도
      // 동일), 여기 ordTd는 rowSpan으로 그룹당 한 번만 그려지니 배지도 한 번만.
      // 클릭 시 실패 사유 팝업 - main(라벨 편집)의 contentEditable 영역과
      // 겹치지 않게 별도 span으로 붙이고 클릭 시 stopPropagation.
      const errMsg=row.portalError ? row.portalErrorMsg : (row.podError ? row.podErrorMsg : '');
      if(errMsg){
        const badge=document.createElement('span');
        badge.className='order-error-badge';
        badge.textContent='⚠';
        badge.title='클릭: 실패 사유 보기';
        badge.addEventListener('click', e=>{
          e.stopPropagation();
          window.alert((row.portalError ? 'Portal 등록(Serial/QR) 실패:\n\n' : 'POD(배송완료) 처리 실패:\n\n')+errMsg);
        });
        ordTd.appendChild(badge);
      }

      // secondary order-code line(s): Bloomberg-side codes (OBD/ORD/SDSK/...), one per
      // line, same small muted style, same free-text edit as everything else - not just
      // a bare number anymore, and more lines can be typed in directly if more than one
      // secondary code applies to this order. Fully persisted as of 2026-08-06 via
      // persistDeliveryCodes()/update_delivery_codes() - the server splits whichever
      // line starts with "OBD" into delivery_no (the real matching key) and joins
      // everything else (ORD/SDSK/etc) into extra_codes (plain text, no matching-key
      // role) - see update_delivery_codes()'s own comment.
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
        persistDeliveryCodes(orderItemIds, v)
          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
      });
      ordTd.appendChild(extra);
      applyHighlight(ordTd, row.itemId, 'order', row);

      tr.appendChild(ordTd);
    }

    if(!row.orderNo && n===1){
      // 2026-08-25: 메모 행은 addManualRow() 모달에서 고른 개수만큼만
      // 제목/M-N/S-N/Customer/Phone/Address/Memo 7칸 중 일부에 값이 들어있고
      // 나머지는 빈 문자열로 저장돼 있다(manualRowGroupStarts() 참고, 정확히
      // 각 그룹의 "시작 열"에만 값이 있음). 2026-08-27 수정: 원래 이 조건이
      // orderLabel==='memo' 리터럴 문자열 매치였는데, Order#를 사용자가 직접
      // 입력할 수 있게 되면서(add_manual_note()의 order_label 참고) 라벨이
      // "memo"가 아닌 메모 행("SDSK1333608359" 등)에서 이 압축 렌더링이
      // 깨지는 문제가 생겨 !row.orderNo(실제 SAP 오더번호가 없는 행 전부)로
      // 바꿈 - 어차피 order_no는 메모 행에서 항상 빈 값이므로 라벨 텍스트와
      // 무관하게 정확히 같은 행 집합을 가리킨다. 다른 오더와 병합된 경우
      // (n>1, 드문 케이스)는 아래 일반 렌더링으로 그대로 둠.
      const fields=[
        ['description','item-cell', v=>row.description=v, v=>persistItemField(row.itemId,'description',v)],
        ['material','mn-cell', v=>row.material=v, v=>persistItemField(row.itemId,'material',v)],
        ['serial','sn-cell', v=>row.serial=v, v=>persistItemField(row.itemId,'serial',v)],
        ['customer','cust-cell customer-cell', v=>row.customer=v, v=>persistOrderField([row.itemId],'customer',v)],
        ['phone','cust-cell phone-cell', v=>row.phone=v, v=>persistOrderField([row.itemId],'phone',v)],
        ['address','cust-cell addr-cell', v=>row.address=v, v=>persistOrderField([row.itemId],'address',v)],
        ['memo','cust-cell memo-cell', v=>row.memo=v, v=>persistOrderField([row.itemId],'memo',v)],
      ];
      // 2026-08-27 버그 수정: 칸 경계를 "값이 있는 열부터 시작해서 뒤따르는
      // 빈 열들을 흡수" 방식으로만 정하면, 사용자가 일부러 7칸으로 만들고
      // 몇 칸은 선택사항이라 비워둔 경우까지 그 빈 칸들을 앞 칸에 합쳐버려
      // (신고 사례: 제목만 채운 7칸짜리 메모가 1칸으로 보임) 실제로 몇 칸을
      // 만들었는지와 화면이 안 맞았다. addManualRow()가 고른 칸 수를
      // manual_cell_count로 저장해두고(add_manual_note() 참고), 있으면 그
      // 값 그대로 manualRowGroupStarts()로 경계를 재현 - 내용 유무와 무관하게
      // 항상 정확히 그 칸 수만큼 그려짐. 그 값이 없는(이 저장 방식 이전에
      // 만들어진) 옛 메모 행만 예전처럼 내용 기준으로 추정.
      const storedCount=parseInt(row.manualCellCount, 10);
      const starts=(storedCount>=1 && storedCount<=fields.length)
        ? manualRowGroupStarts(storedCount)
        : null;
      if(starts){
        starts.forEach((start,gi)=>{
          const end=gi+1<starts.length ? starts[gi+1] : fields.length;
          const span=end-start;
          const [field, cellClass, setLocal, persist]=fields[start];
          const td=makeEditableTd(row[field]||'', cellClass, setLocal, persist);
          if(span>1) td.colSpan=span;
          applyHighlight(td, row.itemId, field, row);
          tr.appendChild(td);
        });
      } else {
        let idx=0;
        while(idx<fields.length){
          const [field, cellClass, setLocal, persist]=fields[idx];
          let span=1;
          while(idx+span<fields.length && !(row[fields[idx+span][0]]||'').trim()) span++;
          const td=makeEditableTd(row[field]||'', cellClass, setLocal, persist);
          if(span>1) td.colSpan=span;
          applyHighlight(td, row.itemId, field, row);
          tr.appendChild(td);
          idx+=span;
        }
      }
    } else {
    const descTd=makeEditableTd(row.description, 'item-cell', v=>row.description=v, v=>persistItemField(row.itemId,'description',v));
    applyHighlight(descTd, row.itemId, 'description', row);
    tr.appendChild(descTd);
    const matTd=makeEditableTd(row.material, 'mn-cell', v=>row.material=v, v=>persistItemField(row.itemId,'material',v));
    applyHighlight(matTd, row.itemId, 'material', row);
    tr.appendChild(matTd);
    // 2026-08-18: 토글 배지는 UI가 번거롭다는 사용자 피드백으로 제거함 -
    // 그냥 원래대로 편집만 되는 셀. "serial을 안 적으면 시리얼 없는
    // 물품"이라는 사용자의 실제 워크플로우는 저장 시점이 아니라 Portal
    // 등록 버튼(run_portal())이 클릭되는 시점에 그때의 serial 값을 보고
    // 판단함 - item["mode"]가 아니라 item["serial"]이 비었는지로 직접
    // 판정하도록 run_portal()을 고쳤음(위쪽 참고, ZINP 551059036
    // "COPYHOLDER FELLOWES BOOKLIFT" qty=1 계기). 그래서 여기 클라이언트
    // 쪽엔 mode를 건드릴 이유가 없음 - serial 저장만 하면 끝.
    const serTd=makeEditableTd(row.serial, 'sn-cell', v=>row.serial=v, v=>persistItemField(row.itemId,'serial',v));
    applyHighlight(serTd, row.itemId, 'serial', row);
    // S/N 오더 반영(paper relo 판별) 실패 배지 (2026-08-19) - portalError/
    // podError의 order-error-badge와 같은 스타일이지만 이건 항목 단위라
    // ordTd(그룹당 1회)가 아니라 이 항목의 sn-cell에 직접 붙인다.
    if(row.snReloError){
      const snBadge=document.createElement('span');
      snBadge.className='order-error-badge';
      snBadge.textContent='⚠';
      snBadge.title='클릭: S/N 오더 반영 실패 사유 보기';
      snBadge.addEventListener('click', e=>{
        e.stopPropagation();
        window.alert('S/N 오더 반영 실패 (paper relo 필요할 수 있음):\n\n'+row.snReloErrorMsg);
      });
      serTd.appendChild(snBadge);
    }
    tr.appendChild(serTd);

    if(i===0){
      // first non-blank value across the group, not just group.rows[0] - a
      // merged pickup leg with no customer/phone/address of its own (e.g.
      // ZRX 67066585's 회수 record) should still show/inherit the delivery
      // leg's values rather than rendering blank. Each field's own span
      // (allSpans[field][gi]) can extend across OTHER orders independently
      // when that field's been manually merged (mergeChecked()) - e.g.
      // address merged across 3 orders while customer stays per-order -
      // editing a merged cell then updates every row in THAT field's merged
      // block, not just this order-group's own rows.
      [['customer','customer-cell'],['phone','phone-cell'],['address','addr-cell'],['memo','memo-cell']]
        .forEach(([field, cellClass])=>{
          const fs=allSpans[field][gi];
          if(!fs.isFirst) return;
          const td=makeEditableTd(
            resolvedField(group, field), 'cust-cell '+cellClass, v=>fs.mergedRows.forEach(r=>r[field]=v),
            v=>persistOrderField(fs.mergedRows.map(r=>r.itemId), field, v)
          );
          td.rowSpan=fs.span;
          applyHighlight(td, row.itemId, field, row);
          tr.appendChild(td);
        });
    }
    }

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
    exportCb.title='엑셀로 내보낼 날짜로 선택 (이 날짜의 행 체크박스가 전부 함께 체크/해제됨 - 마감 전에 아직 해결 안 된 행만 개별로 체크 해제하면 이번 마감에서 빠짐)';
    exportCb.checked=EXPORT_DATES.has(iso);
    exportCb.addEventListener('change', ()=>{
      if(exportCb.checked) EXPORT_DATES.add(iso); else EXPORT_DATES.delete(iso);
      // 날짜 체크박스 = "이 날짜 전체를 기본 선택" 편의 기능 (2026-08-27 user
      // request) - 개별 행(.row-check)을 전부 이 상태로 맞춘 뒤, 특정 행만
      // 다시 체크 해제해서 이번 마감에서 뺄 수 있게 한다. 반대 방향(행을
      // 개별로 건드려도 날짜 체크박스 자체는 되돌리지 않음)은 order-select-all
      // 패턴과 달리 의도적으로 안 함 - 이건 "부분 상태"를 표시하는 용도가
      // 아니라 순수 일괄-선택 버튼이라 단순하게 유지.
      //
      // 2026-09-02 실사고로 수정: 예전엔 상태 무관하게 그 날짜 행을 전부
      // 체크했는데, 오더 67083939처럼 같은 오더의 항목이 날짜별로 갈려있고
      // (배송 9/1 완료=파랑, 회수는 9/8로 옮겨졌지만 아직 미확정) 사용자가
      // 그 날짜 체크박스를 개별 언체크하는 걸 깜빡하면, 아직 완료 안 된
      // 행까지 마감 파일에 그대로 들어가버렸다("파란색 마감된 것만 들어가야
      // 함" 사용자 지적). 이제 체크(on)할 때는 board_status가 'done'(파랑)인
      // 행만 자동으로 체크하고, 나머지(대기/취소/미지정)는 그대로 둔다 -
      // 끌 때(off)는 상태 무관하게 전부 해제(선택 취소는 항상 안전하므로).
      tbody.querySelectorAll('.row-check').forEach(cb=>{
        if(!exportCb.checked){ cb.checked=false; return; }
        const row=ROWS.find(r=>r.rowKey===cb.dataset.rowKey);
        cb.checked = !!(row && row.uiStatus==='done');
      });
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
      const groups=groupRows(rows);
      const allSpans=computeAllFieldSpans(groups);
      groups.forEach((g,gi)=>buildGroupRows(g, allSpans, gi).forEach(tr=>tbody.appendChild(tr)));
    }
    table.appendChild(tbody);
  });
}
renderAll();

// Shift-click range select on row checkboxes (Excel-style): click one row's
// checkbox, then shift-click another, and every row in between is set to
// match the shift-clicked box's new state. Plain clicks need no extra
// handling for "select several" - checkboxes are already additive, each
// click just toggles that one row on top of whatever else is checked.
// Delegated on the table (which itself is never recreated - only its
// tbody.date-group sections are, on every renderAll()) so this keeps working
// across re-renders instead of needing to be re-attached per row.
{
  let lastCheckedCb=null;
  document.getElementById('boardTable').addEventListener('click', e=>{
    const cb=e.target.closest('.row-check');
    if(!cb) return;
    if(e.shiftKey && lastCheckedCb && lastCheckedCb.isConnected){
      const all=[...document.querySelectorAll('.row-check')];
      const from=all.indexOf(lastCheckedCb), to=all.indexOf(cb);
      if(from!==-1 && to!==-1){
        const [lo,hi]=from<to?[from,to]:[to,from];
        const state=cb.checked;
        for(let i=lo;i<=hi;i++){
          if(all[i].checked!==state){ all[i].checked=state; all[i].dispatchEvent(new Event('change')); }
        }
      }
    }
    lastCheckedCb=cb;
  });
}

// -- Item-level persistence helpers, shared by status/move actions below.
// Each posts to workbench.db first and only touches ROWS/renders on success,
// same pattern as deleteChecked() - never leave the screen out of sync with
// what actually got saved. --
async function persistItemStatus(itemIds, status){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/status', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, status})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'상태 저장 실패');
  return true;
}
// Done/Ready/Cancel/초기화 버튼은 사용자가 그 행을 직접 확인/처리했다는
// 신호이므로, set_items_board_status()가 서버(orders/order_items 테이블)에서
// 이미 지운 portal_error/pod_error/sn_relo_error를 화면에도 그 자리에서 함께
// 지운다 - 안 그러면 최대 1.5초짜리 board_version 폴링(5301행)이 돌 때까지
// 주황(row-error)이 그대로 남아 "버튼 눌러도 에러 표시가 안 없어진다"는 오해를
// 만든다(2026-09-02 사용자 지적: 수동으로 문제를 고치고 이 버튼들을 눌러도
// 자동화가 스스로 재시도해서 성공할 때만 지워지는 것처럼 보였음). portalError/
// podError는 orders 테이블 값이라 같은 오더의 다른 행에도 동일하게 붙어있으므로
// (board_row()의 해당 주석 참고) 체크된 항목뿐 아니라 같은 orderId를 가진 모든
// 행에서 지워야 서버 상태와 어긋나지 않는다; snReloError는 항목 단위라 체크된
// 항목 자신만 지운다.
function clearRowErrorFlags(itemIds){
  const idSet=new Set(itemIds);
  const orderIds=new Set(ROWS.filter(r=>idSet.has(r.itemId)).map(r=>r.orderId));
  ROWS.forEach(r=>{
    if(orderIds.has(r.orderId)){ r.portalError=false; r.portalErrorMsg=''; r.podError=false; r.podErrorMsg=''; }
    if(idSet.has(r.itemId)){ r.snReloError=false; r.snReloErrorMsg=''; }
  });
}
async function persistItemDate(itemIds, iso){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/move_date', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, date: iso})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'날짜 이동 저장 실패');
  return true;
}
async function persistItemReorder(itemIds){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/reorder', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'순서 저장 실패');
  return true;
}
// -- Inline-edit persistence: description/material/serial (per item) and
// customer/phone/address/memo (per order, resolved from the edited item ids -
// a shared cell can span multiple orders once merged). These cells used to
// be screen-only (ROWS-only, lost on reload/restart) - exactly the gap that
// silently ate a typed-in serial number. --
async function persistItemField(itemId, field, value){
  const res=await fetch('/api/items/edit_field', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemId, field, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}
async function persistOrderField(itemIds, field, value){
  const res=await fetch('/api/orders/edit_field', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, field, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}
async function persistCustomType(name){
  const res=await fetch('/api/custom_types/add', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}
async function persistDeliveryCodes(itemIds, value){
  const res=await fetch('/api/orders/edit_delivery_codes', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}

// -- Undo: reverts the most recent delete/move/status batch in workbench.db
// (server-side history, capped to the last 20 actions). Reloads the page
// afterward instead of patching ROWS in place - undo can touch items that
// aren't even on the currently-loaded board (e.g. undoing a delete brings an
// item's date section back), so a full reload is the simple, always-correct
// way to pick up whatever state the server ends up in. --
async function undoLastAction(){
  try{
    const res=await fetch('/api/undo', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행취소 실패');
    if(!data.undone){ toast(data.message||'되돌릴 작업이 없습니다'); return }
    toast('실행취소: '+data.summary);
    location.reload();
  }catch(err){
    toast('실행취소 실패: '+err.message);
  }
}
document.addEventListener('keydown', e=>{
  if((e.ctrlKey || e.metaKey) && e.key.toLowerCase()==='z'){
    const el=document.activeElement;
    const isEditing = el && (el.tagName==='INPUT' || el.tagName==='TEXTAREA' || el.isContentEditable);
    if(isEditing) return; // let the browser handle undo inside a text field being edited
    e.preventDefault();
    undoLastAction();
  }
});

// -- Trash: browse/restore soft-deleted items outside Ctrl+Z's
// UNDO_HISTORY_LIMIT-action reach - anything the server still returns here
// is, by definition, within DELETED_ITEM_RETENTION_DAYS (not yet hard-purged).
// Built with createElement/textContent throughout, not innerHTML, same as
// every other cell in this app - order/customer/memo text can contain
// anything typed into SAP or the sheet, so it's never treated as HTML. --
async function openTrash(){
  document.getElementById('trashOverlay').classList.add('show');
  await refreshTrashList();
}
function closeTrash(){
  document.getElementById('trashOverlay').classList.remove('show');
}
async function refreshTrashList(){
  const container=document.getElementById('trashList');
  container.textContent='불러오는 중...';
  try{
    const res=await fetch('/api/items/deleted');
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
    renderTrashList(data.items||[]);
  }catch(err){
    container.textContent='조회 실패: '+err.message;
  }
}
function daysLeftText(deletedAt){
  const deleted=new Date(String(deletedAt||'').replace(' ','T'));
  if(isNaN(deleted.getTime())) return '';
  const daysLeft=Math.max(0, DELETED_ITEM_RETENTION_DAYS - Math.floor((Date.now()-deleted.getTime())/86400000));
  return daysLeft+'일 남음';
}
function renderTrashList(items){
  const container=document.getElementById('trashList');
  container.textContent='';
  if(!items.length){
    const empty=document.createElement('div');
    empty.className='trash-empty';
    empty.textContent='삭제된 항목이 없습니다.';
    container.appendChild(empty);
    return;
  }
  const table=document.createElement('table');
  table.className='trash-table';
  const thead=document.createElement('thead');
  const headRow=document.createElement('tr');
  ['', 'Order#', 'Customer', 'Item', 'M/N', 'S/N', '삭제 일시', '남은 기간'].forEach(label=>{
    const th=document.createElement('th'); th.textContent=label; headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);
  const tbody=document.createElement('tbody');
  items.forEach(item=>{
    const tr=document.createElement('tr');
    const checkTd=document.createElement('td');
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.className='trash-check'; cb.value=item.id;
    checkTd.appendChild(cb);
    tr.appendChild(checkTd);
    [
      `${item.order_type||''} ${item.order_no||''}`.trim(),
      item.customer||'',
      item.description||'',
      item.material||'',
      item.serial||'',
      item.deleted_at||'',
      daysLeftText(item.deleted_at),
    ].forEach(text=>{
      const td=document.createElement('td'); td.textContent=text; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  container.appendChild(table);
}
async function restoreCheckedTrash(){
  const itemIds=[...document.querySelectorAll('.trash-check:checked')].map(cb=>parseInt(cb.value, 10));
  if(!itemIds.length){toast('복구할 항목을 선택하세요');return}
  try{
    const res=await fetch('/api/items/restore', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'복구 실패');
    toast(data.restored+'건 복구됨');
  }catch(err){toast('복구 실패: '+err.message);return}
  location.reload();
}

// -- Process: acts on whichever rows are checked, same checkboxes Portal uses --
async function processDone(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, 'done');
  }catch(err){toast('완료 처리 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='done'});
  clearRowErrorFlags(itemIds);
  renderAll();
  toast(keys.length+'건 완료 처리(파랑) - DB에 반영됨');
}
async function processPickupReady(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>keys.includes(r.rowKey));
  const pendingIds=checkedRows.filter(r=>r.itemType==='회수').map(r=>r.itemId);
  const doneIds=checkedRows.filter(r=>r.itemType!=='회수').map(r=>r.itemId);
  try{
    await Promise.all([persistItemStatus(pendingIds,'pending'), persistItemStatus(doneIds,'done')]);
  }catch(err){toast('상태 저장 실패: '+err.message);return}
  checkedRows.forEach(row=>{row.uiStatus = row.itemType==='회수' ? 'pending' : 'done'});
  clearRowErrorFlags([...pendingIds, ...doneIds]);
  renderAll();
  toast(keys.length+'건 처리: 회수는 대기(빨강), 나머지는 완료(파랑) - DB에 반영됨');
}
// 2026-08-07: 취소된 행 표시 (회색, 줄 전체) - 우클릭 셀 강조가 아니라
// Done/Ready to process와 같은 행 단위 상태로 만들어달라는 사용자 요청.
// board_status에 'cancelled'를 저장 - 삭제는 아니므로 행은 그대로 board에
// 남고(복구는 Done/Ready to process를 다시 눌러 원래 상태로 되돌리면 됨),
// undo(Ctrl+Z)도 다른 status 변경과 동일하게 적용됨.
async function processCancel(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, 'cancelled');
  }catch(err){toast('취소 처리 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='cancelled'});
  clearRowErrorFlags(itemIds);
  renderAll();
  toast(keys.length+'건 취소 처리(회색) - DB에 반영됨');
}
// 2026-08-20: "초기화" - Done/Ready to process/Cancel을 누른 적 없는
// 원래 상태(흰색, board_status='')로 되돌림. 실사고 계기: ZRX 67079036 -
// 배송은 완료했지만 회수는 다음날로 미뤄졌는데 같은 오더 안에서 함께
// Done 처리돼 회수 행도 파랑으로 남아있던 걸 발견 (오더 단위가 아니라
// 행/아이템 단위 상태이므로, 되돌릴 대상도 체크된 행만 - 오더 전체가
// 아니라 회수 행 하나만 체크해서 초기화할 수 있음). set_items_board_status()는
// 이미 임의의 문자열(빈 문자열 포함, 로그 메시지도 '(초기화)'로 이미
// 대비돼 있었음)을 그대로 저장하므로 백엔드 변경 없이 프론트에서만 추가.
async function processReset(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, '');
  }catch(err){toast('초기화 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus=''});
  clearRowErrorFlags(itemIds);
  renderAll();
  toast(keys.length+'건 초기화(흰색) - DB에 반영됨');
}
// 2026-08-20: 실제 SAP 오더 수집 없이 텍스트만으로 board에 새 행을 만드는
// 기능 - 사용자 요청("메모용으로도 한줄씩 추가하고 싶을때가 있다"), 이후
// "구분/날짜도 고를 수 있게 해달라"는 후속 요청으로 단순 prompt() 대신
// 모달로 교체 - trashOverlay와 같은 오버레이 패턴 재사용. 새로 생긴 행은
// 서버가 만든 것이라 로컬 ROWS 배열에 없으므로(다른 버튼들처럼 로컬 상태만
// 패치할 수 없음) 저장 후 전체 새로고침.
// Type/check를 제외한 나머지 7개 board 열, 왼쪽부터의 순서 (add_manual_note()의
// cells 매핑과 반드시 같은 순서여야 함) - MANUAL_ROW_COLUMNS 선언은 파일
// 위쪽(ROWS 바로 아래)으로 옮겨짐: buildGroupRows()의 메모 행 렌더링이
// manualRowGroupStarts()를 통해 이 값을 쓰는데, 그 렌더링은 최초
// renderAll() 호출(이 지점보다 훨씬 앞)때부터 실행되므로 const가 원래
// 자리(여기)에 있으면 TDZ(temporal dead zone)로 "Cannot access
// 'MANUAL_ROW_COLUMNS' before initialization"가 나며 전체 스크립트가
// 죽는다 - 2026-08-27 실사고: 메모 행 추가 버튼을 누르자마자 보드 전체가
// (마지막 렌더된 날짜 한 줄만 남기고) 하얗게 사라짐, 콘솔에 정확히 이
// 에러가 찍힘.
// 2026-08-25 정정: "몇 칸을 쓸지" 물어서 그 수만큼 칸을 만든다는 게, 칸마다
// 원래 열 하나씩(좁게) 쓴다는 뜻이 아니라 - 예를 들어 1칸만 쓰면 나머지 6개
// 열 공간을 전부 병합해서 하나의 넓은 칸으로 쓴다는 뜻이었음(사용자 정정).
// 그래서 7개 열을 N개로 최대한 고르게 나눠 각 칸이 여러 열 너비를 차지하게
// 함 - manualRowGroupStarts(n)이 그 경계를 계산하고, 실제 저장은 각 그룹의
// "시작 열" 자리에만 값을 넣고 나머지는 빈 문자열로 둔다(add_manual_note()의
// cells와 동일한 7칸짜리 배열, 위치만 건너뛰며 채움). 렌더링(buildGroupRows)
// 쪽은 그 결과를 "값이 있는 열부터 시작해서 뒤따르는 빈 열들을 흡수해
// colspan으로 병합"하는 방식으로 다시 그리므로, 이 함수가 계산한 경계와
// 항상 자동으로 일치한다 - 즉 N을 별도로 저장할 필요가 없다(빈 칸이 아닌
// 열의 개수 자체가 곧 N이므로, 나중에 셀을 직접 편집해서 비우면 다음
// 렌더링 때 그 칸만큼 자연스럽게 옆 칸에 합쳐짐).
function manualRowGroupStarts(n){
  const total=MANUAL_ROW_COLUMNS.length;
  const base=Math.floor(total/n), rem=total%n;
  const starts=[]; let acc=0;
  for(let i=0;i<n;i++){ starts.push(acc); acc += base+(i<rem?1:0); }
  return starts;
}
function renderManualRowCells(){
  const countInput=document.getElementById('manualRowCellCount');
  const n=Math.max(1, Math.min(MANUAL_ROW_COLUMNS.length, parseInt(countInput.value)||1));
  countInput.value=n;
  const starts=manualRowGroupStarts(n);
  const wrap=document.getElementById('manualRowCells');
  const prevValues=[...wrap.querySelectorAll('.manualRowCellInput')].map(i=>i.value);
  wrap.innerHTML='';
  starts.forEach((start,i)=>{
    const end=(i+1<starts.length ? starts[i+1] : MANUAL_ROW_COLUMNS.length)-1;
    const rangeLabel=end>start ? MANUAL_ROW_COLUMNS[start]+'~'+MANUAL_ROW_COLUMNS[end] : MANUAL_ROW_COLUMNS[start];
    const label=document.createElement('label');
    label.style.cssText='display:block;margin-bottom:8px;font-size:13px;color:var(--muted,#555)';
    label.textContent=rangeLabel+' 칸'+(i===0?' (필수)':' (선택)');
    const input=document.createElement('input');
    input.type='text';
    input.className='manualRowCellInput';
    input.dataset.start=start;
    input.style.cssText='display:block;width:100%;margin-top:4px;padding:6px 8px;box-sizing:border-box';
    input.value=prevValues[i]||'';
    if(i===0) input.placeholder='제목 (필수)';
    label.appendChild(input);
    wrap.appendChild(label);
  });
  wrap.querySelector('.manualRowCellInput').focus();
}
function addManualRow(){
  document.getElementById('manualRowOrderLabel').value='';
  // 기본값 = 일반 오더 행과 같은 7칸 전부(칸마다 정확히 열 하나씩) - 2026-08-27
  // user request. 필요하면 이 값을 줄여서 병합된 넓은 칸 몇 개로 바꿔 쓸 수
  // 있음(기존 동작 그대로, manualRowGroupStarts() 참고).
  document.getElementById('manualRowCellCount').value=MANUAL_ROW_COLUMNS.length;
  renderManualRowCells();
  const sel=document.getElementById('manualRowType');
  sel.innerHTML='';
  allTypeOptions().forEach(v=>{
    const opt=document.createElement('option');
    opt.value=v; opt.textContent=v;
    sel.appendChild(opt);
  });
  sel.value='배송';
  document.getElementById('manualRowDate').value=new Date().toISOString().slice(0,10);
  document.getElementById('manualRowOverlay').classList.add('show');
}
function closeManualRowModal(){
  document.getElementById('manualRowOverlay').classList.remove('show');
}
async function submitManualRow(){
  const cells=new Array(MANUAL_ROW_COLUMNS.length).fill('');
  document.querySelectorAll('.manualRowCellInput').forEach(inp=>{
    cells[parseInt(inp.dataset.start)]=inp.value.trim();
  });
  const itemType=document.getElementById('manualRowType').value;
  const date=document.getElementById('manualRowDate').value;
  const orderLabel=document.getElementById('manualRowOrderLabel').value.trim();
  // 선택한 칸 수 그대로 서버에 저장(manual_cell_count) - 렌더링이 내용 유무가
  // 아니라 이 값으로 칸 경계를 그리도록(아래 buildGroupRows 참고, 2026-08-27
  // 버그 수정: 7칸으로 만들고 제목만 채워도 나머지 6칸이 빈 칸인 채로 남아있지
  // 않고 1칸으로 합쳐져 보이던 문제).
  const cellCount=parseInt(document.getElementById('manualRowCellCount').value)||cells.length;
  if(!cells[0]){toast('제목이 비어 있습니다');return}
  try{
    const res=await fetch('/api/orders/add_manual', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({cells, itemType, date, orderLabel, cellCount})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'추가 실패');
  }catch(err){toast('메모 행 추가 실패: '+err.message);return}
  closeManualRowModal();
  location.reload();
}

// -- Merge: manual "Excel merge cells", independently per field, for
// customer/phone/address/memo across different order numbers that are
// really the same customer (e.g. ZOR 7825897 / ZOR 7825832 / ZOR 7825838
// sharing one address even when the name on file differs per order - so
// each field is asked for and merged separately, not all 4 as one block).
// Which fields to merge comes from the persistent .merge-field-cb checkbox
// row (Customer/Phone/Address/Memo checked by default; Order# is present
// for layout symmetry but not wired to anything yet - filtered out
// server-side against MERGE_FIELDS regardless). Works on whichever rows are
// checked, same checkboxes as everything else - resolves them to their
// distinct orders server-side. Only renders merged once those orders are
// also adjacent on the board (drag rows next to each other first via the
// drag handle, same-date reorder). --
function selectedMergeFields(){
  return [...document.querySelectorAll('.merge-field-cb:checked')]
    .map(cb=>cb.value).filter(f=>MERGE_FIELDS.includes(f));
}
async function mergeChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const fields=selectedMergeFields();
  if(!fields.length){toast('병합할 칸을 하나 이상 체크하세요');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  let data;
  try{
    const res=await fetch('/api/items/merge', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, fields})
    });
    data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'병합 실패');
  }catch(err){toast('병합 실패: '+err.message);return}
  if(!data.merged){toast(data.message||'서로 다른 오더에 속한 행을 2건 이상 선택하세요');return}
  // mergeGroups lives on rows fetched fresh from the server (order_payload),
  // not something ROWS can patch itself - simplest correct way to pick it up.
  location.reload();
}
async function unmergeChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const fields=selectedMergeFields();
  if(!fields.length){toast('해제할 칸을 하나 이상 체크하세요');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    const res=await fetch('/api/items/unmerge', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, fields})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'병합 해제 실패');
    toast(data.unmerged+'건 병합 해제 - DB에 반영됨');
  }catch(err){toast('병합 해제 실패: '+err.message);return}
  location.reload();
}

// -- Export: the one action in this app that actually writes a real file
// (the live shipping Excel workbook) instead of just changing what's on
// screen. Builds the current board state (post move/edit/merge/delete) for
// just the checked rows (`checkedRowKeys`) of a given date, and sends it to
// the server to write/append into that date's sheet. Checking a date's
// banner checkbox checks every row under it by default (see renderAll's
// exportCb listener) - unchecking a specific row before hitting 마감 excludes
// only that row, leaving it on the board for a later 마감 once it's actually
// resolved (2026-08-27 user request). --
function buildExportRows(iso, checkedRowKeys){
  const rows=ROWS.filter(r=>r.date===iso && checkedRowKeys.has(r.rowKey));
  const groups=groupRows(rows);
  const allSpans=computeAllFieldSpans(groups);
  const out=[];
  groups.forEach((group,gi)=>{
    const customer=resolvedField(group,'customer');
    const phone=resolvedField(group,'phone');
    const address=resolvedField(group,'address');
    const memo=resolvedField(group,'memo');
    group.rows.forEach((row, i)=>{
      out.push({
        // each field merges its own E-H column down its own independent run
        // (see _write_export_sheet) - a manually-merged field (mergeChecked())
        // can span a different set of orders than another merged field.
        mergeFirst: {
          customer: i===0 && allSpans.customer[gi].isFirst,
          phone: i===0 && allSpans.phone[gi].isFirst,
          address: i===0 && allSpans.address[gi].isFirst,
          memo: i===0 && allSpans.memo[gi].isFirst,
        },
        itemType: row.itemType,
        orderLabel: row.orderLabel,
        deliveryNo: row.deliveryNo,
        description: row.description,
        material: row.material,
        serial: row.serial,
        customer, phone, address, memo,
        // carries this row's manual cell highlight (see cell_highlights) into
        // the Excel export - only meaningful on the row that actually owns
        // each field's <td> (i===0 && isFirst, same rows mergeFirst already
        // marks), which is exactly where the highlight was ever recorded.
        highlights: row.highlights || {},
      });
    });
  });
  return out;
}

// 비상용 엑셀 → workbench 불러오기. 평상시엔 SAP 수집이 직접 workbench.db에
// 반영되므로(import_sap_rows) 이 버튼은 자동 동기화를 대신하지 않는다 - 엑셀에
// 직접 손댄 내용을 다시 가져와야 하는 드문 경우에만, 시트+행 범위를 명시해서
// 딱 그만큼만 1회성으로 긁어온다 (통째로 다시 스캔하지 않음).
async function importExcelRange(){
  const sheet=document.getElementById('excelRangeSheet').value.trim();
  const rowStart=document.getElementById('excelRangeStart').value;
  const rowEnd=document.getElementById('excelRangeEnd').value;
  if(!sheet){toast('시트명을 입력하세요 (예: 8-6)');return}
  if(!rowStart||!rowEnd){toast('시작행/종료행을 입력하세요');return}
  if(!window.confirm(`"${sheet}" 시트의 ${rowStart}~${rowEnd}행을 workbench로 불러옵니다. 계속할까요?`))return;
  toast('엑셀에서 불러오는 중...');
  try{
    const res=await fetch('/api/import/range', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({sheet, rowStart:Number(rowStart), rowEnd:Number(rowEnd)})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'불러오기 실패')}
    toast(`불러오기 완료 - ${data.imported}개 오더 반영`);
    location.reload();
  }catch(err){
    toast('불러오기 실패: '+err.message);
  }
}

// -- 배송장 인쇄: Type/Order#/Item/M-N/S-N/Customer/Phone/Address/Memo만 뽑아
// A4 가로/흑백 전용 테이블을 새로 만들어 인쇄한다. 선택은 새 UI를 만들지 않고
// 이미 있는 두 체크박스를 그대로 재사용한다 - 날짜 배너 체크박스(EXPORT_DATES,
// 엑셀 내보내기와 같은 것)는 "그 날짜 전체", 행 체크박스(.row-check)는 체크된
// 행만(한쪽 leg만 체크했으면 그 leg만 인쇄되고 나머지 leg는 안 찍힘) - 다만
// 같은 오더의 체크된 행이 여럿이면 그 체크된 것들끼리는 Customer/Phone/
// Address/Memo를 보드 화면과 똑같이 rowSpan으로 한 번만 찍는다.
// computeAllFieldSpans는 인쇄 대상으로 뽑힌(체크된) 행들만 놓고 다시 돌려서 -
// 화면엔 붙어 있어도 체크 안 된 오더/행까지 병합 셀에 끼어들지 않게 한다. --
function printChecked(){
  const checkedRowKeys=new Set([...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey));
  const sections=[];
  allDatesSorted().forEach(iso=>{
    const rows=ROWS.filter(r=>r.date===iso);
    const groups=groupRows(rows);
    const wholeDate=EXPORT_DATES.has(iso);
    const selectedGroups=wholeDate ? groups : groups
      .map(g=>({key:g.key, rows:g.rows.filter(r=>checkedRowKeys.has(r.rowKey))}))
      .filter(g=>g.rows.length);
    if(!selectedGroups.length) return;
    sections.push({iso, groups:selectedGroups, allSpans:computeAllFieldSpans(selectedGroups)});
  });
  if(!sections.length){
    toast('인쇄할 날짜 또는 행이 선택되지 않았습니다 (날짜 줄 체크박스 또는 행 체크박스를 먼저 선택하세요)');
    return;
  }

  const printArea=document.getElementById('printArea');
  printArea.textContent='';
  const title=document.createElement('h2');
  title.textContent='배송장';
  printArea.appendChild(title);

  sections.forEach(({iso, groups, allSpans}, si)=>{
    const h3=document.createElement('h3');
    // 인쇄물에는 "(Today)"/"(Tomorrow)" 같은 화면 전용 문구를 빼고 순수 날짜만.
    h3.textContent=labelFor(iso).replace(/\s*\((Today|Tomorrow)\)\s*$/,'');
    if(si>0) h3.style.pageBreakBefore='always';
    printArea.appendChild(h3);

    const table=document.createElement('table');
    table.className='print-table';
    const colgroup=document.createElement('colgroup');
    [5,10,15,8,9,8,9,21,15].forEach(w=>{
      const col=document.createElement('col'); col.style.width=w+'%'; colgroup.appendChild(col);
    });
    table.appendChild(colgroup);

    const thead=document.createElement('thead');
    const headTr=document.createElement('tr');
    ['Type','Order #','Item','M/N','S/N','Customer','Phone','Address','Memo'].forEach(h=>{
      const th=document.createElement('th'); th.textContent=h; headTr.appendChild(th);
    });
    thead.appendChild(headTr);
    table.appendChild(thead);

    const tbody=document.createElement('tbody');
    groups.forEach((group, gi)=>{
      const n=group.rows.length;
      group.rows.forEach((row, i)=>{
        const tr=document.createElement('tr');
        const typeTd=document.createElement('td'); typeTd.textContent=row.itemType||'';
        if(row.itemType==='배송') typeTd.className='print-type-delivery';
        tr.appendChild(typeTd);
        if(i===0){
          const ordTd=document.createElement('td'); ordTd.rowSpan=n; ordTd.textContent=row.orderLabel||'';
          tr.appendChild(ordTd);
        }
        const itemTd=document.createElement('td'); itemTd.textContent=row.description||''; tr.appendChild(itemTd);
        const mnTd=document.createElement('td'); mnTd.textContent=row.material||''; tr.appendChild(mnTd);
        const snTd=document.createElement('td'); snTd.textContent=row.serial||''; tr.appendChild(snTd);
        if(i===0){
          ['customer','phone','address','memo'].forEach(field=>{
            const fs=allSpans[field][gi];
            if(!fs.isFirst) return;
            const td=document.createElement('td');
            td.textContent=resolvedField(group, field);
            td.rowSpan=fs.span;
            tr.appendChild(td);
          });
        }
        tbody.appendChild(tr);
      });
    });
    table.appendChild(tbody);
    printArea.appendChild(table);
  });

  window.print();
}

// -- Finalize ("마감"): the actual end-of-day close-out - export the checked
// ROWS (not the whole date - 2026-08-27 user request) to Excel, and ONLY if
// that succeeds, also clear those specific rows out of workbench.db (same
// soft-delete every other delete goes through - 7 day recovery window,
// undoable). Doing this as one action means there's no way to end up
// half-done: delete-before-export-succeeded, or export-then-forget-to-delete
// leaving stale "done" rows cluttering the board.
//
// Granularity is per-row, not per-date: a date's banner checkbox checks
// every row under it by default (see renderAll's exportCb listener), but any
// row can be unchecked first to leave it out of THIS 마감 - e.g. one order
// that isn't actually resolved yet while the rest of that day's orders are.
// That row just stays on the board, unchecked, date unchanged. Whenever it's
// later resolved, checking just that row and hitting 마감 again exports it
// alone - export_dates_to_excel() targets that row's own `date` field, so it
// lands appended to the bottom of the SAME sheet the rest of that date
// already went to, not a new sheet. --
// -- "기존 배송장 엑셀로 동기화": workbench.db에 있는데 아직 기존 배송장
// 엑셀(EXCEL_PATH)에는 없는 오더를 그 오더의 날짜 섹션에 채워 넣는다
// (board_sync.py 참고). 마감과 달리 workbench.db는 전혀 건드리지 않고 -
// 그냥 기존 배송장 엑셀 쪽에 누락분을 채워 넣기만 하는 편도 동기화.
// 2026-08-28: workbench가 완전히 메인이 될 때까지, 기존 배송장 엑셀도
// 손으로 안 맞추기 위한 임시 버튼.
async function syncDeliveryExcel(){
  toast('배송장 엑셀에 동기화 중...');
  let result;
  try{
    const res=await fetch('/api/sync_delivery_excel', {method:'POST'});
    result=await res.json();
    if(!res.ok || result.error){throw new Error(result.error||'동기화 실패')}
  }catch(err){
    toast('동기화 실패: '+err.message);
    return;
  }
  toast(result.added>0 ? `배송장 ${result.sheet}시트에 ${result.added}건 반영 완료` : '이미 최신 상태 (추가할 오더 없음)');
}

async function finalizeChecked(){
  const checkedRowKeys=new Set([...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey));
  const checkedRows=ROWS.filter(r=>checkedRowKeys.has(r.rowKey));
  if(!checkedRows.length){toast('마감할 행을 먼저 체크하세요 (날짜 줄 체크박스로 한 번에 선택 후, 필요하면 개별 행 체크를 해제하세요)');return}
  const isos=[...new Set(checkedRows.map(r=>r.date))];
  const payload={};
  isos.forEach(iso=>{payload[iso]=buildExportRows(iso, checkedRowKeys)});
  toast('마감 처리 중 (엑셀로 내보내는 중)...');
  let exportData;
  try{
    const res=await fetch('/api/export', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({dates:payload})
    });
    exportData=await res.json();
    if(!res.ok || exportData.error){throw new Error(exportData.error||'내보내기 실패')}
  }catch(err){
    toast('마감 실패 (엑셀로 내보내지 못해 workbench.db는 그대로 둠): '+err.message);
    return;
  }
  const itemIds=checkedRows.map(r=>r.itemId);
  const summary=Object.entries(exportData.exported).map(([iso,info])=>`${labelFor(iso)}: ${info.sheet}시트 ${info.rows}행`).join(' / ');
  try{
    const res=await fetch('/api/items/delete', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'삭제 실패')}
  }catch(err){
    toast('엑셀 저장은 됐지만('+summary+') workbench.db 정리는 실패: '+err.message);
    return;
  }
  ROWS = ROWS.filter(r=>!checkedRowKeys.has(r.rowKey));
  isos.forEach(iso=>EXPORT_DATES.delete(iso));
  renderAll();
  toast(checkedRows.length+'건 마감 완료 - '+summary+' / workbench.db 정리됨 (7일간 복구 가능)');
}

// -- Delete: removes checked rows from the board AND from workbench.db
// (order_items rows, cascading to their parent order if that was its last
// item). A checked date-export checkbox also counts as "delete this whole
// date" (it already checks every row underneath it - see renderAll's
// exportCb listener - this dateKeys union is a belt-and-suspenders fallback
// in case some row somehow doesn't have a live checkbox in the DOM). --
async function deleteChecked(){
  const rowKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  const checkedDates=[...document.querySelectorAll('.date-export-check:checked')]
    .map(cb=>cb.closest('tbody.date-group').dataset.date);
  const dateKeys=ROWS.filter(r=>checkedDates.includes(r.date)).map(r=>r.rowKey);
  const keys=[...new Set([...rowKeys, ...dateKeys])];
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    const res=await fetch('/api/items/delete', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'삭제 실패')}
  }catch(err){
    toast('삭제 실패 (DB 반영 안 됨): '+err.message);
    return;
  }
  ROWS = ROWS.filter(r=>!keys.includes(r.rowKey));
  checkedDates.forEach(iso=>EXPORT_DATES.delete(iso));
  renderAll();
  toast(keys.length+'건 삭제 (DB에 반영됨)');
}

// -- Drag-and-drop between date sections (nearby dates) --
function onDragOver(e){e.preventDefault();e.currentTarget.classList.add('drag-over')}
document.addEventListener('dragleave',e=>{
  const grp=e.target.closest && e.target.closest('.date-group');
  if(grp && !grp.contains(e.relatedTarget)) grp.classList.remove('drag-over');
});
document.addEventListener('dragend',e=>{
  document.querySelectorAll('.date-group.drag-over').forEach(g=>g.classList.remove('drag-over'));
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
});
async function onDrop(e){
  // Fallback for dropping into empty space in a date section (not on a
  // specific row) - e.g. the section is empty, or the gap below its last
  // row. Row-level drops (onRowDrop) handle the precise-position case and
  // stopPropagation so this doesn't double-fire on top of them. Same
  // comma-joined multi-row payload onRowDrop reads (see grip's dragstart) -
  // moves every dragged row's whole group at once, keeping their relative
  // order, to the bottom of this date section.
  e.preventDefault();
  const tbody=e.currentTarget;
  tbody.classList.remove('drag-over');
  const sourceKeys=(e.dataTransfer.getData('text/plain')||'').split(',').filter(Boolean);
  if(!sourceKeys.length)return;
  const sourceSet=new Set(sourceKeys);
  const movedRows=ROWS.filter(r=>sourceSet.has(r.rowKey));
  if(!movedRows.length)return;
  const iso=tbody.dataset.date;
  if(movedRows.every(r=>r.date===iso))return;
  try{
    await persistItemDate(movedRows.map(r=>r.itemId), iso);
  }catch(err){toast('날짜 이동 저장 실패: '+err.message);return}
  movedRows.forEach(r=>{r.date=iso});
  renderAll();
  toast(movedRows.length+'건 다른 날짜로 이동 - DB에 반영됨');
}

// -- Same-date (and cross-date-to-a-specific-spot) reordering: drop directly
// on a row to insert the dragged group(s) immediately before/after the
// HOVERED ROW'S WHOLE GROUP - never mid-group, even when hovering the middle
// row of a 3-row order (2026-08-27 user report: the old row-by-row boundary
// could land inside a completely unrelated order's block, making the preview
// look like it was about to mix into it). Persists as board_pos for every
// item currently in the destination date section (see reorder_items()). --
function onRowDragOver(e){
  e.preventDefault();
  const tr=e.currentTarget;
  const tbody=tr.closest('tbody.date-group');
  const groupTrs=[...tbody.querySelectorAll('tr.item-row')].filter(el=>el.dataset.groupKey===tr.dataset.groupKey);
  const firstTr=groupTrs[0], lastTr=groupTrs[groupTrs.length-1];
  const groupTop=firstTr.getBoundingClientRect().top;
  const groupBottom=lastTr.getBoundingClientRect().bottom;
  const before=(e.clientY-groupTop) < (groupBottom-groupTop)/2;
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
  (before?firstTr:lastTr).classList.add(before?'drop-before':'drop-after');
}
async function onRowDrop(e){
  e.preventDefault();
  e.stopPropagation();
  const tr=e.currentTarget;
  const tbody=tr.closest('tbody.date-group');
  const groupTrs=[...tbody.querySelectorAll('tr.item-row')].filter(el=>el.dataset.groupKey===tr.dataset.groupKey);
  const firstTr=groupTrs[0], lastTr=groupTrs[groupTrs.length-1];
  const before=firstTr.classList.contains('drop-before');
  const after=lastTr.classList.contains('drop-after');
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
  if(!before && !after) return;

  const sourceKeys=(e.dataTransfer.getData('text/plain')||'').split(',').filter(Boolean);
  if(!sourceKeys.length) return;
  const sourceSet=new Set(sourceKeys);
  // Target is the whole hovered group, not just the one <tr> the browser
  // fired the drop on - anchor on its LAST row so a drop lands after every
  // row of that group when `after`, not spliced in after just its first row.
  const targetKey = before ? firstTr.dataset.rowKey : lastTr.dataset.rowKey;
  if(sourceSet.has(targetKey))return; // dropped back onto (part of) itself
  const movedRows=ROWS.filter(r=>sourceSet.has(r.rowKey)); // keeps original relative order
  if(!movedRows.length)return;
  ROWS=ROWS.filter(r=>!sourceSet.has(r.rowKey));

  const targetIso=tbody.dataset.date;
  const dateChanged=movedRows.some(r=>r.date!==targetIso);
  movedRows.forEach(r=>{r.date=targetIso});
  let targetIdx=ROWS.findIndex(r=>r.rowKey===targetKey);
  if(targetIdx===-1) targetIdx=ROWS.length;
  ROWS.splice(before?targetIdx:targetIdx+1, 0, ...movedRows);

  try{
    if(dateChanged) await persistItemDate(movedRows.map(r=>r.itemId), targetIso);
    await persistItemReorder(ROWS.filter(r=>r.date===targetIso).map(r=>r.itemId));
  }catch(err){toast('순서 저장 실패: '+err.message);return}
  renderAll();
  toast((dateChanged ? '날짜 이동 + 순서 변경' : '순서 변경')+' - '+movedRows.length+'건 - DB에 반영됨');
}

// -- Excel-panel date picker: check rows below, pick a date up top, they all
// jump there. Handles dates that don't have a section yet (e.g. board only
// shows 7/28 and 7/30 but the user wants 7/29) by creating a fresh date
// section in the right sorted position (renderAll() rebuilds every date in
// sorted order every time, so a brand-new iso just needs to exist on a
// row's `date` field to get its own section). --
async function moveCheckedToDate(iso){
  if(!iso)return;
  if(!/^\d{4}-\d{2}-\d{2}$/.test(iso) || isNaN(new Date(iso+'T00:00:00').getTime())){
    toast('날짜 형식이 올바르지 않습니다. 달력에서 다시 선택해주세요');
    return;
  }
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const toMove=ROWS.filter(r=>keys.includes(r.rowKey) && r.date!==iso);
  if(!toMove.length){toast('이미 해당 날짜에 있습니다');return}
  try{
    await persistItemDate(toMove.map(r=>r.itemId), iso);
  }catch(err){toast('날짜 이동 저장 실패: '+err.message);return}
  toMove.forEach(row=>row.date=iso);
  renderAll();
  toast(toMove.length+'건을 '+labelFor(iso)+'(으)로 이동 - DB에 반영됨');
}
</script>
</body>
</html>
'''


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

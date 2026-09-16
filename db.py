"""Shared foundation for workbench_app.py: paths, logging, and the SQLite
layer (schema, connection, event log, undo history).

Split out of workbench_app.py on 2026-09-16 as part of breaking that
7400-line single file into modules - see BLOOMBERG_HANDOFF.md. Pure code
motion: every name below is unchanged from its original definition.
"""

from __future__ import annotations

import json
import logging
import re
import sqlite3
import sys
import threading
import traceback
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path

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


@contextmanager
def connect():
    """모든 호출부가 `with connect() as con:` 형태로만 쓰고 있음(전수 확인) -
    sqlite3.Connection을 그냥 context manager로 쓰면 __exit__가 commit/
    rollback만 하고 실제로 connection을 close()하지는 않아서, 24시간 도는
    이 프로세스(20분 루프 + HTTP 요청마다)에서 매 호출이 새 커넥션/파일
    핸들을 계속 새로 열기만 하고 하나도 안 닫는 채 쌓이고 있었다.
    @contextmanager로 감싸 with 블록을 빠져나갈 때(정상/예외 어느 쪽이든)
    항상 close()까지 되도록 한다 - 호출부 코드는 전혀 안 바뀜."""
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    try:
        with con:
            yield con
    finally:
        con.close()


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



"""Create a monthly Vendor Activities Dashboard copy/paste sheet.

Reads the "마감" export workbook (FINALIZE_EXPORT_PATH - the day is only
final, with its colors settled, once it's gone through workbench's 마감
button) and writes a month sheet such as VAD_05. The generated sheet mirrors
the Bloomberg dashboard input area so the filled rows can be copied into the
online workbook. Not the live daily-collection file (EXCEL_PATH) - that one
keeps changing after the fact (backfills, corrections), so counting from it
would drift from what 마감 actually froze for a given day.
"""

import argparse
import calendar
import logging
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import xlwings as xw

from config import FINALIZE_EXPORT_PATH
from excel_handler import _com_retry
from order_lock import ExcelWriteLock


BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"
SOURCE_EXCEL_PATH = FINALIZE_EXPORT_PATH
DASHBOARD_EXCEL_PATH = str(
    Path(r"C:\1")
    / "\ubc30\uc1a1\uc7a5"
    / "Vendor Activities Dashboard"
    / "2026 Vendor Activities Dashboard.xlsx"
)

ORDER_RE = re.compile(r"\b(\d{6,})\b")
QTY_RE = re.compile(r"\bQty\s*[:#-]?\s*(\d+)\b|\bx\s*(\d+)\b", re.IGNORECASE)

ITEM_ROWS = [
    ("keyboard", "Keyboard"),
    ("bunit", "B-Unit"),
    ("pc", "Int + Ext PC"),
    ("fp", "FP"),
    ("router", "Router/Server"),
    ("others", "Others"),
    ("sto", "STO"),
]

logger = logging.getLogger(__name__)


def _digits(value):
    text = str(value or "")
    if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def _norm_text(value):
    return str(value or "").strip()


def _extract_order(text):
    match = ORDER_RE.search(str(text or ""))
    return match.group(1).lstrip("0") if match else ""


def _line_kind(order_text):
    text = str(order_text or "").upper()
    if "회수" in text:
        return "pickup"
    if "배송" in text:
        return "delivery"
    return ""


def _extract_qty(description):
    match = QTY_RE.search(str(description or ""))
    if not match:
        return 1
    try:
        return max(1, int(match.group(1) or match.group(2)))
    except ValueError:
        return 1


def _classify_item(description, material=""):
    text = str(description or "").upper()
    mat = _digits(material)
    if "LUGGAGE TAG" in text:
        return "ignore"
    if "STO" in text:
        return "sto"
    if "KEYBOARD" in text or "KB" in text:
        return "keyboard"
    if "BUNIT" in text or "B-UNIT" in text or "AUTHENTICATION" in text or mat in {"10026155", "10051377"}:
        return "bunit"
    if "MONITOR" in text or "모니터" in str(description or "") or "STAND" in text or "FP" in text:
        return "fp"
    if mat in {"10057033"} or any(word in text for word in ("ROUTER", "SERVER", "CISCO", "SWITCH", "POWEREDGE")):
        return "router"
    if text.startswith("PC ") or any(word in text for word in ("LAPTOP", "DESKTOP", " PC", "PERSONAL COMPUTER")):
        return "pc"
    return "others"


def _open_book(path):
    filename = Path(path).name
    try:
        for book in xw.books:
            if book.name == filename:
                return book, False
    except Exception:
        pass
    return xw.Book(path), True


def _sheet_values(sheet):
    used = sheet.used_range
    values = used.value
    if not values:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def _row_value(row, idx):
    try:
        return row[idx]
    except Exception:
        return None


def _is_close_color(color, target, tolerance=12):
    if not color:
        return False
    try:
        return all(abs(int(color[idx]) - target[idx]) <= tolerance for idx in range(3))
    except Exception:
        return False


def _is_completed_row(colors):
    return any(_is_close_color(color, (0, 176, 240)) for color in colors[:8])


def _row_values(sheet, row_idx):
    values = sheet.range((row_idx, 1), (row_idx, 8)).value
    if not isinstance(values, list):
        values = [values]
    return values


def _row_colors(sheet, row_idx):
    return [sheet.cells(row_idx, col).color for col in range(1, 9)]


def _collect_day(sheet):
    stats = {
        "delivery_orders": set(),
        "delivery_quick_orders": set(),
        "pickup_orders": set(),
        "delivery_items": defaultdict(int),
        "pickup_items": defaultdict(int),
    }
    current_kind = ""
    current_order = ""
    current_quick = False
    max_row = sheet.used_range.last_cell.row
    for row_idx in range(1, max_row + 1):
        row = _row_values(sheet, row_idx)
        joined_row = " ".join(_norm_text(value) for value in row if value is not None).lower()
        if "scheduled" in joined_row or "delayed" in joined_row:
            break

        if not _is_completed_row(_row_colors(sheet, row_idx)):
            continue

        order_text = _norm_text(_row_value(row, 0))
        kind = _line_kind(order_text)
        memo = _norm_text(_row_value(row, 7))
        if kind:
            order_num = _extract_order(order_text)
            if not order_num:
                order_num = f"{kind}:{row_idx}:{order_text}"
            current_kind = kind
            current_order = order_num
            current_quick = "퀵" in memo
        else:
            kind = current_kind
            order_num = current_order
            if not kind or not order_num:
                continue

        description = _norm_text(_row_value(row, 1))
        material = _row_value(row, 2)
        if not description and not _digits(material):
            continue

        qty = _extract_qty(description)
        category = _classify_item(description, material)
        if category == "ignore":
            continue

        if kind == "delivery":
            stats["delivery_orders"].add(order_num)
            if current_quick or "퀵" in memo:
                stats["delivery_quick_orders"].add(order_num)
            stats["delivery_items"][category] += qty
        elif kind == "pickup":
            stats["pickup_orders"].add(order_num)
            stats["pickup_items"][category] += qty
    return stats


def _find_day_sheet(sheet_names, month, day):
    """Look up the sheet for one day, tolerant of a trailing suffix such as
    the "8-4+" naming seen in real sheets (someone appends "+" or similar
    when a day's sheet gets a late addition) - an exact "{month}-{day}"
    match wins, otherwise fall back to any sheet name that starts with that
    key followed by a non-digit (so "8-4" matches "8-4+" but not "8-41")."""
    exact_keys = (f"{month}-{day}", f"{month}.{day}")
    for key in exact_keys:
        sheet = sheet_names.get(key)
        if sheet is not None:
            return sheet
    for key in exact_keys:
        for name, sheet in sheet_names.items():
            if name.startswith(key) and (len(name) == len(key) or not name[len(key)].isdigit()):
                return sheet
    return None


def collect_month(month, year=None):
    year = year or datetime.today().year
    book, should_close = _open_book(SOURCE_EXCEL_PATH)
    try:
        result = {}
        sheet_names = {sheet.name: sheet for sheet in book.sheets}
        _, last_day = calendar.monthrange(year, month)
        for day in range(1, last_day + 1):
            sheet = _find_day_sheet(sheet_names, month, day)
            result[day] = _collect_day(sheet) if sheet else {
                "delivery_orders": set(),
                "delivery_quick_orders": set(),
                "pickup_orders": set(),
                "delivery_items": defaultdict(int),
                "pickup_items": defaultdict(int),
            }
        return result
    except Exception:
        raise
    finally:
        if should_close:
            try:
                book.close()
            except Exception:
                logger.warning("소스 워크북 close 실패 (무시 - 읽기만 했으므로 데이터 손실 없음)", exc_info=True)


def _set_cell(sheet, row, col, value=None, color=None, bold=False, font_color=None):
    cell = sheet.cells(row, col)
    if value is not None:
        cell.value = value
    if color is not None:
        cell.color = color
    if font_color is not None:
        cell.api.Font.Color = font_color
    cell.api.Font.Bold = bool(bold)
    return cell


def _clear_or_create_sheet(book, name):
    names = [sheet.name for sheet in book.sheets]
    if name in names:
        sheet = book.sheets[name]
        sheet.clear()
        return sheet
    return book.sheets.add(name, after=book.sheets[-1])


def _write_dashboard(book, month, stats_by_day, year=None):
    return _write_dashboard_full_template(book, month, stats_by_day, year)


def _write_dashboard_full_template(book, month, stats_by_day, year=None):
    year = year or datetime.today().year
    _, last_day = calendar.monthrange(year, month)
    sheet_name = f"VAD_{month:02d}"
    sheet = _clear_or_create_sheet(book, sheet_name)

    max_row = 44
    start_col = 3
    last_col = start_col + last_day - 1

    labels = {
        1: (None, "Date"),
        2: ("OUTBOUND", None),
        3: ("No of Deliveries", None),
        4: (None, "Same day order"),
        5: (None, "Express Delivery"),
        6: (None, None),
        7: (None, "No of Failure"),
        8: (None, "Customer request"),
        9: (None, "Capacity issue"),
        10: (None, None),
        11: ("Delivered Quantity", None),
        12: ("Delivered CPE Quantity", None),
        13: (None, "Keyboard"),
        14: (None, "B-Unit"),
        15: (None, "Int + Ext PC"),
        16: (None, "FP"),
        17: (None, "Router/Server"),
        18: (None, "Others"),
        19: (None, "STO"),
        20: (None, None),
        21: ("PRODUCTION", None),
        22: ("No of Production", None),
        23: (None, "Keyboard"),
        24: (None, "B-Unit"),
        25: (None, "PC"),
        26: (None, "FP"),
        27: (None, "Router/Server"),
        28: (None, None),
        29: ("Backlog Quantity", None),
        30: (None, "Keyboard"),
        31: (None, "B-Unit"),
        32: (None, "PC"),
        33: (None, "FP"),
        34: (None, "Router/Server"),
        35: (None, None),
        36: ("RMA", None),
        37: ("No of Pick Up", None),
        38: ("Pick Up Quantity", None),
        39: (None, "Keyboard"),
        40: (None, "B-Unit"),
        41: (None, "PC"),
        42: (None, "FP"),
        43: (None, "Router/Server"),
        44: (None, "No of Failure"),
    }

    for row, (col_a, col_b) in labels.items():
        sheet.cells(row, 1).value = col_a
        sheet.cells(row, 2).value = col_b

    for day in range(1, last_day + 1):
        col = start_col + day - 1
        sheet.cells(1, col).value = datetime(year, month, day)
        sheet.cells(1, col).number_format = "d-mmm"
        if datetime(year, month, day).weekday() >= 5:
            sheet.range((1, col), (max_row, col)).color = (217, 217, 217)

    color_rows = {
        3: (255, 192, 0),
        11: (0, 176, 80),
        12: (146, 208, 80),
        22: (0, 0, 0),
        29: (255, 230, 153),
        37: (226, 239, 218),
        38: (192, 0, 0),
    }
    for row, color in color_rows.items():
        sheet.range((row, 1), (row, last_col)).color = color
        sheet.range((row, 1), (row, 2)).api.Font.Bold = True
    for row in (22, 38):
        sheet.range((row, 1), (row, last_col)).api.Font.Color = 16777215
    for row in (2, 21, 36):
        sheet.range((row, 1), (row, 2)).api.Font.Bold = True

    for day in range(1, last_day + 1):
        col = start_col + day - 1
        stats = stats_by_day[day]
        delivery_count = len(stats["delivery_orders"])
        express_count = len(stats["delivery_quick_orders"])
        same_day_count = max(0, delivery_count - express_count)
        pickup_count = len(stats["pickup_orders"])

        sheet.cells(3, col).value = delivery_count if delivery_count else "-"
        sheet.cells(4, col).value = same_day_count if same_day_count else ""
        sheet.cells(5, col).value = express_count if express_count else ""

        cpe_sum = sum(stats["delivery_items"][key] for key, _ in ITEM_ROWS[:5])
        delivered_sum = cpe_sum + stats["delivery_items"]["others"] + stats["delivery_items"]["sto"]
        sheet.cells(11, col).value = delivered_sum if delivered_sum else "-"
        sheet.cells(12, col).value = cpe_sum if cpe_sum else "-"
        for offset, (key, _) in enumerate(ITEM_ROWS):
            value = stats["delivery_items"][key]
            sheet.cells(13 + offset, col).value = value if value else ""

        pickup_sum = sum(stats["pickup_items"][key] for key, _ in ITEM_ROWS[:5])
        sheet.cells(37, col).value = pickup_count if pickup_count else "-"
        sheet.cells(38, col).value = pickup_sum if pickup_sum else "-"
        for offset, (key, _) in enumerate(ITEM_ROWS[:5]):
            value = stats["pickup_items"][key]
            sheet.cells(39 + offset, col).value = value if value else ""

    used = sheet.range((1, 1), (max_row, last_col))
    used.api.Font.Name = "Calibri"
    used.api.Font.Size = 10
    used.api.HorizontalAlignment = -4108
    used.api.VerticalAlignment = -4108
    used.api.Borders.LineStyle = 1
    sheet.range((1, 1), (max_row, 2)).api.HorizontalAlignment = -4131
    sheet.range("A:A").column_width = 18
    sheet.range("B:B").column_width = 18
    for col in range(start_col, last_col + 1):
        sheet.range((1, col), (1, col)).column_width = 8.5
    sheet.range((1, start_col), (1, last_col)).api.Orientation = 0
    sheet.range((1, start_col), (max_row, last_col)).api.WrapText = False
    try:
        sheet.book.app.api.ActiveWindow.SplitColumn = 1
        sheet.book.app.api.ActiveWindow.SplitRow = 1
        sheet.book.app.api.ActiveWindow.FreezePanes = True
    except Exception:
        pass
    return sheet


def _generate_once(month, year):
    stats = collect_month(month, year)
    book, should_close = _open_book(DASHBOARD_EXCEL_PATH)
    try:
        sheet = _write_dashboard(book, month, stats, year)
        book.save()
        logger.info("Vendor Dashboard 작성 완료: %s / sheet=%s", DASHBOARD_EXCEL_PATH, sheet.name)
        print(f"[Dashboard] {sheet.name} 작성 완료: {DASHBOARD_EXCEL_PATH}")
        return sheet.name
    finally:
        # book.save() above already landed on disk - a close() hiccup here
        # (COM is finicky right after a save) must not turn a successful
        # generate() into a reported failure.
        if should_close:
            try:
                book.close()
            except Exception:
                logger.warning("Dashboard 워크북 close 실패 (무시 - 이미 저장됨)", exc_info=True)


def generate(month=None, year=None):
    today = datetime.today()
    month = month or today.month
    year = year or today.year
    # Both collect_month (reads FINALIZE_EXPORT_PATH) and the dashboard write
    # below go through Excel COM, same as main.py's live order writes - and
    # that COM session doesn't tolerate concurrent access from another
    # process (automation.log has a history of OLE error 0x800ac472 /
    # RPC_E_CALL_REJECTED from exactly this, 2026-08-24 writeup). Serialize
    # with the same cross-process lock everyone else uses, and retry the
    # whole attempt a couple times in case Excel is just momentarily busy.
    with ExcelWriteLock():
        return _com_retry(lambda: _generate_once(month, year))


def refresh_for_dates(iso_dates):
    """Regenerate the VAD_MM sheet for every month touched by iso_dates.

    Meant to be called right after workbench's 마감 (export) finishes
    writing those dates into FINALIZE_EXPORT_PATH - one call covers however
    many dates/months were just closed out. Never raises: a dashboard
    refresh failing must not be mistaken for the 마감 export itself having
    failed, so every error is logged and swallowed, and callers get back
    which months made it (sheet name) vs. which didn't (None).
    """
    months = set()
    for iso in iso_dates:
        try:
            d = datetime.fromisoformat(str(iso)).date()
        except ValueError:
            logger.warning("Dashboard 갱신: 날짜 형식 인식 못함 (건너뜀): %s", iso)
            continue
        months.add((d.year, d.month))

    results = {}
    for year, month in sorted(months):
        key = f"{year}-{month:02d}"
        try:
            results[key] = generate(month, year)
        except Exception:
            logger.error("Dashboard 자동 갱신 실패 (%s)", key, exc_info=True)
            results[key] = None
    return results


def main():
    parser = argparse.ArgumentParser(description="Create monthly Vendor Activities Dashboard copy/paste sheet.")
    parser.add_argument("--month", type=int, default=datetime.today().month)
    parser.add_argument("--year", type=int, default=datetime.today().year)
    parser.add_argument(
        "--dates",
        default="",
        help="Comma-separated ISO dates (e.g. 2026-08-25,2026-08-24) - if given, "
        "overrides --month/--year and refreshes every month those dates touch. "
        "This is what workbench's 마감 button passes.",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )

    iso_dates = [part.strip() for part in args.dates.split(",") if part.strip()]

    try:
        if iso_dates:
            results = refresh_for_dates(iso_dates)
            ok = [k for k, v in results.items() if v]
            failed = [k for k, v in results.items() if not v]
            _notify_dashboard_result(ok, failed)
            return 1 if failed and not ok else 0
        generate(args.month, args.year)
    except Exception as exc:
        logger.error("Vendor Dashboard 작성 실패: %s", exc, exc_info=True)
        _notify_dashboard_result([], [f"{args.year}-{args.month:02d}"])
        return 1
    return 0


def _notify_dashboard_result(ok_months, failed_months):
    """Best-effort Windows toast so a 마감-triggered background refresh is
    still visible even with no console window (mirrors the win_notify usage
    elsewhere in this project). Never lets a notification failure surface."""
    try:
        from win_notify import send_windows_notification
    except Exception:
        return
    if ok_months and not failed_months:
        title = "Vendor Dashboard 갱신 완료"
        msg = ", ".join(ok_months) + " 로컬 대시보드 갱신됨 - 온라인 대시보드에 붙여넣기 필요"
    elif ok_months and failed_months:
        title = "Vendor Dashboard 일부만 갱신됨"
        msg = f"성공: {', '.join(ok_months)} / 실패: {', '.join(failed_months)} (automation.log 확인)"
    else:
        title = "Vendor Dashboard 갱신 실패"
        msg = f"{', '.join(failed_months)} 갱신 실패 - automation.log 확인 필요"
    send_windows_notification(title, msg)


if __name__ == "__main__":
    sys.exit(main())

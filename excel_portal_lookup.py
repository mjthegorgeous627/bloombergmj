"""Portal 자동화(launcher.py, portal_ship_and_print.py, portal_update_pod.py)가
오더번호로 항목/고객/OBD 기본값을 찾는 조회 모듈.

2026-10-07: 예전엔 "실시간 동기화" 엑셀(REALTIME_SYNC_PATH)의 오늘 시트를 읽었지만
그 파일을 폐지하고 workbench.db를 직접 읽는다 (원본이 DB 하나가 되도록). 함수 이름
(find_excel_order)과 반환 구조(rows[].order_text/obd/description/qty/material/serial/
customer/...)는 호출부를 안 건드리려고 그대로 뒀다. workbench 화면에서 시작하는
Portal/POD는 원래 항목을 직접 넘겨서(--items-json / items_override) 이 조회를 안 타고,
launcher.py와 CLI 대체 경로만 여기로 온다.
"""

import re
from collections import Counter

from db import connect
from order_domain import resolve_item_type

ORDER_RE = re.compile(r"\b(\d{7,})\b")
OBD_RE = re.compile(r"\bOBD\s*[:#-]?\s*(\d{7,})\b", re.IGNORECASE)
QTY_RE = re.compile(r"\bQty\s*[:#-]?\s*(\d+)\b|\bx\s*(\d+)\b", re.IGNORECASE)


def _norm_digits(value):
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _item_label(description):
    text = str(description or "").upper()
    if "MONITOR" in text:
        return "Monitor"
    if "AUTHENTICATION" in text or "BUNIT" in text:
        return "BUNIT"
    if "LAPTOP" in text or "PC" in text:
        return "PC"
    if "KEYBOARD" in text or "KB" in text:
        return "KB"
    return "Item"


def _extract_qty(description):
    match = QTY_RE.search(str(description or ""))
    if not match:
        return 1
    try:
        value = match.group(1) or match.group(2)
        return max(1, int(value))
    except ValueError:
        return 1


def _format_item_counts(pairs):
    counts = Counter()
    for label, qty in pairs:
        try:
            qty_int = max(1, int(qty))
        except (ValueError, TypeError):
            qty_int = 1
        counts[label] += qty_int
    return ", ".join(f"{count}x{label}" for label, count in counts.items())


def _line_kind(order_text):
    text = str(order_text or "").upper()
    if "회수" in text:
        return "pickup"
    return "delivery"


def _order_types(rows):
    joined = "\n".join(str(row.get("order_text") or "") for row in rows).upper()
    return sorted(set(re.findall(r"\b(ZOR|SDSK|ZRX|ZRE|ZINP|ZINT|ORD)\b", joined)))


def _extract_obd(order_text):
    match = OBD_RE.search(str(order_text or ""))
    return match.group(1) if match else ""


def _filter_rows_by_delivery(rows, delivery):
    delivery_digits = _norm_digits(delivery)
    if not delivery_digits:
        return rows

    selected = []
    current_obd = ""
    for row in rows:
        row_obd = _norm_digits(row.get("obd"))
        if row_obd:
            current_obd = row_obd
        if current_obd == delivery_digits:
            selected.append(row)
    return selected


def build_default_remarks(rows, signed_by, pickup_mode=None):
    deliveries = []
    pickups = []
    pickup_expected = False
    for row in rows:
        label = _item_label(row.get("description"))
        qty = row.get("qty") or _extract_qty(row.get("description"))
        if _line_kind(row.get("order_text")) == "pickup":
            pickup_expected = True
            if pickup_mode == "collected":
                pickups.append((label, qty))
        else:
            deliveries.append((label, qty))

    delivered = _format_item_counts(deliveries) or "1xItem"
    collected = _format_item_counts(pickups)
    signed_by = signed_by or ""
    order_types = _order_types(rows)

    if "ZRX" in order_types or pickup_expected:
        if collected:
            return f"Delivered {delivered}, and collected {collected}, to/from {signed_by}."
        return f"Delivered {delivered} to {signed_by}, and will collect later."
    return f"Delivered {delivered} to {signed_by}."


def _build_result(order_num, rows, pickup_mode=None):
    customer = next((r["customer"] for r in rows if r["customer"]), "")
    obd = next((r["obd"] for r in rows if r["obd"]), "")
    return {
        "sheet": "(workbench)",
        "start_row": rows[0]["row"],
        "end_row": rows[-1]["row"],
        "order": str(order_num),
        "obd": obd,
        "customer": customer,
        "remarks": build_default_remarks(rows, customer, pickup_mode=pickup_mode),
        "rows": rows,
        "source": "workbench.db",
    }


def _rows_from_db(order_num):
    """workbench.db에서 이 오더번호의 항목들을 옛 엑셀 행 구조로 돌려준다.
    같은 오더번호가 여러 orders 행(배송/회수가 따로 수집된 경우 등)으로 있으면 전부
    이어 붙인다. 살아있는 항목이 하나도 없으면(이미 마감돼 소프트삭제된 뒤) 삭제된
    항목이라도 보여줘서 마감 직후 POD 처리도 막히지 않게 한다."""
    target = _norm_digits(order_num).lstrip("0")
    if not target:
        return []
    rows = []
    with connect() as con:
        orders = [o for o in con.execute("SELECT * FROM orders ORDER BY id").fetchall()
                  if _norm_digits(o["order_no"]).lstrip("0") == target]
        for include_deleted in (False, True):
            for order in orders:
                od = dict(order)
                sql = "SELECT * FROM order_items WHERE order_id=?"
                if not include_deleted:
                    sql += " AND deleted_at IS NULL"
                items = [dict(i) for i in con.execute(sql + " ORDER BY line_no", (order["id"],)).fetchall()]
                for idx, item in enumerate(items):
                    kind = resolve_item_type(od, item, idx, len(items)) or "배송"
                    delivery_no = str(od.get("delivery_no") or "").strip()
                    code_lines = ([f"OBD {delivery_no}"] if delivery_no else []) + [
                        ln for ln in str(od.get("extra_codes") or "").splitlines() if ln.strip()
                    ]
                    order_text = "\n".join([f"{kind} {od.get('order_type') or ''} {od.get('order_no') or ''}".strip()] + code_lines)
                    description = str(item.get("description") or "")
                    try:
                        qty = max(1, int(float(item.get("qty"))))
                    except (TypeError, ValueError):
                        qty = _extract_qty(description)
                    rows.append({
                        "row": len(rows) + 1,
                        "order_text": order_text,
                        "obd": delivery_no,
                        "description": description,
                        "qty": qty,
                        "material": str(item.get("material") or "").strip(),
                        "serial": str(item.get("serial") or "").strip(),
                        "customer": str(od.get("customer") or "").strip(),
                        "phone": str(od.get("phone") or "").strip(),
                        "address": str(od.get("address") or "").strip(),
                        "memo": str(od.get("memo") or "").strip(),
                    })
            if rows:
                break
    return rows


def find_excel_order(order_num, pickup_mode=None, delivery=None, use_live=False):
    """오더번호로 Portal 기본값을 찾는다. (use_live는 호출부 호환용 - 이제 항상 DB)"""
    rows = _filter_rows_by_delivery(_rows_from_db(order_num), delivery)
    if not rows:
        raise LookupError(f"workbench에서 오더 {order_num}을 찾지 못했습니다.")
    return _build_result(order_num, rows, pickup_mode=pickup_mode)


if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser()
    parser.add_argument("order")
    parser.add_argument("--delivery")
    args = parser.parse_args()
    print(json.dumps(find_excel_order(args.order, delivery=args.delivery), ensure_ascii=False, indent=2))

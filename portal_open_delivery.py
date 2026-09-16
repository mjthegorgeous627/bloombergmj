"""Open Bloomberg Portal delivery page for a SAP Pur.Doc order.

Read-only helper for the first portal automation step:
1. Read VL06O session 0.
2. Find Pur.Doc / EBELN matching the order number.
3. Open the matching Bloomberg delivery detail page by Delivery# / VBELN.
"""

import argparse
import logging
import subprocess
from pathlib import Path

from sap_handler import get_all_rows_from_list, get_sap_session, group_by_order


PORTAL_DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
CHROME_PATHS = [
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


def find_chrome() -> Path:
    for path in CHROME_PATHS:
        if path.exists():
            return path
    raise FileNotFoundError("Chrome 실행 파일을 찾지 못했습니다.")


def normalize_order(order: str) -> str:
    value = "".join(ch for ch in str(order) if ch.isdigit())
    if not value:
        raise ValueError("오더번호에 숫자가 없습니다.")
    return value.lstrip("0") or value


def find_delivery_for_order(order: str):
    target = normalize_order(order)
    session = get_sap_session(0)
    rows = get_all_rows_from_list(session)
    orders = group_by_order(rows)

    if target in orders:
        return target, orders[target]

    for ebeln, info in orders.items():
        if normalize_order(ebeln) == target:
            return ebeln, info

    available = ", ".join(sorted(orders.keys())[:20])
    raise LookupError(f"VL06O에서 Pur.Doc {target}을 찾지 못했습니다. 현재 일부 목록: {available}")


def open_delivery_page(delivery_num: str):
    url = PORTAL_DELIVERY_BASE + delivery_num
    subprocess.Popen([str(find_chrome()), url], close_fds=True)
    return url


def main():
    parser = argparse.ArgumentParser(description="Open Bloomberg Portal delivery page for SAP order.")
    parser.add_argument("order", help="SAP Pur.Doc order number, e.g. 67030018")
    parser.add_argument("--no-open", action="store_true", help="Only print matched Delivery#; do not open Chrome.")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    ebeln, info = find_delivery_for_order(args.order)
    delivery_num = info["vbeln"]

    print(f"Pur.Doc: {ebeln}")
    print(f"Delivery#: {delivery_num}")
    print(f"Ship To: {info.get('name_we', '')}")
    print("Items:")
    for item in info.get("items", []):
        print(f"  - {item.get('matnr', '')} / {item.get('arktx', '')} / qty {item.get('lfimg', '')}")

    if not args.no_open:
        url = open_delivery_page(delivery_num)
        print(f"Opened: {url}")


if __name__ == "__main__":
    main()

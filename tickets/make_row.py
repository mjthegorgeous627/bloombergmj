"""티켓 → 배송장(workbench 보드) 행 생성 도구 (TICKET_MANUAL 5-7).

사용 예 (기본은 미리보기만, --yes 를 줘야 실제로 쓴다):
  python make_row.py --label "SDSK 1331133687" --item "MACBOOK PRO 14\\" ..." --sn C9W7VQY93Y \\
      --customer "BOS Team" --address "Bloomberg Seoul Office (6523)" \\
      --memo "6507 → 6523/0016 sloc to sloc" [--mn 10062370] [--phone ...] [--date 2026-10-12] [--type 배송] [--yes]
- --mn을 안 주면 S/N으로 이력(history.db)에서 M/N을 찾아 제안한다.
- workbench 서버(127.0.0.1:8765)가 떠 있어야 한다. order_no는 빈값으로 만들어져 자동화가 건드리지 않는다.
"""
import argparse, json, os, sqlite3, sys, urllib.request
from datetime import date

HISTORY_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "history.db")
API = "http://127.0.0.1:8765/api/orders/add_manual"


def lookup_mn(serial):
    """S/N이 들어간 과거 배송장 행에서 M/N 후보를 찾는다 (cells_json: 0=오더,1=Item,2=M/N,3=S/N)."""
    if not serial or not os.path.exists(HISTORY_DB):
        return []
    con = sqlite3.connect(HISTORY_DB)
    found = []
    for (cj,) in con.execute("SELECT cells_json FROM rows WHERE cells_json LIKE ?", (f"%{serial}%",)):
        c = [str(x or "").strip() for x in json.loads(cj)] + [""] * 8
        if serial in c[3] and c[2] and c[2] not in found:
            found.append(c[2])
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--label", required=True, help="Order# 칸에 들어갈 값, 예: 'SDSK 1331133687'")
    ap.add_argument("--item", required=True)
    ap.add_argument("--mn", default="")
    ap.add_argument("--sn", default="")
    ap.add_argument("--customer", default="")
    ap.add_argument("--phone", default="")
    ap.add_argument("--address", default="")
    ap.add_argument("--memo", default="")
    ap.add_argument("--date", default=date.today().isoformat())
    ap.add_argument("--type", default="배송")
    ap.add_argument("--yes", action="store_true", help="미리보기 후 실제로 보드에 쓴다")
    a = ap.parse_args()

    mn = a.mn
    if not mn and a.sn:
        cands = lookup_mn(a.sn)
        if cands:
            print(f"[이력에서 찾은 M/N 후보] {cands}")
            mn = cands[0]
    cells = [a.item, mn, a.sn, a.customer, a.phone, a.address, a.memo]
    print(f"Type={a.type} | Order#={a.label} | date={a.date}")
    for name, v in zip(["Item", "M/N", "S/N", "Customer", "Phone", "Address", "Memo"], cells):
        print(f"  {name}: {v!r}")
    if not a.yes:
        print("(미리보기 — 쓰지 않음. --yes 를 주면 보드에 작성)")
        return
    body = {"cells": cells, "date": a.date, "itemType": a.type, "orderLabel": a.label, "cellCount": 7}
    req = urllib.request.Request(API, json.dumps(body).encode("utf-8"), {"Content-Type": "application/json"})
    print(urllib.request.urlopen(req).read().decode())


if __name__ == "__main__":
    main()

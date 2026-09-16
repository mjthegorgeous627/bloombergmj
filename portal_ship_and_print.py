"""Run the normal Bloomberg Portal shipping flow in one command.

Steps:
1. Register serial and click Run ShipERP.
2. Pack and Post.
3. Download and print QR/ZPL labels unless --skip-qr is used.
"""

import argparse
import json
import logging
import subprocess
import sys
import time
from pathlib import Path

from excel_portal_lookup import _line_kind, find_excel_order
from portal_open_delivery import find_delivery_for_order

BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"
logger = logging.getLogger(__name__)


def _digits(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def _serial_text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
        text = text[:-2]
    return text


def _resolve_from_sap(order):
    ebeln, info = find_delivery_for_order(order)
    delivery = info.get("vbeln")
    items = info.get("items", [])
    numeric_items = [item for item in items if _digits(item.get("matnr"))]
    material = numeric_items[0].get("matnr") if numeric_items else ""
    qty = numeric_items[0].get("lfimg") if numeric_items else "1"
    return {
        "order": ebeln,
        "delivery": delivery,
        "material": _digits(material),
        "qty": str(qty or "1").strip(),
    }


def _is_no_serial_item(description):
    text = str(description or "").upper()
    return "BUNIT" in text or "AUTHENTICATION" in text


def _row_qty(row):
    value = row.get("qty") or "1"
    try:
        return str(max(1, int(float(value))))
    except (ValueError, TypeError):
        return "1"


def _resolve_from_excel(order, delivery=None, force_no_serial=False):
    logger.info("Excel에서 오더 조회 시작: order=%s", order)
    data = find_excel_order(order, delivery=delivery)
    logger.info(
        "Excel 오더 조회 완료: sheet=%s rows=%s-%s obd=%s source=%s",
        data.get("sheet"),
        data.get("start_row"),
        data.get("end_row"),
        data.get("obd", ""),
        data.get("source", "file"),
    )
    delivery_obds = []
    for row in data.get("rows", []):
        if _line_kind(row.get("order_text")) == "delivery":
            row_obd = _digits(row.get("obd"))
            if row_obd:
                delivery_obds.append(row_obd)
    unique_obds = list(dict.fromkeys(delivery_obds))
    if not delivery and len(unique_obds) > 1:
        raise RuntimeError(
            "이 오더는 OBD가 여러 개입니다. Delivery# / OBD를 지정해서 하나씩 처리하세요: "
            + ", ".join(unique_obds)
        )
    result = {"delivery": data.get("obd", "")}
    items = []
    for row in data.get("rows", []):
        material = _digits(row.get("material"))
        serial = _serial_text(row.get("serial"))
        qty = _row_qty(row)
        no_serial = bool(force_no_serial) or _is_no_serial_item(row.get("description"))
        if not material:
            continue
        if _line_kind(row.get("order_text")) != "delivery":
            continue
        items.append({
            "material": material,
            "qty": qty,
            "serials": [serial] if serial else [],
            "no_serial": no_serial,
            "description": str(row.get("description") or ""),
        })
    if items:
        total_qty = sum(int(item["qty"]) for item in items)
        result["items"] = items
        result.update({
            "material": items[0]["material"],
            "serial": items[0]["serials"][0] if items[0]["serials"] else "",
            "qty": str(total_qty or 1),
            "no_serial": all(item["no_serial"] for item in items),
            "description": items[0].get("description", ""),
        })
    return result


def resolve_inputs(args):
    resolved = {
        "order": args.order,
        "delivery": args.delivery,
        "material": args.material,
        "qty": args.qty,
        "serial": args.serial,
        "no_serial": bool(getattr(args, "material_only", False) or getattr(args, "no_serial", False)),
        "items": [],
    }
    logger.info(
        "입력값 확인: order=%s delivery=%s material=%s qty=%s serial=%s",
        resolved["order"],
        resolved["delivery"] or "",
        resolved["material"] or "",
        resolved["qty"] or "",
        resolved["serial"] or "",
    )

    items_json = getattr(args, "items_json", "") or ""
    if items_json:
        # Workbench already resolved the real per-item breakdown from
        # workbench.db (multiple distinct materials, e.g. monitor + stand +
        # PC + keyboard in one order) - use it as-is and skip Excel lookup
        # entirely. Previously this argument was accepted but never actually
        # read here, so a multi-material order run through the workbench
        # would have silently fallen through to the single-item logic below
        # and been mis-registered as "N units of just the first material."
        try:
            resolved["items"] = json.loads(items_json)
        except (ValueError, TypeError) as exc:
            raise RuntimeError(f"--items-json 파싱 실패: {exc}") from exc
        if resolved["no_serial"]:
            for item in resolved["items"]:
                item["serials"] = []
                item["no_serial"] = True
        logger.info("Workbench items-json 사용 (엑셀 조회 생략): %d개 항목", len(resolved["items"]))
    elif not resolved["serial"] or not resolved["material"]:
        try:
            excel = _resolve_from_excel(args.order, delivery=resolved["delivery"], force_no_serial=resolved["no_serial"])
        except Exception as exc:
            if "OBD가 여러 개" in str(exc):
                raise
            if not resolved["serial"]:
                raise RuntimeError(f"Serial Number를 엑셀에서 찾지 못했습니다. 런처에 직접 입력하세요. ({exc})") from exc
            excel = {}
        resolved["delivery"] = resolved["delivery"] or excel.get("delivery")
        resolved["serial"] = resolved["serial"] or excel.get("serial")
        resolved["material"] = resolved["material"] or excel.get("material")
        if excel.get("qty") and (not resolved["qty"] or str(resolved["qty"]).strip() == "1"):
            resolved["qty"] = excel.get("qty")
        resolved["no_serial"] = bool(resolved["no_serial"] or excel.get("no_serial"))
        resolved["items"] = excel.get("items") or []
        if resolved["no_serial"] and resolved["items"]:
            for item in resolved["items"]:
                item["serials"] = []
                item["no_serial"] = True
            resolved["serial"] = ""
        logger.info(
            "Excel 반영 후: delivery=%s material=%s serial=%s no_serial=%s",
            resolved["delivery"] or "",
            resolved["material"] or "",
            resolved["serial"] or "",
            resolved["no_serial"],
        )

    if not resolved["delivery"] or not resolved["material"] or not resolved["qty"]:
        sap = _resolve_from_sap(args.order)
        resolved["order"] = resolved["order"] or sap.get("order")
        resolved["delivery"] = resolved["delivery"] or sap.get("delivery")
        resolved["material"] = resolved["material"] or sap.get("material")
        resolved["qty"] = resolved["qty"] or sap.get("qty")

    if not resolved["items"] and resolved["material"]:
        serials = [resolved["serial"]] if resolved["serial"] else []
        resolved["items"] = [{
            "material": _digits(resolved["material"]),
            "qty": str(resolved["qty"] or len(serials) or 1),
            "serials": serials,
            "no_serial": bool(resolved["no_serial"]),
        }]

    if resolved["no_serial"] and resolved.get("items"):
        resolved["serial"] = ""
        for item in resolved["items"]:
            item["serials"] = []
            item["no_serial"] = True

    if resolved.get("serial") and resolved.get("items"):
        for item in resolved["items"]:
            if not item.get("no_serial") and not item.get("serials"):
                item["serials"] = [resolved["serial"]]
                break

    required = ["order", "delivery", "material", "qty"]
    if not resolved["no_serial"]:
        required.append("serial")
    missing = [name for name in required if not resolved.get(name)]
    if missing:
        if "serial" in missing:
            raise RuntimeError(
                "오늘자 엑셀에서 Serial Number가 비어 있습니다. "
                "런처의 Serial Number 칸에 직접 입력하거나 엑셀 Serial 칸을 채운 뒤 다시 실행하세요."
            )
        if "delivery" in missing:
            logger.error(
                "OBD/Delivery#를 찾지 못했습니다. Excel A열의 해당 오더 줄에 'OBD 9...' 형식으로 입력되어 있어야 합니다."
            )
        raise RuntimeError("자동 조회 실패. 런처에서 직접 입력 필요: " + ", ".join(missing))

    resolved["delivery"] = _digits(resolved["delivery"])
    resolved["material"] = _digits(resolved["material"])
    resolved["serial"] = _serial_text(resolved["serial"])
    resolved["qty"] = str(resolved["qty"]).strip() or "1"
    for item in resolved["items"]:
        item["material"] = _digits(item.get("material"))
        item["qty"] = str(item.get("qty") or len(item.get("serials", [])) or 1)
        item["serials"] = [_serial_text(sn) for sn in item.get("serials", []) if _serial_text(sn)]
        item["no_serial"] = bool(item.get("no_serial"))
    if resolved["delivery"] == _digits(resolved["order"]):
        raise RuntimeError(
            "Excel의 OBD 값이 오더번호와 같습니다. "
            "OBD/Delivery#는 SAP 오더번호가 아니라 Bloomberg Portal Delivery#입니다. 예: 92118524"
        )
    if not resolved["delivery"].startswith("9"):
        raise RuntimeError(
            f"OBD/Delivery# 값이 이상합니다: {resolved['delivery']}. "
            "Bloomberg Portal Delivery#는 보통 9로 시작합니다."
        )
    return resolved


def run_step(label, args):
    print("\n" + "=" * 70, flush=True)
    print(label, flush=True)
    print("=" * 70, flush=True)
    logger.info(label)
    cmd = [sys.executable] + args
    result = subprocess.run(
        cmd, cwd=str(BASE_DIR),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if result.returncode != 0:
        logger.error("%s 실패: exit code %s", label, result.returncode)
        raise SystemExit(result.returncode)
    logger.info("%s 완료", label)


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    parser = argparse.ArgumentParser(description="Serial registration, packing post, and label print")
    parser.add_argument("--order", required=True)
    parser.add_argument("--delivery")
    parser.add_argument("--material")
    parser.add_argument("--qty")
    parser.add_argument("--serial")
    parser.add_argument("--items-json", help="JSON item list from Workbench; bypasses Excel item lookup")
    parser.add_argument("--material-only", action="store_true", help="Treat Excel/material rows as no-serial material-only items")
    parser.add_argument("--no-serial", action="store_true", help="Alias for --material-only")
    parser.add_argument("--allow-sap-lookup", action="store_true", help="Fallback to reading VL06O when Excel/manual values are missing.")
    parser.add_argument("--skip-qr", action="store_true", help="Stop after serial registration and Packing Post; do not download/print QR labels.")
    args = parser.parse_args()
    logger.info("=======================================================")
    logger.info("[Portal 배송 처리 시작] order=%s skip_qr=%s", args.order, args.skip_qr)
    try:
        resolved = resolve_inputs(args)
    except Exception as exc:
        # Without this, a validation failure here (e.g. a missing serial)
        # was an UNCAUGHT exception - never reached any logger.error() call,
        # so it never made it into automation.log, only a traceback on
        # stderr of the launched console window. When started from
        # launcher.py that window stays open for the operator to read; when
        # started from workbench (run_portal(), no console attached) it was
        # invisible - looked like the whole automation silently did nothing.
        logger.error("[입력값 확인 실패] %s", exc)
        print(f"\n[실패] {exc}", flush=True)
        raise SystemExit(1) from exc

    logger.info("자동 조회 결과")
    logger.info("Order: %s", resolved["order"])
    logger.info("Delivery#: %s", resolved["delivery"])
    logger.info("Material: %s", resolved["material"])
    logger.info("Qty: %s", resolved["qty"])
    logger.info("Serial: %s", resolved["serial"])
    logger.info("No Serial Item: %s", resolved["no_serial"])
    logger.info("Items: %s", resolved["items"])
    print("자동 조회 결과", flush=True)
    print(f"  Order:    {resolved['order']}", flush=True)
    print(f"  Delivery: {resolved['delivery']}", flush=True)
    print(f"  Material: {resolved['material']}", flush=True)
    print(f"  Qty:      {resolved['qty']}", flush=True)
    print(f"  Serial:   {resolved['serial']}", flush=True)
    print(f"  NoSerial: {resolved['no_serial']}", flush=True)

    # Speed fix (2026-08-28, user-reported slowness): this used to always
    # run_step("0. Portal 로그인 확인", ["portal_login.py"]) here first - a
    # whole separate subprocess launch, and even its fast/common path
    # (already logged in) does a full page RELOAD of whatever tab it finds
    # just to double-check the session (portal_login.py's
    # _find_logged_in_page(), added 2026-08-14 for a real stale-cached-URL
    # bug). Timed live: ~3.5s on every single run, QR and skip-qr alike,
    # even when nothing was ever wrong. That safety net is redundant now -
    # step 1 below (portal_register_serial.py) does a real goto() to the
    # delivery page as its very first action and already detects/reports an
    # expired session itself (SSO/login-page checks in its _connect_page),
    # so an actually-expired session still gets caught, just ~3.5s sooner
    # in the pipeline instead of before it even starts.

    serial_args = [
        "portal_register_serial.py",
        "--delivery", resolved["delivery"],
        "--order", resolved["order"],
        "--items-json", json.dumps(resolved["items"], ensure_ascii=False),
    ]
    if resolved["material"]:
        serial_args.extend(["--material", resolved["material"], "--qty", resolved["qty"]])
    if resolved["serial"]:
        serial_args.extend(["--serial", resolved["serial"]])
    run_step(
        "1. Pick Qty + Run ShipERP" if resolved["no_serial"] else "1. Serial registration + Run ShipERP",
        serial_args,
    )
    run_step(
        "2. Packing Post",
        [
            "portal_pack_post.py",
            "--delivery", resolved["delivery"],
            "--material", resolved["material"],
            "--qty", resolved["qty"],
            "--items-json", json.dumps(resolved["items"], ensure_ascii=False),
            "--post",
        ],
    )
    if not args.skip_qr:
        # Real bug hit live (2026-08-03): clicking Post only settles the
        # delivery on Bloomberg Portal's own server after a few seconds -
        # portal_pack_post.py's post-click wait is short (just long enough to
        # scrape the page for an error message), and jumping straight into
        # Download Labels right after landed on the delivery page while it
        # was still "locked" server-side mid-processing. User had to back out
        # and wait before retrying by hand. A flat pause here before the QR
        # step starts is cheap insurance against that race.
        PACK_POST_SETTLE_SECONDS = 8
        logger.info("Packing Post 완료 - Portal 서버 반영 대기 %s초", PACK_POST_SETTLE_SECONDS)
        time.sleep(PACK_POST_SETTLE_SECONDS)
        run_step(
            "3. QR label download + print",
            [
                "portal_download_labels.py",
                "--delivery", resolved["delivery"],
                "--print",
            ],
        )
        done_msg = "Done: Pick/serial, Packing Post, and QR print completed."
    else:
        done_msg = "Done: Pick/serial and Packing Post completed. QR skipped."
    logger.info("[Portal processing complete] delivery=%s order=%s", resolved["delivery"], resolved["order"])
    print(f"\n{done_msg}", flush=True)
    return

if __name__ == "__main__":
    main()

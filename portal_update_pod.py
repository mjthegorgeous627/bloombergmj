"""Fill and optionally update Bloomberg Portal POD fields.

Default mode fills the POD form only. Use --update to click Update POD.
"""

import argparse
import json
import logging
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout, sync_playwright

from excel_portal_lookup import find_excel_order
from portal_lock import portal_browser_lock


CDP_URL = "http://localhost:9222"
DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
SHIPMENT_SEARCH_URL = "https://bsp.btogo.com/supplier/warehouse/shipment/index"
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"
PENDING_FILE = BASE_DIR / "pod_pending.json"

# Delivery-record lock retry (2026-08-25): right after Packing Post, Bloomberg
# Portal's own backend sync (a separate account, "REST_SVC_USR" - not our
# session) briefly holds the delivery record - going straight back in with
# page.goto() at that moment shows "[SY 530] Delivery N is locked by user
# REST_SVC_USR" and the POD tab click then times out. Confirmed live this
# isn't about which URL/link we use to get in (the Serial-registration step
# just before this hits the SAME delivery the SAME way and doesn't get
# locked) - it's purely about how much time has passed since Packing Post.
# Manually, the user backs out, waits ~5s, hits Search on the Process
# Shipment list to refresh it, then clicks back into the delivery - and that
# always clears it. _reenter_via_search() below does exactly that instead of
# just re-running page.goto() on the same URL.
#
# Real incident (2026-08-31, order 7849282): the lock doesn't always show up
# as the "is locked by" body text _is_locked() checks for. Here _connect_page
# reused the SAME tab Packing Post just finished on (URL already matched, so
# no reload/re-navigation happened) - _is_locked() passed clean, but the POD
# tab link simply hadn't rendered yet (Portal's SPA was still mid-sync with
# REST_SVC_USR) and `a:has-text('POD')` timed out after 10s with no lock
# text ever logged. Staying on the same tab never gives the SPA a reason to
# re-render its tab bar. Fix: treat a POD-tab-click timeout the same as a
# detected lock - back out via _reenter_via_search() (the same "click back
# in" the user does by hand) and retry, instead of only doing that for the
# explicit "is locked by" text case.
LOCK_RETRY_WAIT = 5      # seconds to wait before each re-entry attempt (user-specified)
LOCK_RETRY_MAX = 6       # ~30s of retrying via the search list before giving up

logger = logging.getLogger(__name__)


def _format_portal_datetime(dt=None):
    dt = dt or datetime.now()
    hour = dt.hour % 12 or 12
    ampm = "AM" if dt.hour < 12 else "PM"
    return f"{dt.strftime('%B')} {dt.day}, {dt.year} {hour}:{dt.minute:02d} {ampm}"


def _parse_datetime(value):
    if not value:
        return _format_portal_datetime()
    value = value.strip()
    for fmt in (
        "%Y-%m-%d %H:%M",
        "%Y/%m/%d %H:%M",
        "%m/%d/%Y %H:%M",
        "%B %d, %Y %I:%M %p",
        "%b %d, %Y %I:%M %p",
    ):
        try:
            return _format_portal_datetime(datetime.strptime(value, fmt))
        except ValueError:
            continue
    return value


def _serial_text(value):
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].replace(".", "", 1).isdigit():
        text = text[:-2]
    return text


def _is_locked(page):
    try:
        text = page.locator("body").inner_text(timeout=3000)
    except Exception:
        return False
    return "is locked by" in text.lower()


def _reenter_via_search(page, delivery_num):
    """Re-enter the delivery through Warehouse > Process Shipment > Search,
    same path portal_download_labels.py's _open_process_shipment() uses and
    the same path the user takes by hand - instead of goto()'ing straight
    back to the same delivery URL. See LOCK_RETRY_WAIT comment above."""
    page.goto(SHIPMENT_SEARCH_URL, wait_until="domcontentloaded", timeout=30000)
    try:
        page.locator('input[name="quickFinder"]').first.fill(str(delivery_num), timeout=8000)
        page.get_by_role("button", name="Search").click(timeout=10000)
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)
    except Exception as exc:
        logger.warning("검색 재진입 중 Search 실패(무시하고 계속): %s", exc)
    link = page.locator(f"a:has-text('{delivery_num}')").first
    link.wait_for(timeout=15000)
    link.click(timeout=10000)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    page.wait_for_timeout(1000)


def _connect_page(pw, delivery_num):
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0] if browser.contexts else None
    if context is None:
        raise RuntimeError("Chrome context를 찾지 못했습니다.")
    # Reuse the tab that's already on this delivery's page if one exists,
    # same as portal_register_serial.py/portal_pack_post.py's _connect_page -
    # this file used to always grab context.pages[-1] ("the last tab") and
    # force-navigate IT to the target URL regardless of what was open there,
    # which could hijack an unrelated tab the user was looking at (e.g. the
    # manual-login tab, or a Process Shipment listing from the QR step just
    # before this one) instead of reusing the correct one that may already
    # be open.
    page = None
    for candidate in context.pages:
        if f"/warehouse/delivery/{delivery_num}" in candidate.url or SHIPMENT_SEARCH_URL in candidate.url:
            page = candidate
            break
    if page is None:
        page = context.pages[-1] if context.pages else context.new_page()

    # Real bug (2026-08-28, order 7847017): entering by re-navigating the
    # SAME URL the reused tab was already on (a bare page.goto()) forces a
    # full reload of what's a client-side SPA - Serial registration -> Run
    # ShipERP -> Packing Post all advance an in-memory wizard step without
    # changing the URL - which threw away that progress and dropped the page
    # back to its default view, timing out the POD click that follows since
    # that link only exists once Packing Post (just undone client-side) has
    # completed.
    #
    # Real incident (2026-08-31, order 7849268/7849282): fixing that by
    # skipping the goto() when the tab was ALREADY sitting on this delivery
    # (i.e. no navigation at all right after Packing Post) traded that bug
    # for this one - Portal's own backend sync ("REST_SVC_USR") briefly locks
    # the just-Packed delivery record, and a tab that's simply left sitting
    # there never gives the SPA a reason to re-render past that lock. The
    # user confirmed live: going straight in (reusing the tab / goto'ing the
    # delivery URL directly) hit the lock screen every time right after
    # Packing Post, but going in via Warehouse > Process Shipment > Search
    # (same as the LOCK_RETRY fallback below, and the same "click in, don't
    # jump straight to the link" path the user takes by hand) worked in one
    # shot. So: always enter via the search list, not a direct goto/reuse -
    # by the time POD runs, Packing Post has already committed server-side,
    # so a real search-driven re-navigation costs a few seconds but doesn't
    # lose any progress the way the plain reload above did.
    _reenter_via_search(page, delivery_num)

    # Real incident (2026-08-28, order 7847017): Bloomberg's corporate SSO
    # session had quietly expired between Packing Post and this POD step, so
    # entry above landed on bsso.blpprofessional.com's login page instead of
    # the delivery detail page. This file had no check for that - it fell
    # straight into _open_pod_tab()'s `a:has-text('POD')` click, which of
    # course isn't on a login page, and failed with an opaque 10s timeout
    # that looked identical to the "stuck on Serial registration" symptom
    # this function's docstring above was written to fix. Same SSO check
    # portal_register_serial.py's _connect_page() already has, so this fails
    # fast with an actionable message instead.
    if "bsso.blpprofessional.com" in (page.url or "") or "/idp/startSSO" in (page.url or "") \
            or "/supplier/login" in (page.url or ""):
        raise RuntimeError(
            "Bloomberg Portal/SSO 세션이 만료되어 재로그인이 필요합니다. "
            "자동화 Chrome 창(포트 9222)에서 직접 로그인을 완료한 뒤 다시 실행하세요."
        )

    if f"/warehouse/delivery/{delivery_num}" not in (page.url or ""):
        # Real incident (2026-08-28, order 67083934): same SPA quirk
        # portal_register_serial.py's _connect_page() already retries for -
        # entry above sometimes bounces to the bare filters/list page
        # (.../warehouse/delivery/index) instead of the requested delivery's
        # detail page, not a login redirect so the SSO check above doesn't
        # catch it. That file retries the same goto() once before giving up;
        # this one had no such retry and fell straight into _open_pod_tab()'s
        # `a:has-text('POD')` click on the list page, timing out with the
        # same opaque symptom as the reload bug above and the SSO bug just
        # above it - three different causes, identical-looking failure.
        logger.warning("Portal 상세 페이지 대신 다른 화면(%s)으로 이동됨 - 재시도", page.url)
        _reenter_via_search(page, delivery_num)
        logger.info("재시도 후 Portal URL: %s", page.url)
        if f"/warehouse/delivery/{delivery_num}" not in (page.url or ""):
            raise RuntimeError(
                f"Delivery {delivery_num} 상세 화면으로 진입하지 못했습니다 (현재 URL: {page.url}). "
                "자동화 Chrome 창에서 상태를 확인한 뒤 다시 실행하세요."
            )

    attempt = 0
    while _is_locked(page):
        attempt += 1
        if attempt > LOCK_RETRY_MAX:
            raise RuntimeError(
                f"Delivery {delivery_num} 잠금(REST_SVC_USR)이 재시도 {LOCK_RETRY_MAX}회 후에도 풀리지 않았습니다."
            )
        logger.warning(
            "Delivery %s 잠금 감지(REST_SVC_USR) - %d초 대기 후 검색 목록으로 재진입 (%d/%d)",
            delivery_num, LOCK_RETRY_WAIT, attempt, LOCK_RETRY_MAX,
        )
        page.wait_for_timeout(LOCK_RETRY_WAIT * 1000)
        _reenter_via_search(page, delivery_num)

    return browser, page


def _open_pod_tab(page):
    pod = page.locator("a:has-text('POD')").first
    pod.click(timeout=10000)
    page.locator('select[name="trackingStatus"]').wait_for(timeout=10000)
    page.wait_for_timeout(500)


def _fill_react_input(page, selector, value):
    page.locator(selector).fill(str(value), timeout=8000)


def _collect_errors(page):
    text = page.locator("body").inner_text(timeout=5000)
    lines = []
    for line in text.splitlines():
        lower = line.lower()
        if any(word in lower for word in ["error", "invalid", "must", "required", "failed", "cannot"]):
            lines.append(line.strip())
    return [line for line in lines if line][:10]


def _build_remarks_from_items(items, signed_by, pickup_done):
    """Same phrasing excel_portal_lookup.build_default_remarks() produces,
    but built from a plain item list (workbench.db rows) instead of an Excel
    lookup - reuses its exact label/count-formatting helpers so the wording
    stays identical either way."""
    from excel_portal_lookup import _format_item_counts, _item_label

    deliveries, pickups, pickup_expected = [], [], False
    for item in items:
        label = _item_label(item.get("description"))
        qty = item.get("qty") or 1
        item_type = str(item.get("item_type") or "")
        if item_type in ("회수", "회수 Delayed"):
            pickup_expected = True
            if pickup_done:
                pickups.append((label, qty))
        else:
            deliveries.append((label, qty))
    delivered = _format_item_counts(deliveries) or "1xItem"
    collected = _format_item_counts(pickups)
    signed_by = signed_by or ""
    if pickup_expected:
        if collected:
            return f"Delivered {delivered}, and collected {collected}, to/from {signed_by}."
        return f"Delivered {delivered} to {signed_by}, and will collect later."
    return f"Delivered {delivered} to {signed_by}."


def prepare_values(order, signed_by=None, delivery_datetime=None, remarks=None, pickup_done=False, items_override=None):
    """items_override: optional list of {"description","serial","qty","item_type"}
    dicts sourced from workbench.db - bypasses the Excel lookup entirely when
    given. Real bug this fixes (2026-08-03): find_excel_order() only searches
    TODAY's dated sheet (_recent_sheet_names() ignores its own `days` param
    and just returns today's sheet name), so an order whose 회수 leg was
    moved to a different day on the board - a real, common case, e.g. a ZRX's
    배송 leg ships today but its 회수 leg is picked up days later - can
    silently fail to find that leg in Excel at all. That made the
    auto-built remarks wrongly say "will collect later" even when the pickup
    WAS just completed and checked off on the board, or could raise a hard
    "Excel에서 오더를 찾지 못했습니다" lookup error and block the whole POD
    action even though workbench already has everything needed. workbench.py
    always supplies items_override (+ signed_by from orders.customer) so this
    path is the one actually used from the dashboard; the Excel path remains
    for launcher.py, unchanged."""
    if items_override is not None:
        delivery_serials, pickup_serials = [], []
        for item in items_override:
            serial = _serial_text(item.get("serial"))
            if not serial or serial.upper() == "X":
                continue
            if str(item.get("item_type") or "") in ("회수", "회수 Delayed"):
                pickup_serials.append(serial)
            else:
                delivery_serials.append(serial)
        signed = (signed_by or "").strip()
        final_remarks = (remarks or _build_remarks_from_items(items_override, signed, pickup_done)).strip()
        return {
            "order": str(order),
            "signed_by": signed,
            "delivery_datetime": _parse_datetime(delivery_datetime),
            "remarks": final_remarks,
            "excel": {"sheet": "(workbench)", "start_row": 0, "end_row": 0},
            "pickup_done": bool(pickup_done),
            "delivery_serials": delivery_serials,
            "pickup_serials": pickup_serials,
        }

    pickup_mode = "collected" if pickup_done else None
    info = find_excel_order(order, pickup_mode=pickup_mode)
    signed = (signed_by or info["customer"]).strip()
    final_remarks = (remarks or info["remarks"]).strip()
    delivery_serials = []
    pickup_serials = []
    for row in info.get("rows", []):
        serial = _serial_text(row.get("serial"))
        if not serial or serial.upper() == "X":
            continue
        order_text = str(row.get("order_text") or "")
        if "회수" in order_text or "PICK" in order_text.upper():
            pickup_serials.append(serial)
        else:
            delivery_serials.append(serial)
    values = {
        "order": str(order),
        "signed_by": signed,
        "delivery_datetime": _parse_datetime(delivery_datetime),
        "remarks": final_remarks,
        "excel": info,
        "pickup_done": bool(pickup_done),
        "delivery_serials": delivery_serials,
        "pickup_serials": pickup_serials,
    }
    return values


def write_pending(delivery_num, values):
    payload = {
        "delivery": str(delivery_num),
        "order": str(values["order"]),
        "tracking_status": "DTD-DELIVERED",
        "delivery_datetime": values["delivery_datetime"],
        "signed_by": values["signed_by"],
        "remarks": values["remarks"],
        "pickup_done": values.get("pickup_done", False),
        "delivery_serials": values.get("delivery_serials", []),
        "pickup_serials": values.get("pickup_serials", []),
        "excel_sheet": values["excel"]["sheet"],
        "excel_rows": f"{values['excel']['start_row']}-{values['excel']['end_row']}",
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }
    PENDING_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return payload


def log_preview(payload):
    logger.info("=======================================================")
    logger.info("[POD 확인 필요]")
    logger.info("Delivery#: %s", payload["delivery"])
    logger.info("Order: %s", payload["order"])
    logger.info("Tracking Status: %s", payload["tracking_status"])
    logger.info("Delivery Date & Time: %s", payload["delivery_datetime"])
    logger.info("Signed By: %s", payload["signed_by"])
    logger.info("Remarks: %s", payload["remarks"])
    logger.info("Excel: %s rows %s", payload["excel_sheet"], payload["excel_rows"])
    logger.info("문제 없으면 런처에서 'POD 최종 저장'을 누르세요.")


def fill_pod(delivery_num, values, update=False):
    with portal_browser_lock(f"update_pod:{delivery_num}"):
        with sync_playwright() as pw:
            browser, page = _connect_page(pw, delivery_num)
            try:
                body = page.locator("body").inner_text(timeout=5000)
                if delivery_num not in body:
                    raise RuntimeError(f"Delivery# {delivery_num}을 화면에서 확인하지 못했습니다.")

                attempt = 0
                while True:
                    try:
                        _open_pod_tab(page)
                        break
                    except PlaywrightTimeout:
                        attempt += 1
                        if attempt > LOCK_RETRY_MAX:
                            raise RuntimeError(
                                f"Delivery {delivery_num}: POD 탭 진입 재시도 {LOCK_RETRY_MAX}회 후에도 "
                                "실패했습니다 (REST_SVC_USR 동기화 지연 추정)."
                            )
                        logger.warning(
                            "Delivery %s: POD 탭 클릭 실패(잠금 동기화 지연 추정, 락 문구 없음) - "
                            "%d초 대기 후 검색 목록으로 재진입 (%d/%d)",
                            delivery_num, LOCK_RETRY_WAIT, attempt, LOCK_RETRY_MAX,
                        )
                        page.wait_for_timeout(LOCK_RETRY_WAIT * 1000)
                        _reenter_via_search(page, delivery_num)

                page.locator('select[name="trackingStatus"]').select_option("DTD")
                _fill_react_input(page, 'input[name="deliveryDate"]', values["delivery_datetime"])
                _fill_react_input(page, 'input[name="signedBy"]', values["signed_by"])
                _fill_react_input(page, 'input[name="remarks"]', values["remarks"])
                logger.info("POD 입력 완료")
                logger.info("Tracking Status: DTD-DELIVERED")
                logger.info("Delivery Date & Time: %s", values["delivery_datetime"])
                logger.info("Signed By: %s", values["signed_by"])
                logger.info("Remarks: %s", values["remarks"])

                if update:
                    page.get_by_role("button", name="Update POD").click(timeout=10000)
                    page.wait_for_timeout(3000)
                    errors = _collect_errors(page)
                    if errors:
                        raise RuntimeError("Update POD 후 에러 감지: " + " / ".join(errors))
                    logger.info("Update POD 클릭 완료")
                else:
                    logger.info("Dry-run: Update POD는 누르지 않았습니다.")
                try:
                    page.screenshot(path="portal_update_pod_result.png", full_page=True, timeout=10000)
                except Exception as exc:
                    logger.warning("결과 스크린샷 저장 실패(무시): %s", exc)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass


def main():
    parser = argparse.ArgumentParser(description="Fill Bloomberg Portal POD fields from Excel defaults.")
    parser.add_argument("--delivery", required=True)
    parser.add_argument("--order", required=True)
    parser.add_argument("--signed-by", default="")
    parser.add_argument("--datetime", default="")
    parser.add_argument("--remarks", default="")
    parser.add_argument("--pickup-done", action="store_true", help="For ZRX, build remarks as delivery and pickup both completed.")
    parser.add_argument("--items-json", default="", help="Item list from Workbench (workbench.db); bypasses the Excel lookup entirely.")
    parser.add_argument("--update", action="store_true")
    parser.add_argument("--preview", action="store_true", help="Only print Excel-derived values; do not touch portal.")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    try:
        items_override = None
        if args.items_json:
            try:
                items_override = json.loads(args.items_json)
            except (ValueError, TypeError) as exc:
                raise RuntimeError(f"--items-json 파싱 실패: {exc}") from exc
        values = prepare_values(
            args.order, args.signed_by, args.datetime, args.remarks,
            pickup_done=args.pickup_done, items_override=items_override,
        )
        payload = write_pending(args.delivery, values)
        log_preview(payload)
        if not args.preview:
            fill_pod(args.delivery, values, update=args.update)
        if args.update:
            try:
                PENDING_FILE.unlink()
            except FileNotFoundError:
                pass
    except Exception as exc:
        logger.error("실패: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

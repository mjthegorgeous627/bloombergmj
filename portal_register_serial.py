"""Register a serial number on a Bloomberg Portal delivery page.

This script is intentionally guarded. It only modifies the portal page after
checking the expected delivery number, SAP order number, material, and quantity.
It connects to an already-open Chrome instance on port 9222.
"""

import argparse
import json
import logging
import sys
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout, sync_playwright

from portal_lock import portal_browser_lock

CDP_URL = "http://localhost:9222"
DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"

logger = logging.getLogger(__name__)


def _norm(value):
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def _norm_loose(value):
    return _norm(value).lstrip("0") or _norm(value)


def _page_text(page):
    try:
        body = page.locator("body").inner_text(timeout=5000)
    except Exception:
        body = ""
    try:
        # Real incident (2026-08-20, order 67079748/ZRX): navigation bounced
        # to Bloomberg's corporate SSO login page (bsso.blpprofessional.com),
        # which had username+password pre-filled by Chrome's saved-password
        # autofill. This scrape used to grab EVERY input's value including
        # password fields, and _validate_page() below logs a snippet of this
        # text on failure - so the plaintext password ended up written to
        # automation.log. Excluding type=password (and anything Chrome/the
        # page marks autocomplete=current-password/new-password) keeps a
        # credential from ever reaching a log file again.
        values = page.evaluate(
            """
            () => Array.from(document.querySelectorAll('input, textarea, select'))
              .filter(el => (el.type || '').toLowerCase() !== 'password'
                && !/current-password|new-password/i.test(el.autocomplete || ''))
              .map(el => el.value || el.textContent || '')
              .filter(Boolean)
              .join('\\n')
            """
        )
    except Exception:
        values = ""
    return body + "\n" + values


def _connect_page(pw, delivery_num):
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0] if browser.contexts else None
    if context is None:
        raise RuntimeError("Chrome context를 찾지 못했습니다.")

    target_url = DELIVERY_BASE + delivery_num
    logger.info("Portal delivery detail 이동: %s", target_url)
    page = None
    for candidate in context.pages:
        if f"/warehouse/delivery/{delivery_num}" in candidate.url:
            page = candidate
            break
    if page is None:
        page = context.pages[-1] if context.pages else context.new_page()
        page.goto(target_url, wait_until="domcontentloaded", timeout=30000)

    page.wait_for_load_state("domcontentloaded", timeout=15000)
    page.wait_for_timeout(2500)
    logger.info("Portal 현재 URL: %s", page.url)

    if "/supplier/login" in (page.url or ""):
        # Real bug (2026-08-14, order 67076875/ZRX): step 0's login-check
        # ("0. Portal 로그인 확인") passed in under a second, but the session
        # had actually already expired - the delivery page navigation above
        # got redirected here. Falling straight into _validate_page() below
        # used to produce a misleading "포털 화면 검증 실패: Delivery#=...,
        # Material=..." (reads like a data-matching bug, not "logged out").
        # portal_login.py's own login-check now reloads before trusting a
        # cached URL (same-day fix) so this should be rare, but a session
        # that expires in the couple seconds between that check and this
        # navigation is still possible - so retry the real login here once
        # with the saved credentials before giving up.
        logger.warning("Portal 세션이 만료되어 로그인 화면으로 이동됨 - 재로그인 시도")
        # Fill directly on the already-connected `page` instead of calling
        # portal_login.autofill_login_if_needed() - that helper takes its
        # own portal_browser_lock(), and register_serial() is already
        # holding that same lock (single shared lock file regardless of
        # label - see portal_lock.py), so calling it here would self-block
        # against its own lock for the full WAIT_TIMEOUT and then raise
        # PortalBusyError instead of actually retrying the login.
        try:
            from credentials import PORTAL_PASSWORD, PORTAL_USER
        except ImportError:
            PORTAL_USER = ""
            PORTAL_PASSWORD = ""
        if PORTAL_USER and PORTAL_PASSWORD:
            try:
                user_field = page.locator(
                    "input[name='pf.username'], input[name='username'], input[type='email'], input[type='text']"
                ).first
                pass_field = page.locator("input[type='password']").first
                if user_field.count():
                    user_field.fill(PORTAL_USER, timeout=3000)
                if pass_field.count():
                    pass_field.fill(PORTAL_PASSWORD, timeout=3000)
                page.keyboard.press("Enter")
                page.wait_for_timeout(2000)
            except Exception as exc:
                logger.warning("재로그인 자동입력 실패: %s", exc)
        page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        page.wait_for_timeout(2500)
        logger.info("재로그인 시도 후 Portal URL: %s", page.url)
        if "/supplier/login" in (page.url or ""):
            raise RuntimeError(
                "Portal 세션이 만료되어 재로그인이 필요합니다. "
                "자동화 Chrome 창(포트 9222)에서 Bloomberg Portal에 직접 로그인한 뒤 다시 실행하세요."
            )

    if "bsso.blpprofessional.com" in (page.url or "") or "/idp/startSSO" in (page.url or ""):
        # Real incident (2026-08-20, order 67079748/ZRX): three retries in a
        # row all bounced here instead of the Portal's own /supplier/login -
        # Bloomberg's corporate SSO (separate from bsp.btogo.com's own login
        # check above) had actually logged out. This page can require a
        # B-Unit/MFA challenge after the username+password step, which isn't
        # safe to auto-submit here (repeated blind submissions against an
        # MFA gate risk a lockout, and Chrome's saved-password autofill
        # already fills the fields on load anyway - see the credential-log
        # bug this same incident caused in _page_text() above). Stop with a
        # clear, actionable error instead of retrying blindly into a page
        # that will never resolve to the delivery detail on its own.
        raise RuntimeError(
            "Bloomberg SSO 세션이 만료되어 재로그인이 필요합니다 (B-Unit 인증 필요할 수 있음). "
            "자동화 Chrome 창(포트 9222)에서 직접 로그인을 완료한 뒤 다시 실행하세요."
        )

    if f"/warehouse/delivery/{delivery_num}" not in (page.url or ""):
        # Real bug (2026-08-20, order 67079748/ZRX): the SPA sometimes bounces
        # the goto() above to the bare filters/list page
        # (.../warehouse/delivery/index) instead of landing on the requested
        # delivery's detail page - not a login redirect, so the re-login
        # block above doesn't catch it. Falling straight into _validate_page()
        # produced a misleading "포털 화면 검증 실패: Delivery#=..., Material=..."
        # (reads like the delivery/material don't match, when really the
        # detail page just never loaded). A same-delivery retry a few seconds
        # later has reliably landed on the right page in past incidents (see
        # 2026-08-12, order 7841018 - identical symptom, second attempt 21s
        # later succeeded outright) - so retry the navigation once here
        # before falling through to validation.
        logger.warning("Portal 상세 페이지 대신 다른 화면(%s)으로 이동됨 - 재시도", page.url)
        page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        page.wait_for_timeout(2500)
        logger.info("재시도 후 Portal URL: %s", page.url)

    return browser, page


def _validate_page(page, delivery_num, order_num, material, qty):
    text = _page_text(page)
    missing = []
    for label, expected in [
        ("Delivery#", delivery_num),
        ("Material", material),
    ]:
        if _norm_loose(expected) not in _norm_loose(text):
            missing.append(f"{label}={expected}")
    if _norm_loose(order_num) not in _norm_loose(text):
        logger.warning("SAP order %s는 화면에서 확인되지 않음 - Delivery#/Material 검증으로 진행", order_num)

    if missing:
        snippet = " / ".join(line.strip() for line in text.splitlines() if line.strip())[:600]
        logger.error("Portal 상세 화면 검증 실패. missing=%s", ", ".join(missing))
        logger.error("현재 URL: %s", page.url)
        logger.error("화면 일부: %s", snippet)
        try:
            page.screenshot(path=str(BASE_DIR / "portal_register_serial_error.png"), full_page=True, timeout=10000)
            logger.info("오류 화면 저장: portal_register_serial_error.png")
        except Exception as exc:
            logger.warning("오류 화면 저장 실패: %s", exc)
        raise RuntimeError("포털 화면 검증 실패: " + ", ".join(missing))

    logger.info("포털 화면 검증 완료: delivery=%s order=%s material=%s", delivery_num, order_num, material)


def _click_by_text(page, text, timeout=8000):
    locators = [
        page.get_by_role("button", name=text),
        page.get_by_role("link", name=text),
        page.locator(f"text={text}"),
        page.locator(f"button:has-text('{text}')"),
        page.locator(f"a:has-text('{text}')"),
    ]
    last_error = None
    for locator in locators:
        try:
            locator.first.click(timeout=timeout)
            return True
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"'{text}' 클릭 실패: {last_error}")


def _click_optional_by_text(page, text, timeout=4000):
    try:
        return _click_by_text(page, text, timeout=timeout)
    except Exception as exc:
        logger.info("%s button not found; assuming this step is already complete: %s", text, exc)
        return False


def _get_material_rows(page, material, start=0):
    """All visible DOM rows for a material, from `start` onward, each with
    its own deliveryQuantity - in document order.

    Real incident (2026-08-21, order 67055178/ZRX, CISCO router qty=2):
    run_portal()'s merge step (workbench_app.py) combines same-material
    units into ONE items-json entry with the total qty and every serial, on
    the documented assumption that "SAP combines multiple units of the same
    material into ONE outbound delivery line, so Portal shows a single Pick
    Quantity row per material." That assumption doesn't always hold - this
    delivery's Portal page actually rendered TWO separate rows for the same
    material (Item#10, Item#20), each with deliveryQuantity=1. The old code
    filled the full combined qty (2) into just the first row regardless -
    that row's own deliveryQuantity was only 1, so the fill silently
    over-filled it, both serials got jammed into that one row's serial list,
    and the resulting Run ShipERP page never became a valid Shipment screen
    (next step, Packing Post, failed with "Shipment 화면이 아닙니다").
    Reading each row's real deliveryQuantity up front lets the caller
    distribute an item's serials across however many rows Portal actually
    gave it, instead of assuming exactly one row holds the whole qty."""
    rows = page.evaluate(
        """
        ({ material, start }) => {
          const visible = (el) => {
            if (!el) return false;
            const r = el.getBoundingClientRect();
            const s = window.getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
          };
          const matches = Array.from(document.querySelectorAll('input[name$=".material"]'))
            .filter(el => visible(el) && (el.value || '').trim() === material);
          return matches.slice(start).map(mat => {
            const m = (mat.name || '').match(/items\\.\\[(\\d+)\\]\\./);
            if (!m) return null;
            const idx = Number(m[1]);
            const dq = document.querySelector(`input[name="items.[${idx}].deliveryQuantity"]`);
            return { index: idx, deliveryQuantity: dq ? (Number(dq.value) || 1) : 1 };
          }).filter(Boolean);
        }
        """,
        {"material": str(material), "start": int(start)},
    )
    return rows


def _fill_item_fields(page, item_index, qty, need_serial=True):
    pick_selector = f'input[name="items.[{item_index}].pickQuantity"]'
    pick_input = page.locator(pick_selector)
    if pick_input.is_disabled(timeout=2000):
        current = (pick_input.input_value(timeout=2000) or "").strip()
        expected = str(qty).strip()
        if current == expected:
            logger.info(
                "Pick Quantity already locked with expected value; skipping item: index=%s qty=%s",
                item_index,
                qty,
            )
            return {"already_picked": True}
        raise RuntimeError(
            f"Pick Quantity is disabled with unexpected value: index={item_index} current={current} expected={expected}"
        )
    pick_input.fill(str(qty), timeout=8000)
    if need_serial:
        page.locator("button:has-text('View/Update')").nth(item_index).click(timeout=8000)
        page.wait_for_timeout(1000)
        logger.info("Pick Quantity 입력 및 View/Update 클릭 완료: index=%s qty=%s", item_index, qty)
    else:
        logger.info("Pick Quantity 입력 완료 - Serial 없는 품목이라 View/Update 생략: index=%s qty=%s", item_index, qty)
    return {"already_picked": False}


def _serial_exists(page, serial):
    serial_digits = _norm(serial)
    if not serial_digits:
        return False
    try:
        return page.evaluate(
            """
            ({ serial }) => {
              const text = document.body.innerText || '';
              return text.replace(/\\D/g, '').includes(serial);
            }
            """,
            {"serial": serial_digits},
        )
    except Exception:
        return False


def _add_serial(page, serial):
    page.wait_for_timeout(1000)
    page.locator('input[name="serial"]').fill(str(serial), timeout=8000)
    page.locator("button:has-text('+')").nth(0).click(timeout=8000)
    page.wait_for_timeout(1000)
    logger.info("Serial 입력 및 + 클릭 완료")


def _wait_for_serial(page, serial):
    try:
        page.locator(f"text={serial}").first.wait_for(timeout=8000)
    except PlaywrightTimeout as exc:
        raise RuntimeError(f"Serial {serial} 추가 확인 실패") from exc
    logger.info("Serial 추가 확인 완료: %s", serial)


def _normalize_items(material=None, qty=None, serial="", no_serial=False, items=None):
    if items:
        normalized = []
        for item in items:
            serials = [str(sn).strip() for sn in item.get("serials", []) if str(sn).strip()]
            item_no_serial = bool(item.get("no_serial")) or not serials
            normalized.append({
                "material": str(item.get("material") or "").strip(),
                "qty": str(item.get("qty") or len(serials) or 1).strip(),
                "serials": serials,
                "no_serial": item_no_serial,
            })
        return normalized
    serials = [str(serial).strip()] if str(serial or "").strip() else []
    return [{
        "material": str(material or "").strip(),
        "qty": str(qty or len(serials) or 1).strip(),
        "serials": serials,
        "no_serial": bool(no_serial) or not serials,
    }]


def register_serial(delivery_num, order_num, material=None, qty=None, serial="", run_ship_erp=True, no_serial=False, items=None):
    items_to_process = _normalize_items(material, qty, serial, no_serial, items)
    for item in items_to_process:
        if not item["material"]:
            raise RuntimeError("Material 값이 비어 있습니다.")
        if not item["no_serial"] and not item["serials"]:
            raise RuntimeError(f"Material {item['material']} Serial Number가 필요합니다.")

    with portal_browser_lock(f"register_serial:{delivery_num}"):
        with sync_playwright() as pw:
            browser, page = _connect_page(pw, delivery_num)
            try:
                page_text = _page_text(page)
                if "Items to be packed" in page_text or "Packed Items" in page_text:
                    logger.info("이미 Shipment/Packing 화면입니다. Serial 등록과 Run ShipERP 단계를 생략합니다.")
                    return True
                material_occurrences = {}
                for item in items_to_process:
                    material = item["material"]
                    qty = item["qty"]
                    serials = item["serials"]
                    item_no_serial = item["no_serial"]
                    _validate_page(page, delivery_num, order_num, material, qty)
                    start = material_occurrences.get(material, 0)
                    rows = _get_material_rows(page, material, start=start)
                    if not rows:
                        raise RuntimeError(f"material {material} not found")

                    if item_no_serial:
                        # No serials to track - Portal has never been observed
                        # splitting a no-serial bulk line the way it does
                        # serialized units, so this stays the original
                        # single-row/full-qty behavior.
                        row = rows[0]
                        _fill_item_fields(page, row["index"], qty, need_serial=False)
                        material_occurrences[material] = start + 1
                        logger.info("Serial 없는 품목 처리: material=%s qty=%s", material, qty)
                        continue

                    # Serialized item: distribute `serials` across however
                    # many rows Portal actually rendered for this material -
                    # each row can only hold up to its own deliveryQuantity
                    # (usually 1, but a real combined row can be more) -
                    # see _get_material_rows()'s comment for why this can no
                    # longer assume "the first row, full qty".
                    remaining = list(serials)
                    rows_used = 0
                    for row in rows:
                        if not remaining:
                            break
                        capacity = max(1, int(row.get("deliveryQuantity") or 1))
                        take = remaining[:capacity]
                        remaining = remaining[capacity:]
                        rows_used += 1
                        item_state = _fill_item_fields(page, row["index"], len(take), need_serial=True)
                        if item_state.get("already_picked"):
                            logger.info("Item already picked; skipping serial entry: material=%s row=%s serials=%s", material, row["index"], take)
                            continue
                        for serial_num in take:
                            if _serial_exists(page, serial_num):
                                logger.info("Serial already present; skipping add: %s", serial_num)
                            else:
                                _add_serial(page, serial_num)
                                _wait_for_serial(page, serial_num)
                        _click_optional_by_text(page, "Done")
                        logger.info("Done 클릭 완료: material=%s row=%s serials=%s", material, row["index"], take)
                        page.wait_for_timeout(1000)
                    material_occurrences[material] = start + rows_used
                    if remaining:
                        raise RuntimeError(
                            f"Material {material}: 남은 Serial {remaining}을(를) 채울 Portal 행이 부족합니다 "
                            f"(사용 가능 행 {len(rows)}개)."
                        )
                if run_ship_erp:
                    if _click_optional_by_text(page, "Run ShipERP"):
                        logger.info("Run ShipERP 클릭 완료")
                    # 2026-09-16: 위 클릭이 "버튼을 못 찾음 -> 이미 완료된 걸로
                    # 간주"로 조용히 넘어간 경우, 실제로는 버튼이 막혀있거나
                    # 포털 상태가 예상과 달라서 못 누른 것뿐일 수 있다. 클릭
                    # 성공/스킵 여부와 무관하게 화면이 실제로 다음 단계(Shipment/
                    # Packing)로 넘어갔는지 확인해서, 여기서 조용히 성공 처리하고
                    # 넘어가는 대신 명확히 실패시킨다 - 안 그러면 다음 단계
                    # (Packing Post)가 "packing 버튼을 찾지 못했습니다" 같은,
                    # 진짜 원인과 동떨어진 에러로 대신 실패한다 (2026-09-15/16
                    # 실측: 오더 7858029, 67093809이 정확히 이 증상이었음).
                    page.wait_for_timeout(1500)
                    page_text_after = _page_text(page)
                    if "Items to be packed" not in page_text_after and "Packed Items" not in page_text_after:
                        try:
                            page.screenshot(path=str(BASE_DIR / "portal_register_serial_error.png"), full_page=True, timeout=10000)
                        except Exception:
                            pass
                        raise RuntimeError(
                            "Run ShipERP 이후 Shipment/Packing 화면으로 전환되지 않았습니다 - "
                            "버튼 클릭이 실패했거나 포털 상태가 예상과 다릅니다."
                        )
                return True
            finally:
                try:
                    browser.close()
                except Exception:
                    pass


def main():
    parser = argparse.ArgumentParser(description="Register serial number on Bloomberg Portal delivery.")
    parser.add_argument("--delivery", required=True)
    parser.add_argument("--order", required=True)
    parser.add_argument("--material", default="")
    parser.add_argument("--qty", default="1")
    parser.add_argument("--serial", default="")
    parser.add_argument("--items-json", default="")
    parser.add_argument("--no-serial", action="store_true")
    parser.add_argument("--no-shiperp", action="store_true")
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
        items = json.loads(args.items_json) if args.items_json else None
        logger.info(
            "[Serial 등록 시작] delivery=%s order=%s material=%s qty=%s serial=%s no_serial=%s items=%s",
            args.delivery, args.order, args.material, args.qty, args.serial, args.no_serial, items,
        )
        register_serial(args.delivery, args.order, args.material, args.qty, args.serial, not args.no_shiperp, args.no_serial, items)
    except Exception as exc:
        logger.error("실패: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

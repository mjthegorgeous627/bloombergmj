"""Shared Bloomberg Portal navigation: how every portal_*.py script gets
from "some CDP-connected Chrome" to "sitting on delivery #N's detail page".

Real incident (2026-08-28, orders 7847017 / 67083934): every portal_*.py
script used to reach a delivery's detail page with a raw
page.goto(DELIVERY_BASE + delivery_num) - an address-bar-style deep link.
This SPA's client router does not reliably resolve that: on a fresh full
page load it sometimes lands correctly on the delivery detail, and
sometimes bounces back to the bare Process Shipment filter page
(.../warehouse/delivery/index) instead - a race in the app's own bundle
init, unrelated to the delivery or session being fine. Comparing against
how a person actually uses the portal (confirmed live via screenshots,
2026-08-28): Warehouse > Process Shipment > Search > click the OBD number
link - that path has never been seen to bounce, because it's real in-app
client-side routing (an actual link click), not a fresh navigation to a
deep URL. portal_register_serial.py used to retry the same goto() once
after a bounce, and portal_update_pod.py was patched the same day to do
the same - that only reduces how often the failure surfaces, it's still a
coin flip. This module makes the reliable path (search then click) the
DEFAULT way in for every script, and keeps reusing an already-open,
already-correct tab (no navigation at all) as the fast path for the common
case of moving from one step to the next against the same delivery in the
same tab (Serial registration -> Packing Post -> POD).
"""

import logging

CDP_URL = "http://localhost:9222"
DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
SHIPMENT_SEARCH_URL = "https://bsp.btogo.com/supplier/warehouse/shipment/index"

logger = logging.getLogger(__name__)


def _delivery_path(delivery_num):
    return f"/warehouse/delivery/{delivery_num}"


def _is_sso_or_login(url):
    url = url or ""
    return (
        "bsso.blpprofessional.com" in url
        or "/idp/startSSO" in url
        or "/supplier/login" in url
    )


def find_open_delivery_tab(context, delivery_num):
    """Return the context's page already sitting on this delivery's detail
    page, or None. Checked first everywhere so an in-progress multi-step
    flow (Serial -> Pack Post -> POD) never re-navigates at all."""
    target = _delivery_path(delivery_num)
    for candidate in context.pages:
        if target in (candidate.url or ""):
            return candidate
    return None


def _retry_login_if_expired(page):
    """Same-day session expiring between steps (2026-08-14, order
    67076875/ZRX): auto-fill the saved credentials once and retry, since
    this is the portal's own quick login form (not the corporate SSO/MFA
    gate handled separately below)."""
    if "/supplier/login" not in (page.url or ""):
        return
    logger.warning("Portal 세션이 만료되어 로그인 화면으로 이동됨 - 재로그인 시도")
    try:
        from credentials import PORTAL_PASSWORD, PORTAL_USER
    except ImportError:
        PORTAL_USER = PORTAL_PASSWORD = ""
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


def ensure_session_alive(page):
    """Raise a clear, actionable error the moment the page is sitting on any
    login/SSO screen, instead of letting the caller's own element-lookup
    time out with a confusing "button/link not found" a few seconds later
    (2026-08-28 incident: portal_update_pod.py had no such check and
    reported "POD 링크를 못 찾음" for what was actually an expired SSO
    session)."""
    _retry_login_if_expired(page)
    if "/supplier/login" in (page.url or ""):
        raise RuntimeError(
            "Portal 세션이 만료되어 재로그인이 필요합니다. "
            "자동화 Chrome 창(포트 9222)에서 Bloomberg Portal에 직접 로그인한 뒤 다시 실행하세요."
        )
    if "bsso.blpprofessional.com" in (page.url or "") or "/idp/startSSO" in (page.url or ""):
        raise RuntimeError(
            "Bloomberg SSO 세션이 만료되어 재로그인이 필요합니다 (B-Unit 인증 필요할 수 있음). "
            "자동화 Chrome 창(포트 9222)에서 직접 로그인을 완료한 뒤 다시 실행하세요."
        )


def navigate_to_delivery_via_search(page, delivery_num):
    """Reach a delivery's detail page the same way a person does by hand:
    Process Shipment search list -> click the OBD number link. Reliable
    where a raw goto() straight to the delivery URL is not (see module
    docstring)."""
    page.goto(SHIPMENT_SEARCH_URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(1000)
    ensure_session_alive(page)
    try:
        page.locator('input[name="quickFinder"]').first.fill(str(delivery_num), timeout=8000)
        page.get_by_role("button", name="Search").click(timeout=10000)
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)
    except Exception as exc:
        logger.warning("Process Shipment 검색 실패(무시하고 목록에서 바로 찾아봄): %s", exc)
    link = page.locator(f"a:has-text('{delivery_num}')").first
    link.wait_for(timeout=15000)
    link.click(timeout=10000)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    page.wait_for_timeout(1500)
    ensure_session_alive(page)
    target = _delivery_path(delivery_num)
    if target not in (page.url or ""):
        raise RuntimeError(
            f"Process Shipment 목록에서 Delivery# {delivery_num}으로 진입하지 못했습니다 "
            f"(현재 URL: {page.url})."
        )
    logger.info("Process Shipment 검색 경유로 Delivery# %s 진입 완료", delivery_num)


def open_delivery_page(pw, delivery_num, cdp_url=CDP_URL):
    """Connect to the shared automation Chrome and land on this delivery's
    detail page. Returns (browser, page).

    Fast path: an already-open tab already on this exact delivery is reused
    untouched - no navigation at all. Otherwise navigates through Process
    Shipment search + link click, never a raw goto() to the delivery URL.
    """
    browser = pw.chromium.connect_over_cdp(cdp_url)
    context = browser.contexts[0] if browser.contexts else None
    if context is None:
        raise RuntimeError("Chrome context를 찾지 못했습니다.")

    page = find_open_delivery_tab(context, delivery_num)
    if page is not None:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        return browser, page

    page = context.pages[-1] if context.pages else context.new_page()
    ensure_session_alive(page)
    navigate_to_delivery_via_search(page, delivery_num)
    return browser, page

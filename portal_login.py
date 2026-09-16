"""Open Bloomberg supplier portal SSO in Chrome.

The portal automation uses a dedicated Chrome profile so Chrome allows
remote-debugging on port 9222. After the first login, this profile keeps the
Bloomberg session for future automation runs.
"""

import subprocess
import sys
import time
from pathlib import Path

from portal_lock import portal_browser_lock


PORTAL_LOGIN_URL = "https://bsp.btogo.com/supplier/login"
PORTAL_SSO_URL = "https://bsso.blpprofessional.com/idp/startSSO.ping?PartnerSpId=bsp.bloomberg.com&ACSIdx=1"
PROFILE_DIR = Path.home() / "Documents" / "MJSuh" / "chrome_profiles" / "bloomberg_portal"
CHROME_PATHS = [
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


def find_chrome() -> Path:
    for path in CHROME_PATHS:
        if path.exists():
            return path
    raise FileNotFoundError("Chrome 실행 파일을 찾지 못했습니다.")


def open_portal_login():
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    chrome = find_chrome()
    subprocess.Popen(
        [
            str(chrome),
            "--remote-debugging-port=9222",
            f"--user-data-dir={PROFILE_DIR}",
            PORTAL_SSO_URL,
        ],
        close_fds=True,
    )


def _find_login_page(context):
    """Pick the actual SSO/login tab out of context.pages, instead of just
    trusting page ordering. Real bug (2026-08-12): the automation Chrome
    profile almost always has other real tabs already open (leftover
    delivery list/detail pages from earlier register_serial/pack_post/QR
    runs) by the time a login attempt happens - open_portal_login() opens
    the SSO URL as a NEW tab in that same already-running Chrome instance
    (Chrome's normal behavior when a URL is passed to an instance that's
    already up), so it does not land at a predictable position. Blindly
    using context.pages[-1] (the previous version of this function) could
    grab one of those unrelated leftover tabs instead - username/password
    would then get filled/clicked against a page with no login form at all,
    which explains exactly what was observed live: the real SSO tab sits
    there fully autofilled (user only ever needed to click Next by hand),
    but the automated fill+click loop never touches it and just burns the
    whole timeout. Same class of fix portal_register_serial.py's own
    _connect_page() already uses (match by URL, don't trust ordering)."""
    if context is None:
        return None
    for candidate in context.pages:
        url = candidate.url or ""
        if "bsso.blpprofessional.com" in url or "/supplier/login" in url:
            return candidate
    return context.pages[-1] if context.pages else None


def _connect_page(timeout_seconds=20):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None, None, None

    pw = sync_playwright().start()
    deadline = time.time() + timeout_seconds
    last_error = None
    while time.time() < deadline:
        try:
            browser = pw.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0] if browser.contexts else None
            page = _find_login_page(context)
            if page is not None:
                return pw, browser, page
            browser.close()
        except Exception as exc:
            last_error = exc
        time.sleep(1)
    try:
        pw.stop()
    except Exception:
        pass
    if last_error:
        print(f"Portal Chrome 연결 실패: {last_error}")
    return None, None, None


def _click_first(page, selectors, timeout=2000):
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.count():
                loc.click(timeout=timeout)
                return True
        except Exception:
            continue
    return False


def _find_logged_in_page(timeout_seconds=5):
    """Check whatever's already open in the automation Chrome profile
    (port 9222, same --user-data-dir open_portal_login() uses) for a tab
    that's already past login - bsp.btogo.com, not the /supplier/login page.
    Returns (pw, browser, page) for the first match, or (None, None, None)
    if nothing's open yet or nothing there is logged in. Does NOT open or
    navigate anything itself - purely a look-before-you-leap check."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None, None, None
    try:
        pw = sync_playwright().start()
    except Exception:
        return None, None, None
    try:
        browser = pw.chromium.connect_over_cdp("http://localhost:9222", timeout=timeout_seconds * 1000)
    except Exception:
        try:
            pw.stop()
        except Exception:
            pass
        return None, None, None
    context = browser.contexts[0] if browser.contexts else None
    if context is not None:
        for candidate in context.pages:
            url = candidate.url or ""
            if "bsp.btogo.com" in url and "/supplier/login" not in url:
                # Real bug (2026-08-14, order 67076875/ZRX): a tab's cached
                # URL keeps showing a logged-in-looking bsp.btogo.com path
                # long after the session has actually expired server-side -
                # the tab only gets redirected to /supplier/login on its
                # NEXT real navigation, not just from sitting open/idle.
                # Trusting the cached URL alone made this always report
                # "이미 로그인됨" in well under a second (every single run in
                # automation.log's history, including the one that then
                # failed): step 1's later page.goto() to the delivery page
                # hit the actually-expired session and got redirected to
                # /supplier/login, which surfaced downstream as a confusing
                # "포털 화면 검증 실패: Delivery#=..., Material=..." instead
                # of the real "not logged in" cause. Reload the candidate tab
                # here and re-check where it actually lands - a stale/expired
                # session redirects on reload just like it would on any other
                # real navigation, so this catches it before reporting success.
                try:
                    candidate.reload(wait_until="domcontentloaded", timeout=15000)
                except Exception:
                    continue
                reload_url = candidate.url or ""
                if "bsp.btogo.com" in reload_url and "/supplier/login" not in reload_url:
                    return pw, browser, candidate
    try:
        browser.close()
    except Exception:
        pass
    try:
        pw.stop()
    except Exception:
        pass
    return None, None, None


def ensure_portal_login(timeout_seconds=35):
    """Open the portal profile and advance the SSO flow when possible - but
    first checks whether some tab in the SAME automation Chrome profile is
    already logged in (e.g. the user used "Portal 로그인 (수동)" first, or a
    previous run's session is still valid) and, if so, does nothing else at
    all: no new tab, no autofill/auto-click attempt. Real bug fixed
    2026-08-03: this used to unconditionally call open_portal_login() (which
    opens a brand-new SSO tab in the same profile - Chrome's single-instance
    behavior means it lands in the SAME window as whatever's already open)
    and then ran the autofill/click-Next/BUIT-wait loop against THAT new tab
    every single "1. Serial 등록 & QR 인쇄" run, even right after the user had
    already logged in by hand - a redundant login attempt racing against the
    serial-registration step that immediately followed, against a different
    tab than the one the user was actually looking at."""
    pw, browser, page = _find_logged_in_page()
    if page is not None:
        print("이미 로그인된 Portal 탭을 찾았습니다 - 로그인 절차를 건너뜁니다.")
        try:
            browser.close()
        except Exception:
            pass
        try:
            pw.stop()
        except Exception:
            pass
        return True

    open_portal_login()
    with portal_browser_lock("login"):
        pw, browser, page = _connect_page(timeout_seconds=20)
        if page is None:
            return False

        try:
            try:
                page.wait_for_load_state("domcontentloaded", timeout=15000)
            except Exception:
                pass

            try:
                from credentials import PORTAL_PASSWORD, PORTAL_USER
            except ImportError:
                PORTAL_USER = ""
                PORTAL_PASSWORD = ""

            deadline = time.time() + timeout_seconds
            while time.time() < deadline:
                url = page.url or ""
                if "bsp.btogo.com" in url and "/supplier/login" not in url:
                    print("Portal 로그인 확인 완료")
                    return True

                try:
                    user_field = page.locator(
                        "input[name='pf.username'], input[name='username'], input[type='email'], input[type='text']"
                    ).first
                    pass_field = page.locator("input[type='password']").first
                    if user_field.count() and PORTAL_USER:
                        try:
                            if not (user_field.input_value(timeout=1000) or "").strip():
                                user_field.fill(PORTAL_USER, timeout=3000)
                        except Exception:
                            pass
                    if pass_field.count() and PORTAL_PASSWORD:
                        try:
                            if not (pass_field.input_value(timeout=1000) or "").strip():
                                pass_field.fill(PORTAL_PASSWORD, timeout=3000)
                        except Exception:
                            pass
                except Exception:
                    pass

                clicked = _click_first(
                    page,
                    [
                        "button:has-text('Next')",
                        "input[type='submit'][value*='Next']",
                        "button:has-text('Sign In')",
                        "button:has-text('Login')",
                        "button:has-text('Continue')",
                        "input[type='submit']",
                    ],
                )
                if not clicked:
                    try:
                        page.keyboard.press("Enter")
                    except Exception:
                        pass
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=5000)
                except Exception:
                    pass
                time.sleep(1)

            print("Portal 로그인 확인 실패: 인증 화면에서 멈춰 있습니다. Chrome에서 인증을 완료한 뒤 다시 실행하세요.")
            return False
        finally:
            try:
                browser.close()
            except Exception:
                pass
            try:
                pw.stop()
            except Exception:
                pass


def press_next():
    """Press Enter on the Bloomberg SSO page.

    The SSO page already has the saved username/password filled in Chrome, and
    Enter submits the same form as the Next button.
    """
    try:
        import win32com.client
    except ImportError:
        print("win32com이 없어 Next 자동 입력은 건너뜁니다.")
        return False

    shell = win32com.client.Dispatch("WScript.Shell")
    for title in ("Bloomberg", "Chrome"):
        try:
            if shell.AppActivate(title):
                time.sleep(0.5)
                shell.SendKeys("{ENTER}")
                print("Next 입력(Enter)을 보냈습니다.")
                return True
        except Exception:
            continue

    print("Chrome 창을 찾지 못해 Next 자동 입력은 건너뜁니다.")
    return False


def autofill_login_if_needed():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return False

    try:
        from credentials import PORTAL_PASSWORD, PORTAL_USER
    except ImportError:
        PORTAL_USER = ""
        PORTAL_PASSWORD = ""

    with portal_browser_lock("login_autofill"):
        with sync_playwright() as pw:
            for _ in range(10):
                try:
                    browser = pw.chromium.connect_over_cdp("http://localhost:9222")
                    context = browser.contexts[0] if browser.contexts else None
                    page = _find_login_page(context)
                    if page is None:
                        time.sleep(1)
                        continue

                    page.wait_for_load_state("domcontentloaded", timeout=10000)
                    user_field = page.locator(
                        "input[name='pf.username'], input[name='username'], input[type='email'], input[type='text']"
                    ).first
                    pass_field = page.locator("input[type='password']").first

                    if user_field.count() and PORTAL_USER:
                        try:
                            user_field.fill(PORTAL_USER, timeout=3000)
                        except Exception:
                            pass
                    if pass_field.count() and PORTAL_PASSWORD:
                        try:
                            pass_field.fill(PORTAL_PASSWORD, timeout=3000)
                        except Exception:
                            pass

                    try:
                        page.keyboard.press("Enter")
                        print("로그인 화면에 Enter를 보냈습니다.")
                        return True
                    except Exception:
                        return False
                except Exception:
                    time.sleep(1)
    return False


if __name__ == "__main__":
    if "--open-only" in sys.argv:
        # Just opens the automation Chrome profile (right debug port, right
        # user-data-dir) to the SSO start page and stops there - no
        # autofill, no auto-clicking Next, no BUIT wait-loop. For manually
        # logging in by hand while the automated login flow above is being
        # sorted out; the rest of the Portal automation (register_serial /
        # pack_post / download_labels) just connects to whatever's already
        # open in this same profile via CDP, so a fully manual login here
        # works exactly as well as the automated one once it lands on the
        # real logged-in portal page.
        open_portal_login()
        print("Chrome을 열었습니다 (자동화 프로필). 직접 로그인하세요 - Next 자동클릭/BUIT 대기는 하지 않습니다.")
        sys.exit(0)
    ok = ensure_portal_login()
    if not ok:
        if not autofill_login_if_needed():
            press_next()
        sys.exit(1)
    print("Bloomberg SSO 로그인 크롬을 열었습니다.")
    print("저장된 계정으로 Next를 누르고, BUIT가 뜨면 직접 인증하세요.")

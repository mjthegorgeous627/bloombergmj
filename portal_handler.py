"""
Bloomberg Vendor Portal - Process Shipment 라벨 자동 출력.

흐름:
  1. bsp.btogo.com 로그인 (세션 쿠키 재사용, 만료 시 재로그인)
  2. Warehouse > Process Shipment > Search
  3. 목록에서 배송번호(9xxxxxxx) 링크 수집
  4. 미출력 배송번호만: 배송 상세 페이지 → Download Labels → ZPL(PDF) 다운로드
  5. PDF 파싱 (QR + 이름 + 회사) → Zebra 라벨 출력
  6. 출력 완료 번호 portal_printed.json에 저장 (중복 방지)

사용법:
  python portal_handler.py         → 미출력 라벨 전체 출력
  python portal_handler.py --test  → 첫 번째 건만 출력 (테스트)

bunit 로그인 (2주마다):
  브라우저가 자동으로 열리므로 화면에서 직접 처리 후 Enter
"""

import os
import json
import time
import glob
import logging
import sys
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

logger = logging.getLogger(__name__)

PORTAL_LOGIN_URL    = "https://bsp.btogo.com/supplier/login"
PORTAL_SHIPMENT_URL = "https://bsp.btogo.com/supplier/warehouse/shipment/index"


def _wait(page, seconds):
    """느린 페이지 전환 대기 (networkidle + 추가 sleep)."""
    try:
        page.wait_for_load_state("networkidle", timeout=seconds * 1000 + 5000)
    except PlaywrightTimeout:
        pass
    time.sleep(max(0, seconds - 2))
DOWNLOADS_DIR      = r"C:\Users\bloomberg\Downloads"
SESSION_FILE  = os.path.join(os.path.dirname(__file__), "portal_session.json")
PRINTED_FILE  = os.path.join(os.path.dirname(__file__), "portal_printed.json")


# ── 출력 이력 관리 ────────────────────────────────────────────────────────────

def load_printed():
    if os.path.exists(PRINTED_FILE):
        with open(PRINTED_FILE, encoding='utf-8') as f:
            return set(json.load(f))
    return set()


def save_printed(printed):
    with open(PRINTED_FILE, 'w', encoding='utf-8') as f:
        json.dump(sorted(printed), f, ensure_ascii=False, indent=2)


# ── 세션 쿠키 관리 ───────────────────────────────────────────────────────────

def _load_session(context):
    if os.path.exists(SESSION_FILE):
        with open(SESSION_FILE, encoding='utf-8') as f:
            cookies = json.load(f)
        context.add_cookies(cookies)
        logger.info("세션 쿠키 로드")
        return True
    return False


def _save_session(context):
    cookies = context.cookies()
    with open(SESSION_FILE, 'w', encoding='utf-8') as f:
        json.dump(cookies, f, ensure_ascii=False, indent=2)
    logger.info("세션 쿠키 저장")


# ── 로그인 ────────────────────────────────────────────────────────────────────

def _is_logged_in(page):
    """WAREHOUSE 메뉴 존재 여부로 로그인 상태 확인."""
    try:
        return page.locator("text=WAREHOUSE").count() > 0
    except Exception:
        return False


def _login(page, user, password):
    """
    Bloomberg 포털 로그인.
    1) bsp.btogo.com/supplier/login → 주황색 LOG IN 버튼 클릭
    2) bsso.blpprofessional.com SSO → 아이디 입력 → Next → 비밀번호 입력 → Next
    3) bunit 등 추가 인증이 뜨면 브라우저에서 수동 처리 후 Enter
    """
    logger.info("로그인 페이지 접속...")
    page.goto(PORTAL_LOGIN_URL)
    _wait(page, 3)

    # 이미 로그인된 상태
    if _is_logged_in(page):
        logger.info("이미 로그인 상태")
        return True

    # 1단계: 주황색 LOG IN 버튼 클릭 → SSO 리디렉트
    try:
        page.click("text=LOG IN")
        _wait(page, 4)
        logger.info(f"SSO 페이지: {page.url}")
    except Exception as e:
        logger.error(f"LOG IN 버튼 클릭 실패: {e}")
        return False

    # 2단계: 아이디(username) 입력 → Next
    try:
        field = page.locator(
            "input[name='pf.username'], input[name='username'], input[type='email'], input[type='text']"
        ).first
        field.wait_for(timeout=10000)
        field.click()
        field.type(user, delay=50)          # 키입력 시뮬레이션으로 JS 검증 트리거
        time.sleep(1)
        field.press("Enter")                # Enter로 제출 (disabled 버튼 우회)
        _wait(page, 4)
        logger.info("아이디 입력 완료")
    except Exception as e:
        logger.error(f"아이디 입력 실패: {e}")
        return False

    # 3단계: 비밀번호 입력 → Next/Sign In
    try:
        pw_field = page.locator("input[type='password']").first
        pw_field.wait_for(timeout=10000)
        pw_field.click()
        pw_field.type(password, delay=50)   # 키입력 시뮬레이션
        time.sleep(1)
        pw_field.press("Enter")             # Enter로 제출
        _wait(page, 5)
        logger.info("비밀번호 입력 완료")
    except Exception as e:
        logger.error(f"비밀번호 입력 실패: {e}")
        return False

    # 4단계: bunit 등 추가 인증 대기
    if not _is_logged_in(page):
        logger.warning("추가 인증(bunit 등) 필요 - 브라우저에서 직접 완료 후 Enter")
        input("  인증 완료 후 Enter: ")
        _wait(page, 3)

    if _is_logged_in(page):
        logger.info("로그인 성공")
        return True

    logger.error(f"로그인 실패 (현재 URL: {page.url})")
    return False


# ── Process Shipment 배송번호 수집 ───────────────────────────────────────────

def _get_delivery_numbers(page):
    """
    Process Shipment Search 결과에서 배송번호(9로 시작) 링크 수집.
    각 행: 위줄 TJCK... (shipment ref), 아래줄 9xxxxxxx (delivery #) → 아래줄만 필요.
    """
    # Search 폼 대신 파라미터 URL로 직접 이동 (±7일 날짜 범위)
    from_date = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
    to_date   = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")
    search_url = (
        f"{PORTAL_SHIPMENT_URL}"
        f"?carrierCode=TJCK&fromDate={from_date}&plant=6507"
        f"&responsibleParty=&shipmentDirection=2&status=N&toDate={to_date}"
    )
    logger.info(f"Process Shipment 검색: {search_url}")
    page.goto(search_url)
    _wait(page, 5)

    # 디버깅: 전체 페이지 스크린샷 저장
    shot_path = os.path.join(os.path.dirname(__file__), "portal_debug.png")
    page.screenshot(path=shot_path, full_page=True)
    logger.info(f"스크린샷 저장: {shot_path} (URL: {page.url})")

    # 페이지 내 모든 링크 중 /warehouse/delivery/ 포함 링크 수집
    links = page.locator("a").all()
    delivery_nums = []
    for link in links:
        try:
            href = link.get_attribute("href") or ""
            if "warehouse/delivery/" not in href:
                continue
            num = href.rstrip("/").split("/")[-1]
            if num.isdigit() and num not in delivery_nums:
                delivery_nums.append(num)
        except Exception:
            continue

    logger.info(f"배송번호 {len(delivery_nums)}개: {delivery_nums}")
    return delivery_nums


# ── ZPL(PDF) 다운로드 ─────────────────────────────────────────────────────────

def _download_label(page, delivery_num):
    """
    배송 상세 페이지 → Download Labels 클릭 → SHIPLABEL_1.ZPL 다운로드.
    반환: 저장된 파일 경로 (실패 시 None)
    """
    url = f"https://bsp.btogo.com/supplier/warehouse/delivery/{delivery_num}"
    logger.info(f"  배송 상세 페이지: {url}")
    page.goto(url)
    _wait(page, 4)

    # Download Labels 버튼 클릭 → SHIPLABEL 링크 생성 대기
    try:
        page.click("button:has-text('Download Labels')")
        page.wait_for_selector(
            "a:has-text('SHIPLABEL'), text=SHIPLABEL_1.ZPL",
            timeout=12000,
        )
        time.sleep(1)
    except Exception as e:
        logger.error(f"  Download Labels 실패 ({delivery_num}): {e}")
        return None

    # SHIPLABEL 링크 다운로드
    try:
        with page.expect_download(timeout=15000) as dl_info:
            page.click("a:has-text('SHIPLABEL'), text=SHIPLABEL_1.ZPL")
        dl = dl_info.value

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"SHIPLABEL_{delivery_num}_{ts}.ZPL"
        save_path = os.path.join(DOWNLOADS_DIR, filename)
        dl.save_as(save_path)
        logger.info(f"  다운로드 완료: {filename}")
        return save_path
    except Exception as e:
        logger.error(f"  ZPL 다운로드 실패 ({delivery_num}): {e}")
        return None


# ── 메인 실행 ────────────────────────────────────────────────────────────────

def _process_delivery_list(page, delivery_nums, printed, force=False):
    """delivery_nums 목록 처리 공통 로직. force=True면 이미 출력된 번호도 재출력."""
    from zebra_handler import parse_pdf_label, print_label

    target = delivery_nums if force else [d for d in delivery_nums if d not in printed]
    logger.info(f"처리 대상: {len(target)}개 → {target}")

    for delivery_num in target:
        logger.info(f"=== 배송 {delivery_num} 처리 ===")

        label_path = _download_label(page, delivery_num)
        if not label_path:
            continue

        info = parse_pdf_label(label_path)
        logger.info(f"  이름: {info['name']}")
        logger.info(f"  회사: {info['company']}")
        logger.info(f"  QR:   {info['qr_data'][:60]}{'...' if len(info['qr_data']) > 60 else ''}")

        if not info['qr_data']:
            logger.warning(f"  QR 데이터 없음 - 건너뜀")
            continue

        ok = print_label(info['qr_data'], info['name'], info['company'])
        if ok:
            printed.add(delivery_num)
            save_printed(printed)
            logger.info(f"  라벨 출력 완료")
        else:
            logger.error(f"  라벨 출력 실패")


def _open_browser_with_session(pw):
    """브라우저 + 로그인 세션 준비. (page, context, browser) 반환."""
    try:
        from credentials import PORTAL_USER, PORTAL_PASSWORD
    except ImportError:
        logger.error("credentials.py에 PORTAL_USER, PORTAL_PASSWORD 없음")
        return None, None, None

    browser = pw.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True, viewport={"width": 1280, "height": 900})
    page = context.new_page()

    session_loaded = _load_session(context)
    if session_loaded:
        page.goto(PORTAL_SHIPMENT_URL)
        _wait(page, 4)
        if not _is_logged_in(page):
            logger.info("세션 만료 → 재로그인")
            session_loaded = False

    if not session_loaded:
        if not _login(page, PORTAL_USER, PORTAL_PASSWORD):
            browser.close()
            return None, None, None
        _save_session(context)

    return page, context, browser


def run_portal_labels(test_mode=False):
    """
    Bloomberg 포털 Process Shipment → 미출력 라벨 전체 자동 출력.
    test_mode=True: 첫 번째 건만 처리.
    """
    printed = load_printed()

    with sync_playwright() as pw:
        page, context, browser = _open_browser_with_session(pw)
        if page is None:
            return

        delivery_nums = _get_delivery_numbers(page)
        if test_mode and delivery_nums:
            delivery_nums = delivery_nums[:1]
            logger.info(f"[테스트] 첫 번째만 처리: {delivery_nums}")

        _process_delivery_list(page, delivery_nums, printed)
        _save_session(context)
        browser.close()

    logger.info("포털 라벨 처리 완료")


def run_portal_label_by_order(order_num: str, force=False):
    """
    특정 배송번호(오더넘버) 라벨만 출력.
    force=True: 이미 출력한 번호도 재출력.

    사용법:
      python portal_handler.py --order 9123456789
      python portal_handler.py --order 9123456789 --force
    """
    printed = load_printed()

    with sync_playwright() as pw:
        page, context, browser = _open_browser_with_session(pw)
        if page is None:
            return

        _process_delivery_list(page, [order_num], printed, force=force)
        _save_session(context)
        browser.close()

    logger.info("완료")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[logging.StreamHandler()],
    )

    args = sys.argv[1:]

    if "--order" in args:
        idx = args.index("--order")
        if idx + 1 >= len(args):
            print("사용법: python portal_handler.py --order <배송번호>")
            sys.exit(1)
        order = args[idx + 1]
        force = "--force" in args
        logger.info(f"오더 지정 모드: {order}  force={force}")
        run_portal_label_by_order(order, force=force)
    else:
        test = "--test" in args
        run_portal_labels(test_mode=test)

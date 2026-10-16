"""
포털 라벨 인쇄 단독 테스트.

사용법:
  1. start_chrome_portal.bat 으로 Chrome 열고 로그인
  2. python test_portal_print.py              → 오더 페이지 직접 이동 후 Enter
     python test_portal_print.py 92090849    → 배송번호 직접 지정

흐름:
  1. 이미 열린 Chrome에 CDP로 연결 (로그인 상태 그대로 유지)
  2. 배송 상세 페이지로 이동
  3. "Download Labels" → "SHIPLABEL 1.ZPL" 링크 클릭 → 다운로드
  4. ZPL(실제 PDF) 파싱 → QR + 이름 + 회사
  5. Zebra 라벨 출력
"""

import os
import sys
import time
import logging
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

from zebra_handler import parse_pdf_label, print_label

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

CDP_URL       = "http://localhost:9222"
DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
DOWNLOADS_DIR = r"C:\Users\bloomberg\Downloads"


def _download_label(page, delivery_num):
    """Download Labels 클릭 → SHIPLABEL 1.ZPL 링크 클릭 → 파일 저장."""

    logger.info("'Download Labels' 버튼 클릭...")
    try:
        page.click("button:has-text('Download Labels')", timeout=8000)
    except PlaywrightTimeout:
        logger.error("'Download Labels' 버튼을 찾지 못했습니다.")
        btns = page.locator("button").all_text_contents()
        logger.info(f"페이지 버튼 목록: {btns}")
        return None

    # SHIPLABEL 링크 대기 (텍스트에 SHIPLABEL 포함하는 <a>)
    logger.info("SHIPLABEL 링크 대기 중...")
    try:
        shiplabel = page.locator("a", has_text="SHIPLABEL").first
        shiplabel.wait_for(timeout=15000)
        logger.info(f"링크 확인: {shiplabel.text_content()}")
    except PlaywrightTimeout:
        logger.error("SHIPLABEL 링크가 나타나지 않음")
        links = [a.text_content() or a.get_attribute("href") for a in page.locator("a").all()[:30]]
        logger.info(f"페이지 링크 목록: {links}")
        return None

    # 다운로드
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"SHIPLABEL_{delivery_num}_{ts}.ZPL"
    save_path = os.path.join(DOWNLOADS_DIR, filename)

    try:
        with page.expect_download(timeout=15000) as dl_info:
            shiplabel.click()
        dl = dl_info.value
        dl.save_as(save_path)
        logger.info(f"다운로드 완료: {filename}")
        return save_path
    except Exception as e:
        logger.error(f"다운로드 실패: {e}")
        return None


def run():
    delivery_num_arg = sys.argv[1] if len(sys.argv) > 1 else None

    with sync_playwright() as pw:
        # ── 기존 Chrome에 CDP로 연결 시도 ──────────────────────────────────
        browser = None
        try:
            browser = pw.chromium.connect_over_cdp(CDP_URL)
            logger.info("기존 Chrome에 연결 성공 (CDP)")
            # 열린 컨텍스트/페이지 사용
            context = browser.contexts[0] if browser.contexts else None
            page = context.pages[0] if (context and context.pages) else None
            if page is None:
                page = context.new_page()
        except Exception as e:
            logger.warning(f"CDP 연결 실패: {e}")
            logger.info("새 Chrome 창을 열겠습니다 (로그인 필요할 수 있음)")
            browser = pw.chromium.launch(headless=False)
            context = browser.new_context(accept_downloads=True, viewport={"width": 1280, "height": 900})
            page = context.new_page()

        # ── 배송 페이지 이동 ───────────────────────────────────────────────
        if delivery_num_arg:
            delivery_num = delivery_num_arg
            url = f"{DELIVERY_BASE}{delivery_num}"
            logger.info(f"배송 페이지로 이동: {url}")
            page.goto(url)
            page.wait_for_load_state("networkidle", timeout=15000)
        else:
            logger.info("브라우저에서 오더 상세 페이지로 직접 이동하세요.")
            input("이동 완료 후 Enter: ")
            # URL에서 배송번호 추출
            current_url = page.url
            logger.info(f"현재 URL: {current_url}")
            if "/warehouse/delivery/" in current_url:
                delivery_num = current_url.rstrip("/").split("/")[-1]
            else:
                delivery_num = input("배송번호를 직접 입력하세요: ").strip()

        logger.info(f"배송번호: {delivery_num}")

        # ── 로그인 확인 ────────────────────────────────────────────────────
        time.sleep(1)
        if "login" in page.url.lower() or "sso" in page.url.lower():
            logger.warning("로그인 페이지로 리디렉트됨 - 로그인 후 Enter:")
            input("  로그인 완료 후 Enter: ")
            page.goto(f"{DELIVERY_BASE}{delivery_num}")
            page.wait_for_load_state("networkidle", timeout=15000)

        # ── Download Labels → ZPL 다운로드 ────────────────────────────────
        label_path = _download_label(page, delivery_num)
        if not label_path:
            logger.error("라벨 다운로드 실패. 종료합니다.")
            return

        # ── ZPL(PDF) 파싱 ──────────────────────────────────────────────────
        logger.info("라벨 파싱 중...")
        info = parse_pdf_label(label_path)
        logger.info(f"  이름:  {info['name']}")
        logger.info(f"  전화:  {info['phone']}")
        logger.info(f"  회사:  {info['company']}")
        logger.info(f"  QR:    {info['qr_data'][:80]}{'...' if len(info['qr_data']) > 80 else ''}")

        if not info['qr_data']:
            logger.error("QR 데이터 없음. 파일 확인: " + label_path)
            return

        # ── Zebra 출력 ─────────────────────────────────────────────────────
        ok = print_label(info['qr_data'], info['name'], info['company'])
        logger.info("라벨 출력 완료!" if ok else "라벨 출력 실패")


if __name__ == "__main__":
    run()

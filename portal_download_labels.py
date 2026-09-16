"""Download and optionally print Bloomberg Portal SHIPLABEL ZPL files."""

import argparse
import logging
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeout, sync_playwright

from portal_lock import portal_browser_lock

CDP_URL = "http://localhost:9222"
PORTAL_SHIPMENT_URL = "https://bsp.btogo.com/supplier/warehouse/shipment/index"
DELIVERY_BASE = "https://bsp.btogo.com/supplier/warehouse/delivery/"
DOWNLOADS_DIR = r"C:\Users\bloomberg\Downloads"
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "automation.log"

logger = logging.getLogger(__name__)


def _search_url(days_before=14, days_after=14):
    from_date = (datetime.now() - timedelta(days=days_before)).strftime("%Y-%m-%d")
    to_date = (datetime.now() + timedelta(days=days_after)).strftime("%Y-%m-%d")
    return (
        f"{PORTAL_SHIPMENT_URL}"
        f"?carrierCode=TJCK&fromDate={from_date}&plant=6507"
        f"&responsibleParty=&shipmentDirection=2&status=N&toDate={to_date}"
    )


def _connect_page(pw, delivery_num=None):
    """2026-09-16: 이 파일만 예전 "context.pages[-1]"(=지금 열려있는 마지막
    탭을 무조건 가로챈다) 방식을 그대로 쓰고 있었다 - portal_register_serial.py
    /portal_pack_post.py/portal_update_pod.py는 이미 이 delivery 관련 URL에
    있는 탭을 먼저 찾아 재사용하도록 고쳐졌는데(portal_update_pod.py의
    2026-08-28/08-31 실사고 주석 참고 - 무관한 탭을 가로채면 그 탭이
    진행 중이던 작업/사용자가 보던 화면이 날아간다) 이 파일만 빠져있었다."""
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0] if browser.contexts else None
    if context is None:
        raise RuntimeError("Chrome context를 찾지 못했습니다.")
    page = None
    if delivery_num:
        for candidate in context.pages:
            if DELIVERY_BASE + str(delivery_num) in candidate.url or PORTAL_SHIPMENT_URL in candidate.url:
                page = candidate
                break
    if page is None:
        page = context.pages[-1] if context.pages else context.new_page()
    return browser, page


def _body_snippet(page, limit=800):
    try:
        text = page.locator("body").inner_text(timeout=5000)
    except Exception:
        text = ""
    return " / ".join(line.strip() for line in text.splitlines() if line.strip())[:limit]


def _save_error_shot(page):
    try:
        page.screenshot(path=str(BASE_DIR / "portal_download_labels_error.png"), full_page=True, timeout=10000)
        logger.info("오류 화면 저장: portal_download_labels_error.png")
    except Exception as exc:
        logger.warning("오류 화면 저장 실패: %s", exc)


def _open_process_shipment(page, delivery_num):
    url = _search_url()
    logger.info("Process Shipment 이동: %s", url)
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    try:
        page.locator('input[name="quickFinder"]').first.fill(str(delivery_num), timeout=8000)
        page.get_by_role("button", name="Search").click(timeout=10000)
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(2500)
    except Exception as exc:
        logger.warning("Quick Finder search failed, continuing with default list: %s", exc)
    try:
        page.locator("text=Delivery List").wait_for(timeout=20000)
    except PlaywrightTimeout:
        logger.warning("Delivery List 텍스트 대기 시간 초과 - Search 버튼 클릭 시도")
        try:
            page.get_by_role("button", name="Search").click(timeout=10000)
        except Exception:
            page.locator("button:has-text('Search'), input[type='submit'][value='Search']").first.click(timeout=10000)
        page.wait_for_load_state("domcontentloaded", timeout=30000)
        page.wait_for_timeout(2500)
        try:
            page.locator("text=Delivery List").wait_for(timeout=20000)
        except PlaywrightTimeout:
            logger.warning("Search 후에도 Delivery List 텍스트 대기 시간 초과")
    try:
        page.locator(f"text={delivery_num}").first.wait_for(timeout=20000)
    except PlaywrightTimeout:
        logger.warning("Process Shipment 목록에서 Delivery# 대기 시간 초과: %s", delivery_num)

    page.wait_for_timeout(1500)
    logger.info("Process Shipment 현재 URL: %s", page.url)
    text = page.locator("body").inner_text(timeout=8000)
    if delivery_num not in text:
        logger.error("Process Shipment 목록에서 Delivery# %s 없음", delivery_num)
        logger.error("화면 일부: %s", _body_snippet(page))
        _save_error_shot(page)
        raise RuntimeError(f"Process Shipment 목록에서 Delivery# {delivery_num}을 찾지 못했습니다.")

    link = page.locator(f"a:has-text('{delivery_num}')").first
    link.click(timeout=10000)
    page.wait_for_load_state("domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)
    logger.info("Delivery 상세 진입 후 URL: %s", page.url)


def _download_all_labels(page, delivery_num):
    if f"/warehouse/delivery/{delivery_num}" not in page.url:
        target = DELIVERY_BASE + delivery_num
        logger.info("Delivery 상세 이동: %s", target)
        page.goto(target, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)

    text = page.locator("body").inner_text(timeout=8000)
    if delivery_num not in text:
        logger.error("Delivery 상세 화면에서 Delivery# %s 확인 실패", delivery_num)
        logger.error("현재 URL: %s", page.url)
        logger.error("화면 일부: %s", _body_snippet(page))
        _save_error_shot(page)
        raise RuntimeError(f"Delivery 상세 화면에서 {delivery_num}을 확인하지 못했습니다.")

    try:
        page.get_by_role("button", name="Download Labels").click(timeout=20000)
    except Exception:
        page.locator("button:has-text('Download Labels')").first.click(timeout=20000)
    logger.info("Download Labels 클릭 완료")

    label_locator = page.locator("a:has-text('SHIPLABEL'), button:has-text('SHIPLABEL')")
    try:
        label_locator.first.wait_for(timeout=20000)
    except PlaywrightTimeout as exc:
        logger.error("SHIPLABEL 링크 대기 실패. 현재 URL: %s", page.url)
        logger.error("화면 일부: %s", _body_snippet(page))
        _save_error_shot(page)
        raise RuntimeError("SHIPLABEL 링크가 나타나지 않았습니다.") from exc

    count = label_locator.count()
    if count == 0:
        raise RuntimeError("다운로드할 SHIPLABEL 링크가 없습니다.")

    saved = []
    for index in range(count):
        label_text = (label_locator.nth(index).inner_text(timeout=5000) or f"SHIPLABEL_{index + 1}.ZPL").strip()
        safe_label = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in label_text)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{delivery_num}_{index + 1}_{timestamp}_{safe_label}"
        if not filename.lower().endswith(".zpl"):
            filename += ".ZPL"
        save_path = os.path.join(DOWNLOADS_DIR, filename)

        with page.expect_download(timeout=30000) as dl_info:
            label_locator.nth(index).click(timeout=10000)
        download = dl_info.value
        download.save_as(save_path)
        saved.append(save_path)
        logger.info("라벨 다운로드 완료: %s", save_path)
        page.wait_for_timeout(500)
    return saved


def _print_files(paths):
    from print_zpl_file import is_pdf, print_as_pdf, print_as_zpl

    ok_all = True
    for path in paths:
        logger.info("라벨 인쇄 시작: %s", path)
        ok = print_as_pdf(path) if is_pdf(path) else print_as_zpl(path)
        ok_all = ok_all and ok
    return ok_all


# 2026-08-27 실제 장애 (오더 67082500 / delivery 92165893): Process Shipment를
# 경유한 재진입까지는 성공했는데, 그 직후 Download Labels 클릭이 "Target page,
# context or browser has been closed"로 실패 - 자동화가 잡고 있던 공유 Chrome
# 탭이 실행 도중 (사용자 조작이든 포털/브라우저 쪽 문제든) 닫혀버린 케이스.
# 기존 코드는 이 경우를 구분하지 않고 그대로 실패 처리했음. 이 클래스의 에러만
# 골라 재연결(_connect_page) 후 전체 플로우를 한 번 더 시도한다 - 무한 재시도는
# 아니고 딱 1회만.
def _is_target_closed(exc):
    msg = str(exc)
    return "has been closed" in msg or "Target closed" in msg


def _attempt_download(page, delivery_num):
    # 2026-09-01: delivery URL로 바로(콜드) goto하면 Packing Post가 실제로 일어난
    # 탭과 다른 탭을 잡거나, 서버 반영이 덜 된 상태로 로드되어 시리얼 입력 화면이
    # 뜨는 경우가 있었다 (오더 7849305 실사용 확인). Process Shipment 목록을 거쳐
    # 델리버리 상세로 재진입하는 경로는 항상 Packed 상태의 올바른 화면으로 들어가
    # 므로, 이를 최초 시도부터 기본 경로로 사용한다 (기존의 "실패 시에만 우회"
    # 방식 폐기).
    _open_process_shipment(page, delivery_num)
    return _download_all_labels(page, delivery_num)


def run(delivery_num, do_print=False):
    with portal_browser_lock(f"download_labels:{delivery_num}"):
        with sync_playwright() as pw:
            for attempt in range(2):
                browser, page = _connect_page(pw, delivery_num)
                try:
                    paths = _attempt_download(page, delivery_num)
                    if do_print and not _print_files(paths):
                        raise RuntimeError("라벨 인쇄 중 실패가 있었습니다.")
                    try:
                        page.screenshot(path=str(BASE_DIR / "portal_download_labels_result.png"), full_page=True, timeout=10000)
                    except Exception as exc:
                        logger.warning("결과 스크린샷 저장 실패(무시): %s", exc)
                    return paths
                except Exception as exc:
                    if _is_target_closed(exc) and attempt == 0:
                        logger.warning(
                            "공유 Chrome 탭이 처리 도중 닫힌 것으로 보임 - 재연결 후 처음부터 재시도: %s", exc
                        )
                        continue
                    raise
                finally:
                    try:
                        browser.close()
                    except Exception:
                        pass


def main():
    parser = argparse.ArgumentParser(description="Download Bloomberg Portal SHIPLABEL files.")
    parser.add_argument("--delivery", required=True)
    parser.add_argument("--print", action="store_true", help="Print all downloaded label files.")
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
        logger.info("[QR 라벨 다운로드 시작] delivery=%s print=%s", args.delivery, getattr(args, "print"))
        paths = run(args.delivery, do_print=getattr(args, "print"))
    except Exception as exc:
        logger.error("실패: %s", exc)
        return 1

    print("Downloaded labels:")
    for path in paths:
        print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())

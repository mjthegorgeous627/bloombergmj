"""Pack a Bloomberg Portal shipment and optionally post it.

Default mode is safe: it prepares/validates packing but does not click Post.
Use --post only when the user explicitly wants to complete the shipment.
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


def _connect_page(pw, delivery_num):
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    context = browser.contexts[0] if browser.contexts else None
    if context is None:
        raise RuntimeError("Chrome context를 찾지 못했습니다.")

    page = None
    for candidate in context.pages:
        if f"/warehouse/delivery/{delivery_num}" in candidate.url:
            page = candidate
            break
    if page is None:
        page = context.pages[-1] if context.pages else context.new_page()
        page.goto(DELIVERY_BASE + delivery_num)

    page.wait_for_load_state("domcontentloaded", timeout=15000)
    time.sleep(1)
    return browser, page


def _page_values_text(page):
    body = page.locator("body").inner_text(timeout=5000)
    values = page.evaluate(
        """
        () => Array.from(document.querySelectorAll('input, textarea, select'))
          .map(el => el.value || el.textContent || '')
          .filter(Boolean)
          .join('\\n')
        """
    )
    return body + "\n" + values


def _validate_shipment_page(page, delivery_num, material):
    text = _page_values_text(page)
    if "Shipment" not in text or "Items to be packed" not in text:
        raise RuntimeError("Shipment 화면이 아닙니다. 먼저 Run ShipERP 화면까지 진입해야 합니다.")
    if _norm(delivery_num) not in _norm(text):
        raise RuntimeError(f"Delivery# {delivery_num}을 화면에서 확인하지 못했습니다.")
    if _norm(material) not in _norm(text):
        raise RuntimeError(f"Material {material}을 화면에서 확인하지 못했습니다.")
    logger.info("Shipment 화면 검증 완료: delivery=%s material=%s", delivery_num, material)


def _get_pack_state(page, material):
    return page.evaluate(
        """
        ({ material }) => {
          const inputs = Array.from(document.querySelectorAll('input'));
          const materialFound = inputs.some(el => (el.value || '').includes(material))
            || (document.body.innerText || '').includes(material);
          const packQty = inputs.find(el => /items\\.\\[\\d+\\]\\.packQuantity/.test(el.name || ''));
          const weights = inputs
            .filter(el => /handlingUnits\\.\\[\\d+\\]\\.grossWeight/.test(el.name || ''))
            .map(el => ({ name: el.name, value: el.value || '', disabled: !!el.disabled }));
          const packedCount = weights.length;
          const remainingText = (document.body.innerText || '').match(/Remaining\\s*Qty[\\s\\S]{0,120}/i)?.[0] || '';
          return {
            materialFound,
            packQtyName: packQty ? packQty.name : '',
            packQtyValue: packQty ? packQty.value : '',
            packQtyDisabled: packQty ? !!packQty.disabled : null,
            packedCount,
            weights,
            remainingText,
          };
        }
        """,
        {"material": str(material)},
    )


def _set_pack_qty(page, material, qty):
    state = _get_pack_state(page, material)
    if state["packedCount"] > 0:
        logger.info("이미 Packed Items가 있어 Pack Qty 입력 생략")
        return
    if not state["packQtyName"]:
        raise RuntimeError("Pack Qty 입력칸을 찾지 못했습니다.")
    selector = f'input[name="{state["packQtyName"]}"]'
    if page.locator(selector).is_disabled():
        raise RuntimeError("Pack Qty 입력칸이 비활성화되어 있고 Packed Items도 없습니다.")
    page.locator(selector).fill(str(qty), timeout=8000)
    logger.info("Pack Qty 입력 완료: %s", qty)


def _set_pack_qty_for_material(page, material, qty):
    result = page.evaluate(
        """
        ({ material }) => {
          const visible = (el) => {
            if (!el) return false;
            const r = el.getBoundingClientRect();
            const s = window.getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
          };
          const mat = Array.from(document.querySelectorAll('input[name$=".material"]'))
            .find(el => visible(el) && (el.value || '').trim() === String(material));
          if (!mat) return { ok: false, error: `material ${material} not found` };
          const match = (mat.name || '').match(/items\\.\\[(\\d+)\\]\\./);
          if (!match) return { ok: false, error: `item index not found from ${mat.name}` };
          return { ok: true, index: Number(match[1]) };
        }
        """,
        {"material": str(material)},
    )
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Pack Qty item lookup failed"))
    selector = f'input[name="items.[{int(result["index"])}].packQuantity"]'
    if page.locator(selector).is_disabled():
        logger.info("Pack Qty 이미 비활성화됨: material=%s", material)
        return
    page.locator(selector).fill(str(qty), timeout=8000)
    logger.info("Pack Qty 입력 완료: material=%s qty=%s", material, qty)


def _set_pack_qty_by_order(page, item_index, material, qty):
    inputs = page.locator('input[name$=".packQuantity"]')
    count = inputs.count()
    if item_index >= count:
        raise RuntimeError(f"Pack Qty 입력칸 순번 {item_index + 1}을 찾지 못했습니다. 현재 {count}개")
    target = inputs.nth(item_index)
    if target.is_disabled():
        logger.info("Pack Qty 이미 비활성화됨: index=%s material=%s", item_index, material)
        return
    target.fill(str(qty), timeout=8000)
    logger.info("Pack Qty 순서 입력 완료: index=%s material=%s qty=%s", item_index, material, qty)


def _get_pack_rows(page, material):
    """All visible 'Items to be packed' rows for a material, in document
    order, each with its own Remaining Qty capacity.

    Real incident (2026-08-21, order 67055178/ZRX, CISCO router qty=2):
    Bloomberg Portal rendered this material as TWO separate rows (each
    Delivery/Remaining Qty=1) instead of one combined row - same thing
    portal_register_serial.py's _get_material_rows() hit one step earlier
    in the same delivery (see its comment). `_set_pack_qty()` filled the
    whole combined qty (2) into just the first row, which only had Remaining
    Qty=1, and Portal rejected it: "Pack quantity greater than remaining
    quantity". Reading each row's own Remaining Qty up front lets the caller
    split an item's qty across as many rows as Portal actually gave it."""
    rows = page.evaluate(
        """
        ({ material }) => {
          const visible = (el) => {
            if (!el) return false;
            const r = el.getBoundingClientRect();
            const s = window.getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
          };
          const inputs = Array.from(document.querySelectorAll('input[name$=".packQuantity"]')).filter(visible);
          return inputs.map(inp => {
            const m = (inp.name || '').match(/items\\.\\[(\\d+)\\]\\./);
            if (!m) return null;
            const row = inp.closest('tr');
            const text = row ? (row.innerText || '') : '';
            if (!text.includes(String(material))) return null;
            const nums = (text.match(/\\d+(?:\\.\\d+)?/g) || []).map(Number);
            // columns in order: Delivery Qty, Pick Qty, Remaining Qty[, Pack Qty if already filled]
            const remaining = nums.length >= 3 ? nums[2] : (nums.length ? nums[nums.length - 1] : 1);
            return { index: Number(m[1]), remainingQty: remaining, disabled: inp.disabled };
          }).filter(Boolean);
        }
        """,
        {"material": str(material)},
    )
    return rows


def _fill_pack_qty(page, items):
    """Fill Pack Qty for each item, splitting its qty across however many
    rows Portal actually rendered for that material (see _get_pack_rows()'s
    comment) instead of assuming one row per item / one combined row per
    material."""
    consumed = {}
    for item in items:
        material = item["material"]
        qty = int(float(item.get("qty") or 1))
        start = consumed.get(material, 0)
        rows = _get_pack_rows(page, material)[start:]
        if not rows:
            raise RuntimeError(f"Pack Qty 입력칸을 찾지 못했습니다: material={material}")
        remaining_qty = qty
        rows_used = 0
        for row in rows:
            if remaining_qty <= 0:
                break
            rows_used += 1
            if row["disabled"]:
                logger.info("Pack Qty 이미 비활성화됨: material=%s row=%s", material, row["index"])
                continue
            take = min(remaining_qty, max(1, int(row["remainingQty"] or 1)))
            selector = f'input[name="items.[{row["index"]}].packQuantity"]'
            page.locator(selector).fill(str(take), timeout=8000)
            logger.info("Pack Qty 입력 완료: material=%s row=%s qty=%s", material, row["index"], take)
            remaining_qty -= take
        consumed[material] = start + rows_used
        if remaining_qty > 0:
            raise RuntimeError(
                f"Material {material}: Pack Qty {qty} 중 {remaining_qty}만큼 채울 행이 부족합니다 "
                f"(사용 가능 행 {len(rows)}개)."
            )


def _list_pack_options(page, query="pack"):
    page.locator('input[role="combobox"]').nth(0).click(timeout=8000)
    page.locator('input[role="combobox"]').nth(0).fill(query, timeout=8000)
    page.wait_for_timeout(1000)
    options = page.evaluate(
        """
        () => Array.from(document.querySelectorAll('[role="option"], .css-10wo9uf-option, .css-d7l1ni-option'))
          .map((el, index) => ({ index, text: (el.innerText || el.textContent || '').trim() }))
          .filter(x => x.text)
        """
    )
    if not options:
        # React-select often renders options as plain divs near the control.
        options = page.evaluate(
            """
            () => Array.from(document.querySelectorAll('div'))
              .map((el, index) => ({ index, text: (el.innerText || el.textContent || '').trim() }))
              .filter(x => /^BOX-|^PACK/i.test(x.text))
              .slice(0, 20)
            """
        )
    logger.info("Packing Material 후보 %s개: %s", len(options), [o["text"] for o in options[:8]])
    return options


def _select_pack_option_by_index(page, option_index=0):
    options = page.locator('[role="option"], .css-10wo9uf-option, .css-d7l1ni-option')
    try:
        options.nth(option_index).click(timeout=5000)
        return
    except Exception:
        pass
    # Fallback: keyboard select from current react-select menu.
    for _ in range(option_index):
        page.keyboard.press("ArrowDown")
    page.keyboard.press("Enter")


def _click_pack_button(page):
    clicked = page.evaluate(
        """
        () => {
          const visible = (el) => {
            const r = el.getBoundingClientRect();
            const s = window.getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
          };
          const combos = Array.from(document.querySelectorAll('input[role="combobox"]')).filter(visible);
          const buttons = Array.from(document.querySelectorAll('button.btn-secondary')).filter(visible);
          if (!combos.length || !buttons.length) return false;
          const comboBox = combos[0].getBoundingClientRect();
          const comboMidY = comboBox.y + comboBox.height / 2;
          const candidates = buttons
            .map((btn) => ({ btn, r: btn.getBoundingClientRect() }))
            .filter(x => x.r.y < comboBox.y + 80 && x.r.x > comboBox.x)
            .sort((a, b) => Math.abs((a.r.y + a.r.height / 2) - comboMidY) - Math.abs((b.r.y + b.r.height / 2) - comboMidY));
          const target = (candidates[0] || buttons[0]);
          target.btn.click();
          return true;
        }
        """
    )
    if not clicked:
        raise RuntimeError("packing 버튼을 찾지 못했습니다.")
    page.wait_for_timeout(1500)
    logger.info("packing 버튼 클릭 완료")


def _wait_for_packed_items(page, timeout_ms=8000, poll_ms=500):
    """packing 버튼 클릭 후 Packed Item(HU weight 입력칸)이 나타날 때까지 폴링.

    실제 사고 (2026-08-31 키보드/시리얼 건, 2026-09-01 오더 551060846 no-serial
    2건): 클릭 직후 _get_pack_state()를 한 번만 즉시 확인했더니, Portal 서버가
    아직 Packed Item/HU를 만들지 않은 찰나를 잡아 "Packed Items가 생성되지
    않았습니다"로 오판했다. no_serial 여부와 무관하게 재현되는 순수 타이밍
    문제였다 - 고정 지연 1회 체크 대신 짧은 간격으로 최대 timeout_ms까지
    재확인한다."""
    deadline = time.time() + timeout_ms / 1000
    state = _get_pack_state(page, "")
    while state["packedCount"] == 0 and time.time() < deadline:
        page.wait_for_timeout(poll_ms)
        state = _get_pack_state(page, "")
    return state


def _attempt_pack_with_fallback(page, items, option_index=0, max_attempts=5):
    """Fill Pack Qty and try packing-material options in order until one
    actually produces a Packed Item, instead of trusting option_index blindly.

    Real incident (2026-09-01, order 551060846/delivery 92168758, no-serial
    2-item case): index 0 ("BOX-10045448 BOX H-PACKAGING MATERIAL") turned
    out to be a broken Portal master-data record - selecting it and clicking
    pack calls Portal's own material-lookup API, which 404s
    (materials/BOX-10045448%20BOX%20H, confirmed via the browser's network
    log), so the pack action silently no-ops no matter how many times or how
    long you wait/re-click *that same option*. Picking any other,
    correctly-formatted option (e.g. "BOX-10036467-BOX-10036467 (CR COMBO
    PACK)") worked immediately. This tries option_index first, then the
    remaining options in order, so a future bad default option degrades to
    "try the next one" instead of hard-failing the whole order."""
    _fill_pack_qty(page, items)
    options = _list_pack_options(page, "pack")
    if not options:
        raise RuntimeError("Packing Material 후보를 찾지 못했습니다.")
    order = [option_index] + [i for i in range(len(options)) if i != option_index]
    for attempt_num, idx in enumerate(order[:max_attempts]):
        if idx >= len(options):
            continue
        _select_pack_option_by_index(page, idx)
        _click_pack_button(page)
        state = _wait_for_packed_items(page)
        if state["packedCount"] > 0:
            if attempt_num > 0:
                logger.info("Packing Material 후보 %s번째(%s)로 성공 - 앞선 후보 실패", idx, options[idx]["text"])
            return
        logger.info("Packing Material 후보 %s번째(%s) 실패 - 다음 후보 시도", idx, options[idx]["text"])
        options = _list_pack_options(page, "pack")
    raise RuntimeError("Packed Items가 생성되지 않았습니다 (Packing Material 후보를 모두 시도했지만 실패).")


def _ensure_weight(page, weight="1"):
    state = _wait_for_packed_items(page)
    if state["packedCount"] == 0:
        raise RuntimeError("Packed Items가 생성되지 않았습니다.")
    for item in state["weights"]:
        selector = f'input[name="{item["name"]}"]'
        current = str(item["value"]).strip()
        if current in ("", "0", "0.0", "0.00"):
            page.locator(selector).fill(str(weight), timeout=8000)
            logger.info("Weight 보정: %s -> %s", item["name"], weight)
        else:
            logger.info("Weight 확인: %s=%s", item["name"], current)


def _collect_errors(page):
    text = page.locator("body").inner_text(timeout=5000)
    lines = []
    for line in text.splitlines():
        lower = line.lower()
        if any(word in lower for word in ["error", "invalid", "must", "required", "failed", "cannot"]):
            lines.append(line.strip())
    return [line for line in lines if line][:10]


def prepare_pack(delivery_num, material, qty, post=False, option_index=0, weight="1", items=None):
    with portal_browser_lock(f"pack_post:{delivery_num}"):
        with sync_playwright() as pw:
            browser, page = _connect_page(pw, delivery_num)
            try:
                items = items or [{"material": material, "qty": qty}]
                for item in items:
                    _validate_shipment_page(page, delivery_num, item["material"])
                state = _get_pack_state(page, material)
                if state["packedCount"] == 0:
                    _attempt_pack_with_fallback(page, items, option_index)
                else:
                    logger.info("기존 Packed Items 재사용: %s개", state["packedCount"])

                _ensure_weight(page, weight)

                if post:
                    post_button = page.get_by_role("button", name="Post")
                    if post_button.count() == 0:
                        logger.info("Post 버튼 없음 - 이미 Post 완료된 상태로 보고 생략")
                    elif post_button.first.is_disabled(timeout=2000):
                        logger.info("Post 버튼 비활성화 - 이미 Post 완료 또는 진행 불가 상태로 보고 생략")
                    else:
                        post_button.first.click(timeout=10000)
                        page.wait_for_timeout(3000)
                        errors = _collect_errors(page)
                        if errors:
                            raise RuntimeError("Post 후 에러 감지: " + " / ".join(errors))
                        logger.info("Post 클릭 완료")
                else:
                    logger.info("Dry-run: Post는 누르지 않았습니다.")
                page.screenshot(path="portal_pack_post_result.png", full_page=True)
            finally:
                try:
                    browser.close()
                except Exception:
                    pass


def main():
    parser = argparse.ArgumentParser(description="Prepare Bloomberg Portal packing and optionally post.")
    parser.add_argument("--delivery", required=True)
    parser.add_argument("--material", required=True)
    parser.add_argument("--qty", required=True)
    parser.add_argument("--items-json", default="")
    parser.add_argument("--option-index", type=int, default=0)
    parser.add_argument("--weight", default="1")
    parser.add_argument("--post", action="store_true", help="Actually click Post after packing validation.")
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
        items = None
        if args.items_json:
            try:
                items = json.loads(args.items_json)
            except json.JSONDecodeError as exc:
                logger.warning("items-json 파싱 실패 → 단일 material 방식으로 진행: %s", exc)
        prepare_pack(args.delivery, args.material, args.qty, args.post, args.option_index, args.weight, items)
    except Exception as exc:
        logger.error("실패: %s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

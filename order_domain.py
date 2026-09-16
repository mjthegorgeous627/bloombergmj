"""
공통 오더 도메인 로직 - 여러 스크립트에 거의 그대로 중복돼 있던 것들을 한 곳으로
모음 (2026-09-08, 퀵 배송 문서 기능 추가하면서 quick_courier.py도 이 매핑이
필요해져서 세 번째 복사본을 만드는 대신 통합).

- 품목명 매핑(PRODUCT_MAP/get_product_name): 예전엔 quick_handler.py와
  kakao_handler.py에 거의 동일하게 복사돼 있었음.
- 배송/회수 판정(infer_item_type/resolve_item_type): 예전엔
  workbench_app.py에만 있었음 - quick_courier.py가 workbench.db 항목의
  방향을 판정할 때 workbench_app.py 전체(http.server 앱)를 import하지
  않도록 여기로 분리.
"""

import re

# ── 품목명 매핑 ──────────────────────────────────────────────────────────────

PRODUCT_MAP = [
    (['KEYBOARD', '5'],  '키보드5'),
    (['KEYBOARD'],       '일반키보드'),
    (['MONITOR'],        '모니터'),
    (['PC'],             'PC'),
    (['ROUTER'],         '라우터'),
    (['SERVER'],         '서버'),
]

US_KEYBOARD_MATERIALS = {'10045246'}


def get_product_name(description, material=''):
    desc = str(description).upper()
    material = str(material or '').strip()
    if material in US_KEYBOARD_MATERIALS or ('KEYBOARD' in desc and re.search(r'\bUS\b', desc)):
        return 'US일반키보드'
    for keywords, name in PRODUCT_MAP:
        if all(k in desc for k in keywords):
            return name
    return str(description).split()[0] if description else ''


# ── 배송/회수 방향 판정 ──────────────────────────────────────────────────────

def infer_item_type(order_type, idx, total):
    """
    배송/회수 fallback heuristic, used only for legacy rows imported before
    load_today_from_excel started capturing the real per-row marker from
    column A (item_type). ZRE-only orders are pure pickups, everything else
    defaults to delivery, and ZRX (exchange - genuinely carries both) is
    split in half just so the layout has something to show. Do not trust
    ZRX's split for real portal actions.
    """
    order_type = (order_type or "").upper()
    if order_type == "ZRE":
        return "회수"
    if order_type == "ZRX" and total > 1:
        return "회수" if idx >= (total + 1) // 2 else "배송"
    return "배송"


def resolve_item_type(order, item, idx, total):
    """2026-08-26 fix: this used to only trust a stored item_type of exactly
    '배송' or '회수', silently discarding anything else (Bloomberg/Delayed/
    회수 Delayed/any custom type - see allTypeOptions() on the client) back
    to the infer_item_type() heuristic on every reload. That heuristic exists
    only as a fallback for legacy rows with no item_type recorded at all
    (see its own docstring) - it was never meant to override a real,
    explicitly-set value. Real symptom (사용자 신고, 8/26 메모 행
    "Scrap,Resale_26August"): Type dropdown set to Bloomberg kept reverting
    to 배송 on every page reload. Now any non-blank stored value wins outright;
    infer_item_type() only fires for rows where item_type is genuinely empty."""
    real = str(item.get("item_type") or "").strip()
    if real:
        return real
    return infer_item_type(order.get("order_type"), idx, total)

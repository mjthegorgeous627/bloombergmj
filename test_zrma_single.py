"""
66999156 단일 오더 테스트 - ZRMA_Q RLKR (세션2).
processed 체크 없이 바로 Excel에 기록.
"""
import logging
import sys
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s',
                    handlers=[logging.StreamHandler(sys.stdout)])

from sap_handler import get_sap_session, get_ship_to_address_zrma, get_text_content, parse_extra_orders, parse_contact_from_text, parse_memo_for_display, ZRMA_MENU_TEXTS
from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order, get_items_from_zrma_order, build_excel_rows_zrma
from excel_handler import write_orders_to_excel

TARGET = '66999156'

session = get_sap_session(2)
print(f"화면: {session.findById('wnd[0]').Text}")

# 목록에서 해당 오더 찾기
rows = get_all_rows_from_zrma(session)
order_map = group_zrma_by_order(rows)

if TARGET not in order_map:
    print(f"[ERROR] {TARGET} 목록에 없음. 현재 오더: {list(order_map.keys())}")
    sys.exit(1)

info = order_map[TARGET]
print(f"오더 정보: {info}")

# 오더 진입
if not navigate_to_zrma_order(session, info['grid_idx']):
    print("[ERROR] 오더 진입 실패")
    sys.exit(1)

# 주소
address = get_ship_to_address_zrma(session)
print(f"주소: {address}")

# 텍스트
text = get_text_content(session, menu_id=ZRMA_MENU_TEXTS)
all_extra = parse_extra_orders(text) if text else []
extra_orders = [eo for eo in all_extra if TARGET not in eo]
text_contact = parse_contact_from_text(text) if text else {'found': False}
memo = parse_memo_for_display(text) if text else ""
if text_contact.get('found') and not address.get('phone'):
    address['phone'] = text_contact['phone']

# 아이템
items = get_items_from_zrma_order(session, TARGET, info['order_type'])
print(f"\n아이템 수: {len(items)}")
for it in items:
    print(f"  {it}")

if not items:
    print("[ERROR] 아이템 없음")
    session.findById("wnd[0]").sendVKey(3)
    sys.exit(1)

# Excel 행 생성
excel_rows = build_excel_rows_zrma(items, address, extra_orders, memo)
print(f"\nExcel 행 수: {len(excel_rows)}")
for r in excel_rows:
    print(f"  {r}")

# Excel 기록
write_orders_to_excel(excel_rows)
print(f"\n[완료] {TARGET} Excel 기록 완료")

# 목록으로 복귀
session.findById("wnd[0]").sendVKey(3)

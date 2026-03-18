"""
ZRMA_Q 핸들러 - RLKR (회수 오더, 키보드 제외) 및 Q2 (키보드 회수) 처리.
"""

import time
import re
import logging
from datetime import date
import calendar

from sap_handler import (
    run_transaction,
    get_ship_to_address_zrma,
    get_text_content,
    parse_extra_orders,
    parse_contact_from_text,
    ZRMA_MENU_TEXTS,
)

logger = logging.getLogger(__name__)

# 확인 필요한 element ID (discover_sap.py로 확인 후 업데이트)
ZRMA_GRID_PATH = "wnd[0]/usr/cntlCUST_CONT/shellcont/shell"


def _end_of_next_month_str():
    """다음달 말일 (SAP 날짜 입력 형식)."""
    today = date.today()
    year = today.year + 1 if today.month == 12 else today.year
    month = 1 if today.month == 12 else today.month + 1
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, last_day).strftime('%m/%d/%Y')


def _end_of_year_str():
    """올해 말일."""
    return date(date.today().year, 12, 31).strftime('%m/%d/%Y')


def navigate_to_zrma_q(session, variant, date_mode='next_month'):
    """
    ZRMA_Q 진입.
    variant: 'rlkr' 또는 'q2'
    date_mode: 'next_month' (RLKR) 또는 'year_end' (Q2)
    """
    logger.info(f"ZRMA_Q 진입 중 (variant={variant})...")
    run_transaction(session, "ZRMA_Q")
    time.sleep(1.5)

    # Get Variants: Shift+F5
    session.findById("wnd[0]").sendVKey(25)
    time.sleep(1)

    # 팝업: Variant 입력
    for fid in ["wnd[1]/usr/txtV-LOW", "wnd[1]/usr/txtVARIANT"]:
        try:
            session.findById(fid).text = variant.upper()
            logger.info(f"Variant 입력 ({fid})")
            break
        except Exception:
            continue

    # Created By 공백으로
    for fid in ["wnd[1]/usr/txtENAME-LOW", "wnd[1]/usr/txtUSER-LOW"]:
        try:
            session.findById(fid).text = ""
            logger.info(f"Created By 공백 ({fid})")
            break
        except Exception:
            continue

    # F8: Execute (variant 선택)
    session.findById("wnd[1]").sendVKey(8)
    time.sleep(1.5)

    # Creation Date to 필드 수정
    date_to = _end_of_next_month_str() if date_mode == 'next_month' else _end_of_year_str()
    for fid in [
        "wnd[0]/usr/ctxtS_ERDAT-HIGH",
        "wnd[0]/usr/txtS_ERDAT-HIGH",
        "wnd[0]/usr/ctxtSEL_ERDAT-HIGH",
        "wnd[0]/usr/ctxtS_CDAT-HIGH",
    ]:
        try:
            session.findById(fid).text = date_to
            logger.info(f"날짜(to) 설정 ({fid}): {date_to}")
            break
        except Exception:
            continue

    # F8: Execute (목록 화면)
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)

    # 레이아웃 /6507 PETER 적용
    _apply_layout(session)
    logger.info(f"ZRMA_Q ({variant}) 준비 완료")


def _apply_layout(session):
    """Manage Layouts → Choose Layouts → /6507 PETER 더블클릭."""
    try:
        grid = session.findById(ZRMA_GRID_PATH)
        grid.pressToolbarContextButton("&MB_VARIANT")
        time.sleep(0.5)
        grid.selectContextMenuItem("&LOAD_VARIANT")
        time.sleep(1)
        # 팝업에서 /6507 PETER 찾아 더블클릭
        # TODO: 팝업 구조 확인 후 업데이트
        popup_grid_candidates = [
            "wnd[1]/usr/cntlGRID1/shellcont/shell",
            "wnd[1]/usr/cntlALV_GRID/shellcont/shell",
        ]
        for pgid in popup_grid_candidates:
            try:
                pg = session.findById(pgid)
                for row_i in range(pg.RowCount):
                    try:
                        val = pg.GetCellValue(row_i, pg.ColumnOrder[0])
                        if "/6507" in str(val) or "PETER" in str(val):
                            pg.setCurrentCell(row_i, pg.ColumnOrder[0])
                            pg.doubleClickCurrentCell()
                            time.sleep(1)
                            logger.info("레이아웃 /6507 PETER 적용 완료")
                            return
                    except Exception:
                        continue
            except Exception:
                continue
        logger.warning("레이아웃 /6507 PETER 자동 선택 실패 (수동 필요)")
    except Exception as e:
        logger.warning(f"레이아웃 적용 실패: {e}")


def is_on_zrma_list(session):
    """ZRMA_Q 목록 화면인지 확인."""
    try:
        title = session.findById("wnd[0]").Text
        return "ZRMA" in title.upper() or "RMA" in title.upper()
    except Exception:
        return False


def refresh_zrma_list(session):
    """F5 새로고침."""
    session.findById("wnd[0]").sendVKey(5)
    time.sleep(2)


def get_all_rows_from_zrma(session):
    """
    ZRMA_Q 목록 그리드에서 전체 행 읽기.
    컬럼명은 discover_sap.py 옵션2로 확인 후 업데이트 필요.
    반환: [{'grid_idx': i, 'order_num': '6***', 'order_type': 'ZRX', 'name': '...', 'rma_status': '...'}, ...]
    """
    rows = []
    try:
        grid = session.findById(ZRMA_GRID_PATH)
        row_count = grid.RowCount
        logger.info(f"ZRMA 그리드 행 수: {row_count}")

        COL_ORDER_NUM  = "T_VBELN"      # Sales Document (오더번호)
        COL_ORDER_TYPE = "T_AUART"      # Sales Document Type (ZRE/ZRX/ZINX...)
        COL_NAME       = "T_NAME1"      # 회사명
        COL_RMA_STATUS = "T_KVGR4_NAME"  # RMA Status (1st Attempt 등)

        for i in range(row_count):
            try:
                def safe_get(col):
                    try:
                        return grid.GetCellValue(i, col).strip()
                    except Exception:
                        return ''

                order_num = safe_get(COL_ORDER_NUM)
                if not order_num:
                    continue

                rows.append({
                    'grid_idx':   i,
                    'order_num':  order_num,
                    'order_type': safe_get(COL_ORDER_TYPE),
                    'name':       safe_get(COL_NAME),
                    'rma_status': safe_get(COL_RMA_STATUS),
                })
            except Exception as e:
                logger.warning(f"ZRMA 행 {i} 읽기 오류: {e}")

    except Exception as e:
        logger.error(f"ZRMA 그리드 읽기 실패: {e}")

    return rows


def navigate_to_zrma_order(session, grid_idx):
    """목록에서 오더 더블클릭으로 진입."""
    try:
        grid = session.findById(ZRMA_GRID_PATH)
        # TODO: Sales Document 컬럼명 확인 후 업데이트
        grid.setCurrentCell(grid_idx, "T_VBELN")
        grid.doubleClickCurrentCell()
        time.sleep(2)
        logger.info(f"ZRMA 오더 진입 (행 {grid_idx})")
        return True
    except Exception as e:
        logger.error(f"ZRMA 오더 진입 실패: {e}")
        return False


def get_items_from_zrma_order(session, order_num, order_type):
    """
    ZRMA_Q 개별 오더에서 All Items 표 읽기.
    Item이 00으로 끝나는 것(개괄명) 제외.
    Route=Return → 회수, Route=Nextday → 배송
    반환: [{'prefix': '배송'/'회수', 'order_type': 'ZRX', 'order_num': '6***',
            'matnr': '...', 'arktx': '...', 'qty': 1}, ...]
    """
    items = []
    try:
        table = session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
            "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
            "/subSUBSCREEN_TC:SAPMV45A:4900"
            "/tblSAPMV45ATCTRL_U_ERF_AUFTRAG"
        )

        # GuiTableControl: 컬럼 인덱스로 접근
        # 0=POSNR, 1=MABNR(Material), 2=ARKTX(Desc), 3=KWMENG(Qty), 12=ROUTE
        row_count = table.RowCount

        for i in range(row_count):
            try:
                item_no = table.GetCell(i, 0).Text.strip()
                matnr   = table.GetCell(i, 1).Text.strip()
                arktx   = table.GetCell(i, 2).Text.strip()
                qty_str = table.GetCell(i, 3).Text.strip()
                route   = table.GetCell(i, 12).Text.strip()

                # 00으로 끝나는 개괄명 제외
                if not item_no or item_no.endswith('00') or item_no == '000000':
                    continue

                try:
                    qty = max(1, int(float(qty_str)))
                except (ValueError, TypeError):
                    qty = 1

                # Route로 배송/회수 판단
                route_upper = route.upper()
                if route_upper == 'RETURN':
                    prefix = '회수'
                elif route_upper == 'NEXDAY':
                    prefix = '배송'
                else:
                    prefix = '배송'  # 기본값
                    logger.warning(f"Route 판단 불명확: '{route}' → 배송으로 처리")

                items.append({
                    'prefix':     prefix,
                    'order_type': order_type,
                    'order_num':  order_num,
                    'matnr':      matnr,
                    'arktx':      arktx,
                    'qty':        qty,
                })

            except Exception as e:
                logger.warning(f"ZRMA 아이템 {i} 읽기 오류: {e}")

    except Exception as e:
        logger.error(f"ZRMA 아이템 읽기 실패: {e}")

    return items


def build_excel_rows_zrma(items, address, extra_orders, memo):
    """
    ZRMA_Q 아이템 리스트 → Excel 행 리스트.
    아이템별 qty만큼 행 생성. 첫 번째 행에만 고객/주소 정보.
    배송/회수 행이 섞여있을 수 있으므로 prefix별로 그룹화하지 않고 순서대로.
    """
    rows = []
    is_first = True

    # 배송(NEXDAY) 먼저, 회수(RETURN) 나중
    items = sorted(items, key=lambda x: 0 if x['prefix'] == '배송' else 1)

    for item in items:
        for _ in range(item['qty']):
            rows.append({
                'order_prefix': item['prefix'],
                'order_type':   item['order_type'],
                'order_num':    item['order_num'],
                'extra_orders': extra_orders if is_first else [],
                'material':     item['matnr'],
                'description':  item['arktx'],
                'customer':     address.get('customer', '') if is_first else '',
                'phone':        address.get('phone', '') if is_first else '',
                'company':      address.get('company', '') if is_first else '',
                'street':       address.get('street', '') if is_first else '',
                'street2':      address.get('street2', '') if is_first else '',
                'memo':         memo if is_first else '',
                'is_first_item': is_first,
            })
            is_first = False

    return rows


def process_zrma_orders(session, new_order_nums, order_map):
    """
    새 ZRMA 오더 처리.
    new_order_nums: ['6123456', ...]
    order_map: get_all_rows_from_zrma() 결과를 order_num으로 그룹핑한 dict
    반환: {order_num: [excel_row, ...], ...}
    """
    results = {}

    for order_num in new_order_nums:
        logger.info(f"=== ZRMA 오더 {order_num} 처리 중 ===")
        order_info = order_map.get(order_num)
        if not order_info:
            continue

        grid_idx   = order_info['grid_idx']
        order_type = order_info['order_type']

        # 1st Attempt 오더는 건너뜀 (이미 엑셀에 있을 가능성 높음)
        if 'attempt' in order_info.get('rma_status', '').lower():
            logger.info(f"오더 {order_num} 1st Attempt → 건너뜀")
            continue

        if not navigate_to_zrma_order(session, grid_idx):
            logger.error(f"오더 {order_num} 진입 실패")
            continue

        # 주소 읽기
        address = get_ship_to_address_zrma(session)

        # 텍스트 읽기
        text = get_text_content(session, menu_id=ZRMA_MENU_TEXTS)
        extra_orders = parse_extra_orders(text) if text else []
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = f"[Text] {text[:150]}" if text and text_contact['found'] else ""
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # 아이템 읽기
        items = get_items_from_zrma_order(session, order_num, order_type)
        if not items:
            logger.warning(f"오더 {order_num} 아이템 없음")
            session.findById("wnd[0]").sendVKey(3)
            time.sleep(0.5)
            continue

        rows = build_excel_rows_zrma(items, address, extra_orders, memo)
        results[order_num] = rows
        logger.info(f"ZRMA 오더 {order_num}: {len(rows)}행 생성")

        # 목록으로 복귀
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)

    return results


def group_zrma_by_order(rows):
    """
    ZRMA 목록 행을 order_num별로 그룹화.
    반환: {'6123456': {'grid_idx': 0, 'order_type': 'ZRX', 'name': '...', 'rma_status': '...'}, ...}
    """
    orders = {}
    for row in rows:
        num = row['order_num']
        if num not in orders:
            orders[num] = {
                'grid_idx':   row['grid_idx'],
                'order_type': row['order_type'],
                'name':       row['name'],
                'rma_status': row['rma_status'],
            }
    return orders

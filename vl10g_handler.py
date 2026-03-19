"""
VL10G 핸들러.
- 배송 미완성 오더 → Background로 VL06O 이관 + ZZ/ZS Delivery Block 해제
- ZRE(회수만) 오더 → 회수 정보 수집
"""

import time
import logging
from datetime import date
import calendar

from sap_handler import (
    run_transaction, get_ship_to_address, get_text_content,
    parse_extra_orders, parse_contact_from_text,
    get_ship_to_address_zrma, parse_memo_for_display, ZRMA_MENU_TEXTS,
)
from zrma_handler import get_items_from_zrma_order, collect_serial_numbers, build_excel_rows_zrma

logger = logging.getLogger(__name__)

# TODO: discover_sap.py로 확인 후 업데이트
VL10G_GRID_PATH = "wnd[0]/usr/cntlGRID1/shellcont/shell"

# 그리드 컬럼명 (discover_sap.py 옵션2로 확인 필요)
COL_ORIG_DOC    = "VBELV"    # OriginDoc (오더번호) — discover 확인
COL_DOC_TYPE    = "AUART"    # Sales Document Type
COL_DELIV_BLOCK = "LIFSP"    # Delivery Block (ZZ/ZS 여부) — discover 확인
COL_VBELN       = "VBELN"    # 배송 문서번호


def _end_of_next_month_str():
    today = date.today()
    year = today.year + 1 if today.month == 12 else today.year
    month = 1 if today.month == 12 else today.month + 1
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, last_day).strftime('%m/%d/%Y')


def navigate_to_vl10g(session):
    """
    VL10G 진입:
    - Shipping Point: 6507
    - Deliv. Creation Date from: 공백
    - Deliv. Creation Date to: 다음달 말일
    - F8
    """
    logger.info("VL10G 진입 중...")
    run_transaction(session, "/nVL10G")
    time.sleep(1.5)

    # Shipping Point/Receiving Pt 입력
    session.findById("wnd[0]/usr/ctxtST_VSTEL-LOW").text = "6507"
    logger.info("Shipping Point 입력: 6507")

    # Deliv. Creation Date from: 공백
    session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").text = ""
    logger.info("날짜(from) 공백")

    # Deliv. Creation Date to: 다음달 말일
    date_to = _end_of_next_month_str()
    session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").text = date_to
    logger.info(f"날짜(to) 설정: {date_to}")

    # F8 Execute
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    logger.info("VL10G 목록 화면 준비 완료")


def is_on_vl10g_list(session):
    try:
        title = session.findById("wnd[0]").Text
        return ("VL10" in title.upper()
                or "Delivery" in title
                or "Activities Due for Shipping" in title)
    except Exception:
        return False


def refresh_vl10g(session):
    session.findById("wnd[0]").sendVKey(5)
    time.sleep(2)


def get_all_rows_from_vl10g(session):
    """
    VL10G 목록 전체 행 읽기.
    반환: [{'grid_idx': i, 'orig_doc': '7***', 'doc_type': 'ZOR',
            'deliv_block': 'ZZ', 'vbeln': '...', 'name1': '회사명'}, ...]
    """
    rows = []
    try:
        grid = session.findById(VL10G_GRID_PATH)
        row_count = grid.RowCount
        logger.info(f"VL10G 그리드 행 수: {row_count}")

        for i in range(row_count):
            try:
                def safe_get(col):
                    try:
                        return grid.GetCellValue(i, col).strip()
                    except Exception:
                        return ''

                orig_doc = safe_get(COL_ORIG_DOC)
                if not orig_doc:
                    continue

                rows.append({
                    'grid_idx':    i,
                    'orig_doc':    orig_doc,
                    'doc_type':    safe_get(COL_DOC_TYPE),
                    'deliv_block': safe_get(COL_DELIV_BLOCK),
                    'vbeln':       safe_get(COL_VBELN),
                    'name1':       safe_get("NAME1"),
                })
            except Exception as e:
                logger.warning(f"VL10G 행 {i} 읽기 오류: {e}")

    except Exception as e:
        logger.error(f"VL10G 그리드 읽기 실패: {e}")

    return rows


def _read_status_bar(session):
    """SAP 상태바 메시지 읽기."""
    try:
        return session.findById("wnd[0]/sbar").Text.strip()
    except Exception:
        return ''


def build_blocked_order_row(orig_doc, doc_type, deliv_block, name1, error_msg):
    """
    블록/에러 오더용 Excel 행 생성 (아이템 없음).
    A열: doc_type + orig_doc, G열: 회사명, H열: 블록 정보 + 에러메시지
    """
    memo = f"[Delivery Block: {deliv_block}]"
    if error_msg:
        memo += f"\n{error_msg}"
    return [{
        'order_prefix':  '',
        'order_type':    doc_type,
        'order_num':     orig_doc,
        'extra_orders':  [],
        'material':      '',
        'description':   '',
        'customer':      '',
        'phone':         '',
        'company':       name1,
        'street':        '',
        'street2':       '',
        'memo':          memo,
        'is_first_item': True,
    }]


_BLOCK_FIELD_CANDIDATES = [
    # ZRE (RMA Order) - 실제 확인된 경로
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK",
    # 일반 Sales Order 경로 (fallback)
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/cmbVBAK-LIFSK",
    "wnd[0]/usr/cmbVBAK-LIFSK",
    "wnd[0]/usr/cmbVBAK-LIFSP",
]


def _try_release_block(session):
    """ZS/ZZ Delivery Block 필드를 공백으로 변경. 반환: True(성공) / False(필드 없음)"""
    # 아이템 테이블 선택 상태 해제 (setCurrentCell 후 헤더 필드가 read-only가 됨)
    try:
        session.findById("wnd[0]").sendVKey(0)  # Enter
        time.sleep(0.3)
    except Exception:
        pass

    # Sales 탭 선택
    try:
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01").select()
        time.sleep(0.5)
    except Exception as e:
        logger.warning(f"  Sales 탭 select 실패: {e}")

    for fid in _BLOCK_FIELD_CANDIDATES:
        try:
            session.findById(fid).key = ""
            logger.info(f"  블록 해제 성공: {fid.split('/')[-1]}")
            return True
        except Exception as e:
            logger.warning(f"  후보 실패 {fid.split('/')[-1]}: {e}")
    return False


def release_delivery_block(session, grid_idx, orig_doc):
    """
    VL10G에서 ZZ/ZS Delivery Block 해제 시도.
    OriginDoc 클릭 → Delivery Block 공백 선택 → 저장 → 복귀.
    반환: (True, '') 성공 / (False, 에러메시지) 실패
    """
    logger.info(f"Delivery Block 해제 중: {orig_doc} (행 {grid_idx})")
    try:
        grid = session.findById(VL10G_GRID_PATH)
        grid.setCurrentCell(grid_idx, COL_ORIG_DOC)
        grid.clickCurrentCell()
        time.sleep(1.5)

        title = session.findById("wnd[0]").Text
        sbar_msg = _read_status_bar(session)
        logger.info(f"진입 후 화면: '{title}' / 상태바: '{sbar_msg}'")

        # 상태바에 에러 메시지가 있으면 진입 실패
        if sbar_msg:
            logger.warning(f"오더 {orig_doc} Block 해제 불가: {sbar_msg}")
            session.findById("wnd[0]").sendVKey(3)
            time.sleep(1)
            return False, sbar_msg

        # 여전히 VL10G 목록 화면이면 진입 실패
        if is_on_vl10g_list(session):
            msg = f"오더 진입 실패 (목록 화면 유지): {title}"
            logger.warning(f"오더 {orig_doc}: {msg}")
            return False, msg

        # Delivery Block 필드 → 공백 선택
        if not _try_release_block(session):
            msg = "Delivery Block 필드 찾기 실패"
            logger.warning(f"오더 {orig_doc}: {msg}")
            session.findById("wnd[0]").sendVKey(3)
            return False, msg

        # 저장 후 F3 복귀
        session.findById("wnd[0]").sendVKey(11)
        time.sleep(1.5)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)
        logger.info(f"오더 {orig_doc} Block 해제 저장 완료")
        return True, ''

    except Exception as e:
        msg = str(e)
        logger.error(f"Block 해제 실패 ({orig_doc}): {msg}")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass
        return False, msg


def _enter_zre_order(session, grid_idx, orig_doc):
    """
    ZRE 오더 진입: VL10G 그리드 클릭 → 실패시 VA02 fallback.
    반환: ('grid'|'va02'|None, True|False)
    """
    # 1) 그리드 클릭 시도
    try:
        grid = session.findById(VL10G_GRID_PATH)
        grid.setCurrentCell(grid_idx, COL_ORIG_DOC)
        grid.clickCurrentCell()
        time.sleep(1.5)
        sbar_msg = _read_status_bar(session)
        if not sbar_msg and not is_on_vl10g_list(session):
            logger.info(f"  오더 {orig_doc} 그리드 클릭 진입 성공")
            return 'grid', True
        logger.warning(f"  그리드 클릭 진입 실패 (sbar='{sbar_msg}') → VA02 시도")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass
    except Exception as e:
        logger.warning(f"  그리드 클릭 오류: {e} → VA02 시도")

    # 2) VA02 fallback
    try:
        run_transaction(session, "VA02")
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = orig_doc
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        sbar_msg = _read_status_bar(session)
        if not sbar_msg and not is_on_vl10g_list(session):
            logger.info(f"  오더 {orig_doc} VA02 진입 성공")
            return 'va02', True
        logger.warning(f"  VA02 진입 실패: sbar='{sbar_msg}'")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass
    except Exception as e:
        logger.error(f"  VA02 진입 오류: {e}")

    return None, False


def _back_to_vl10g(session, entry_method):
    """오더 화면에서 VL10G 목록으로 복귀. /n으로 강제 이동해 팝업/중간화면 무시."""
    try:
        run_transaction(session, "/nVL10G")
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtST_VSTEL-LOW").text = "6507"
        session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").text = _end_of_next_month_str()
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2)
        logger.info("  VL10G 목록 복귀 완료")
    except Exception as e:
        logger.warning(f"  VL10G 복귀 실패: {e}")


def process_zre_with_block_release(session, grid_idx, orig_doc, order_type, name1):
    """
    ZRE 회수 오더: 오더 진입 → 데이터 수집 → ZS/ZZ 블록 해제 시도 → 저장.
    - 블록 해제 실패해도 수집된 데이터는 반환 (메모에 [ZS미해제-수동확인] 표기)
    - 그리드 클릭 진입 실패시 VA02 fallback
    반환: (True, [excel_rows]) | (False, 에러메시지)
    """
    logger.info(f"ZRE 오더 {orig_doc} 진입 (데이터 수집 + ZS 해제 시도)...")

    entry_method, entered = _enter_zre_order(session, grid_idx, orig_doc)
    if not entered:
        return False, "오더 진입 실패 (grid + VA02 모두 불가)"

    try:
        title = session.findById("wnd[0]").Text
        logger.info(f"  오더 화면: {title}")

        # 1. 주소 읽기
        address = get_ship_to_address_zrma(session)

        # 2. 텍스트/메모 읽기
        text = get_text_content(session, menu_id=ZRMA_MENU_TEXTS)
        all_extra = parse_extra_orders(text) if text else []
        extra_orders = [eo for eo in all_extra if orig_doc not in eo]
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = parse_memo_for_display(text) if text else ''
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # 3. 아이템 읽기
        items = get_items_from_zrma_order(session, orig_doc, order_type)

        # 4. S/N 수집
        if items:
            collect_serial_numbers(session, items)

        # 5. Excel 행 생성
        if items:
            excel_rows = build_excel_rows_zrma(items, address, extra_orders, memo)
        else:
            logger.warning(f"  아이템 없음 → fallback 행")
            excel_rows = build_blocked_order_row(orig_doc, order_type, 'ZS', name1, '아이템 없음')

        # 6. ZS/ZZ 블록 해제 시도
        block_released = _try_release_block(session)

        if block_released:
            # 저장 후 복귀
            session.findById("wnd[0]").sendVKey(11)
            time.sleep(1.5)
            logger.info(f"  오더 {orig_doc} ZS 해제 + 저장 완료")
            _back_to_vl10g(session, entry_method)
        else:
            # 해제 실패 → 저장 없이 복귀, 메모에 수동 확인 표기
            logger.warning(f"  오더 {orig_doc} ZS 해제 불가 → 데이터만 수집")
            for row in excel_rows:
                existing = row.get('memo', '')
                row['memo'] = ("[ZS미해제-수동확인]\n" + existing).strip() if existing else "[ZS미해제-수동확인]"
            _back_to_vl10g(session, entry_method)

        return True, excel_rows

    except Exception as e:
        logger.error(f"ZRE 오더 {orig_doc} 처리 실패: {e}")
        try:
            _back_to_vl10g(session, entry_method)
        except Exception:
            pass
        return False, str(e)


def push_to_vl06o(session, grid_idx):
    """
    VL10G 목록에서 해당 행 선택 → Background 버튼 클릭 → VL06O로 이관.
    반환: True(성공) / False(실패)
    """
    try:
        grid = session.findById(VL10G_GRID_PATH)
        # selectedRows로 행 선택 후 Background 버튼
        grid.selectedRows = str(grid_idx)
        time.sleep(0.3)

        # Background 버튼 클릭 (tbar[1]/btn[19] = Shift+F7)
        session.findById("wnd[0]/tbar[1]/btn[19]").press()
        time.sleep(2)
        logger.info("Background 버튼 클릭 (tbar[1]/btn[19])")
        return True

    except Exception as e:
        logger.error(f"VL06O 이관 실패: {e}")
        return False


def process_vl10g(session, processed):
    """
    VL10G 전체 처리:
    1. ZZ/ZS Block 해제 시도 → 실패시 실패목록 반환
    2. 배송 오더(ZRE 아닌 것) → Background로 VL06O 이관
    3. ZRE(회수만) 오더 → 회수 정보 수집

    반환: {
        'failed_blocks': [{'orig_doc': '...', 'doc_type': '...'}],  # 수동 해제 필요
        'pushed_to_vl06o': ['7***', ...],                           # VL06O로 이관됨
        'zre_orders': {'6***': [excel_row, ...]}                    # ZRE 회수 오더
    }
    """
    result = {
        'blocked_excel_rows': {},   # {orig_doc: [excel_row]} 블록/에러 오더
        'pushed_to_vl06o':    [],   # Background 성공 오더
        'zre_orders':         {},   # ZRE 회수 오더 (추후 처리)
    }

    rows = get_all_rows_from_vl10g(session)
    if not rows:
        logger.info("VL10G 목록 비어있음")
        return result

    logger.info(f"VL10G 행 수: {len(rows)}")

    new_rows = [r for r in rows if r['orig_doc'] not in processed]
    logger.info(f"새 VL10G 오더: {len(new_rows)}개")

    for row in new_rows:
        orig_doc    = row['orig_doc']
        doc_type    = row['doc_type']
        deliv_block = row['deliv_block'].upper()
        grid_idx    = row['grid_idx']
        name1       = row.get('name1', '')

        # ── ZRE + ZS/ZZ 블록: 데이터 수집 + 블록 해제 한 번에 ──────────
        if doc_type == 'ZRE' and deliv_block in ('ZZ', 'ZS'):
            success, data = process_zre_with_block_release(
                session, grid_idx, orig_doc, doc_type, name1
            )
            if success:
                result['zre_orders'][orig_doc] = data
                logger.info(f"ZRE 오더 {orig_doc} 처리 완료 ({len(data)}행)")
            else:
                excel_rows = build_blocked_order_row(orig_doc, doc_type, deliv_block, name1, data)
                result['blocked_excel_rows'][orig_doc] = excel_rows
                logger.warning(f"ZRE 오더 {orig_doc} 처리 실패 → Excel 기록")
            refresh_vl10g(session)
            rows = get_all_rows_from_vl10g(session)
            continue

        # ── Non-ZRE ZZ/ZS Block → 해제만 ───────────────────────────────
        if deliv_block in ('ZZ', 'ZS'):
            success, err_msg = release_delivery_block(session, grid_idx, orig_doc)
            if not success:
                excel_rows = build_blocked_order_row(orig_doc, doc_type, deliv_block, name1, err_msg)
                result['blocked_excel_rows'][orig_doc] = excel_rows
                logger.warning(f"오더 {orig_doc} Block 해제 불가 → Excel 기록")
                continue
            refresh_vl10g(session)
            rows = get_all_rows_from_vl10g(session)

        # ── ZRE (블록 없음): 추후 처리 ──────────────────────────────────
        if doc_type == 'ZRE':
            logger.info(f"ZRE 회수 오더 {orig_doc} (블록 없음) - 추후 처리")
            result['zre_orders'][orig_doc] = []
        else:
            # 배송 오더 → Background VL06O 이관 (현재 홀드)
            logger.info(f"오더 {orig_doc} VL06O 이관 홀드 (스킵)")

    return result

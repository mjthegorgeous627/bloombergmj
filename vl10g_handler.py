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
    get_ship_to_address_zrma, parse_memo_for_display, refresh_sap_list, ZRMA_MENU_TEXTS,
)
from zrma_handler import get_items_from_zrma_order, collect_serial_numbers, build_excel_rows_zrma

logger = logging.getLogger(__name__)

# TODO: discover_sap.py로 확인 후 업데이트
VL10G_GRID_PATH = "wnd[0]/usr/cntlGRID1/shellcont/shell"
VL10G_SHIPPING_POINT_FIELD = "wnd[0]/usr/ctxtST_VSTEL-LOW"
VL10G_DATE_FROM_FIELD = "wnd[0]/usr/ctxtST_LEDAT-LOW"
VL10G_DATE_TO_FIELD = "wnd[0]/usr/ctxtST_LEDAT-HIGH"

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


def _safe_find(session, elem_id):
    try:
        return session.findById(elem_id)
    except Exception:
        return None


def _screen_title(session):
    try:
        return session.findById("wnd[0]").Text
    except Exception:
        return ""


def _close_common_popup(session):
    """
    /n 이동 중 경고/확인 팝업이 뜨면 기본값으로 닫는다.
    저장 확인처럼 위험할 수 있는 팝업은 Enter가 기본 동작이므로 별도 선택을 강제하지 않는다.
    """
    try:
        popup = session.findById("wnd[1]")
    except Exception:
        return False

    logger.warning(f"  SAP 팝업 감지: '{getattr(popup, 'Text', '')}' → Enter")
    try:
        popup.sendVKey(0)
        time.sleep(0.8)
        return True
    except Exception as e:
        logger.warning(f"  팝업 Enter 처리 실패: {e}")
        return False


def _wait_for_element(session, elem_id, timeout=8, interval=0.4):
    deadline = time.time() + timeout
    while time.time() < deadline:
        _close_common_popup(session)
        elem = _safe_find(session, elem_id)
        if elem is not None:
            return elem
        time.sleep(interval)
    return None


def _open_vl10g_selection(session):
    """
    현재 화면이 무엇이든 /nVL10G로 선택화면까지 이동한다.
    SAP가 늦게 그리거나 팝업이 끼는 경우를 감안해 필드 출현을 기다린다.
    """
    logger.info("VL10G 선택화면 진입 중...")
    for attempt in range(2):
        run_transaction(session, "/nVL10G")
        field = _wait_for_element(session, VL10G_SHIPPING_POINT_FIELD, timeout=8)
        if field is not None:
            return field

        logger.warning(
            "VL10G 선택화면 필드 대기 실패 "
            f"(attempt={attempt + 1}, title='{_screen_title(session)}', sbar='{_read_status_bar(session)}')"
        )
        time.sleep(1)

    raise RuntimeError(
        "VL10G 선택화면 진입 실패: "
        f"title='{_screen_title(session)}', sbar='{_read_status_bar(session)}'"
    )


def _execute_vl10g_selection(session):
    shipping_point = _wait_for_element(session, VL10G_SHIPPING_POINT_FIELD, timeout=4)
    if shipping_point is None:
        raise RuntimeError(
            "VL10G 선택화면 필드를 찾지 못함: "
            f"title='{_screen_title(session)}', sbar='{_read_status_bar(session)}'"
        )

    shipping_point.text = "6507"
    logger.info("Shipping Point 입력: 6507")

    date_from = _wait_for_element(session, VL10G_DATE_FROM_FIELD, timeout=2)
    if date_from is not None:
        date_from.text = ""
        logger.info("날짜(from) 공백")
    else:
        logger.warning("VL10G 날짜(from) 필드를 찾지 못함")

    date_to = _wait_for_element(session, VL10G_DATE_TO_FIELD, timeout=2)
    if date_to is None:
        raise RuntimeError("VL10G 날짜(to) 필드를 찾지 못함")

    to_value = _end_of_next_month_str()
    date_to.text = to_value
    logger.info(f"날짜(to) 설정: {to_value}")

    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    logger.info("VL10G 목록 화면 준비 완료")


def navigate_to_vl10g(session):
    """
    VL10G 진입:
    - 이미 목록 화면이면 새로고침만 실행
    - 아니면 /nVL10G + 조건 입력 + F8
    """
    if is_on_vl10g_list(session) and _safe_find(session, VL10G_GRID_PATH) is not None:
        logger.info("VL10G already on list; refreshing current list")
        if refresh_sap_list(session, VL10G_GRID_PATH, "VL10G"):
            return
        logger.warning("VL10G refresh failed; reopening selection")

    _open_vl10g_selection(session)
    _execute_vl10g_selection(session)


def is_on_vl10g_list(session):
    try:
        title = session.findById("wnd[0]").Text
        return ("VL10" in title.upper()
                or "Activities Due for Shipping" in title)
    except Exception:
        return False


def refresh_vl10g(session):
    """현재 목록 화면이면 새로고침만, 아니면 선택화면 재진입 후 실행."""
    if _safe_find(session, VL10G_GRID_PATH) is not None:
        if refresh_sap_list(session, VL10G_GRID_PATH, "VL10G"):
            return
        logger.warning("VL10G refresh failed; reopening selection")

    try:
        session.findById("wnd[0]").sendVKey(3)  # F3 = Back
        time.sleep(1.5)
    except Exception as e:
        logger.warning(f"VL10G F3 복귀 실패 → /nVL10G 재진입: {e}")

    if _safe_find(session, VL10G_SHIPPING_POINT_FIELD) is None:
        _open_vl10g_selection(session)

    _execute_vl10g_selection(session)


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

    2026-08-07: process_vl10g()에서 더 이상 호출하지 않음 - 사용자 지시로
    블록 해제 자체를 시도하지 않고 바로 VA02/VA03 직접 수집으로 넘어가도록
    바뀜(권한 부족으로 인한 실패가 반복 에러의 원인이었음). 함수는 필요할 때
    다시 쓸 수 있게 남겨둠.
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
    ZRE 회수 오더 (ZZ/ZS 블록 있는 것): 오더 진입 → 데이터 수집 → ZS/ZZ 해제 → 저장.
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
            session.findById("wnd[0]").sendVKey(11)
            time.sleep(1.5)
            logger.info(f"  오더 {orig_doc} ZS 해제 + 저장 완료")
            _back_to_vl10g(session, entry_method)
        else:
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

    2026-08-28 사용자 지시로 process_vl10g()가 이 함수를 호출하지 않는다 -
    이관(배송 문서 생성)은 앞으로도 사용자가 SAP에서 직접 처리하고, 자동화는
    건드리지 않는다(오더 수집 자체는 run_manual_order()로 이관과 무관하게
    이미 처리됨 - process_vl10g() 참고). 필요해지면 다시 연결할 수 있게
    함수는 남겨둠.
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
    1. ZRE + ZZ/ZS 블록 → 데이터 수집 + 블록 해제 한 번에 (process_zre_with_block_release)
    2. 그 외 배송 오더(ZRE 블록없음 포함, 블록 있든 없든) → run_manual_order()로
       VA02/VA03 직접 수집. "이관"(Background, Shift+F7로 VL06O에 배송 문서
       생성)은 자동으로 누르지 않는다(2026-08-28 사용자 지시 - 사용자가 SAP
       에서 직접 처리) - push_to_vl06o()는 그래서 호출하지 않지만, 수집
       자체는 이관 여부와 무관하게 항상 실행된다(이관 전이라 VL06O에 아직
       안 뜬 오더가 자동 수집 자체에서도 방치되던 문제 수정, 2026-08-28).
       2026-09-01: ZRE 블록없음을 "ZRMA_Q 담당이니 스킵"으로 특별취급하던
       것도 없애고 여기로 합류시켰다 - VL06O/VL10G에 보이는 오더는 ZRMA_Q
       스캔 결과와 무관하게 항상 수집되어야 한다는 원칙(사용자 지시).

    반환: {
        'blocked_excel_rows': {'7***': [excel_row]},  # 수집 실패 시 에러 placeholder
        'pushed_to_vl06o': [],                        # 더 이상 안 씀 (항상 빈 리스트)
        'zre_orders': {'6***': [excel_row, ...]}       # ZRE 회수 오더
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

        # ── ZRE(블록 없음)를 포함한 그 외 배송 오더 → VA02/VA03 직접 수집 ──
        # 2026-09-01 사용자 지시로 "ZRE 블록없음은 ZRMA_Q가 담당하니 스킵"
        # 특별취급을 없앴다. ZRMA_Q RLKR/Q2가 이 오더를 놓치는 경우가
        # 실제로 있었고(2026-08-19 오더 67072626 미수집, 원인 미확정), 그럴
        # 때 VL10G에 분명히 떠 있는데도 VL06O(6번대 필터 제외)·VL10G(이
        # 스킵) 어느 쪽도 안 잡아서 완전한 사각지대에 빠졌었다.
        # run_manual_order()는 화면 제목으로 ZRE도 이미 정상 인식하므로
        # (manual_order_handler._order_type_for) 바로 아래 분기로 그냥
        # 흘려보내면 된다 - ZRMA_Q가 먼저 잡아도 공유 processed로 dedup되니
        # 중복 수집 걱정은 없다(기존 "그 외 배송 오더" 분기와 동일 안전성).
        # 2026-08-07 사용자 지시로 ZZ/ZS 블록은 release_delivery_block()
        # (권한 부족으로 자주 실패했음)로 해제 시도하지 않고 run_manual_order()
        # (오더번호로 VA02/VA03에 직접 진입 - 블록과 무관하게 정상 동작)로
        # 바로 수집하게 됐었다. 블록 없는 정상 배송 오더는 그동안 별도
        # 취급이었다 - VL10G→VL06O "이관"(Background, Shift+F7) 자동 실행을
        # 시도하는 코드였는데, 실제로는 그 실행 자체가 "(현재 홀드)"라는
        # 주석과 함께 아예 호출되지 않고 로그만 남기고 스킵하는 상태로
        # 남아있었다(git 최초 커밋 시점부터 - 만든 사람도 잊었을 만큼 오래됨).
        # 그 결과 이관 전 오더는 SAP 자체 백그라운드 job이 알아서 배송을
        # 만들어줄 때까지(보통 20분~1시간, 가끔 몇 시간) 매 사이클 이 로그만
        # 남기고 방치됐다 - 실측: 오더 67083939가 2시간 넘게 이 상태.
        #
        # 2026-08-28 사용자 지시: "이관"(Background) 버튼은 앞으로도 자동으로
        # 누르지 않는다(사용자가 SAP에서 직접 처리) - push_to_vl06o()는
        # 그래서 그대로 안 쓴다. 그런데 오더 "수집" 자체는 이관 여부를 기다릴
        # 이유가 전혀 없다 - run_manual_order()가 VA02/VA03으로 직접 들어가서
        # 이관/블록과 무관하게 동작하므로, 블록 있는 오더와 똑같이 바로 수집한다.
        # 나중에 실제로 이관되어 VL06O에도 뜨더라도, 이미 이 경로에서
        # processed로 마킹돼있으므로 _run_vl06o()가 중복으로 다시 수집/카톡
        # 전송하지 않는다(order_tracker의 공유 processed 파일 기준 dedup).
        logger.info(f"오더 {orig_doc}: VA02/VA03 직접 수집 (block={deliv_block or '없음'}, 이관은 수동으로 진행 예정)")
        collected = False
        for attempt in range(2):
            try:
                from manual_order_handler import run_manual_order

                if run_manual_order(orig_doc, close_window=True):
                    logger.info(f"오더 {orig_doc} 직접 수집 완료 ({doc_type}, block={deliv_block or '없음'})")
                    collected = True
                    break
            except Exception as exc:
                logger.error(f"오더 {orig_doc} 직접 수집 실패 (시도 {attempt + 1}/2): {exc}")
                if attempt == 0:
                    time.sleep(2)
        if not collected:
            excel_rows = build_blocked_order_row(orig_doc, doc_type, deliv_block, name1, "VA02/VA03 직접 수집 실패")
            result['blocked_excel_rows'][orig_doc] = excel_rows
            logger.warning(f"오더 {orig_doc} 직접 수집 실패 → Block 정보만 Excel 기록")

    return result

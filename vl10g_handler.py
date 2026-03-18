"""
VL10G 핸들러.
- 배송 미완성 오더 → Background로 VL06O 이관 + ZZ/ZS Delivery Block 해제
- ZRE(회수만) 오더 → 회수 정보 수집
"""

import time
import logging
from datetime import date
import calendar

from sap_handler import run_transaction, get_ship_to_address, get_text_content, parse_extra_orders, parse_contact_from_text

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
    run_transaction(session, "VL10G")
    time.sleep(1.5)

    # Shipping Point/Receiving Pt 입력
    for fid in [
        "wnd[0]/usr/ctxtLIKP-VSTEL",
        "wnd[0]/usr/ctxtSHIPPING_POINT",
        "wnd[0]/usr/ctxtS_VSTEL-LOW",
    ]:
        try:
            session.findById(fid).text = "6507"
            logger.info(f"Shipping Point 입력 ({fid})")
            break
        except Exception:
            continue

    # Deliv. Creation Date from: 공백
    for fid in [
        "wnd[0]/usr/ctxtS_ERDAT-LOW",
        "wnd[0]/usr/txtS_ERDAT-LOW",
        "wnd[0]/usr/ctxtSEL_ERDAT-LOW",
    ]:
        try:
            session.findById(fid).text = ""
            logger.info(f"날짜(from) 공백 ({fid})")
            break
        except Exception:
            continue

    # Deliv. Creation Date to: 다음달 말일
    date_to = _end_of_next_month_str()
    for fid in [
        "wnd[0]/usr/ctxtS_ERDAT-HIGH",
        "wnd[0]/usr/txtS_ERDAT-HIGH",
        "wnd[0]/usr/ctxtSEL_ERDAT-HIGH",
    ]:
        try:
            session.findById(fid).text = date_to
            logger.info(f"날짜(to) 설정 ({fid}): {date_to}")
            break
        except Exception:
            continue

    # F8 Execute
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    logger.info("VL10G 목록 화면 준비 완료")


def is_on_vl10g_list(session):
    try:
        title = session.findById("wnd[0]").Text
        return "VL10" in title.upper() or "Delivery" in title
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

        # 진입 실패 (에러 화면) → 상태바 메시지 캡처 후 복귀
        if sbar_msg or "Change Sales Documents" in title or "Overview" not in title:
            if not sbar_msg:
                sbar_msg = f"화면 진입 실패: {title}"
            logger.warning(f"오더 {orig_doc} Block 해제 불가: {sbar_msg}")
            session.findById("wnd[0]").sendVKey(3)
            time.sleep(1)
            return False, sbar_msg

        # Delivery Block 필드 → 공백 선택
        block_field_candidates = [
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/cmbVBAK-LIFSK",
            "wnd[0]/usr/cmbVBAK-LIFSK",
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/cmbVBAK-LIFSP",
            "wnd[0]/usr/cmbVBAK-LIFSP",
        ]
        cleared = False
        for fid in block_field_candidates:
            try:
                field = session.findById(fid)
                field.key = ""
                logger.info(f"Delivery Block 공백 선택 ({fid})")
                cleared = True
                break
            except Exception:
                continue

        if not cleared:
            msg = "Delivery Block 필드 찾기 실패"
            logger.warning(f"오더 {orig_doc}: {msg}")
            session.findById("wnd[0]").sendVKey(3)
            return False, msg

        # 저장
        session.findById("wnd[0]").sendVKey(11)
        time.sleep(1.5)
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

        # ZZ/ZS Block → 해제 시도
        if deliv_block in ('ZZ', 'ZS'):
            success, err_msg = release_delivery_block(session, grid_idx, orig_doc)
            if not success:
                # 해제 불가 → Excel에 블록 정보 + 에러 기록
                excel_rows = build_blocked_order_row(orig_doc, doc_type, deliv_block, name1, err_msg)
                result['blocked_excel_rows'][orig_doc] = excel_rows
                logger.warning(f"오더 {orig_doc} Block 해제 불가 → Excel 기록")
                continue
            # 해제 성공 → 목록 새로고침
            refresh_vl10g(session)
            rows = get_all_rows_from_vl10g(session)

        # ZRE(회수만): 향후 처리 (placeholder)
        if doc_type == 'ZRE':
            logger.info(f"ZRE 회수 오더 {orig_doc} - 추후 처리")
            result['zre_orders'][orig_doc] = []

        else:
            # 배송 오더 → Background로 VL06O 이관
            success = push_to_vl06o(session, grid_idx)
            if success:
                result['pushed_to_vl06o'].append(orig_doc)
                logger.info(f"오더 {orig_doc} VL06O 이관 완료")
            else:
                logger.warning(f"오더 {orig_doc} VL06O 이관 실패")

    return result

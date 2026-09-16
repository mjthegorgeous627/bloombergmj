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
    refresh_sap_list,
    get_scripting_engine,
    get_ship_to_address_zrma,
    get_text_content,
    parse_extra_orders,
    parse_contact_from_text,
    parse_memo_for_display,
    collect_order_extra_fields,
    add_extra_order_dedup,
    ZRMA_MENU_TEXTS,
)

logger = logging.getLogger(__name__)

# 확인 필요한 element ID (discover_sap.py로 확인 후 업데이트)
ZRMA_GRID_PATH = "wnd[0]/usr/cntlCUST_CONT/shellcont/shell"

# All Items 테이블 경로
ZRMA_ITEMS_TABLE_PATH = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
    "/subSUBSCREEN_TC:SAPMV45A:4900"
    "/tblSAPMV45ATCTRL_U_ERF_AUFTRAG"
)

# Extras > Technical objects 메뉴 (menu[3]=Extras, menu[9]=Technical objects)
EXTRAS_TECH_OBJ = "wnd[0]/mbar/menu[3]/menu[9]"

# Maintain/Display Serial Numbers 팝업 — GuiTableControl 경로 후보
_SN_TABLE_IDS = [
    "wnd[1]/usr/tblSAPLIPW1TC_SERIAL_NUMBERS",
    "wnd[1]/usr/tblSAPLAQSE_TC_SERIAL",
    "wnd[1]/usr/tblSAPLAQS1TC_SERIAL",
    "wnd[1]/usr/sub0001/tblSAPLAQSE_TC_SERIAL",
]

# Display 모드에서 열리는 ALV Grid(GuiGridView) 경로 후보
_SN_GRID_IDS = [
    "wnd[1]/usr/cntlGRID1/shellcont/shell",
    "wnd[1]/usr/cntlSERIAL/shellcont/shell",
    "wnd[1]/usr/cntlCUSTOM/shellcont/shell",
    "wnd[1]/usr/cntlALV_GRID/shellcont/shell",
]


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


def navigate_to_zrma_q(session, variant, date_mode='next_month', apply_layout=False):
    """
    ZRMA_Q 진입.
    variant: 'rlkr' 또는 'q2'
    date_mode: 'next_month' (RLKR) 또는 'year_end' (Q2)
    """
    logger.info(f"ZRMA_Q 진입 중 (variant={variant})...")
    run_transaction(session, "/nZRMA_Q")
    time.sleep(1.5)

    # Get Variants 버튼 클릭 (tbar[1]/btn[17] = Shift+F5)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(1)

    # 팝업: Variant 입력
    for fid in ["wnd[1]/usr/txtV-LOW", "wnd[1]/usr/txtVARIANT"]:
        try:
            session.findById(fid).text = variant.upper()
            logger.info(f"Variant 입력 ({fid}): {variant.upper()}")
            break
        except Exception:
            continue

    # Created By 공백으로
    for fid in ["wnd[1]/usr/txtENAME-LOW", "wnd[1]/usr/txtUSER-LOW"]:
        try:
            session.findById(fid).text = ""
            break
        except Exception:
            continue

    # F8: Execute (variant 선택)
    session.findById("wnd[1]").sendVKey(8)
    time.sleep(1.5)

    # Creation Date to 필드 수정 (discover 확인된 ID)
    date_to = _end_of_next_month_str() if date_mode == 'next_month' else _end_of_year_str()
    session.findById("wnd[0]/usr/ctxtS_ERDAT-HIGH").text = date_to
    logger.info(f"날짜(to) 설정: {date_to}")

    # F8: Execute (목록 화면)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(2)

    # 레이아웃 /6507 PETER 적용
    if apply_layout:
        _apply_layout(session)
    else:
        logger.info("ZRMA_Q layout apply skipped")
    logger.info(f"ZRMA_Q ({variant}) 준비 완료")


def _iter_layout_popup_grids(session):
    """Choose Layout 팝업 안의 GridView/Shell 컨트롤을 경로 고정 없이 찾는다."""
    explicit_ids = [
        "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell",
        "wnd[1]/usr/cntlGRID1/shellcont/shell",
        "wnd[1]/usr/cntlALV_GRID/shellcont/shell",
        "wnd[1]/usr/cntlCONTAINER/shellcont/shell",
        "wnd[1]/usr/cntlGRID/shellcont/shell",
    ]
    seen = set()
    for elem_id in explicit_ids:
        try:
            obj = session.findById(elem_id)
            seen.add(obj.Id)
            yield obj
        except Exception:
            pass

    try:
        root = session.findById("wnd[1]/usr")
    except Exception:
        return

    def walk(obj):
        try:
            children = obj.Children
            count = children.Count
        except Exception:
            return
        for i in range(count):
            try:
                child = children(i)
            except Exception:
                continue
            child_id = getattr(child, 'Id', '')
            child_type = getattr(child, 'Type', '')
            child_text = str(getattr(child, 'Text', '') or '')
            if child_id not in seen and (child_type == 'GuiShell' or 'GridView' in child_text):
                seen.add(child_id)
                yield child
            yield from walk(child)

    yield from walk(root)


def _read_layout_popup_rows(session):
    """Choose Layout 팝업의 행 텍스트를 최대한 여러 방식으로 읽는다."""
    column_candidates = (
        "LAYOUT", "VARIANT", "TEXT", "DESCRIPT", "LTDX", "REPORT",
        "S_LAYOUT", "VARIANT_TEXT", "COLTEXT", "SCRTEXT_L",
    )

    for pg in _iter_layout_popup_grids(session):
        row_count = int(getattr(pg, 'RowCount', 0) or getattr(pg, 'VisibleRowCount', 0) or 0)
        logger.info(f"ZRMA_Q layout popup control found: {getattr(pg, 'Id', '')}, rows={row_count}")

        columns = list(column_candidates)
        try:
            for i in range(30):
                try:
                    col = pg.ColumnOrder(i)
                except Exception:
                    try:
                        col = pg.ColumnOrder[i]
                    except Exception:
                        break
                if col and col not in columns:
                    columns.insert(0, col)
        except Exception:
            pass

        for row_i in range(row_count):
            parts = []
            hit_col = columns[0] if columns else 0
            for col in columns:
                try:
                    value = str(pg.GetCellValue(row_i, col)).strip()
                except Exception:
                    try:
                        value = str(pg.GetCell(row_i, col).Text).strip()
                    except Exception:
                        continue
                if value:
                    parts.append(value)
                    hit_col = col
            row_text = " ".join(parts)
            if row_text:
                yield pg, row_i, hit_col, row_text

def _press_layout_popup_ok(session):
    for btn_id in ["wnd[1]/tbar[0]/btn[0]", "wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/usr/btnBUTTON_1"]:
        try:
            session.findById(btn_id).press()
            time.sleep(1)
            return True
        except Exception:
            continue
    try:
        session.findById("wnd[1]").sendVKey(0)
        time.sleep(1)
        return True
    except Exception:
        return False



def _select_second_layout_row_fallback(session):
    """텍스트 읽기가 실패하면 팝업 GridView의 두 번째 행(/6507 PETER)을 직접 선택한다."""
    for grid in _iter_layout_popup_grids(session):
        try:
            row_count = int(getattr(grid, 'RowCount', 0) or getattr(grid, 'VisibleRowCount', 0) or 0)
            if row_count < 2:
                continue

            col = None
            try:
                col = grid.ColumnOrder(0)
            except Exception:
                try:
                    col = grid.ColumnOrder[0]
                except Exception:
                    pass

            try:
                grid.selectedRows = "1"
            except Exception:
                pass
            if col:
                try:
                    grid.setCurrentCell(1, col)
                except Exception:
                    pass

            try:
                grid.doubleClickCurrentCell()
                time.sleep(1)
                logger.info(f"레이아웃 /6507 PETER 적용 완료 (grid row fallback: {getattr(grid, 'Id', '')})")
                return True
            except Exception:
                if _press_layout_popup_ok(session):
                    logger.info(f"레이아웃 /6507 PETER 적용 완료 (grid ok fallback: {getattr(grid, 'Id', '')})")
                    return True
        except Exception as e:
            logger.debug(f"layout grid fallback failed: {e}")

    logger.warning("레이아웃 /6507 PETER fallback 선택 실패")
    return False

def _apply_layout(session):
    """ALV Choose Layout 팝업에서 /6507 PETER를 선택한다."""
    target_keys = ("/6507", "PETER", "Q2 6507")
    try:
        grid = session.findById(ZRMA_GRID_PATH)
        grid.pressToolbarContextButton("&MB_VARIANT")
        time.sleep(0.5)

        menu_opened = False
        for menu_id in ("&LOAD", "&LOAD_VARIANT"):
            try:
                grid.selectContextMenuItem(menu_id)
                menu_opened = True
                logger.info(f"ZRMA_Q layout menu opened ({menu_id})")
                break
            except Exception as e:
                logger.debug(f"ZRMA_Q layout menu id failed ({menu_id}): {e}")
        if not menu_opened:
            logger.warning("ZRMA_Q layout menu open failed")
            return False

        time.sleep(1)
        seen_rows = []
        for popup_grid, row_i, hit_col, row_text in _read_layout_popup_rows(session):
            seen_rows.append(row_text)
            upper_text = row_text.upper()
            if any(key in upper_text for key in target_keys):
                try:
                    popup_grid.setCurrentCell(row_i, hit_col)
                except Exception:
                    pass
                try:
                    popup_grid.selectedRows = str(row_i)
                except Exception:
                    pass
                try:
                    popup_grid.doubleClickCurrentCell()
                    time.sleep(1)
                except Exception:
                    _press_layout_popup_ok(session)
                logger.info(f"레이아웃 /6507 PETER 적용 완료: {row_text}")
                return True

        logger.warning(f"레이아웃 /6507 PETER 자동 선택 실패. 보이는 행: {seen_rows[:8]}")
        return _select_second_layout_row_fallback(session)
    except Exception as e:
        logger.warning(f"레이아웃 적용 실패: {e}")
        return False

def is_on_zrma_list(session):
    """ZRMA_Q 목록 화면인지 확인."""
    try:
        title = session.findById("wnd[0]").Text
        if "RMA LIST" not in title.upper():
            return False
        session.findById(ZRMA_GRID_PATH)
        return True
    except Exception:
        return False


def refresh_zrma_list(session):
    """20분 루프용 ZRMA_Q 재조회. 레이아웃 창을 다시 열지 않는다."""
    try:
        session.findById(ZRMA_GRID_PATH)
    except Exception:
        logger.warning("ZRMA_Q 목록 그리드가 없어 재조회 실패")
        return False

    if refresh_sap_list(session, ZRMA_GRID_PATH, "ZRMA_Q", try_f5=False, toolbar_button_id="REF"):
        return True

    logger.info("ZRMA_Q 일반 Refresh 불가 → F3/F8 재조회 시도")
    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1.5)
        logger.info(f"ZRMA_Q F3 후 화면: '{session.findById('wnd[0]').Text}'")
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2)
        session.findById(ZRMA_GRID_PATH)
        logger.info("ZRMA_Q F3/F8 재조회 완료")
        return True
    except Exception as e:
        logger.warning(f"ZRMA_Q F3/F8 재조회 실패: {e}")
        return False

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
        # Collection Due Date - 회수 마감일자. 개별 오더 화면 안에는 안 보이고
        # ZRMA_Q 목록 그리드 자체에 컬럼으로 있음 (라이브 세션에서 확인,
        # 2026-08-06 - grid.GetColumnTitles()로 발견, project_sap_order_
        # collection_gaps 메모 참고).
        COL_COLL_DUE_DATE = "T_ZZCOLL_DUE_DATE"

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
                    'coll_due_date': safe_get(COL_COLL_DUE_DATE),
                })
            except Exception as e:
                logger.warning(f"ZRMA 행 {i} 읽기 오류: {e}")

    except Exception as e:
        logger.error(f"ZRMA 그리드 읽기 실패: {e}")

    return rows


def _close_tech_obj_popup(session):
    """
    Technical objects 팝업(wnd[1]) 닫기.
    F12 → 'Would you like to terminate processing?' → Yes 클릭.
    """
    try:
        session.findById("wnd[1]").sendVKey(12)  # F12
        time.sleep(0.8)
    except Exception:
        return

    # "Would you like to terminate processing?" 대화상자 처리
    for popup_id in ["wnd[2]", "wnd[1]"]:
        try:
            session.findById(popup_id)  # 팝업 존재 확인
        except Exception:
            continue

        yes_clicked = False
        for btn_id in [
            f"{popup_id}/usr/btnSPOPLI-SELFLAG",  # SPOPLI Yes 버튼
            f"{popup_id}/tbar[0]/btn[0]",           # 첫 번째 툴바 버튼
        ]:
            try:
                session.findById(btn_id).press()
                yes_clicked = True
                time.sleep(0.5)
                logger.info(f"terminate processing → Yes ({btn_id})")
                break
            except Exception:
                continue

        if not yes_clicked:
            # Enter (Yes가 기본값인 경우)
            try:
                session.findById(popup_id).sendVKey(0)
                time.sleep(0.5)
                logger.info("terminate processing → Enter")
            except Exception:
                pass
        break


def _write_sn_to_popup(session, new_sn):
    """
    'Maintain Serial Numbers' 팝업(wnd[1])의 첫 번째 빈 행에 new_sn 입력.
    입력 후 Enter → Yes 팝업 처리.
    반환: True(성공) / False(실패)

    2026-08-19 실측(inspect_return_serial_popup.py로 실제 오더 67072626의
    팝업 구조 라이브 확인) 기반 수정: 실제 컬럼 기술명은 'RIPW0-SERNR'다
    ('SERNR'은 존재하지 않음 - GetCell(i,"SERNR")은 원래도 실패하고
    있었지만 그 예외를 잡아 GetCell(i,0)으로 넘어가서 셀 자체는 잘
    찾고 있었음). 진짜 버그는 그 아래 tbl.setCurrentCell(i, "SERNR")
    호출 - GetCell과 별개로 다시 문자열 "SERNR"을 그대로 넘겨서 실제
    컬럼명과 안 맞아 예외가 났고, 이게 cell.text=new_sn **이전에** 터지는
    바람에 시리얼 값이 아예 입력되지 않은 채(사용자 실측: "serial
    number를 붙여넣지 않고") 예외 핸들러로 빠져 다음 행을 시도하다가
    결국 3행 다 실패 → False 리턴하는 흐름이었다(주문 저장까지는 도달
    안 함 - Ctrl+S는 이 함수가 True를 반환해야만 호출되므로 실제 오더
    데이터 손상은 없었을 것으로 판단됨). setCurrentCell 호출을 제거하고
    _read_serial_numbers_from_popup()과 동일한 3단계 컬럼명 폴백으로
    통일 + 입력 직후 읽어보기(read-back)로 실제로 값이 들어갔는지
    확인한 뒤에만 Enter를 누르도록 강화."""
    for tbl_id in _SN_TABLE_IDS:
        try:
            tbl = session.findById(tbl_id)
        except Exception:
            continue

        for i in range(tbl.RowCount):
            try:
                try:
                    cell = tbl.GetCell(i, "SERNR")
                except Exception:
                    try:
                        cell = tbl.GetCell(i, "RIPW0-SERNR")
                    except Exception:
                        cell = tbl.GetCell(i, 0)

                if cell.Text.strip() != '':
                    continue  # 이미 값 있는 행 건너뜀

                # 빈 행 발견 → S/N 입력 (setCurrentCell 없이 cell에 바로 씀)
                cell.setFocus()
                cell.text = new_sn
                time.sleep(0.3)

                # 실제로 입력이 반영됐는지 읽어보기 - 반영 안 됐으면 Enter를
                # 누르지 않고 안전하게 다음 행/테이블로 넘어간다(빈 값으로
                # 확정해버리는 사고 방지).
                written = ''
                try:
                    written = cell.Text.strip()
                except Exception:
                    pass
                if written != str(new_sn).strip():
                    logger.warning(f"S/N 입력 검증 실패 (행 {i}): 입력값이 반영 안 됨 (읽은 값='{written}')")
                    continue

                # Enter로 입력 확정
                session.findById("wnd[1]").sendVKey(0)
                time.sleep(1.2)

                # 팝업(wnd[2])이 뜨면 Yes → wnd[1] 자동 닫힘
                # 팝업 없으면 그냥 진행
                try:
                    popup = session.findById("wnd[2]")
                    yes_clicked = False
                    for btn_id in [
                        "wnd[2]/usr/btnSPOPLI-SELFLAG",
                        "wnd[2]/tbar[0]/btn[0]",
                    ]:
                        try:
                            session.findById(btn_id).press()
                            yes_clicked = True
                            time.sleep(0.5)
                            logger.info(f"S/N 입력 팝업 Yes ({btn_id})")
                            break
                        except Exception:
                            continue
                    if not yes_clicked:
                        session.findById("wnd[2]").sendVKey(0)
                        time.sleep(0.5)
                except Exception:
                    pass  # 팝업 없음 → 그냥 진행

                logger.info(f"S/N '{new_sn}' 입력 완료 (행 {i})")
                return True

            except Exception as e:
                logger.warning(f"S/N 입력 실패 (행 {i}): {e}")
                continue

    logger.error("S/N 입력할 빈 셀 없음 또는 테이블 미발견")
    return False


def open_order_via_va02(session, order_num):
    """VA02로 오더 직접 열기. 성공 시 True 반환."""
    from sap_handler import run_transaction
    try:
        run_transaction(session, "/nVA02")
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order_num
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        logger.info(f"VA02 오더 {order_num} 열기")
        return True
    except Exception as e:
        logger.error(f"VA02 오더 {order_num} 열기 실패: {e}")
        return False


def input_sn_to_return_item(session, order_num, new_sn):
    """
    오더의 회수 아이템 Technical objects에 S/N 입력.

    흐름:
      VA02 오더 열기 → Item Overview 탭 → 회수 아이템 행 찾기(RETURN route, BUNIT 제외)
      → Extras > Technical objects → 기존 S/N 확인
      → S/N 없으면 new_sn 입력 → Enter → 팝업 Yes → Technical objects 닫기 → 오더 저장

    반환: {
      'status': 'skipped' | 'saved' | 'error',
      'existing_sns': [...],   # skipped일 때 기존 S/N 목록
      'message': '...',
    }
    """
    # 1. VA02로 오더 열기
    if not open_order_via_va02(session, order_num):
        return {'status': 'error', 'existing_sns': [], 'message': f'VA02 오더 {order_num} 열기 실패'}

    # 2. Item Overview 탭 활성화
    try:
        session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
        ).select()
        time.sleep(0.5)
    except Exception:
        pass

    # 3. 회수 아이템 행 찾기 (RETURN route, BUNIT 제외)
    return_row_i = None
    try:
        table = session.findById(ZRMA_ITEMS_TABLE_PATH)
        for i in range(table.RowCount):
            try:
                item_no = table.GetCell(i, 0).Text.strip()
                if not item_no or item_no.endswith('00') or item_no == '000000':
                    continue
                arktx = table.GetCell(i, 2).Text.strip()
                route = table.GetCell(i, 12).Text.strip().upper()
                if route == 'RETURN' and 'BUNIT' not in arktx.upper():
                    return_row_i = i
                    logger.info(f"회수 아이템 발견: 행 {i}, '{arktx}'")
                    break
            except Exception:
                break
    except Exception as e:
        return {'status': 'error', 'existing_sns': [], 'message': f'아이템 테이블 읽기 실패: {e}'}

    if return_row_i is None:
        return {'status': 'error', 'existing_sns': [], 'message': '회수 아이템(RETURN route, BUNIT 제외)을 찾을 수 없음'}

    # 4. 행 선택 → Extras > Technical objects
    try:
        table = session.findById(ZRMA_ITEMS_TABLE_PATH)
        table.GetCell(return_row_i, 0).setFocus()
        time.sleep(0.3)
        session.findById(EXTRAS_TECH_OBJ).select()
        time.sleep(1.5)
    except Exception as e:
        return {'status': 'error', 'existing_sns': [], 'message': f'Technical objects 진입 실패: {e}'}

    # 5. 기존 S/N 확인
    existing_sns = _read_serial_numbers_from_popup(session)
    logger.info(f"기존 S/N: {existing_sns if existing_sns else '없음'}")

    if existing_sns:
        _close_tech_obj_popup(session)
        return {
            'status': 'skipped',
            'existing_sns': existing_sns,
            'message': f'이미 S/N 있음: {", ".join(existing_sns)}',
        }

    # 6. S/N 입력
    if not _write_sn_to_popup(session, new_sn):
        _close_tech_obj_popup(session)
        return {'status': 'error', 'existing_sns': [], 'message': 'S/N 입력 실패'}

    # 7. 오더 저장 (Ctrl+S) - 실제로 저장됐는지 확인까지 해야 한다. S/N의
    # firm/cust와 오더의 firm/cust가 다르면 예외 없이 SAP 상태바에
    # 에러 메시지만 뜨고 저장은 안 되는 케이스(paper relo 필요)라서,
    # 예외가 안 났다고 무조건 'saved'로 보면 안 된다.
    try:
        session.findById("wnd[0]").sendVKey(11)
        time.sleep(1.5)
    except Exception as e:
        return {'status': 'error', 'existing_sns': [], 'message': f'오더 저장 실패: {e}'}

    # 저장 중 예상 못한 팝업(wnd[1])이 뜨면 - firm/cust 불일치 등으로 저장이
    # 막혔을 가능성. 팝업 문구만 읽고 안전하게 Enter로 닫은 뒤 에러 처리 -
    # paper relo 여부는 이 문구를 보고 사람이 최종 판단해야 한다.
    try:
        popup = session.findById("wnd[1]")
        popup_text = popup.Text
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(0.5)
        except Exception:
            pass
        logger.warning(f"오더 {order_num} 저장 중 팝업: {popup_text}")
        return {
            'status': 'error',
            'existing_sns': [],
            'message': f"저장 중 팝업 발생 - {popup_text} (firm/cust 불일치 가능 - paper relo 확인 필요)",
        }
    except Exception:
        pass  # 팝업 없음 - 정상 진행

    sbar_text, sbar_type = '', ''
    try:
        sbar = session.findById("wnd[0]/sbar")
        sbar_text = (sbar.Text or '').strip()
        sbar_type = (sbar.MessageType or '').upper()
    except Exception:
        pass

    if sbar_type in ('E', 'A'):
        logger.error(f"오더 {order_num} 저장 실패 (상태바 에러): {sbar_text}")
        return {
            'status': 'error',
            'existing_sns': [],
            'message': f"저장 실패 - {sbar_text} (firm/cust 불일치 가능 - paper relo 확인 필요)",
        }

    logger.info(f"오더 {order_num} 저장 완료 - 상태바: {sbar_text}")
    return {
        'status': 'saved',
        'existing_sns': [new_sn],
        'message': f"S/N '{new_sn}' 입력 및 저장 완료" + (f" ({sbar_text})" if sbar_text else ""),
    }


def _read_serial_numbers_from_popup(session, retries=2, retry_delay=0.4):
    """
    'Maintain/Display Serial Numbers' 팝업(wnd[1])에서 S/N 목록 반환.
    GuiTableControl(편집 모드) 및 GuiGridView ALV(표시 모드) 모두 지원.
    반환: ['SN001', 'SN002', ...] 또는 []

    NWBC에서 팝업이 뜬 직후 컨트롤이 아직 완전히 렌더링되지 않아 첫 시도에
    테이블/그리드를 못 찾는 경우가 있어(2026-08-19 실측) 짧게 재시도한다.
    """
    for attempt in range(retries + 1):
        sn_list = _read_serial_numbers_from_popup_once(session)
        if sn_list:
            return sn_list
        if attempt < retries:
            logger.info(f"    S/N 읽기 재시도 ({attempt + 1}/{retries})")
            time.sleep(retry_delay)

    logger.warning("    S/N 읽기 실패: 팝업 테이블/그리드를 찾지 못함")
    return []


def _read_serial_numbers_from_popup_once(session):
    sn_list = []

    # 1) GuiTableControl 시도 (VA02 편집 모드)
    for tbl_id in _SN_TABLE_IDS:
        try:
            tbl = session.findById(tbl_id)
            for i in range(tbl.RowCount):
                try:
                    try:
                        sn = tbl.GetCell(i, "SERNR").Text.strip()
                    except Exception:
                        try:
                            sn = tbl.GetCell(i, "RIPW0-SERNR").Text.strip()
                        except Exception:
                            sn = tbl.GetCell(i, 0).Text.strip()
                    if sn and sn not in sn_list:
                        sn_list.append(sn)
                except Exception:
                    continue
            if sn_list:
                logger.info(f"    S/N TableControl에서 {len(sn_list)}개 읽음")
                return sn_list
            break  # 테이블은 찾았으나 데이터 없음
        except Exception:
            continue

    # 2) GuiGridView ALV 시도 (VA03 표시 모드 또는 신형 UI)
    for grid_id in _SN_GRID_IDS:
        try:
            grid = session.findById(grid_id)
            for i in range(grid.RowCount):
                try:
                    # 컬럼명 순서대로 시도
                    sn = ''
                    for col in ("SERNR", "SERIAL_NO", "SERIALNO"):
                        try:
                            sn = grid.GetCellValue(i, col).strip()
                            if sn:
                                break
                        except Exception:
                            continue
                    if sn and sn not in sn_list:
                        sn_list.append(sn)
                except Exception:
                    continue
            if sn_list:
                logger.info(f"    S/N ALV Grid에서 {len(sn_list)}개 읽음")
                return sn_list
            break
        except Exception:
            continue

    return sn_list


def _open_new_session(session):
    """현재 세션에서 새 SAP 세션 생성 후 반환. 실패 시 None."""
    try:
        session.createSession()
        time.sleep(2)
        conn = get_scripting_engine().Children(0)
        count = conn.Children.Count
        new_sess = conn.Children(count - 1)
        logger.info(f"새 SAP 세션 생성 완료 (세션{count - 1})")
        return new_sess
    except Exception as e:
        logger.error(f"새 SAP 세션 생성 실패: {e}")
        return None


def _close_session(sess):
    """SAP 세션(창) 닫기."""
    try:
        sess.findById("wnd[0]").close()
        time.sleep(0.5)
    except Exception:
        pass


def collect_serial_numbers(session, items):
    """
    RETURN 아이템(Bunit 제외)의 S/N을 수집.
    원본 ZRMA_Q 세션(session)에서 직접 Technical objects에 접근.
    수동 흐름과 동일: ZRMA 오더 뷰 → 행 선택 → Extras > Technical objects → S/N 읽기 → 팝업 닫기.
    """
    needs = [it for it in items if it.get('serial_numbers') is None]
    if not needs:
        return

    order_num = needs[0]['order_num']
    logger.info(f"  S/N 수집 (원본 세션): 오더 {order_num}")

    for idx, item in enumerate(needs):
        row_i = item['table_row_i']
        scroll_top = item.get('table_scroll_top', 0)
        logger.info(f"  S/N 수집: [{item['arktx']}] (행 {row_i}, 스크롤 {scroll_top})")
        try:
            table = session.findById(ZRMA_ITEMS_TABLE_PATH)
            try:
                if table.VerticalScrollbar.Position != scroll_top:
                    table.VerticalScrollbar.Position = scroll_top
                    time.sleep(0.2)
            except Exception as e:
                logger.warning(f"    스크롤 복원 실패 (top={scroll_top}): {e}")
            # 행 구조 탐색 (최초 1회만)
            if idx == 0:
                try:
                    row_obj = table.rows.elementAt(row_i)
                    logger.info(f"    row type={row_obj.Type}  selected={row_obj.selected}")
                    for ci in range(min(row_obj.Children.Count, 6)):
                        cell = row_obj.Children.ElementAt(ci)
                        txt = ''
                        try:
                            txt = cell.Text
                        except Exception:
                            pass
                        logger.info(f"    row.Children[{ci}]: id={cell.Id}  type={cell.Type}  text='{txt}'")
                except Exception as e:
                    logger.info(f"    row 탐색 실패: {e}")
            table.GetCell(row_i, 0).setFocus()
            time.sleep(0.3)
            try:
                table.rows.elementAt(row_i).selected = True
                logger.info(f"    행 {row_i} 선택 완료")
            except Exception as e:
                logger.warning(f"    행 선택 실패 (setFocus만 적용): {e}")

            session.findById(EXTRAS_TECH_OBJ).select()
            time.sleep(1.5)

            # 팝업 확인
            try:
                popup = session.findById("wnd[1]")
                logger.info(f"    팝업 제목: '{popup.Text}'")
                # tbar 버튼 목록 로깅
                try:
                    tbar = session.findById("wnd[1]/tbar[0]")
                    for bi in range(tbar.Children.Count):
                        btn = tbar.Children.ElementAt(bi)
                        tooltip = ''
                        try:
                            tooltip = btn.Tooltip
                        except Exception:
                            pass
                        logger.info(f"    tbar[0] 버튼[{bi}]: id={btn.Id}  type={btn.Type}  tooltip='{tooltip}'")
                except Exception as e:
                    logger.info(f"    tbar[0] 없음: {e}")
                if popup.Text == "Information":
                    logger.info("    Information 팝업 → Enter로 닫기")
                    popup.sendVKey(0)
                    time.sleep(1.5)
            except Exception:
                logger.warning("    팝업(wnd[1]) 없음")

            sn_list = _read_serial_numbers_from_popup(session)
            logger.info(f"    → S/N: {sn_list if sn_list else '없음(X)'}")
            item['serial_numbers'] = sn_list if sn_list else ['X']

        except Exception as e:
            logger.warning(f"  S/N 수집 실패 (행 {row_i}): {e}")
            item['serial_numbers'] = ['X']
        finally:
            _close_tech_obj_popup(session)


def navigate_to_zrma_order(session, grid_idx):
    """목록에서 오더 더블클릭으로 진입."""
    try:
        grid = session.findById(ZRMA_GRID_PATH)
        # TODO: Sales Document 컬럼명 확인 후 업데이트
        grid.setCurrentCell(grid_idx, "T_VBELN")
        grid.doubleClickCurrentCell()
        time.sleep(2)
        logger.info(f"ZRMA 오더 진입 (행 {grid_idx})")
        try:
            logger.info(f"  오더 진입 후 화면: {session.findById('wnd[0]').Text}")
        except Exception:
            pass
        return True
    except Exception as e:
        logger.error(f"ZRMA 오더 진입 실패: {e}")
        return False



def _is_numeric_material(matnr):
    """Excel 반영 대상 material: 숫자로만 된 material number."""
    value = str(matnr or '').strip()
    return bool(value) and value.isdigit()


# PC/모니터/키보드/서버/라우터는 매번 나가는 정규 배송품이라 수량이 많아도
# 유닛마다 시리얼 번호를 따로 입력해야 해서 절대 한 줄(Qty: N)로 합치면 안
# 된다 (사용자 확인, 2026-07-29). 그 외 가끔 나가는 이벤트성 물품만 수량이
# 많고(qty>10) 시리얼이 없을 때 한 줄로 합친다 (기존 로직 유지).
_ALWAYS_SERIALIZED_KEYWORDS = ('PC', 'MONITOR', 'KEYBOARD', 'SERVER', 'ROUTER')


def _is_always_serialized_item(arktx):
    text = str(arktx or '').upper()
    return any(re.search(rf'\b{kw}\b', text) for kw in _ALWAYS_SERIALIZED_KEYWORDS)


def get_items_from_zrma_order(session, order_num, order_type):
    """
    ZRMA_Q 개별 오더에서 All Items 표 읽기.
    Item이 00으로 끝나는 것(개괄명) 제외.
    Route=Return → 회수, Route=Nextday → 배송
    반환: [{'prefix': '배송'/'회수', 'order_type': 'ZRX', 'order_num': '6***',
            'matnr': '...', 'arktx': '...', 'qty': 1}, ...]
    """
    items = []
    table_id = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
        "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
        "/subSUBSCREEN_TC:SAPMV45A:4900"
        "/tblSAPMV45ATCTRL_U_ERF_AUFTRAG"
    )
    try:
        table = session.findById(table_id)

        # GuiTableControl: 컬럼 인덱스로 접근
        # 0=POSNR, 1=MABNR(Material), 2=ARKTX(Desc), 3=KWMENG(Qty), 12=ROUTE
        #
        # 주의 (7825897로 실측 확인, 2026-07-29): table.RowCount는 "화면에
        # 보이는 행 수"가 아니라 전체 아이템 개수(예: 13)를 반환하는데,
        # GetCell(i, ...)은 그 중 실제로 화면에 그려진 상대 인덱스(0~약3)만
        # 읽을 수 있고 그 이상은 COM 예외("invalid argument")가 난다. RowCount를
        # 페이지당 반복 횟수로 쓰면 인덱스 4 근처에서 예외가 나 "다 읽었다"고
        # 착각하고 스크롤 시도조차 못 해보고 멈춘다 - 그래서 화면보다 많은
        # 아이템(BOM 하위 구성품 등)이 조용히 누락됐다. FirstVisibleRow 속성은
        # 이 컨트롤엔 아예 없다(AttributeError) - VerticalScrollbar가 유일한
        # 스크롤 수단이고, 한 페이지에 실제로 몇 줄이 보이는지는 GetCell이
        # 예외를 낼 때까지 직접 세어봐야 한다.
        #
        # 그리고 스크롤(=서버 왕복) 직후에는 예전 table/scrollbar COM 참조가
        # stale해져서 GetCell(0,..)조차 바로 예외를 낸다 (Position 세팅 성공,
        # 경고 없이 조용히 페이지가 비어있는 것처럼 보임) - 그래서 스크롤할
        # 때마다 session.findById로 table을 다시 잡는다.
        try:
            scrollbar = table.VerticalScrollbar
            scroll_max = int(scrollbar.Maximum)
        except Exception:
            scrollbar = None
            scroll_max = 0

        seen_item_nos = set()
        scroll_top = 0
        max_pages = 50  # 안전장치: 스크롤이 안 멈추는 이상 상황 방지

        for _page in range(max_pages):
            if scrollbar is not None and scroll_top > 0:
                try:
                    scrollbar.Position = scroll_top
                    time.sleep(0.3)
                    table = session.findById(table_id)
                    scrollbar = table.VerticalScrollbar
                except Exception as e:
                    logger.warning(f"ZRMA 테이블 스크롤 실패 (top={scroll_top}): {e}")
                    break

            found_new = False
            rows_seen_this_page = 0
            i = 0
            while True:
                # GetCell(i,0) 실패 = 이 페이지에서 화면에 그려진 범위 초과 → 다음 페이지로
                try:
                    item_no = table.GetCell(i, 0).Text.strip()
                except Exception:
                    break
                rows_seen_this_page = i + 1

                # 빈 행 또는 상위/묶음 행 제외.
                # ZRX/ZRE 실제 반영 대상은 101/102/201/202 같은 하위 item이다.
                if not item_no or item_no.endswith('00') or item_no == '000000':
                    i += 1
                    continue

                # 스크롤 겹침으로 이전에 이미 읽은 행이면 스킵 (already-seen이면
                # 아래 아이템 처리는 건너뛰고 i만 증가시킨다).
                if item_no in seen_item_nos:
                    i += 1
                    continue
                seen_item_nos.add(item_no)
                found_new = True

                try:
                    matnr = table.GetCell(i, 1).Text.strip()
                    if not _is_numeric_material(matnr):
                        logger.info(f"오더 {order_num} material 제외 (숫자 아님): {matnr}")
                        i += 1
                        continue
                    arktx   = table.GetCell(i, 2).Text.strip()
                    qty_str = table.GetCell(i, 3).Text.strip()
                    try:
                        route = table.GetCell(i, 12).Text.strip()
                    except Exception:
                        route = ''

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

                    # 회수 아이템 S/N 초기 설정
                    is_return = (prefix == '회수')
                    is_bunit  = 'BUNIT' in arktx.upper()
                    if is_return and is_bunit:
                        sn_list = ['X']   # Bunit: S/N 없음 (Technical objects 진입 불필요)
                    elif is_return:
                        sn_list = None    # 나중에 collect_serial_numbers에서 수집
                    else:
                        sn_list = []      # 배송: S/N 불필요

                    items.append({
                        'prefix':          prefix,
                        'order_type':      order_type,
                        'order_num':       order_num,
                        'matnr':           matnr,
                        'arktx':           arktx,
                        'qty':             qty,
                        'table_row_i':     i,          # Technical objects 진입용 행 인덱스 (해당 스크롤 위치 기준)
                        'table_scroll_top': scroll_top,  # 위 행 인덱스가 유효한 스크롤 위치
                        'serial_numbers':  sn_list,    # None=미수집, []=S/N없음, ['SN...']=수집됨
                    })

                except Exception as e:
                    logger.warning(f"ZRMA 아이템 {i} 읽기 오류: {e}")

                i += 1

            if not found_new:
                break
            if scrollbar is None or scroll_top >= scroll_max:
                break
            scroll_top = min(scroll_top + rows_seen_this_page, scroll_max)

        if scrollbar is not None:
            try:
                scrollbar.Position = 0
            except Exception:
                pass

    except Exception as e:
        logger.error(f"ZRMA 아이템 읽기 실패: {e}")

    return items


def build_excel_rows_zrma(items, address, extra_orders, memo, obd_map=None):
    """
    ZRMA_Q 아이템 리스트 → Excel 행 리스트.
    아이템별 qty만큼 행 생성. 첫 번째 행에만 고객/주소 정보.
    배송/회수 행이 섞여있을 수 있으므로 prefix별로 그룹화하지 않고 순서대로.
    serial_numbers: qty 수에 맞게 분배 (부족하면 마지막 값 반복, 없으면 '')

    PC/모니터/키보드/서버/라우터(_is_always_serialized_item)는 수량이 많아도
    한 줄로 합치지 않고 항상 유닛별로 한 줄씩 생성한다 - 나중에 시리얼 번호를
    한 줄씩 입력해야 하기 때문. 그 외 가끔 나가는 이벤트 물품만 수량>10이고
    시리얼이 없을 때 "Qty: N" 한 줄로 합친다.
    """
    rows = []
    is_first = True
    obd_map = obd_map or {}
    obd_cursor = {}

    def next_obd_for_item(item):
        if item.get('prefix') != '배송':
            return ''
        value = obd_map.get(item.get('order_num'), '')
        if isinstance(value, (list, tuple)):
            key = item.get('order_num')
            idx = obd_cursor.get(key, 0)
            obd_cursor[key] = idx + 1
            if not value:
                return ''
            return str(value[idx] if idx < len(value) else value[-1]).strip()
        return str(value or '').strip()

    # 배송(NEXDAY) 먼저, 회수(RETURN) 나중
    items = sorted(items, key=lambda x: 0 if x['prefix'] == '배송' else 1)

    for item in items:
        sn_list = item.get('serial_numbers') or []  # None → []
        large_no_serial = (
            item.get('qty', 1) > 10
            and not _is_always_serialized_item(item.get('arktx'))
            and (not sn_list or all(str(sn or '').strip().upper() in ('', 'X') for sn in sn_list))
        )
        unit_count = 1 if large_no_serial else item['qty']

        for unit_idx in range(unit_count):
            # S/N 분배: unit_idx에 맞는 S/N 선택, 부족하면 마지막 값, 없으면 ''
            if large_no_serial:
                sn = ''
            elif sn_list:
                sn = sn_list[unit_idx] if unit_idx < len(sn_list) else sn_list[-1]
            else:
                sn = ''

            rows.append({
                'order_prefix':  item['prefix'],
                'order_type':    item['order_type'],
                'order_num':     item['order_num'],
                'obd':           next_obd_for_item(item),
                'extra_orders':  extra_orders,
                'material':      item['matnr'],
                'description':   item['arktx'],
                'quantity':      item['qty'] if large_no_serial else 1,
                'serial_number': sn,
                'customer':      address.get('customer', ''),
                'phone':         address.get('phone', ''),
                'company':       address.get('company', ''),
                'street':        address.get('street', ''),
                'street2':       address.get('street2', ''),
                'cust_no':       address.get('cust_no', ''),
                'memo':          memo,
                'is_first_item': is_first,
            })
            is_first = False

    return rows


def _safe_return_to_zrma_list(session, variant=None, date_mode=None):
    """상세 화면에서 목록으로 복귀한다. F3이 막히면 ZRMA_Q를 다시 연다."""
    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)
        if is_on_zrma_list(session):
            return True
    except Exception as e:
        logger.warning(f"ZRMA 목록 복귀(F3) 실패: {e}")

    if variant and date_mode:
        try:
            logger.info(f"ZRMA_Q ({variant}) 목록 재진입")
            navigate_to_zrma_q(session, variant, date_mode, apply_layout=False)
            return True
        except Exception as e:
            logger.error(f"ZRMA_Q ({variant}) 목록 재진입 실패: {e}")
    return False


def _read_zrma_order_detail(session, order_num, order_type, obd_map=None, coll_due_date=""):
    """현재 열린 RMA 오더 화면에서 Excel 입력용 row를 만든다."""
    address = get_ship_to_address_zrma(session)

    text = get_text_content(session, menu_id=ZRMA_MENU_TEXTS)
    all_extra = parse_extra_orders(text) if text else []
    extra_orders = [eo for eo in all_extra if order_num not in eo]
    text_contact = parse_contact_from_text(text) if text else {'found': False}
    memo = parse_memo_for_display(text) if text else ""
    if text_contact['found'] and not address.get('phone'):
        address['phone'] = text_contact['phone']

    # SDSK(대부분) 또는 ORD(ZOR인 경우) 번호 / Ship-to Party 번호(Cust#) /
    # Delivery Date - 오더가 이미 열려있는 이 session을 그대로 재사용 (VL06O/ZOR
    # 경로와 달리 세션을 새로 열 필요 없음). 사용자 확인, 2026-08-06: ZRE/ZRX/
    # ZINX 전부 동일 규칙 - PO Number 필드는 ZOR만 ORD, 나머지는 SDSK.
    extra_fields = collect_order_extra_fields(order_num, order_type, existing_session=session)
    if extra_fields.get('cust_no'):
        address['cust_no'] = extra_fields['cust_no']
    if extra_fields.get('po_number'):
        extra_orders = add_extra_order_dedup(extra_orders, extra_fields['po_label'], extra_fields['po_number'])
    memo_notes = []
    if extra_fields.get('delivery_date'):
        memo_notes.append(f"(delivery date: {extra_fields['delivery_date']})")
    if coll_due_date:
        memo_notes.append(f"(coll due date: {coll_due_date})")
    if memo_notes:
        memo = f"{memo}\n" + "\n".join(memo_notes) if memo else "\n".join(memo_notes)

    items = get_items_from_zrma_order(session, order_num, order_type)
    if not items:
        return []

    collect_serial_numbers(session, items)
    return build_excel_rows_zrma(items, address, extra_orders, memo, obd_map=obd_map)


def process_zrma_orders(session, new_order_nums, order_map, variant=None, date_mode=None, obd_map=None):
    """
    새 ZRMA 오더 처리.
    목록 더블클릭이 Header Data 등 예상과 다른 화면으로 들어가면 VA02 직접 진입으로 재시도한다.
    """
    results = {}

    for order_num in new_order_nums:
        logger.info(f"=== ZRMA 오더 {order_num} 처리 중 ===")
        order_info = order_map.get(order_num)
        if not order_info:
            continue

        grid_idx = order_info['grid_idx']
        order_type = order_info['order_type']
        coll_due_date = order_info.get('coll_due_date', '')

        if 'attempt' in order_info.get('rma_status', '').lower():
            logger.info(f"오더 {order_num} 1st Attempt → 건너뜀")
            continue

        rows = []
        opened_from_list = navigate_to_zrma_order(session, grid_idx)
        if opened_from_list:
            rows = _read_zrma_order_detail(session, order_num, order_type, obd_map=obd_map, coll_due_date=coll_due_date)
            if not rows:
                logger.warning(f"오더 {order_num} 목록 진입 화면에서 아이템 없음 → VA02 직접조회 재시도")
        else:
            logger.error(f"오더 {order_num} 목록 진입 실패 → VA02 직접조회 재시도")

        if not rows:
            if order_num.startswith('6') and open_order_via_va02(session, order_num):
                rows = _read_zrma_order_detail(session, order_num, order_type, obd_map=obd_map, coll_due_date=coll_due_date)
            else:
                logger.error(f"오더 {order_num} VA02 직접조회 실패")

        if rows:
            results[order_num] = rows
            logger.info(f"ZRMA 오더 {order_num}: {len(rows)}행 생성")
        else:
            logger.warning(f"오더 {order_num} Excel 행 생성 실패")

        _safe_return_to_zrma_list(session, variant, date_mode)

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
                'coll_due_date': row.get('coll_due_date', ''),
            }
    return orders

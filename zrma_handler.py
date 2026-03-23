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
    parse_memo_for_display,
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

# Extras > Technical objects 메뉴 (menu[3]=Extras, menu[7]=Technical objects)
EXTRAS_TECH_OBJ = "wnd[0]/mbar/menu[3]/menu[7]"

# Maintain/Display Serial Numbers 팝업 — GuiTableControl 경로 후보
_SN_TABLE_IDS = [
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


def navigate_to_zrma_q(session, variant, date_mode='next_month'):
    """
    ZRMA_Q 진입.
    variant: 'rlkr' 또는 'q2'
    date_mode: 'next_month' (RLKR) 또는 'year_end' (Q2)
    """
    logger.info(f"ZRMA_Q 진입 중 (variant={variant})...")
    run_transaction(session, "ZRMA_Q")
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
    """F3으로 선택화면 복귀 → F8 재실행 (DB 재조회)."""
    session.findById("wnd[0]").sendVKey(3)   # F3 = Back
    time.sleep(1.5)
    title = session.findById("wnd[0]").Text
    logger.info(f"ZRMA_Q F3 후 화면: '{title}'")
    session.findById("wnd[0]").sendVKey(8)   # F8 = Execute (sendVKey로 통일)
    time.sleep(2)
    logger.info("ZRMA_Q F3+F8 재실행 완료")


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
    """
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
                    cell = tbl.GetCell(i, 0)

                if cell.Text.strip() != '':
                    continue  # 이미 값 있는 행 건너뜀

                # 빈 행 발견 → S/N 입력
                tbl.setCurrentCell(i, "SERNR")
                cell.text = new_sn
                time.sleep(0.3)

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
        run_transaction(session, "VA02")
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

    # 7. 오더 저장 (Ctrl+S)
    try:
        session.findById("wnd[0]").sendVKey(11)
        time.sleep(1.5)
        logger.info(f"오더 {order_num} 저장 완료")
    except Exception as e:
        return {'status': 'error', 'existing_sns': [], 'message': f'오더 저장 실패: {e}'}

    return {
        'status': 'saved',
        'existing_sns': [new_sn],
        'message': f"S/N '{new_sn}' 입력 및 저장 완료",
    }


def _read_serial_numbers_from_popup(session):
    """
    'Maintain/Display Serial Numbers' 팝업(wnd[1])에서 S/N 목록 반환.
    GuiTableControl(편집 모드) 및 GuiGridView ALV(표시 모드) 모두 지원.
    반환: ['SN001', 'SN002', ...] 또는 []
    """
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

    if not sn_list:
        logger.warning("    S/N 읽기 실패: 팝업 테이블/그리드를 찾지 못함")
    return sn_list


def _open_new_session(session):
    """현재 세션에서 새 SAP 세션 생성 후 반환. 실패 시 None."""
    try:
        import win32com.client
        session.createSession()
        time.sleep(2)
        sap = win32com.client.GetObject("SAPGUI")
        conn = sap.GetScriptingEngine.Children(0)
        new_sess = conn.Children(conn.Count - 1)
        logger.info(f"새 SAP 세션 생성 완료 (세션{conn.Count - 1})")
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
    새 SAP 세션에서 VA02를 열어 Technical objects 접근.
    원본 ZRMA_Q 세션(session)은 그대로 유지.
    """
    needs = [it for it in items if it.get('serial_numbers') is None]
    if not needs:
        return

    order_num = needs[0]['order_num']
    logger.info(f"  VA02 새 세션으로 S/N 수집: 오더 {order_num}")

    va02 = _open_new_session(session)
    if not va02:
        for item in needs:
            item['serial_numbers'] = ['X']
        return

    try:
        # VA02 열기
        run_transaction(va02, "VA02")
        time.sleep(1)
        va02.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order_num
        va02.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)

        # Item Overview 탭
        try:
            va02.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
            ).select()
            time.sleep(0.5)
        except Exception:
            pass

        for item in needs:
            row_i = item['table_row_i']
            logger.info(f"  S/N 수집: [{item['arktx']}] (행 {row_i})")
            try:
                table = va02.findById(ZRMA_ITEMS_TABLE_PATH)
                table.GetCell(row_i, 0).setFocus()
                time.sleep(0.3)

                va02.findById(EXTRAS_TECH_OBJ).select()
                time.sleep(1.5)

                # 팝업 확인
                try:
                    popup = va02.findById("wnd[1]")
                    logger.info(f"    팝업 제목: '{popup.Text}'")
                    try:
                        usr = va02.findById("wnd[1]/usr")
                        for ci in range(min(usr.Children.Count, 8)):
                            child = usr.Children.ElementAt(ci)
                            txt = ''
                            try:
                                txt = child.Text
                            except Exception:
                                pass
                            logger.info(f"    wnd[1]/usr 하위[{ci}]: id={child.Id}  type={child.Type}  text='{txt}'")
                    except Exception:
                        pass

                    if popup.Text == "Information":
                        logger.info("    Information 팝업 → Enter로 닫기")
                        popup.sendVKey(0)
                        time.sleep(1.0)
                except Exception:
                    logger.warning("    팝업(wnd[1]) 없음")

                sn_list = _read_serial_numbers_from_popup(va02)
                logger.info(f"    → S/N: {sn_list if sn_list else '없음(X)'}")
                item['serial_numbers'] = sn_list if sn_list else ['X']

            except Exception as e:
                logger.warning(f"  S/N 수집 실패 (행 {row_i}): {e}")
                item['serial_numbers'] = ['X']
            finally:
                _close_tech_obj_popup(va02)

    except Exception as e:
        logger.error(f"  VA02 S/N 수집 오류: {e}")
        for item in needs:
            if item.get('serial_numbers') is None:
                item['serial_numbers'] = ['X']
    finally:
        _close_session(va02)


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
            # GetCell(i,0) 실패 = 가시 범위 초과 → 루프 종료
            try:
                item_no = table.GetCell(i, 0).Text.strip()
            except Exception:
                break

            # 빈 행 또는 개괄명(00) 제외
            if not item_no or item_no.endswith('00') or item_no == '000000':
                continue

            try:
                matnr   = table.GetCell(i, 1).Text.strip()
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
                    'prefix':         prefix,
                    'order_type':     order_type,
                    'order_num':      order_num,
                    'matnr':          matnr,
                    'arktx':          arktx,
                    'qty':            qty,
                    'table_row_i':    i,        # Technical objects 진입용 행 인덱스
                    'serial_numbers': sn_list,  # None=미수집, []=S/N없음, ['SN...']=수집됨
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
    serial_numbers: qty 수에 맞게 분배 (부족하면 마지막 값 반복, 없으면 '')
    """
    rows = []
    is_first = True

    # 배송(NEXDAY) 먼저, 회수(RETURN) 나중
    items = sorted(items, key=lambda x: 0 if x['prefix'] == '배송' else 1)

    for item in items:
        sn_list = item.get('serial_numbers') or []  # None → []

        for unit_idx in range(item['qty']):
            # S/N 분배: unit_idx에 맞는 S/N 선택, 부족하면 마지막 값, 없으면 ''
            if sn_list:
                sn = sn_list[unit_idx] if unit_idx < len(sn_list) else sn_list[-1]
            else:
                sn = ''

            rows.append({
                'order_prefix':  item['prefix'],
                'order_type':    item['order_type'],
                'order_num':     item['order_num'],
                'extra_orders':  extra_orders,
                'material':      item['matnr'],
                'description':   item['arktx'],
                'serial_number': sn,
                'customer':      address.get('customer', ''),
                'phone':         address.get('phone', ''),
                'company':       address.get('company', ''),
                'street':        address.get('street', ''),
                'street2':       address.get('street2', ''),
                'memo':          memo,
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
        all_extra = parse_extra_orders(text) if text else []
        # 메인 오더번호와 동일한 항목 제거 (중복 방지)
        extra_orders = [eo for eo in all_extra if order_num not in eo]
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = parse_memo_for_display(text) if text else ""
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # 아이템 읽기
        items = get_items_from_zrma_order(session, order_num, order_type)
        if not items:
            logger.warning(f"오더 {order_num} 아이템 없음")
            session.findById("wnd[0]").sendVKey(3)
            time.sleep(0.5)
            continue

        # 회수 아이템 S/N 수집 (Extras > Technical objects)
        collect_serial_numbers(session, items)

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

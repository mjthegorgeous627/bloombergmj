"""
SAP GUI Scripting - VL06O 배송 오더 데이터 추출 모듈.
확인된 element ID 기준으로 작성됨.
"""

import win32com.client
import time
import re
import logging

logger = logging.getLogger(__name__)

# 오더번호 패턴: ZOR/SOR/ZINX/ZINP/SDSK/ZRX/ZRE/ORD + 숫자 (앞의 00 제거)
ORDER_PATTERN = re.compile(
    r'\b(ZOR|SOR|ZINX|ZINP|SDSK|ZRX|ZRE|ORD#?)\s*:?\s*(00)?(\d{6,})\b',
    re.IGNORECASE
)

# 확인된 element ID
GRID_PATH        = "wnd[0]/usr/cntlGRID1/shellcont/shell"
ADDR_BTN         = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/btnBT_WADR_T"
MENU_TEXTS       = "wnd[0]/mbar/menu[2]/menu[1]/menu[9]"   # Goto > Header > Texts (VL06O)
MENU_ENV_SHIP    = "wnd[0]/mbar/menu[4]/menu[0]"           # Environment > Ship-To Party

# ZRMA_Q (VA03 스타일) 메뉴 - VL06O와 인덱스 다름
ZRMA_MENU_PARTNERS = "wnd[0]/mbar/menu[2]/menu[1]/menu[9]"   # Goto > Header > Partners
ZRMA_MENU_TEXTS    = "wnd[0]/mbar/menu[2]/menu[1]/menu[10]"  # Goto > Header > Texts

# ZRMA Texts 탭 텍스트 에디터 경로 (GuiSplitterShell 하위 shellcont[1])
ZRMA_TEXT_EDITOR = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4152"
    "/subSUBSCREEN_TEXT:SAPLV70T:2100"
    "/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
)
ITEM_TABLE       = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER"


def get_sap_session(session_idx=0):
    """
    실행 중인 SAP GUI 세션에 연결.
    session_idx: 0=VL06O, 1=VL10G, 2=ZRMA_Q RLKR, 3=ZRMA_Q Q2
    """
    try:
        sap = win32com.client.GetObject("SAPGUI")
        app = sap.GetScriptingEngine
        conn = app.Children(0)
        session = conn.Children(session_idx)
        logger.info(f"SAP 연결 성공 [세션{session_idx}]: {session.findById('wnd[0]').Text}")
        return session
    except Exception as e:
        raise ConnectionError(f"SAP GUI 연결 실패 (세션 {session_idx}): {e}")


def run_transaction(session, tcode):
    """T-code 실행."""
    session.findById("wnd[0]/tbar[0]/okcd").text = tcode
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)


def is_on_list_screen(session):
    """현재 VL06O 오더 목록 화면인지 확인."""
    try:
        title = session.findById("wnd[0]").Text
        return "List of Outbound Deliveries" in title
    except Exception:
        return False


def navigate_to_vl06o_list(session):
    """
    VL06O → Variant KSCPs1 → F8 실행 → 오더 목록 화면 진입.
    이미 목록 화면이면 F5(새로고침)만 실행.
    """
    if is_on_list_screen(session):
        logger.info("이미 목록 화면 - F5 새로고침")
        session.findById("wnd[0]").sendVKey(5)
        time.sleep(2)
        return

    logger.info("VL06O 진입 중...")
    run_transaction(session, "VL06O")
    time.sleep(1)

    # Ctrl+F7: Display Variants 팝업 열기
    session.findById("wnd[0]").sendVKey(31)
    time.sleep(1)

    # Variant 입력 (My Variant 필드)
    for field_id in [
        "wnd[1]/usr/txtV-LOW",
        "wnd[1]/usr/txtENAME-LOW",
        "wnd[1]/usr/txtVARIANT",
    ]:
        try:
            session.findById(field_id).text = "KSCPs1"
            logger.info(f"Variant 입력 완료 ({field_id})")
            break
        except Exception:
            continue

    time.sleep(0.3)
    # Enter로 선택 확정
    session.findById("wnd[1]").sendVKey(0)
    time.sleep(0.5)

    # F8: Execute
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(2)
    logger.info("목록 화면 진입 완료")


def get_all_rows_from_list(session):
    """
    VL06O 목록 그리드에서 전체 행 데이터 읽기.
    반환: [{'ebeln': '7763549', 'vbeln': '92086300', 'matnr': '10045196',
             'arktx': 'KEYBOARD...', 'lfimg': '2', 'name_we': 'KOREA UNIV'}, ...]
    """
    rows = []
    try:
        grid = session.findById(GRID_PATH)
        row_count = grid.RowCount
        logger.info(f"그리드 행 수: {row_count}")

        for i in range(row_count):
            try:
                row = {
                    'grid_idx': i,                                       # 실제 그리드 행 번호 (더블클릭용)
                    'ebeln':   grid.GetCellValue(i, "EBELN").strip(),   # Pur.Doc (오더번호)
                    'vbeln':   grid.GetCellValue(i, "VBELN").strip(),   # 배송 문서번호
                    'matnr':   grid.GetCellValue(i, "MATNR").strip(),   # Material
                    'arktx':   grid.GetCellValue(i, "ARKTX").strip(),   # Description
                    'lfimg':   grid.GetCellValue(i, "LFIMG").strip(),   # Delivery Qty
                    'name_we': grid.GetCellValue(i, "NAME_WE").strip(), # Ship-to 회사명
                }
                if row['ebeln']:
                    rows.append(row)
            except Exception as e:
                logger.warning(f"행 {i} 읽기 오류: {e}")
                continue

    except Exception as e:
        logger.error(f"그리드 읽기 실패: {e}")

    return rows


def group_by_order(rows):
    """
    그리드 행을 EBELN(오더번호)별로 그룹화.
    row_idx: 그리드에서 해당 오더의 첫 번째 행 인덱스 (더블클릭 진입에 사용)
    반환: {'7763549': {'vbeln': '92086300', 'row_idx': 0, 'items': [...], 'name_we': '...'}, ...}
    """
    orders = {}
    for row in rows:
        ebeln = row['ebeln']
        if ebeln not in orders:
            orders[ebeln] = {
                'vbeln':   row['vbeln'],
                'name_we': row['name_we'],
                'row_idx': row['grid_idx'],  # 실제 그리드 행 인덱스 (enumerate 아님)
                'items':   [],
            }
        orders[ebeln]['items'].append({
            'matnr': row['matnr'],
            'arktx': row['arktx'],
            'lfimg': row['lfimg'],
        })
    return orders


def navigate_to_delivery(session, row_idx, vbeln):
    """
    VL06O 목록 그리드에서 해당 행을 더블클릭하여 배송 상세 화면 진입.
    VL03N을 별도로 열지 않음.
    """
    logger.info(f"배송 문서 {vbeln} 진입 (행 {row_idx} 더블클릭)...")
    try:
        grid = session.findById(GRID_PATH)
        grid.setCurrentCell(row_idx, "VBELN")
        grid.doubleClickCurrentCell()
        time.sleep(2)

        title = session.findById("wnd[0]").Text
        if "Outbound delivery" in title or "Outbound Delivery" in title:
            logger.info(f"상세 화면 진입 성공: {title}")
            return True
        else:
            logger.error(f"예상치 못한 화면: {title}")
            return False
    except Exception as e:
        logger.error(f"더블클릭 진입 실패: {e}")
        return False


def get_text_content(session, menu_id=None):
    """
    Goto → Header → Texts 진입 후 텍스트 내용 추출.
    menu_id: None이면 MENU_TEXTS(VL06O) 사용. ZRMA는 ZRMA_MENU_TEXTS 전달.
    반환: 텍스트 문자열 (없으면 '')
    """
    text_content = ""
    try:
        session.findById(menu_id or MENU_TEXTS).select()
        time.sleep(1)

        # 텍스트 화면에서 내용 읽기 시도 (여러 방식)
        editor_ids = [
            "wnd[0]/usr/cntlSFTXT_0100_EDITOR/shellcont/shell",
            "wnd[0]/usr/txtTDLINES-TDLINE",
        ]
        if menu_id == ZRMA_MENU_TEXTS:
            editor_ids = [ZRMA_TEXT_EDITOR] + editor_ids

        for editor_id in editor_ids:
            try:
                elem = session.findById(editor_id)
                val = getattr(elem, 'Value', '') or getattr(elem, 'Text', '')
                if val and str(val).strip():
                    # \r\n → \n 정규화 (SAP GuiShell Value에 \r 포함됨)
                    text_content = str(val).replace('\r\n', '\n').replace('\r', '\n').strip()
                    break
            except Exception:
                continue

        # 텍스트 화면에서 테이블 방식으로 읽기 (fallback)
        if not text_content:
            try:
                table = session.findById("wnd[0]/usr/tblSAPDOCU_LINES")
                lines = []
                for i in range(table.RowCount):
                    try:
                        line = table.GetCell(i, 0).Text.strip()
                        if line:
                            lines.append(line)
                    except Exception:
                        continue
                text_content = '\n'.join(lines)
            except Exception:
                pass

        # 뒤로가기 (F3)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)

    except Exception as e:
        logger.warning(f"Text 접근 실패 (내용 없을 수 있음): {e}")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass

    return text_content.strip()


def parse_extra_orders(text):
    """
    텍스트에서 오더번호 패턴 추출.
    반환: ['SDSK 001320768583', 'ZOR 7763549', ...]  중복 없이, 앞의 00 제거
    """
    found = []
    seen = set()
    for match in ORDER_PATTERN.finditer(text):
        prefix = match.group(1).upper()
        num    = match.group(3)  # 앞의 00은 group(2)에서 이미 분리됨
        entry  = f"{prefix} {num}"
        if entry not in seen:
            seen.add(entry)
            found.append(entry)
    return found


def normalize_phone(raw):
    """
    전화번호 정규화.
    - 82로 시작 → 한국 번호 형식 변환
    - 81로 시작 → [해외-일본] 표기
    - 기타 국제번호 → [해외] 표기
    - 이미 한국 번호 형식이면 그대로
    """
    if not raw:
        return ''
    digits = re.sub(r'\D', '', raw)
    if not digits:
        return raw.strip()

    if digits.startswith('82'):
        local = '0' + digits[2:]
        return _format_korean(local)
    elif digits.startswith('81'):
        return f'[해외-일본] {raw.strip()}'
    elif len(digits) >= 10 and not digits.startswith('0'):
        # 기타 국제번호 (0으로 시작하지 않는 긴 번호)
        return f'[해외] {raw.strip()}'
    else:
        return _format_korean(digits)


def _format_korean(digits):
    """한국 번호 형식으로 변환 (앞에 0 포함된 숫자열)."""
    if digits.startswith('010') and len(digits) == 11:
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'
    elif digits.startswith('02'):
        rest = digits[2:]
        if len(rest) == 7:
            return f'02-{rest[:3]}-{rest[3:]}'
        elif len(rest) == 8:
            return f'02-{rest[:4]}-{rest[4:]}'
        else:
            return f'02-{rest}'
    elif len(digits) == 11:
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'
    elif len(digits) == 10:
        return f'{digits[:3]}-{digits[3:6]}-{digits[6:]}'
    return digits


def parse_contact_from_text(text):
    """
    텍스트에서 전화번호 등 연락처 추출.
    'Caller Phone:' 명시 패턴 우선, 일반 번호 패턴 fallback.
    반환: {'phone': '', 'found': False}
    """
    result = {'phone': '', 'found': False}

    # 1차: "Caller Phone: +82-2-767-5886" 등 명시 패턴
    caller_re = re.compile(r'Caller\s+Phone\s*[:\-]\s*([+\d][\d\-.\s]{6,})', re.IGNORECASE)
    m = caller_re.search(text)
    if m:
        result['phone'] = normalize_phone(m.group(1).strip())
        result['found'] = True
        return result

    # 2차: 일반 번호 패턴 (+82-2-xxx-xxxx, 010-xxxx-xxxx 등)
    # \d? → 서울(2), 부산(51) 등 한 자리 지역번호 지원
    phone_re = re.compile(
        r'(\+?82[-.\s]?|0)[1-9]\d?[-.\s]?\d{3,4}[-.\s]?\d{4}'
    )
    m = phone_re.search(text)
    if m:
        result['phone'] = normalize_phone(m.group().strip())
        result['found'] = True
    return result


def get_ship_to_address(session):
    """
    Ship-to Party 주소 팝업 읽기.
    1차: ADDR_BTN 버튼 클릭 → 팝업
    2차: Goto > Header > Partners → WE 행 더블클릭 → 팝업
    3차: 헤더 한줄 주소 (txtRV50A-TXTWE) fallback
    반환: {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}
    """
    result = {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}

    def read_field(*ids):
        for fid in ids:
            try:
                val = session.findById(fid).text.strip()
                if val:
                    return val
            except Exception:
                continue
        return ''

    def read_addr_popup():
        """팝업(wnd[1])에서 SAPLSZA1 주소 필드 읽기. (discover_sap.py로 확인된 실제 ID)"""
        base = "wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301"
        phone_raw = read_field(
            f"{base}/txtSZA1_D0100-TEL_NUMBER",
            f"{base}/txtSZA1_D0100-MOB_NUMBER",
        )
        r = {
            'company':  read_field(f"{base}/txtADDR1_DATA-NAME1"),
            'customer': read_field(f"{base}/txtADDR1_DATA-NAME2"),
            'street':   read_field(f"{base}/txtADDR1_DATA-STR_SUPPL1"),
            'street2':  read_field(f"{base}/txtADDR1_DATA-STR_SUPPL2"),
            'phone':    normalize_phone(phone_raw),
        }
        logger.info(f"팝업 주소 읽기 결과: {r}")
        return r

    def close_popup():
        try:
            session.findById("wnd[1]").sendVKey(12)  # F12 닫기
            time.sleep(0.5)
        except Exception:
            pass

    # 1차 시도: Overview 화면 헤더의 Ship-to Party 주소 버튼 (ADDR_BTN)
    try:
        session.findById(ADDR_BTN).press()
        time.sleep(1.5)
        session.findById("wnd[1]")  # 팝업 존재 확인
        r = read_addr_popup()
        close_popup()
        if any(r.values()):
            return r
        logger.warning("ADDR_BTN 팝업 필드 모두 빈값")
    except Exception as e:
        logger.warning(f"ADDR_BTN 시도 실패: {e}")

    # 2차 시도: Goto > Header > Partners → WE(Ship-to) 행 더블클릭
    try:
        session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[8]").select()
        time.sleep(1)

        # Partners 화면: 테이블에서 WE(Ship-to Party) 행 찾아 더블클릭
        partner_table_ids = [
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV50A:1204/tblSAPMV50ATC_VBPA_HEAD_OVER",
            "wnd[0]/usr/tblSAPMV50ATC_VBPA_HEAD_OVER",
        ]
        table = None
        for tid in partner_table_ids:
            try:
                table = session.findById(tid)
                break
            except Exception:
                continue

        we_row = -1
        if table:
            for row_i in range(table.RowCount):
                try:
                    parvw = table.getCellValue(row_i, "PARVW").strip()
                    if parvw == "WE":
                        we_row = row_i
                        break
                except Exception:
                    # 컬럼 이름이 다를 수 있으므로 셀 텍스트로 시도
                    try:
                        cell_val = table.GetCell(row_i, 0).Text.strip()
                        if cell_val == "WE":
                            we_row = row_i
                            break
                    except Exception:
                        continue

        if we_row >= 0 and table:
            try:
                table.setCurrentCell(we_row, "KUNNR")
                table.doubleClickCurrentCell()
                time.sleep(1.5)
                session.findById("wnd[1]")  # 팝업 열렸는지 확인
                r = read_addr_popup()
                close_popup()
                # Partners 화면에서 Overview로 복귀
                session.findById("wnd[0]").sendVKey(3)
                time.sleep(0.5)
                if any(r.values()):
                    return r
            except Exception as e:
                logger.warning(f"Partners WE 더블클릭 실패: {e}")
                try:
                    session.findById("wnd[0]").sendVKey(3)
                    time.sleep(0.5)
                except Exception:
                    pass
        else:
            logger.warning(f"Partners 테이블에서 WE 행 못 찾음 (table={table}, we_row={we_row})")
            try:
                session.findById("wnd[0]").sendVKey(3)
                time.sleep(0.5)
            except Exception:
                pass

    except Exception as e:
        logger.warning(f"Goto>Header>Partners 시도 실패: {e}")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass

    # 3차 fallback: 헤더의 한줄 주소 텍스트 (불완전하지만 없는 것보다 나음)
    try:
        oneline = session.findById(
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/txtRV50A-TXTWE"
        ).text.strip()
        if oneline:
            result['company'] = oneline
            logger.warning(f"주소 fallback(한줄): {oneline}")
    except Exception:
        pass

    return result


_ZRMA_PARTNER_TABLE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
    "/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)
_ZRMA_PARTNER_SUB = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
)


def get_ship_to_address_zrma(session):
    """
    ZRMA_Q 오더에서 Ship-to Party 주소 읽기.
    Goto > Header > Partners → Ship-to 행 포커스 → btnDETAIL 클릭
    → SAPLSZA1 팝업에서 company/customer/street/street2/phone 읽기.
    """
    result = {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}

    try:
        session.findById(ZRMA_MENU_PARTNERS).select()
        time.sleep(1)

        table = session.findById(_ZRMA_PARTNER_TABLE)

        # Ship-to 행 찾기
        ship_to_row = -1
        for row_i in range(min(table.RowCount, 10)):
            try:
                parvw = table.GetCell(row_i, 0).Text.strip()
                if 'Ship' in parvw:
                    ship_to_row = row_i
                    break
            except Exception:
                break

        if ship_to_row < 0:
            logger.warning("ZRMA: Ship-to 행 못 찾음")
            session.findById("wnd[0]").sendVKey(3)
            return result

        # Ship-to 행 포커스 → btnDETAIL → SAPLSZA1 팝업
        table.GetCell(ship_to_row, 1).setFocus()
        time.sleep(0.5)
        session.findById(_ZRMA_PARTNER_SUB + "/btnDETAIL").press()
        time.sleep(1.5)

        base = "wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301"

        def rf(*ids):
            for fid in ids:
                try:
                    v = session.findById(fid).text.strip()
                    if v:
                        return v
                except Exception:
                    continue
            return ''

        phone_raw = rf(
            f"{base}/txtSZA1_D0100-TEL_NUMBER",
            f"{base}/txtSZA1_D0100-MOB_NUMBER",
        )
        result = {
            'company':  rf(f"{base}/txtADDR1_DATA-NAME1"),
            'customer': rf(f"{base}/txtADDR1_DATA-NAME2"),
            'street':   rf(f"{base}/txtADDR1_DATA-STR_SUPPL1"),
            'street2':  rf(f"{base}/txtADDR1_DATA-STR_SUPPL2"),
            'phone':    normalize_phone(phone_raw),
        }
        logger.info(f"ZRMA 주소 읽기: {result}")

        session.findById("wnd[1]").sendVKey(12)
        time.sleep(0.5)

    except Exception as e:
        logger.warning(f"ZRMA Partners 접근 실패: {e}")
        try:
            session.findById("wnd[1]").sendVKey(12)
        except Exception:
            pass

    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)
    except Exception:
        pass

    return result


def build_excel_rows(ebeln, order_info, address, extra_orders, memo):
    """
    하나의 오더에서 Excel 입력용 행 리스트 생성.
    아이템 1개 × 수량 = 행 수
    """
    rows = []
    is_first = True

    for item in order_info['items']:
        try:
            qty = max(1, int(item['lfimg']))
        except (ValueError, TypeError):
            qty = 1

        for _ in range(qty):
            rows.append({
                'order_prefix': '배송',
                'order_type':   'ZOR',
                'order_num':    ebeln,
                'extra_orders': extra_orders,
                'material':     item['matnr'],
                'description':  item['arktx'],
                'customer':     address.get('customer', ''),
                'phone':        address.get('phone', ''),
                'company':      address.get('company', ''),
                'street':       address.get('street', ''),
                'street2':      address.get('street2', ''),
                'memo':         memo if is_first else '',
                'is_first_item': is_first,
            })
            is_first = False

    return rows


def process_new_orders(session, new_ebelns, order_map):
    """
    새 오더 목록을 처리하여 Excel 입력용 데이터 반환.
    new_ebelns: ['7763549', ...] (7-prefix, 미처리)
    order_map:  group_by_order() 결과
    반환: {ebeln: [row, row, ...], ...}
    """
    results = {}

    for ebeln in new_ebelns:
        logger.info(f"=== 오더 {ebeln} 처리 중 ===")
        order_info = order_map.get(ebeln)
        if not order_info:
            continue

        vbeln   = order_info['vbeln']
        row_idx = order_info['row_idx']

        # 상세 화면 진입 (그리드 행 더블클릭)
        if not navigate_to_delivery(session, row_idx, vbeln):
            logger.error(f"오더 {ebeln} 진입 실패 - 건너뜀")
            continue

        # 1. 주소 먼저 읽기 (팝업 - 화면 이동 없음)
        address = get_ship_to_address(session)

        # 2. Text 읽기 (다른 화면으로 이동했다가 복귀)
        text = get_text_content(session)
        extra_orders = parse_extra_orders(text) if text else []
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = f"[Text] {text[:150]}" if text and text_contact['found'] else ""

        # text에서 전화번호 찾은 경우 주소에 없으면 보완
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # Excel 행 생성
        rows = build_excel_rows(ebeln, order_info, address, extra_orders, memo)
        results[ebeln] = rows

        logger.info(f"오더 {ebeln}: {len(rows)}행 생성")

        # 목록 화면으로 복귀 (F3 Back)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)

    return results

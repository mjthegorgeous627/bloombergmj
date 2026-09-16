"""
ZREC (회수 아이템 SAP 입고 처리) 자동화 - ZIH08 조회 + ZREC 접수.

todo c (2026-08-18) 1단계 구현. ZREC_DASH(접수 완료 후 Return Order/Label
RMA 비교)는 그 결과 화면이 SAP GUI Scripting으로 아예 읽히지 않는 임베디드
컨트롤이라(GuiCustomControl, 자식 0개 - "Extract to File" 옵션도 ABAP
런타임 에러로 확인 실패) 일단 보류. 완료 확인은 ZIH08 재조회(Plant/
Location이 6507/0052로 바뀌었는지)로 대체한다 - verify_received() 참고.

**중요 - receive_equipment_zrec()의 dry_run=False(실제 Receive Equipment
클릭 + 성공 팝업 처리) 경로는 실사용에서 버튼 클릭/접수 자체는 매번
성공했지만, 성공 여부 판정 로직에 버그가 있었다 (2026-08-19 확인).**
원래는 팝업 문구에 "success"라는 단어가 있는지로 판정했는데, 실제 팝업
문구는 그 단어를 포함하지 않아 매번 실패로 오판됐다(SAP에는 정상
접수됐는데 workbench는 실패 취급 - done 처리가 안 되는 증상으로 발견).
지금은 상태바 MessageType(S/W/E/A)을 우선 판정 기준으로 쓰도록 고쳤지만,
정확한 팝업 문구/상태바 타입 조합은 여전히 라이브로 한 번도 직접 대조
확인한 적이 없다 - 이 판정 자체가 참고용일 뿐, workbench의 "5. ZREC 처리"
3단계(ZIH08 재조회로 Plant/Location 6507/0052 확인)가 최종 진실이므로
done 처리는 그쪽 결과를 따른다.

사용법:
  python zrec_handler.py lookup --serials 70631395 70499385
  python zrec_handler.py receive --order 67078303 --serial 70631395
      (기본 dry-run: 화면까지만 채우고 Receive Equipment는 누르지 않음)
  python zrec_handler.py receive --order 67078303 --serial 70631395 --commit
      (실제로 Receive Equipment 클릭 + 성공 팝업 확인까지 - 아직 미검증)
  python zrec_handler.py receive --order 67078303 --material 10045196 --qty 1 --non-serial --commit
  python zrec_handler.py verify --serials 70631395 70499385
      (ZREC 이후 Plant/Location이 6507/0052로 바뀌었는지 재조회)
"""

import argparse
import json
import logging
import sys
import time

from sap_handler import get_scripting_engine

logger = logging.getLogger(__name__)

MAX_SESSIONS = 6

# ── ZIH08 요소 ID (2026-08-18 실측) ──────────────────────────────────────
IH08_GRID = "wnd[0]/usr/cntlGRID1/shellcont/shell"
IH08_SERNR_BTN = "wnd[0]/usr/btn%_SERNR_%_APP_%-VALU_PUSH"
IH08_MULTI_TABLE = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
IH08_MULTI_COPY_BTN = "wnd[1]/tbar[0]/btn[8]"       # tooltip "Copy (F8)" - 값 적용 + 팝업 닫기
IH08_MULTI_CANCEL_BTN = "wnd[1]/tbar[0]/btn[12]"    # tooltip "Cancel (F12)"
IH08_EXECUTE_BTN = "wnd[0]/tbar[1]/btn[8]"           # tooltip "Execute (F8)"

# ALV 그리드 컬럼 (전체 102개 중 실제 필요한 것만 - 2026-08-18 실측 확인)
IH08_COLUMNS = [
    "MATNR", "SERNR", "KUNDE", "NAME1", "FIRM", "KDAUF", "SD_AUART",
    "WERK", "LAGER", "TXT30", "ESTAT",
]

# ── ZREC 요소 ID (2026-08-18 실측) ───────────────────────────────────────
ZREC_PLANT_FIELD = "wnd[0]/usr/tabsTABSTRIP_REC/tabpRMA/ssubSUB_RMA:ZSD_RMA_RECEIVE_CONT:3006/ctxtGV_PLANT"
ZREC_REF_FIELD = "wnd[0]/usr/tabsTABSTRIP_REC/tabpRMA/ssubSUB_RMA:ZSD_RMA_RECEIVE_CONT:3006/txtGV_REF"
ZREC_SER_TAB = "wnd[0]/usr/tabsTABSTRIP_3000/tabpTAB1"
ZREC_SER_CELL = (
    "wnd[0]/usr/tabsTABSTRIP_3000/tabpTAB1/ssubSUB1:ZSD_RMA_RECEIVE_CONT:3003"
    "/tblZSD_RMA_RECEIVE_CONTTC_3003/txtGS_SERIAL-SERNR[0,0]"
)
ZREC_NONSER_TAB = "wnd[0]/usr/tabsTABSTRIP_3000/tabpTAB2"
ZREC_NONSER_MAT_CELL = (
    "wnd[0]/usr/tabsTABSTRIP_3000/tabpTAB2/ssubSUB2:ZSD_RMA_RECEIVE_CONT:3002"
    "/tblZSD_RMA_RECEIVE_CONTTC_3002/ctxtGS_MATNR-MATNR[0,0]"
)
ZREC_NONSER_QTY_CELL = (
    "wnd[0]/usr/tabsTABSTRIP_3000/tabpTAB2/ssubSUB2:ZSD_RMA_RECEIVE_CONT:3002"
    "/tblZSD_RMA_RECEIVE_CONTTC_3002/txtGS_MATNR-KWMENG[1,0]"
)
ZREC_RECEIVE_BTN = "wnd[0]/tbar[1]/btn[8]"           # tooltip "Receive Equipment (F8)"
ZREC_PLANT_VALUE = "6507"


# ── 세션 관리 ─────────────────────────────────────────────────────────────

def _connection():
    app = get_scripting_engine()
    return app.Children(0)


def _find_session(tcode):
    """이미 열려있는 세션 중 Info.Transaction이 정확히 이 tcode인 세션을
    찾는다 (없으면 None). VL06O/VL10G/ZRMA_Q 등 메인 루프 고정 세션은
    다른 tcode라 여기 걸리지 않음 - 절대 안 건드림."""
    conn = _connection()
    for i in range(conn.Children.Count):
        s = conn.Children(i)
        try:
            if (s.Info.Transaction or "").upper() == tcode.upper():
                return s
        except Exception:
            continue
    return None


def _open_new_session(tcode):
    """새 SAP 세션을 열어 tcode로 이동. open_session.py의 '마지막
    child = 새 세션' 가정은 방금 닫힌 세션의 SessionNumber가 재사용될 때
    엉뚱한(기존) 세션에 명령을 잘못 보내는 실제 사고를 냈다 (2026-08-18
    실측: ZIH08 세션이 통째로 ZREC로 덮어써짐). 생성 전/후 SessionNumber
    집합을 비교해서 진짜 새로 생긴 번호만 새 세션으로 인정하는 방식으로
    그 버그를 피한다."""
    conn = _connection()
    if conn.Children.Count >= MAX_SESSIONS:
        raise RuntimeError(
            f"SAP GUI 세션이 이미 최대({MAX_SESSIONS}개)입니다 - "
            f"{tcode} 세션을 열려면 안 쓰는 세션을 먼저 닫으세요."
        )
    before = {conn.Children(i).Info.SessionNumber for i in range(conn.Children.Count)}
    conn.Children(0).createSession()
    deadline = time.time() + 10
    new_sess = None
    while time.time() < deadline:
        time.sleep(0.5)
        for i in range(conn.Children.Count):
            s = conn.Children(i)
            if s.Info.SessionNumber not in before:
                new_sess = s
                break
        if new_sess:
            break
    if new_sess is None:
        raise RuntimeError(f"{tcode} 새 세션 생성 실패 (시간 초과)")
    new_sess.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
    new_sess.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)
    return new_sess


def get_session(tcode):
    """해당 tcode 세션을 찾아 반환 - 있으면 재사용(선택화면으로 리셋해서
    이전 조회/입력 상태 제거), 없으면 새로 연다."""
    session = _find_session(tcode)
    if session is None:
        logger.info(f"{tcode} 세션 없음 - 새로 엽니다")
        return _open_new_session(tcode)
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)
    return session


# ── ZIH08 조회 (read-only) ───────────────────────────────────────────────

def _fill_multi_serial(session, serials):
    """Multiple Selection for Serial Number 팝업에 시리얼 리스트 입력.
    한 화면에 보이는 행 수(VisibleRowCount, 실측 11)를 넘으면 스크롤해가며
    채운다."""
    session.findById(IH08_SERNR_BTN).press()
    time.sleep(1.5)
    table = session.findById(IH08_MULTI_TABLE)
    visible = table.VisibleRowCount
    table.VerticalScrollbar.Position = 0

    for idx, serial in enumerate(serials):
        page = idx // visible
        row_on_page = idx % visible
        if row_on_page == 0 and page > 0:
            table.VerticalScrollbar.Position = page * visible
        cell_id = f"{IH08_MULTI_TABLE}/txtRSCSEL_255-SLOW_I[1,{row_on_page}]"
        session.findById(cell_id).text = str(serial)

    session.findById(IH08_MULTI_COPY_BTN).press()
    time.sleep(1)


def _read_ih08_grid(session, serials):
    """이미 결과가 떠 있는 ZIH08 grid를 읽어 {serial: {...}} 로 변환 -
    lookup_serials_ih08()과 verify_received()의 재사용 경로가 공유하는
    추출 로직(조회를 새로 실행하는 부분과 분리)."""
    try:
        grid = session.findById(IH08_GRID)
    except Exception as exc:
        logger.error(f"ZIH08 결과 그리드를 찾지 못함 (조회 결과 없음일 수 있음): {exc}")
        return {}

    results = {}
    for i in range(grid.RowCount):
        row = {}
        for col in IH08_COLUMNS:
            try:
                row[col] = grid.GetCellValue(i, col).strip()
            except Exception:
                row[col] = ""
        sernr = row.get("SERNR", "")
        if not sernr:
            continue
        results[sernr] = {
            "material": row.get("MATNR", ""),
            "cust_no": row.get("KUNDE", ""),
            "customer_name": row.get("NAME1", ""),
            "firm_no": row.get("FIRM", ""),
            "matched_order": row.get("KDAUF", ""),
            "order_type": row.get("SD_AUART", ""),
            "plant": row.get("WERK", ""),
            "location": row.get("LAGER", ""),
            "status": row.get("TXT30", ""),
        }

    missing = [s for s in serials if s not in results]
    if missing:
        logger.warning(f"ZIH08에서 조회 안 된 시리얼: {missing}")

    return results


def lookup_serials_ih08(serials, session=None):
    """시리얼 리스트 → ZIH08 조회 → {serial: {material, cust_no,
    customer_name, firm_no, matched_order, order_type, plant, location,
    status}}. Execute(조회)만 실행 - 데이터 변경 없음, 완전히 안전.

    matched_order(KDAUF)가 그 시리얼이 실제로 회수 매칭된 SAP 오더번호 -
    ZREC의 Reference Document에 넣을 값. plant/location이 비어있으면 아직
    미접수, '6507'/'0052'면 이미 접수 완료된 것."""
    serials = [str(s).strip() for s in serials if str(s or "").strip()]
    if not serials:
        return {}

    session = session or get_session("ZIH08")
    _fill_multi_serial(session, serials)
    session.findById(IH08_EXECUTE_BTN).press()
    time.sleep(2)

    return _read_ih08_grid(session, serials)


def _find_ih08_session_on_results():
    """이미 결과 grid를 보여주고 있는 ZIH08 세션을 찾는다 (get_session()과
    달리 tcode를 재입력해 선택화면으로 리셋하지 않음) - verify_received()가
    lookup_serials_ih08()이 남겨둔 화면을 그대로 재사용하기 위함. 세션이
    없거나 아직 grid가 없으면(선택화면에 머물러 있는 등) None."""
    session = _find_session("ZIH08")
    if session is None:
        return None
    try:
        session.findById(IH08_GRID)
    except Exception:
        return None
    return session


def verify_received(serials, session=None, wait_before=5):
    """ZREC 접수 이후 완료 확인용 - Plant/Location이 6507/0052인지 판정해서
    received_ok로 함께 반환. ZREC_DASH가 보류된 지금은 이게 유일한 자동
    완료 확인 수단.

    사용자 지적(2026-08-31): ZREC 완료 확인을 위해 매번 ZIH08을 tcode부터
    새로 쳐서 들어가고(get_session()의 "/nZIH08" 리셋) Serial 목록을 통째로
    재입력하는 건 낭비다 - ZIH08은 원래 확인용으로, 1차 조회 때 이미 같은
    시리얼로 결과 grid를 띄워놨으니 잠깐 기다렸다가 그 화면에서 Execute를
    다시 누르기만(새로고침) 해도 Plant/Location이 바뀌었는지 보인다(사용자가
    수동으로 하던 방식 그대로). session이 명시적으로 주어지지 않으면 그
    화면이 아직 열려 있는지 먼저 찾아보고, 있으면 그걸 재사용 - tcode
    재입력도, Serial 팝업 재입력도 하지 않는다. 그 화면을 못 찾으면(세션이
    닫혔거나 애초에 lookup을 안 거친 경우) lookup_serials_ih08()로 새로
    연다 - 이 폴백 경로는 기존과 동일하게 동작한다.

    실사용 확인(2026-09-01): 새로고침 재사용 경로 자체가 실패하는 경우가
    있었다(정확한 원인 미확인 - Execute 버튼이 결과 화면에서는 다른
    동작을 하거나 grid 읽기가 안 되는 등, 사용자가 목격한 그 실패는 아직
    재현/원인 파악 못 함). 이 최적화는 어디까지나 "되면 빠른" 경로이지
    완료 확인 자체가 죽어서는 안 되므로, 재사용 시도 중 뭐든 실패하면
    lookup_serials_ih08()로 통째로 재조회하는 기존 방식으로 자동
    폴백한다."""
    reused = session
    if reused is None:
        reused = _find_ih08_session_on_results()
    if reused is not None:
        try:
            time.sleep(wait_before)
            reused.findById(IH08_EXECUTE_BTN).press()
            time.sleep(2)
            results = _read_ih08_grid(reused, serials)
            if not results:
                raise RuntimeError("새로고침 후 grid가 비어 있음")
        except Exception as exc:
            logger.warning(f"ZIH08 새로고침 재사용 실패 - 통째로 재조회로 폴백: {exc}")
            results = lookup_serials_ih08(serials)
    else:
        results = lookup_serials_ih08(serials)
    for row in results.values():
        row["received_ok"] = (row["plant"] == "6507" and row["location"] == "0052")
    return results


# ── ZREC 접수 ─────────────────────────────────────────────────────────────

def _strip_order_prefix(order_no):
    """'ZRX 67076875' / 'ZRX67076875' / 67076875 → '67076875'."""
    digits = "".join(ch for ch in str(order_no) if ch.isdigit())
    return digits or str(order_no).strip()


def _popup_text(popup):
    """팝업(wnd[1]) 안의 모든 텍스트를 모아 반환 - 성공/실패 판정용.
    **미검증**: 실제 성공 팝업을 한 번도 안 눌러봐서 이 구조가 맞는지
    확인 안 됨 (모듈 docstring 참고)."""
    texts = []

    def walk(el):
        try:
            t = getattr(el, "Text", "")
            if t:
                texts.append(str(t))
        except Exception:
            pass
        try:
            for i in range(el.Children.Count):
                walk(el.Children.ElementAt(i))
        except Exception:
            pass

    walk(popup)
    return "\n".join(texts)


def _confirm_popup(session):
    """SAP 성공/확인 팝업을 자동으로 닫는다 - 사람이 SAP 창에서 직접
    클릭할 필요 없게 하는 게 목적(사용자 확정, 2026-08-19: "체크해야하는데
    자동으로 너가 체크하고 나한테 묻지 말아라"). 팝업마다 버튼 구성이
    다를 수 있어 여러 방법을 순서대로 시도:
    1) 팝업 안의 실제 버튼(보통 'OK'/'Continue' 하나뿐인 확인 팝업)을 찾아 클릭
    2) 그게 없으면 Enter(VKey 0)
    3) 그래도 안 닫히면 한 번 더 Enter
    각 시도 후 팝업이 실제로 사라졌는지 확인하고, 안 사라졌으면 다음
    방법으로 넘어간다."""
    def popup_gone():
        try:
            session.findById("wnd[1]")
            return False
        except Exception:
            return True

    if popup_gone():
        return

    def find_buttons(el, found, depth=0):
        if depth > 6:
            return
        try:
            if getattr(el, "Type", "") == "GuiButton":
                found.append(el)
        except Exception:
            return
        try:
            for i in range(el.Children.Count):
                find_buttons(el.Children.ElementAt(i), found, depth + 1)
        except Exception:
            pass

    try:
        popup = session.findById("wnd[1]")
        buttons = []
        find_buttons(popup, buttons)
        for btn in buttons:
            try:
                btn.press()
                time.sleep(1)
                if popup_gone():
                    return
            except Exception:
                continue
    except Exception:
        pass

    for _ in range(2):
        if popup_gone():
            return
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(1)
        except Exception:
            break

    if not popup_gone():
        logger.warning("성공 팝업을 자동으로 닫지 못했습니다 - SAP 창을 직접 확인하세요.")


def _validate_zrec_fields(session, ref, serial=None, material=None, non_serial=False):
    """Receive Equipment를 누르기 직전, 방금 채운 필드를 다시 읽어서 실제
    화면 값이 의도한 값과 일치하는지 확인 - 화면 렌더링 지연 등으로 값이
    실제로 안 들어간 채 클릭하는 사고를 막기 위한 마지막 방어선."""
    plant = session.findById(ZREC_PLANT_FIELD).text.strip()
    ref_on_screen = session.findById(ZREC_REF_FIELD).text.strip()
    if plant != ZREC_PLANT_VALUE:
        raise RuntimeError(f"Plant 필드 검증 실패: 화면='{plant}' 기대값='{ZREC_PLANT_VALUE}'")
    if _strip_order_prefix(ref_on_screen) != ref:
        raise RuntimeError(f"Reference Document 검증 실패: 화면='{ref_on_screen}' 기대값='{ref}'")
    if non_serial:
        mat_on_screen = session.findById(ZREC_NONSER_MAT_CELL).text.strip()
        if mat_on_screen != str(material):
            raise RuntimeError(f"Material 검증 실패: 화면='{mat_on_screen}' 기대값='{material}'")
    else:
        ser_on_screen = session.findById(ZREC_SER_CELL).text.strip()
        if ser_on_screen != str(serial):
            raise RuntimeError(f"Serial 검증 실패: 화면='{ser_on_screen}' 기대값='{serial}'")


def prepare_zrec_item(order_no, serial=None, material=None, qty=None, non_serial=False):
    """workbench "ZREC 준비" 버튼용 - ZIH08 조회 → ZREC 화면 채우기(dry-run,
    Receive Equipment는 절대 안 누름)까지 한 번에. (todo c, 2026-08-18: 사용자
    요청 - "zih 화면을 열어서 serial을 조회하고, zrec 화면을 열어서 그
    오더넘버와 serial을 넣은 것까지만" 하고 멈추는 버튼.)

    시리얼 품목이면 ZIH08에서 조회한 matched_order(KDAUF)를 Reference
    Document에 쓴다 - workbench가 들고 있는 order_no보다 ZIH08의 실시간
    매칭 결과가 더 신뢰할 수 있는 값이기 때문(애초에 ZIH08을 먼저 보는
    이유). matched_order가 없으면(조회 실패 등) workbench의 order_no로
    폴백하되 경고를 함께 반환 - 이 경우 사람이 직접 확인 후 진행해야 함.
    workbench의 order_no와 ZIH08 matched_order가 둘 다 있는데 서로 다르면
    그것도 경고로 반환 (사용자가 원래 하던 수동 대조 작업).

    Non-serialized 품목(non_serial=True)은 ZIH08이 시리얼 기반 조회라
    대조할 게 없어 조회를 건너뛰고 workbench의 order_no를 그대로 쓴다."""
    warnings = []
    ih08_info = {}

    if non_serial:
        if not material:
            raise ValueError("non_serial=True면 material이 필요합니다")
        ref_order = order_no
        warnings.append("Non-serialized 품목이라 ZIH08 대조를 건너뛰었습니다 - Reference Document를 직접 확인하세요.")
    else:
        if not serial:
            raise ValueError("serial이 필요합니다 (non_serial=True가 아니면)")
        ih08_results = lookup_serials_ih08([serial])
        ih08_info = ih08_results.get(str(serial).strip(), {})
        matched_order = ih08_info.get("matched_order", "")
        expected = _strip_order_prefix(order_no) if order_no else ""
        if matched_order:
            ref_order = matched_order
            if expected and matched_order != expected:
                warnings.append(
                    f"workbench 오더번호({expected})와 ZIH08 매칭 오더(KDAUF={matched_order})가 다릅니다 - "
                    "ZIH08 매칭값을 사용합니다. 반드시 직접 확인하세요."
                )
        else:
            ref_order = order_no
            warnings.append("ZIH08에서 매칭된 Sales Order(KDAUF)를 찾지 못했습니다 - workbench 오더번호로 대체했습니다. 반드시 직접 확인하세요.")
        if ih08_info.get("plant") == "6507" and ih08_info.get("location") == "0052":
            warnings.append("ZIH08 조회 결과 이 시리얼은 이미 Plant 6507/Location 0052로 접수되어 있습니다 - 중복 접수 주의.")

    fill_result = receive_equipment_zrec(
        ref_order, serial=serial, material=material, qty=qty,
        non_serial=non_serial, dry_run=True,
    )

    return {
        "ih08": ih08_info,
        "filled_order": _strip_order_prefix(ref_order),
        "serial": serial,
        "material": material,
        "qty": qty,
        "non_serial": non_serial,
        "warnings": warnings,
        "fill_result": fill_result,
    }


def receive_equipment_zrec(order_no, serial=None, material=None, qty=None,
                            non_serial=False, dry_run=True, session=None):
    """ZREC 접수 1건. 기본은 dry_run=True - Plant/Reference Document/
    Serial(또는 Material+수량)까지만 채우고 Receive Equipment는 누르지
    않는다(화면을 직접 확인 후 사람이 클릭). dry_run=False일 때만 실제로
    버튼을 누르고 성공 팝업까지 처리한다 - **이 경로는 미검증**이니
    모듈 docstring의 경고를 반드시 먼저 읽을 것 (todo c, 2026-08-18:
    사용자가 명시적으로 요청한 단계적 검증 방식).

    한 번에 시리얼/품목 1개만 처리한다 (Serial Nos에 여러 개를 한 번에
    넣지 않음 - 기존 수동 작업과 동일하게 안전하게 하나씩)."""
    if non_serial:
        if not material:
            raise ValueError("non_serial=True면 material이 필요합니다")
        qty = qty or 1
    else:
        if not serial:
            raise ValueError("serial이 필요합니다 (non_serial=True가 아니면)")

    ref = _strip_order_prefix(order_no)
    session = session or get_session("ZREC")

    session.findById(ZREC_PLANT_FIELD).text = ZREC_PLANT_VALUE
    session.findById(ZREC_REF_FIELD).text = ref
    # 2026-08-27 실사용 확인: 두 건(67081036의 3HPPPP3, 67039538의 70250317)이
    # _validate_zrec_fields()는 통과했는데도(=그 시점 화면 표시값은 정상) 실제
    # SAP 접수는 Reference Document 연결 없이 처리됐다 - 하나는 아예 접수가
    # 안 됐고(70250317, 여전히 Assigned to Sales Order), 하나는 오더 연결 없이
    # 엉뚱한 Location(0060, 기대값 0052)으로 들어갔다(3HPPPP3, matched_order
    # 빈값으로 바뀜). GuiTextField.text 대입은 화면에는 즉시 반영되지만 SAP
    # 백엔드가 그 입력을 실제로 처리(포커스 이탈/OKCODE 검증)할 시간을 안 주면
    # 화면 검증은 통과해도 백엔드엔 반영이 덜 된 채로 다음 동작(탭 전환,
    # Receive Equipment 클릭)이 넘어갈 수 있다는 가설 - 완전히 확정된 원인은
    # 아니지만, Reference Document 입력 직후 처리 시간을 벌어주는 게 가장
    # 저비용의 완화책이라 추가함. 그래도 100% 방지는 보장 못 하니 이어지는
    # verify_received()의 ZIH08 재조회가 최종 방어선인 건 그대로다.
    time.sleep(0.4)

    if non_serial:
        session.findById(ZREC_NONSER_TAB).select()
        time.sleep(0.5)
        session.findById(ZREC_NONSER_MAT_CELL).text = str(material)
        session.findById(ZREC_NONSER_QTY_CELL).text = str(qty)
        logger.info(f"ZREC 입력 완료 (Non-serialized): order={ref} material={material} qty={qty}")
    else:
        session.findById(ZREC_SER_TAB).select()
        time.sleep(0.5)
        session.findById(ZREC_SER_CELL).text = str(serial)
        logger.info(f"ZREC 입력 완료 (Serialized): order={ref} serial={serial}")

    if dry_run:
        logger.info("[DRY RUN] Receive Equipment는 누르지 않았습니다 - 화면을 직접 확인 후 사람이 클릭하세요.")
        return {"submitted": False, "dry_run": True}

    _validate_zrec_fields(session, ref, serial=serial, material=material, non_serial=non_serial)

    session.findById(ZREC_RECEIVE_BTN).press()
    time.sleep(2)

    # 2026-08-19 실사용 확인: 팝업 문구에 "success"라는 단어가 실제로는
    # 없어서(모듈 docstring이 "Successfully..."로 추정했던 것과 다름)
    # 매번 실패로 오판됐다 - 여러 건이 SAP에는 정상 접수됐는데도 workbench가
    # 전부 실패 취급했음. 팝업 문구는 정확한 성공 문구를 아직도 모르니
    # 계속 참고용으로만 로그에 남기고, 판정은 SAP GUI 스크립팅의 표준
    # 신호인 상태바 MessageType(S=성공, W=경고→성공 취급, E/A=오류)을
    # 우선한다. 상태바를 못 읽으면(팝업이 그 자리를 가림 등) 문구 매칭으로
    # 폴백한다. 그래도 여기 판정은 참고용일 뿐 - workbench가 이어서
    # ZIH08 재조회(Plant/Location 6507/0052)로 실제 접수 여부를 최종
    # 확인하니, 여기서 성공/실패를 잘못 판정해도 done 처리 자체는 안전하다.
    popup_text = ""
    try:
        popup = session.findById("wnd[1]")
        popup_text = _popup_text(popup)
    except Exception:
        pass
    _confirm_popup(session)

    sbar_text = ""
    sbar_type = ""
    try:
        sbar = session.findById("wnd[0]/sbar")
        sbar_text = sbar.Text
        sbar_type = sbar.MessageType
    except Exception:
        pass

    message = popup_text or sbar_text
    if sbar_type:
        ok = sbar_type in ("S", "W")
    elif popup_text:
        ok = "success" in popup_text.lower()
    else:
        logger.info(f"팝업/상태바 모두 확인 불가 - 상태바 텍스트: {sbar_text!r}")
        return {"submitted": True, "success": None, "message": sbar_text}

    if not ok:
        logger.error(f"ZREC 실패로 보이는 응답 (상태바 타입={sbar_type!r}): {message}")
        return {"submitted": True, "success": False, "message": message}
    logger.info(f"ZREC 처리됨 (상태바 타입={sbar_type!r}): {message}")
    return {"submitted": True, "success": True, "message": message}


# ── CLI ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="ZREC 회수 아이템 SAP 입고 처리")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_lookup = sub.add_parser("lookup", help="ZIH08 시리얼 조회 (read-only, 완전히 안전)")
    p_lookup.add_argument("--serials", nargs="+", required=True)

    p_recv = sub.add_parser("receive", help="ZREC 접수 (기본 dry-run)")
    p_recv.add_argument("--order", required=True)
    p_recv.add_argument("--serial")
    p_recv.add_argument("--material")
    p_recv.add_argument("--qty", type=int, default=1)
    p_recv.add_argument("--non-serial", action="store_true")
    p_recv.add_argument(
        "--commit", action="store_true",
        help="실제로 Receive Equipment 클릭 (기본은 화면만 채우고 멈춤 - 아직 미검증 경로이니 신중히)",
    )

    p_verify = sub.add_parser("verify", help="ZREC 후 ZIH08 재조회로 Plant/Location 확인 (read-only)")
    p_verify.add_argument("--serials", nargs="+", required=True)

    p_prep = sub.add_parser(
        "prepare",
        help="ZIH08 조회 + ZREC 화면 채우기(dry-run)를 한 번에 - workbench 'ZREC 준비' 버튼용. "
             "결과를 RESULT_JSON: 로 시작하는 한 줄로 stdout에 출력한다.",
    )
    p_prep.add_argument("--order", help="workbench의 오더번호 (ZIH08 매칭값과 다르면 경고, 매칭값을 우선 사용)")
    p_prep.add_argument("--serial")
    p_prep.add_argument("--material")
    p_prep.add_argument("--qty", type=int, default=1)
    p_prep.add_argument("--non-serial", action="store_true")

    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    if args.cmd == "lookup":
        results = lookup_serials_ih08(args.serials)
        for sernr, row in results.items():
            print(sernr, row)
        print("RESULT_JSON:" + json.dumps(results, ensure_ascii=False))
    elif args.cmd == "receive":
        result = receive_equipment_zrec(
            args.order, serial=args.serial, material=args.material, qty=args.qty,
            non_serial=args.non_serial, dry_run=not args.commit,
        )
        print(result)
        print("RESULT_JSON:" + json.dumps(result, ensure_ascii=False))
    elif args.cmd == "verify":
        results = verify_received(args.serials)
        for sernr, row in results.items():
            status = "OK (6507/0052)" if row.get("received_ok") else "아직 미접수"
            print(sernr, status, row)
        print("RESULT_JSON:" + json.dumps(results, ensure_ascii=False))
    elif args.cmd == "prepare":
        try:
            result = prepare_zrec_item(
                args.order, serial=args.serial, material=args.material, qty=args.qty,
                non_serial=args.non_serial,
            )
            print("RESULT_JSON:" + json.dumps(result, ensure_ascii=False))
        except Exception as exc:
            logger.error(f"prepare 실패: {exc}")
            print("RESULT_JSON:" + json.dumps({"error": str(exc)}, ensure_ascii=False))
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())

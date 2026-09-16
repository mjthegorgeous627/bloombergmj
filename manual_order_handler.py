"""
Manual single-order import.

Usage:
  python order.py 67025693

Rules:
  - 6* orders: open VA02
  - all other orders: open VA03
Then reuse the sales-order extraction flow and write the result to Excel.
"""

import json
import logging
import os
import subprocess
import sys
import tempfile
import time

from excel_handler import write_kakao_sent
from kakao_handler import send_kakao_order, format_kakao_message
from order_lock import OrderLock
from order_tracker import load_processed, mark_processed, save_processed
from win_notify import send_windows_notification
from sap_handler import (
    run_transaction,
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
from zrma_handler import (
    get_items_from_zrma_order,
    collect_serial_numbers,
    build_excel_rows_zrma,
)

logger = logging.getLogger(__name__)


MAX_SESSIONS = 6  # SAP GUI 연결당 세션 수 하드 리밋 - open_session.py/zrec_handler.py/
                   # sn_relo_handler.py와 동일 (2026-09-16, SAP 자주 멈추는 문제 조사).


def _open_new_sap_session():
    """Create a new SAP GUI session and return it.

    2026-09-16: "세션 개수가 늘었다 → 마지막 child가 새 세션"이라는 가정은
    open_session.py/zrec_handler.py에도 있던 것과 같은 위험한 패턴이었다 -
    방금 닫힌 세션의 SessionNumber가 재사용될 때 엉뚱한 기존 세션에 명령을
    잘못 보내는 실제 사고를 낸 적이 있다 (zrec_handler.py의 2026-08-18
    실측 사고: ZIH08 세션이 통째로 ZREC로 덮어써짐). SessionNumber 집합을
    생성 전/후로 비교하는 안전한 방식으로 교체하고, 이미 세션이 다 찼을 때
    시도하다 SAP GUI 자체가 먹통이 되는 것도 미리 막는다."""
    app = get_scripting_engine()
    conn = app.Children(0)
    if conn.Children.Count >= MAX_SESSIONS:
        raise RuntimeError(
            f"SAP GUI 세션이 이미 최대({MAX_SESSIONS}개)입니다 - "
            "새 세션을 열려면 SAP에서 안 쓰는 세션을 먼저 닫으세요."
        )
    before = {conn.Children(i).Info.SessionNumber for i in range(conn.Children.Count)}
    conn.Children(0).createSession()

    deadline = time.time() + 10
    while time.time() < deadline:
        try:
            for i in range(conn.Children.Count):
                s = conn.Children(i)
                if s.Info.SessionNumber not in before:
                    logger.info(f"새 SAP 세션 생성: {s.findById('wnd[0]').Text}")
                    return s
        except Exception:
            pass
        time.sleep(0.5)

    raise RuntimeError("새 SAP 세션 생성 실패")


def _close_session(session):
    try:
        session.findById("wnd[0]").close()
        time.sleep(0.5)
    except Exception:
        pass


def _order_type_for(order_num, title):
    title_upper = (title or "").upper()
    if "ZINT" in title_upper:
        return "ZINT"
    if "ZINP" in title_upper:
        return "ZINP"
    if "EXCHANGE" in title_upper:
        return "ZRX"
    if "RMA" in title_upper or order_num.startswith("6"):
        return "ZRE"
    if order_num.startswith("7"):
        return "ZOR"
    if order_num.startswith(("4", "5")):
        return "ZINP"
    return "ORD"


def _lookup_delivery_obd(order_num):
    """Read VL06O session 0 and return the Bloomberg delivery/OBD number for a ZRX delivery row."""
    try:
        from portal_open_delivery import find_delivery_for_order

        _, info = find_delivery_for_order(order_num)
        obd = str(info.get("vbeln") or "").strip()
        if obd:
            logger.info(f"수동 오더 {order_num}: VL06O OBD 조회 완료 {obd}")
        return obd
    except Exception as exc:
        logger.warning(f"수동 오더 {order_num}: VL06O OBD 조회 실패: {exc}")
        return ""


def _open_order(session, order_num):
    tcode = "VA02" if order_num.startswith("6") else "VA03"
    logger.info(f"수동 오더 진입: {order_num} ({tcode})")
    run_transaction(session, f"/n{tcode}")
    time.sleep(1)
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order_num
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    # 2026-09-02 버그 수정 (오더 67086079, ZRX가 ZRE로 오분류): "special
    # handling notes" 메모가 걸린 오더는 이 시점에 "Information" 팝업
    # (wnd[1])이 뜨는데, 그걸 안 닫고 바로 아래에서 wnd[0].Text를 읽으면
    # 아직 전환 안 된 이전 화면 제목("Change Sales Documents" 등)을 읽어와
    # _order_type_for()가 "EXCHANGE"/"RMA" 둘 다 못 찾고 order_num 앞자리만
    # 보는 fallback으로 잘못 떨어진다. 팝업을 먼저 닫아야 실제 오더 제목
    # ("Change/Display RMA Order w/Exchange ...")을 제대로 읽는다.
    try:
        popup = session.findById("wnd[1]")
        if popup.Text == "Information":
            logger.info(f"수동 오더 {order_num}: Information 팝업 → Enter로 닫기")
            popup.sendVKey(0)
            time.sleep(1)
    except Exception:
        pass

    title = session.findById("wnd[0]").Text
    logger.info(f"오더 화면: {title}")
    return tcode, title


def _push_to_workbench(rows, order_num):
    """Push this order's rows straight into workbench.db right after writing
    it to Excel - see main.py's own _push_to_workbench for the full
    rationale (replaces the old Excel-reimport-based _sync_workbench(),
    removed 2026-08-06). Duplicated rather than imported so this standalone
    script (order.py) doesn't pull in main.py's module-level
    logging.basicConfig() setup."""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8", dir=os.path.dirname(__file__)
        ) as tmp:
            json.dump(rows, tmp, ensure_ascii=False)
            tmp_path = tmp.name
        subprocess.run(
            [sys.executable, os.path.join(os.path.dirname(__file__), "workbench_app.py"), "--import-rows-json", tmp_path],
            cwd=os.path.dirname(__file__),
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=60,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
    except Exception as e:
        logger.warning(f"workbench.db 반영 실패 (무시하고 계속) ({order_num}): {e}")
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def _write_and_notify(rows, order_num):
    # 2026-09-01: workbench가 메인, Excel은 백업일 뿐이다(사용자 지시) - 그런데
    # 이 함수가 write_orders_to_excel()을 먼저 부르고 그게 예외를 던지면(배송장
    # 병합 COM 오류 등) 그 아래 _push_to_workbench()가 아예 실행이 안 돼서
    # "SAP 수집은 성공했는데 workbench에는 안 뜨는" 상황이 났다. 순서를
    # 뒤집어서 workbench 반영을 항상 먼저/독립적으로 하고, Excel은 실패해도
    # 나머지(카카오, processed 마킹)를 막지 않는 순수 백업으로 격리한다.
    _push_to_workbench(rows, order_num)

    # 2026-09-01: main.py._write_and_notify()와 동일하게 윈도우 토스트 알림도
    # 여기서 보낸다 - VL06O(4/5/7-prefix)까지 이 함수(run_manual_order)로
    # 수집하도록 바뀌면서, 이 알림이 없으면 그만큼 새 오더 알림이 조용히
    # 빠지게 된다(2026-08-20에 추가한 기능인데 이 경로만 빠져있었음).
    try:
        message = format_kakao_message(rows)
        if message:
            lines = message.split('\n')
            toast_title, toast_body = lines[0], '\n'.join(lines[1:])
            send_windows_notification(toast_title, toast_body)
    except Exception as e:
        logger.warning(f"윈도우 알림 예외 ({order_num}): {e}")

    # 2026-09-01 사용자 결정: main.py._write_and_notify()와 동일 - 공유파일
    # 직접쓰기 중단, 그 파일은 이제 마감(확정) 데이터만 받는 아카이브로 전환.
    result = None

    try:
        ok = send_kakao_order(rows)
        if ok:
            if result:
                ws, start_row, end_row = result
                write_kakao_sent(ws, start_row, end_row)
            logger.info(f"카카오 전송 완료 ({order_num})")
        else:
            logger.warning(f"카카오 전송 실패 ({order_num}): send_kakao_order returned False")
    except Exception as e:
        logger.warning(f"카카오 전송 예외 ({order_num}): {e}")

    processed = load_processed()
    mark_processed(order_num, processed)
    save_processed(processed)


def run_manual_order(order_num, close_window=True):
    """Open one SAP order in a new session, extract it, and write it to Excel."""
    order_num = str(order_num).strip()
    if not order_num:
        print("[오더입력] 오더번호가 비어 있습니다.")
        return False

    lock = OrderLock(order_num)
    if not lock.__enter__():
        msg = (
            f"[오더입력] {order_num}는 지금 다른 프로세스(자동 수집 루프 또는 "
            f"다른 오더번호 반영)에서 이미 처리 중입니다. 몇 분 후 다시 시도하세요."
        )
        logger.warning(msg)
        print(msg)
        return False

    session = _open_new_sap_session()
    try:
        _, title = _open_order(session, order_num)
        order_type = _order_type_for(order_num, title)

        address = get_ship_to_address_zrma(session)

        text = get_text_content(session, menu_id=ZRMA_MENU_TEXTS)
        all_extra = parse_extra_orders(text) if text else []
        extra_orders = [eo for eo in all_extra if order_num not in eo]
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = parse_memo_for_display(text) if text else ""
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # ORD(ZOR only)/SDSK(everything else)/Cust#/Delivery Date - reuses this
        # same already-open order session, same as the automatic ZRMA_Q path
        # (see zrma_handler._read_zrma_order_detail). No Coll Due Date here -
        # that only exists as a ZRMA_Q *list* column, which a manually-typed
        # single order never scans.
        extra_fields = collect_order_extra_fields(order_num, order_type, existing_session=session)
        if extra_fields.get('cust_no'):
            address['cust_no'] = extra_fields['cust_no']
        if extra_fields.get('po_number'):
            extra_orders = add_extra_order_dedup(extra_orders, extra_fields['po_label'], extra_fields['po_number'])
        if extra_fields.get('delivery_date'):
            note = f"(delivery date: {extra_fields['delivery_date']})"
            memo = f"{memo}\n{note}" if memo else note

        items = get_items_from_zrma_order(session, order_num, order_type)
        if not items:
            logger.error(f"수동 오더 {order_num}: 아이템을 찾을 수 없음")
            print(f"[오더입력] {order_num} 아이템을 찾을 수 없습니다.")
            return False

        collect_serial_numbers(session, items)
        obd_map = {}
        if order_type in ("ZOR", "ZRX", "ZINP", "ZINT"):
            obd = _lookup_delivery_obd(order_num)
            if obd:
                obd_map[order_num] = obd
        rows = build_excel_rows_zrma(items, address, extra_orders, memo, obd_map=obd_map)
        if not rows:
            logger.error(f"수동 오더 {order_num}: Excel 행 생성 실패")
            print(f"[오더입력] {order_num} Excel 행 생성 실패")
            return False

        _write_and_notify(rows, order_num)
        # 2026-09-02: 이 로그가 예전엔 "Excel 입력 완료"였는데, 2026-09-01에
        # _write_and_notify()에서 실제 Excel 쓰기 단계 자체를 뺐으면서(공유
        # 파일 직접쓰기 중단) 더 이상 사실이 아니게 됐음 - 지금 실제로 하는
        # 일(workbench.db 반영)에 맞게 문구 수정.
        logger.info(f"수동 오더 {order_num} workbench 반영 완료 ({len(rows)}행)")
        print(f"[오더입력] {order_num} workbench 반영 완료 ({len(rows)}행)")
        return True
    except Exception as e:
        # 2026-09-01: 이 try에 except가 없어서 write_orders_to_excel() 같은
        # 단계에서 예외가 나면 (예: 배송장.xlsx 병합 COM 오류) order.py
        # 프로세스 전체가 로그 한 줄 없이 조용히 죽었다 - 자동루프
        # (main.py._run_zrma)는 이미 잡아서 "Excel 기록 실패"로 로그를
        # 남기는데 이 수동 경로만 빠져 있었음. SAP 쪽 수집(주소/S·N/OBD)은
        # 이미 끝난 뒤라 유실은 없지만, 버튼을 눌러도 아무 반응이 없어
        # 실패조차 몰랐던 문제를 고친다 - 동작은 "실패, 다음에 재시도"로
        # main.py와 동일하게 맞춤.
        logger.error(f"수동 오더 {order_num} 처리 중 예외: {e}", exc_info=True)
        print(f"[오더입력] {order_num} 처리 실패: {e}")
        return False
    finally:
        if close_window:
            _close_session(session)
        lock.__exit__(None, None, None)

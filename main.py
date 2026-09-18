"""
SAP 배송 오더 자동화 메인.

실행 방법:
  python main.py          → 10분마다 반복 실행
  python main.py --once   → 1회만 실행 (테스트용)

SAP 세션 구성 (4개 창 미리 열어둬야 함):
  세션0: VL06O  - 배송 오더 목록
  세션1: VL10G  - 배송 대기/회수(ZRE) 오더
  세션2: ZRMA_Q - 회수 오더 (RLKR variant, 키보드 제외)
  세션3: ZRMA_Q - 키보드 회수 오더 (Q2 variant)
"""

import json
import sys
import os
import subprocess
import tempfile
import time
import logging
from datetime import datetime

from config import REFRESH_INTERVAL_MINUTES, LOG_FILE
from order_lock import OrderLock
from order_tracker import load_processed, mark_processed, save_processed, purge_old_processed
from sap_handler import (
    get_sap_session,
    navigate_to_vl06o_list,
    get_all_rows_from_list,
    group_by_order,
)
from zrma_handler import (
    navigate_to_zrma_q,
    is_on_zrma_list,
    refresh_zrma_list,
    get_all_rows_from_zrma,
    group_zrma_by_order,
    process_zrma_orders,
)
from vl10g_handler import (
    navigate_to_vl10g,
    is_on_vl10g_list,
    refresh_vl10g,
    process_vl10g,
)
from excel_handler import write_kakao_sent, get_existing_order_numbers
from kakao_handler import send_kakao_order, format_kakao_message
from win_notify import send_windows_notification
from focus_guard import ForegroundLock

# 2026-08-26 사용자 요청으로 ZRMA RLKR/Q2(세션2/3)를 자동 20분 루프에서 뺐었으나,
# 2026-09-02 재검토 후 다시 켬: delivery block이 걸린 교환(ZRX) 오더는 VL06O/VL10G가
# 구조적으로 며칠간 못 보는데, 그 시간 동안 ZRMA 자동스캔마저 꺼져있으면 "즉시조회를
# 기억해서 눌러야만" 잡히는 사각지대가 생김(오더 67086079 실사례). RMA List 2개도
# 다시 매 사이클 자동 순회한다(launcher.py "특정 창만 조회"의 개별 버튼과
# run_now_2.flag/run_now_3.flag 단독 실행 경로는 이 상수와 무관하게 그대로 살아있음).
AUTO_SCAN_ZRMA = True

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)

# board_sync는 자기 import 시점에 logging.basicConfig()를 한 번 호출한다
# (board_sync.log 전용) - basicConfig()는 root logger에 handler가 이미 있으면
# 조용히 무시되므로, 위 main.py 자신의 basicConfig()보다 먼저 import되면
# automation.log가 board_sync.log로 뒤바뀐다(2026-08-28 board_sync.py
# 자체 문서화된 알려진 함정). 그래서 반드시 이 지점(로깅 설정 이후)에서
# import한다 - workbench_app.py가 함수 안에서 늦게 import하는 것과 같은 이유.
import board_sync


def _push_to_workbench(rows, order_id):
    """Push a just-collected order's rows straight into workbench.db - the
    same structured data write_orders_to_excel() just wrote to Excel, sent
    to workbench_app.py's import_sap_rows() instead of relying on Excel to
    be re-scanned/re-parsed later (the old _sync_workbench()/--import-today
    round-trip - removed 2026-08-06, see project_sap_watchdog_fix memory:
    that automatic Excel->Workbench reimport was the actual cause behind
    "workbench randomly loses edits/goes down", not SAP itself). Runs as a
    standalone subprocess (doesn't need the workbench server running) and
    never raises - a push failure must not break the actual SAP/Excel/kakao
    pipeline that the rest of run_once() just did. Excel is still written
    exactly as before (0minjung.bat, kakao, Portal automation's own Excel
    read path are all untouched) - this only adds a second, direct
    destination for the same data, it doesn't remove the first."""
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
        logger.warning(f"workbench.db 반영 실패 (무시하고 계속) ({order_id}): {e}")
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def _write_and_notify(rows, order_id, processed):
    """workbench.db 반영(메인) → 윈도우 알림 → 카카오 전송 → processed 저장
    (2026-09-01부터 Excel 기록 단계는 없음 - 아래 참고).

    2026-09-01: workbench가 메인이고 Excel은 백업이라는 게 애초 설계 의도였다
    (바로 위 _push_to_workbench 독스트링: "Excel is still written exactly as
    before ... this only adds a second, direct destination" - 즉 원래도 서로
    독립이어야 했다). 그런데 실제 코드는 write_orders_to_excel()을 먼저 불러서
    그게 예외를 던지면(배송장 병합 COM 오류 등) 그 아래 _push_to_workbench()가
    아예 실행이 안 됐다 - "SAP 수집은 성공했는데 workbench에는 안 뜨는" 원인이
    바로 이거였다. 순서를 뒤집고 Excel 실패를 격리해서, 메인(workbench)이
    백업(Excel) 상태와 무관하게 항상 반영되게 한다."""
    _push_to_workbench(rows, order_id)

    message = format_kakao_message(rows)
    if message:
        lines = message.split('\n')
        toast_title, toast_body = lines[0], '\n'.join(lines[1:])
        send_windows_notification(toast_title, toast_body)

    # 2026-09-01 사용자 결정: 공유파일(EXCEL_PATH)에 매일 SAP 수집을 직접
    # 쓰던 걸 중단 - workbench가 이미 메인이라 그 파일에 "아직 확정 안 된"
    # 진행중 상태까지 섞이면 지저분해진다는 지적. 공유파일은 이제 마감
    # (확정) 데이터만 받는 1년치 아카이브 역할로 전환(workbench_app.py의
    # export_dates_to_excel 참고), 실시간 상태는 board_sync.py가 20분 자동
    # 루프로 REALTIME_SYNC_PATH에 채운다(main.py run_loop() 참고) - 그
    # 파일이 이제 "실시간 반영"의 실제 대상이라 excel_portal_lookup.py도
    # 거기를 보도록 같이 옮겼다. result=None 유지 - 아래 write_kakao_sent는
    # 쓸 Excel 셀이 없으므로 자연히 스킵된다(기존 None 처리 그대로 재사용).
    result = None

    try:
        ok = send_kakao_order(rows)
        if ok:
            if result:
                ws, start_row, end_row = result
                write_kakao_sent(ws, start_row, end_row)
            logger.info(f"카카오 전송 완료 ({order_id})")
        else:
            logger.warning(f"카카오 전송 실패 ({order_id}): send_kakao_order returned False")
    except Exception as e:
        logger.warning(f"카카오 전송 예외 ({order_id}): {e}")
    mark_processed(order_id, processed)
    save_processed(processed)


def run_once(excel_only=False):
    """SAP GUI Scripting이 이 안에서 여러 세션 창을 계속 조작하면서 사용자
    작업 창의 포커스를 뺏어가는 문제(2026-08-24, 사용자 실측) - ForegroundLock으로
    감싸서 자동화가 프로그램적으로 창을 앞으로 가져오는 것만 막는다. 사용자가
    SAP 창을 직접 클릭하는 건 이 락과 무관하게 항상 그대로 작동한다(focus_guard.py
    설명 참고). 실제 작업은 그대로 _run_once_impl()가 한다."""
    _mark_cycle_state("running")
    try:
        with ForegroundLock():
            _run_once_impl(excel_only=excel_only)
    finally:
        _mark_cycle_state("idle")


def _run_once_impl(excel_only=False):
    logger.info("=" * 55)
    logger.info(f"실행 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    purged = purge_old_processed()
    if purged:
        logger.info(f"오래된 오더 {purged}개 자동 삭제 (45일 초과)")

    # excel_only=True: Excel 기록만 기준 (JSON 무시) → 누락 오더 재처리용
    if excel_only:
        excel_existing = get_existing_order_numbers()
        processed = excel_existing.copy()
        logger.info(f"[catchup] Excel 기준 {len(processed)}개 오더만 기처리로 간주")
    else:
        processed = load_processed()
        # JSON이 비어있을 때만 Excel 스캔 (JSON 초기화/삭제 복구용)
        if not processed:
            excel_existing = get_existing_order_numbers()
            if excel_existing:
                logger.info(f"JSON 비어있음 → Excel에서 {len(excel_existing)}개 오더 복구")
                processed = excel_existing
                save_processed(processed)

    vl06o_obd_map = {}

    # ── 세션0: VL06O ────────────────────────────────────────
    try:
        sess0 = get_sap_session(0)
        vl06o_obd_map = _run_vl06o(sess0, processed)
        # processed는 각 함수 내에서 mark_processed로 업데이트됨
        processed = load_processed()
    except ConnectionError as e:
        logger.error(f"세션0 VL06O 연결 실패: {e}")

    # ── 세션1: VL10G ────────────────────────────────────────
    try:
        sess1 = get_sap_session(1)
        _run_vl10g(sess1, processed)
        processed = load_processed()
    except ConnectionError as e:
        logger.error(f"세션1 VL10G 연결 실패: {e}")

    if AUTO_SCAN_ZRMA:
        # ── 세션2: ZRMA_Q RLKR ─────────────────────────────────
        try:
            sess2 = get_sap_session(2)
            _run_zrma(sess2, processed, variant='rlkr', date_mode='next_month', obd_map=vl06o_obd_map)
            processed = load_processed()
        except ConnectionError as e:
            logger.error(f"세션2 ZRMA_Q RLKR 연결 실패: {e}")

        # ── 세션3: ZRMA_Q Q2 ───────────────────────────────────
        try:
            sess3 = get_sap_session(3)
            _run_zrma(sess3, processed, variant='q2', date_mode='year_end', obd_map=vl06o_obd_map)
            processed = load_processed()
        except ConnectionError as e:
            logger.error(f"세션3 ZRMA_Q Q2 연결 실패: {e}")
    else:
        logger.info("ZRMA RLKR/Q2 세션은 자동 루프 제외 (2026-08-26 설정) - 필요 시 수동 조회")

    logger.info("이번 실행 완료")
    # (workbench.db 반영은 각 오더가 처리될 때 _write_and_notify()에서 바로 됨 -
    # 더 이상 여기서 엑셀 전체를 다시 긁어오지 않음)


def _run_vl06o(session, processed):
    """VL06O: 목록에서 7/5/4-prefix 새 오더를 찾아 run_manual_order()(VA02/VA03
    직접 진입)로 수집한다. 2026-09-01부터 VL06O 자체 그리드 추출
    (process_new_orders) 대신 이 방식을 쓴다 - 아래 new_orders 루프의
    주석 참고."""
    try:
        navigate_to_vl06o_list(session)
    except Exception as e:
        logger.error(f"VL06O 이동 실패: {e}")
        return {}

    all_rows = get_all_rows_from_list(session)
    if not all_rows:
        logger.info("VL06O 목록 비어있음")
        return {}

    order_map = group_by_order(all_rows)
    obd_map = {}
    for ebeln, info in order_map.items():
        vbelns = []
        for item in info.get('items', []):
            vbeln = str(item.get('vbeln') or info.get('vbeln') or '').strip()
            if vbeln:
                vbelns.append(vbeln)
        unique_vbelns = list(dict.fromkeys(vbelns))
        if len(unique_vbelns) > 1:
            obd_map[str(ebeln)] = vbelns
        elif unique_vbelns:
            obd_map[str(ebeln)] = unique_vbelns[0]
    logger.info(f"VL06O 전체 오더 수: {len(order_map)}")

    # 7-prefix는 항상 VL06O에서, 5-prefix는 processed에 없는 것만
    new_orders = [
        ebeln for ebeln in order_map
        if ebeln[0] in ('7', '5', '4') and ebeln not in processed
    ]
    logger.info(f"VL06O 새 오더: {len(new_orders)}개 → {new_orders}")

    if not new_orders:
        logger.info("VL06O 새로운 오더 없음")
        return obd_map

    # 2026-09-01 사용자 지시로 process_new_orders()(VL06O 배송문서 그리드만
    # 읽음) 대신 run_manual_order()(VA02/VA03로 오더 자체를 직접 진입)로
    # 교체했다. VL06O는 "Outbound Delivery" 문서만 보여주는 화면이라, 교환형
    # 오더처럼 배송 라인과 회수 라인이 한 오더에 같이 있는 경우 회수 라인이
    # 통째로 안 보인다(실사용자 확인) - VA02/VA03로 오더 자체에 들어가야
    # 배송/회수 라인이 전부 잡힌다. run_manual_order()는 6-prefix만 VA02,
    # 나머지(7/5/4 포함)는 VA03을 쓰므로 7-prefix(ZOR, display-only)도 그대로
    # 안전하다. VL10G의 "그 외 배송 오더" 분기와 완전히 같은 패턴 - 자체적으로
    # 주소/S·N/Excel/workbench/카카오/processed 마킹까지 다 처리하므로 여기서
    # 별도로 _write_and_notify()를 부르지 않는다.
    #
    # 2026-09-02 버그 수정: 이 아래를 예전엔 `with OrderLock(ebeln) as locked:`로
    # 한 번 더 감싸고 있었다. run_manual_order() 자신도 내부에서 같은
    # order_num으로 OrderLock을 잡는데(order.py의 "오더번호 반영"과 동시
    # 재수집을 막기 위한 원래 목적, order_lock.py 참고) - OrderLock은
    # 파일 존재만으로 배타 판단하는 비-재진입(non-reentrant) 락이라, 같은
    # 프로세스가 같은 오더에 대해 바깥/안쪽에서 두 번 잡으려 하면 안쪽
    # 시도가 "다른 프로세스가 처리 중"이라는 오탐과 함께 항상 실패했다
    # (실제로는 같은 프로세스의 바깥 with가 이미 쥐고 있을 뿐, 다른
    # 프로세스가 아님). 그 결과 2026-09-01 VL06O가 run_manual_order()를
    # 쓰도록 바뀐 뒤로 이 경로로 들어온 신규 오더는 자동/개별 조회 어느
    # 쪽으로도 100% 실패했고(automation.log 실측: 시도 3건/성공 0건),
    # 다음 사이클에도 processed에 안 남아 같은 오더가 계속 재탐지→
    # 재실패를 반복하다가, 결국 사용자가 order.py로 별도 수동 반영해야만
    # 풀렸다. vl10g_handler.py의 동일 패턴(그 외 배송 오더 분기)은 이
    # 이중 래핑 없이 run_manual_order()를 바로 호출하고 있었고 그쪽은
    # 정상 동작했다 - 그것과 똑같이 바깥 OrderLock을 없애고 run_manual_order()
    # 자신의 락에만 맡긴다.
    for ebeln in new_orders:
        collected = False
        for attempt in range(2):
            try:
                from manual_order_handler import run_manual_order

                if run_manual_order(ebeln, close_window=True):
                    logger.info(f"VL06O 오더 {ebeln} VA02/VA03 직접 수집 완료 (workbench 반영됨)")
                    collected = True
                    break
            except Exception as exc:
                logger.error(f"VL06O 오더 {ebeln} 직접 수집 실패 (시도 {attempt + 1}/2): {exc}", exc_info=True)
                if attempt == 0:
                    time.sleep(2)
        if not collected:
            logger.warning(f"VL06O 오더 {ebeln} 직접 수집 실패 → 다음 회차 재시도")
    return obd_map


def _run_vl10g(session, processed):
    """VL10G: Block 해제 + VL06O 이관 + ZRE 수집."""
    try:
        # 항상 재실행 (F8) → DB 재조회로 새 오더 반영
        navigate_to_vl10g(session)
    except Exception as e:
        logger.error(f"VL10G 이동 실패: {e}")
        return

    result = process_vl10g(session, processed)

    # Block 해제 불가 오더 → Excel에 오더번호/회사명/에러 기록
    for order_num, rows in result['blocked_excel_rows'].items():
        with OrderLock(order_num) as locked:
            if not locked:
                logger.warning(f"VL10G Block 오더 {order_num}: 다른 프로세스가 처리 중 → 이번 회차 건너뜀")
                continue
            try:
                _write_and_notify(rows, order_num, processed)
                logger.info(f"VL10G Block 오더 {order_num} workbench 반영 완료")
            except Exception as e:
                logger.error(f"VL10G Block 오더 {order_num} 기록 실패: {e}")

    if result['pushed_to_vl06o']:
        logger.info(f"VL06O로 이관된 오더: {result['pushed_to_vl06o']}")

    # ZRE 회수 오더 처리 (VL10G ZS 해제 + S/N 수집 완료)
    for order_num, rows in result['zre_orders'].items():
        if rows:
            with OrderLock(order_num) as locked:
                if not locked:
                    logger.warning(f"ZRE 오더 {order_num}: 다른 프로세스가 처리 중 → 이번 회차 건너뜀")
                    continue
                try:
                    _write_and_notify(rows, order_num, processed)
                    logger.info(f"ZRE 오더 {order_num} workbench 반영 완료")
                except Exception as e:
                    logger.error(f"ZRE 오더 {order_num} 기록 실패: {e}")


def _run_zrma(session, processed, variant, date_mode, obd_map=None):
    """ZRMA_Q: 6-prefix + 회수 포함 4/5-prefix 처리."""
    try:
        if not is_on_zrma_list(session):
            navigate_to_zrma_q(session, variant, date_mode, apply_layout=False)
        else:
            if not refresh_zrma_list(session):
                logger.info(f"ZRMA_Q ({variant}) refresh failed; reopening selection")
                navigate_to_zrma_q(session, variant, date_mode, apply_layout=False)
    except Exception as e:
        logger.error(f"ZRMA_Q ({variant}) 이동 실패: {e}")
        return

    all_rows = get_all_rows_from_zrma(session)
    if not all_rows:
        logger.info(f"ZRMA_Q ({variant}) 목록 비어있음")
        return

    order_map = group_zrma_by_order(all_rows)
    logger.info(f"ZRMA_Q ({variant}) 전체 오더: {len(order_map)}개")

    # 6-prefix, 4/5-prefix 중 미처리 오더
    new_orders = [
        num for num in order_map
        if num and num[0] in ('4', '5', '6') and num not in processed
    ]
    logger.info(f"ZRMA_Q ({variant}) 새 오더: {len(new_orders)}개 → {new_orders}")

    if not new_orders:
        logger.info(f"ZRMA_Q ({variant}) 새로운 오더 없음")
        return

    results = process_zrma_orders(session, new_orders, order_map, variant, date_mode, obd_map=obd_map)
    for order_num, rows in results.items():
        # see _run_vl06o의 동일 주석 - order.py와의 동시 쓰기 경합 방지.
        with OrderLock(order_num) as locked:
            if not locked:
                logger.warning(f"ZRMA 오더 {order_num}: 다른 프로세스가 처리 중 → 이번 회차 건너뜀")
                continue
            try:
                _write_and_notify(rows, order_num, processed)
                logger.info(f"ZRMA 오더 {order_num} workbench 반영 완료 ({len(rows)}행)")
            except Exception as e:
                logger.error(f"ZRMA 오더 {order_num} 기록 실패: {e}", exc_info=True)


def _get_current_vl06o_obd_map():
    """세션 단독 실행 때 현재 VL06O 목록에서 ZRX 배송 OBD 매핑만 읽는다."""
    try:
        sess0 = get_sap_session(0)
        if not sess0:
            return {}
        rows = get_all_rows_from_list(sess0)
        order_map = group_by_order(rows)
        result = {}
        for ebeln, info in order_map.items():
            vbelns = []
            for item in info.get('items', []):
                vbeln = str(item.get('vbeln') or info.get('vbeln') or '').strip()
                if vbeln:
                    vbelns.append(vbeln)
            unique_vbelns = list(dict.fromkeys(vbelns))
            if len(unique_vbelns) > 1:
                result[str(ebeln)] = vbelns
            elif unique_vbelns:
                result[str(ebeln)] = unique_vbelns[0]
        return result
    except Exception as e:
        logger.warning(f"VL06O OBD 매핑 조회 실패: {e}")
        return {}



TRIGGER_FILE = os.path.join(os.path.dirname(__file__), "run_now.flag")
STOP_FILE = os.path.join(os.path.dirname(__file__), "stop_automation.flag")
TRIGGER_FILES = {
    i: os.path.join(os.path.dirname(__file__), f"run_now_{i}.flag")
    for i in range(4)
}

# 2026-09-17: automation.log의 mtime만 보는 기존 워치독(sap_ops.py)은 사이클
# 사이 정상 유휴 시간(최대 REFRESH_INTERVAL_MINUTES분)까지 감안해야 해서 무갱신
# 기준이 30분으로 느슨하다 - 오늘 실측된 새 증상(get_sap_session() 등 SAP GUI
# Scripting 호출이 예외 없이 그냥 영원히 블로킹, SapGuiServer.exe 자체가
# Responding=False로 멈춤)은 사이클 "진행 중"에 일어나는데, 그걸 30분간 못 잡으면
# 화면이 그만큼 오래 먹통으로 방치된다. 진행 중/유휴 상태를 별도 파일에 기록해서,
# 워치독이 "진행 중인데 오래 멈췄다"만 훨씬 짧은 기준(수 분)으로 따로 잡게 한다 -
# 유휴 구간의 자연스러운 침묵과는 구분되므로 오탐 없이 감지 속도만 개선된다.
CYCLE_STATE_FILE = os.path.join(os.path.dirname(__file__), "cycle_state.json")


def _mark_cycle_state(status):
    try:
        with open(CYCLE_STATE_FILE, "w", encoding="utf-8") as fh:
            json.dump({"status": status, "since": datetime.now().isoformat()}, fh)
    except Exception:
        pass


def _run_session(sess_idx):
    """개별 세션 단독 실행. run_once()와 같은 이유로 ForegroundLock 적용 -
    실제 작업은 _run_session_impl()가 한다."""
    _mark_cycle_state("running")
    try:
        with ForegroundLock():
            _run_session_impl(sess_idx)
    finally:
        _mark_cycle_state("idle")


def _run_session_impl(sess_idx):
    purge_old_processed()
    processed = load_processed()
    logger.info(f"=== 세션{sess_idx} 단독 실행 ===")
    try:
        if sess_idx == 0:
            _run_vl06o(get_sap_session(0), processed)
        elif sess_idx == 1:
            _run_vl10g(get_sap_session(1), processed)
        elif sess_idx == 2:
            _run_zrma(get_sap_session(2), processed, variant='rlkr', date_mode='next_month', obd_map=_get_current_vl06o_obd_map())
        elif sess_idx == 3:
            _run_zrma(get_sap_session(3), processed, variant='q2', date_mode='year_end', obd_map=_get_current_vl06o_obd_map())
    except ConnectionError as e:
        logger.error(f"세션{sess_idx} 연결 실패: {e}")
        # 2026-09-02 버그 수정 (오더 67086095: 개별 조회 버튼을 눌러도 못 잡던
        # 사례): run_loop()의 폴링 루프는 이 함수를 부르기 전에 이미
        # run_now_{sess_idx}.flag를 지워버린다(main.py의 TRIGGER_FILES
        # 처리부). 그 직후 SAP GUI가 하필 이 순간 끊겨(재로그인이 필요한
        # COM 예외 등) ConnectionError가 나면, 사용자가 누른 "개별 조회"
        # 요청이 재시도 없이 그냥 조용히 사라졌다 - 로그에만 남고 화면에는
        # 아무 표시도 없어서 사용자는 "버튼을 눌렀는데 안 잡혔다"고만
        # 느꼈다. SAP가 (워치독에 의해서든 자동 재로그인으로든) 복구된 뒤
        # 다음 폴링 사이클이 이 세션을 다시 시도하도록 flag를 되살려둔다.
        try:
            open(TRIGGER_FILES[sess_idx], "a").close()
            logger.info(f"세션{sess_idx} 트리거 재생성 - SAP 복구 후 재시도됨")
        except Exception as touch_err:
            logger.warning(f"세션{sess_idx} 트리거 재생성 실패: {touch_err}")
    except Exception:
        # run_loop()의 run_once() 호출부와 같은 이유의 최후 방어선 - 이게 없으면
        # 트리거 파일(run_now_{i}.flag)로 단독 실행했을 때도 미지의 예외 하나로
        # 루프 전체가 죽을 수 있었다.
        logger.error(f"세션{sess_idx} 단독 실행 미처리 예외", exc_info=True)
    # (workbench.db 반영은 _write_and_notify()에서 바로 됨 - 위 참고)


def run_loop():
    logger.info(f"자동화 시작 (간격: {REFRESH_INTERVAL_MINUTES}분)")
    logger.info("수동 실행: 런처 버튼 또는 'python now.py'")
    if os.path.exists(STOP_FILE):
        os.remove(STOP_FILE)

    while True:
        if os.path.exists(STOP_FILE):
            # 2026-09-13 버그 수정: 예전엔 여기서 os.remove(STOP_FILE)까지 했는데,
            # 그러면 evening_shutdown.py가 "의도적으로 꺼둔 상태"라고 남겨둔 그
            # 플래그를 main.py 자신이 멈추면서 곧바로 지워버렸다 - watchdog.py/
            # workbench_app.py의 _sap_loop_watchdog_thread()는 둘 다 이 플래그가
            # 있으면 재시작을 건너뛰는데, 지워지고 나면 그냥 automation.log
            # 30분 무갱신만 보고 "죽었다"고 오판해 30분 뒤 SAP를 다시 무인으로
            # 깨웠다(실측: 2026-09-12 19:00 evening_shutdown → 19:00:02 이
            # 코드가 플래그 삭제 → 19:30:39 워치독이 재시작, 그대로 다음날
            # 아침까지 밤새 돌았음). 이 플래그를 "다시 시작할 때 지우는" 책임은
            # 이미 run_loop() 맨 위(바로 위 os.remove)와 sap_start_loop() 등
            # 시작 경로들이 지고 있으므로, 멈추는 쪽에서는 그냥 감지만 하고
            # 파일은 그대로 남겨둔다.
            logger.info("중지 요청 감지 → 자동화 종료")
            return

        try:
            run_once()
        except Exception:
            # 2026-08-24: run_once()/_run_vl06o/_run_zrma 등은 세션 확보 시점의
            # ConnectionError만 잡고 있고, 그 사이 실제 스크래핑/네비게이션
            # 함수들(get_all_rows_from_list, group_by_order, navigate_to_zrma_q
            # 등)은 아무것도 안 감싸져 있었다 - 처음 보는 종류의 SAP GUI
            # Scripting/COM 예외가 그 구간 어디서든 터지면 이 while 루프 전체가
            # 죽고 아무도 재시작하지 않았다(워치독이 30~40분 뒤에나 눈치챔).
            # "loop 돌리던 중 절로 종료"의 유력 원인 - 매번 다른 지점에서
            # 재발했던 이유이기도 함(개별 지점을 고쳐도 다음 미지의 예외가
            # 다른 자리에서 또 터짐). 여기서 잡아 로그만 남기고 다음 사이클로
            # 넘어가는 게 최후의 방어선.
            logger.error("run_once() 미처리 예외 - 이번 사이클 실패, 다음 사이클로 계속", exc_info=True)

        # 2026-09-01 사용자 결정: REALTIME_SYNC_PATH(board_sync.py의 "동기화"
        # 대상)를 SAP 자동수집과 같은 20분 주기로 항상 최신 유지 - 워크벤치
        # 버튼을 눌러야만 갱신되던 걸 여기서도 자동으로 하게 해서, 그 파일을
        # 보는 excel_portal_lookup.py(POD remarks) 등이 항상 최신 상태를
        # 본다. create_new_day_sheet()는 오늘 시트가 이미 있으면 그대로
        # sync()만 하므로 매 사이클 호출해도 안전(--new-day와 동일 동작).
        # run_once()와 별개의 try/except로 격리 - 이게 실패해도 SAP
        # 자동수집 자체는 계속돼야 한다.
        try:
            board_sync.create_new_day_sheet()
        except Exception:
            logger.error("board_sync 자동 동기화 실패 - 무시하고 계속", exc_info=True)

        logger.info(f"{REFRESH_INTERVAL_MINUTES}분 후 다음 실행... (수동: now.py / now.py 0~3)")

        for _ in range(REFRESH_INTERVAL_MINUTES * 60 // 5):
            time.sleep(5)
            if os.path.exists(STOP_FILE):
                # 위 while True 진입부의 동일 수정과 같은 이유 - 플래그 파일은
                # 그대로 남겨서 evening_shutdown.py가 세운 "의도적으로 꺼둠"
                # 표시가 워치독들에게 계속 보이게 한다.
                logger.info("중지 요청 감지 → 자동화 종료")
                return

            if os.path.exists(TRIGGER_FILE):
                os.remove(TRIGGER_FILE)
                logger.info("▶ 수동 전체 실행 트리거 감지 → 즉시 실행")
                break

            for sess_idx, flag_path in TRIGGER_FILES.items():
                if os.path.exists(flag_path):
                    os.remove(flag_path)
                    logger.info(f"▶ 세션{sess_idx} 단독 실행 트리거 감지")
                    _run_session(sess_idx)
                    break

if __name__ == "__main__":
    if "--catchup" in sys.argv:
        run_once(excel_only=True)
    elif "--once" in sys.argv:
        run_once()
    elif "--print-afternoon" in sys.argv:
        from print_handler import print_afternoon
        print_afternoon()
    elif "--print-tomorrow" in sys.argv:
        from print_handler import print_tomorrow
        print_tomorrow()
    elif "--print" in sys.argv:
        idx = sys.argv.index("--print")
        rest = sys.argv[idx + 1:]
        if not rest:
            print("사용법: python main.py --print 시트 [날짜] [오후]")
            print("예시:   python main.py --print 4-2")
            print("        python main.py --print 4-2 오후")
            print("        python main.py --print 4-2 4-3")
            print("        python main.py --print 4-2 4-3 오후")
        else:
            from print_handler import print_section
            sheet_key = rest[0]
            date_key = None
            afternoon = False
            for arg in rest[1:]:
                if arg == "오후":
                    afternoon = True
                elif "-" in arg:
                    date_key = arg
            print_section(sheet_key, date_key, afternoon)
    elif "--quick" in sys.argv:
        idx = sys.argv.index("--quick")
        if idx + 1 < len(sys.argv):
            from quick_handler import run_quick
            run_quick(sys.argv[idx + 1])
        else:
            print("사용법: python main.py --quick [오더번호]")
            print("예시:   python main.py --quick 7780226")
    else:
        run_loop()

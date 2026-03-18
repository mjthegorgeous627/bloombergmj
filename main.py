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

import sys
import time
import logging
from datetime import datetime

from config import REFRESH_INTERVAL_MINUTES, LOG_FILE
from order_tracker import load_processed, mark_processed, save_processed
from sap_handler import (
    get_sap_session,
    navigate_to_vl06o_list,
    get_all_rows_from_list,
    group_by_order,
    process_new_orders,
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
from excel_handler import write_orders_to_excel, get_existing_order_numbers

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)


def run_once():
    logger.info("=" * 55)
    logger.info(f"실행 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 처리 완료 오더 로드 + Excel 기존 오더 합산
    processed = load_processed()
    excel_existing = get_existing_order_numbers()
    if excel_existing - processed:
        logger.info(f"Excel 기존 오더 {len(excel_existing)}개 → processed에 추가")
        processed = processed | excel_existing
        save_processed(processed)

    # ── 세션0: VL06O ────────────────────────────────────────
    try:
        sess0 = get_sap_session(0)
        _run_vl06o(sess0, processed)
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

    # ── 세션2: ZRMA_Q RLKR ─────────────────────────────────
    try:
        sess2 = get_sap_session(2)
        _run_zrma(sess2, processed, variant='rlkr', date_mode='next_month')
        processed = load_processed()
    except ConnectionError as e:
        logger.error(f"세션2 ZRMA_Q RLKR 연결 실패: {e}")

    # ── 세션3: ZRMA_Q Q2 ───────────────────────────────────
    try:
        sess3 = get_sap_session(3)
        _run_zrma(sess3, processed, variant='q2', date_mode='year_end')
        processed = load_processed()
    except ConnectionError as e:
        logger.error(f"세션3 ZRMA_Q Q2 연결 실패: {e}")

    logger.info("이번 실행 완료")


def _run_vl06o(session, processed):
    """VL06O: 7-prefix + 5-prefix(ZRMA에 없는 것) 처리."""
    try:
        navigate_to_vl06o_list(session)
    except Exception as e:
        logger.error(f"VL06O 이동 실패: {e}")
        return

    all_rows = get_all_rows_from_list(session)
    if not all_rows:
        logger.info("VL06O 목록 비어있음")
        return

    order_map = group_by_order(all_rows)
    logger.info(f"VL06O 전체 오더 수: {len(order_map)}")

    # 7-prefix는 항상 VL06O에서, 5-prefix는 processed에 없는 것만
    new_orders = [
        ebeln for ebeln in order_map
        if ebeln[0] in ('7', '5', '4') and ebeln not in processed
    ]
    logger.info(f"VL06O 새 오더: {len(new_orders)}개 → {new_orders}")

    if not new_orders:
        logger.info("VL06O 새로운 오더 없음")
        return

    results = process_new_orders(session, new_orders, order_map)
    for ebeln, rows in results.items():
        try:
            write_orders_to_excel(rows)
            mark_processed(ebeln, processed)
            save_processed(processed)
            logger.info(f"VL06O 오더 {ebeln} Excel 기록 완료 ({len(rows)}행)")
        except Exception as e:
            logger.error(f"VL06O 오더 {ebeln} Excel 기록 실패: {e}", exc_info=True)


def _run_vl10g(session, processed):
    """VL10G: Block 해제 + VL06O 이관 + ZRE 수집."""
    try:
        if not is_on_vl10g_list(session):
            navigate_to_vl10g(session)
        else:
            refresh_vl10g(session)
    except Exception as e:
        logger.error(f"VL10G 이동 실패: {e}")
        return

    result = process_vl10g(session, processed)

    # Block 해제 실패 오더 → Excel에 알림 기록
    if result['failed_blocks']:
        logger.warning(f"수동 Block 해제 필요: {result['failed_blocks']}")
        _write_block_failed_notice(result['failed_blocks'])

    if result['pushed_to_vl06o']:
        logger.info(f"VL06O로 이관된 오더: {result['pushed_to_vl06o']}")

    # ZRE 회수 오더 처리 (TODO: 구현 예정)
    for order_num, rows in result['zre_orders'].items():
        if rows:
            try:
                write_orders_to_excel(rows)
                mark_processed(order_num, processed)
                save_processed(processed)
                logger.info(f"ZRE 오더 {order_num} Excel 기록 완료")
            except Exception as e:
                logger.error(f"ZRE 오더 {order_num} 기록 실패: {e}")


def _run_zrma(session, processed, variant, date_mode):
    """ZRMA_Q: 6-prefix + 회수 포함 4/5-prefix 처리."""
    try:
        if not is_on_zrma_list(session):
            navigate_to_zrma_q(session, variant, date_mode)
        else:
            refresh_zrma_list(session)
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

    results = process_zrma_orders(session, new_orders, order_map)
    for order_num, rows in results.items():
        try:
            write_orders_to_excel(rows)
            mark_processed(order_num, processed)
            save_processed(processed)
            logger.info(f"ZRMA 오더 {order_num} Excel 기록 완료 ({len(rows)}행)")
        except Exception as e:
            logger.error(f"ZRMA 오더 {order_num} Excel 기록 실패: {e}", exc_info=True)


def _write_block_failed_notice(failed_list):
    """Block 해제 실패 오더를 Excel H열에 '수동 해제 필요' 메모로 기록."""
    for item in failed_list:
        rows = [{
            'order_prefix': '확인필요',
            'order_type':   item.get('doc_type', ''),
            'order_num':    item.get('orig_doc', ''),
            'extra_orders': [],
            'material':     '',
            'description':  '',
            'customer':     '',
            'phone':        '',
            'company':      '',
            'street':       '',
            'street2':      '',
            'memo':         '[수동 해제 필요] Delivery Block 자동 해제 실패. VA02에서 확인하세요.',
            'is_first_item': True,
        }]
        try:
            write_orders_to_excel(rows)
            logger.info(f"Block 실패 알림 기록: {item['orig_doc']}")
        except Exception as e:
            logger.error(f"Block 실패 알림 기록 실패: {e}")


def run_loop():
    logger.info(f"자동화 시작 (간격: {REFRESH_INTERVAL_MINUTES}분)")
    while True:
        run_once()
        logger.info(f"{REFRESH_INTERVAL_MINUTES}분 후 다음 실행...")
        time.sleep(REFRESH_INTERVAL_MINUTES * 60)


if __name__ == "__main__":
    if "--once" in sys.argv:
        run_once()
    else:
        run_loop()

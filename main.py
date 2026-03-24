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
import os
import time
import logging
from datetime import datetime

from config import REFRESH_INTERVAL_MINUTES, LOG_FILE
from order_tracker import load_processed, mark_processed, save_processed, purge_old_processed
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
from excel_handler import write_orders_to_excel, write_kakao_sent, get_existing_order_numbers
from kakao_handler import send_kakao_order

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)


def _write_and_notify(rows, order_id, processed):
    """Excel 기록 → 카카오 전송 → J열 기록 → processed 저장."""
    result = write_orders_to_excel(rows)
    if result:
        ws, start_row, end_row = result
        try:
            ok = send_kakao_order(rows)
            if ok:
                write_kakao_sent(ws, start_row, end_row)
                logger.info(f"카카오 전송 완료 ({order_id})")
            else:
                logger.warning(f"카카오 전송 실패 ({order_id}): send_kakao_order returned False")
        except Exception as e:
            logger.warning(f"카카오 전송 예외 ({order_id}): {e}")
    mark_processed(order_id, processed)
    save_processed(processed)


def run_once(excel_only=False):
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
            _write_and_notify(rows, ebeln, processed)
            logger.info(f"VL06O 오더 {ebeln} Excel 기록 완료 ({len(rows)}행)")
        except Exception as e:
            logger.error(f"VL06O 오더 {ebeln} Excel 기록 실패: {e}", exc_info=True)


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
        try:
            _write_and_notify(rows, order_num, processed)
            logger.info(f"VL10G Block 오더 {order_num} Excel 기록 완료")
        except Exception as e:
            logger.error(f"VL10G Block 오더 {order_num} 기록 실패: {e}")

    if result['pushed_to_vl06o']:
        logger.info(f"VL06O로 이관된 오더: {result['pushed_to_vl06o']}")

    # ZRE 회수 오더 처리 (VL10G ZS 해제 + S/N 수집 완료)
    for order_num, rows in result['zre_orders'].items():
        if rows:
            try:
                _write_and_notify(rows, order_num, processed)
                logger.info(f"ZRE 오더 {order_num} Excel 기록 완료")
            except Exception as e:
                logger.error(f"ZRE 오더 {order_num} 기록 실패: {e}")


def _run_zrma(session, processed, variant, date_mode):
    """ZRMA_Q: 6-prefix + 회수 포함 4/5-prefix 처리."""
    try:
        if not is_on_zrma_list(session):
            navigate_to_zrma_q(session, variant, date_mode)
        else:
            # 이미 목록 화면 → F3 복귀 + F8 재실행으로 DB 재조회
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
            _write_and_notify(rows, order_num, processed)
            logger.info(f"ZRMA 오더 {order_num} Excel 기록 완료 ({len(rows)}행)")
        except Exception as e:
            logger.error(f"ZRMA 오더 {order_num} Excel 기록 실패: {e}", exc_info=True)



TRIGGER_FILE = os.path.join(os.path.dirname(__file__), "run_now.flag")
TRIGGER_FILES = {
    i: os.path.join(os.path.dirname(__file__), f"run_now_{i}.flag")
    for i in range(4)
}


def _run_session(sess_idx):
    """개별 세션 단독 실행."""
    purge_old_processed()
    processed = load_processed()
    logger.info(f"=== 세션{sess_idx} 단독 실행 ===")
    try:
        if sess_idx == 0:
            _run_vl06o(get_sap_session(0), processed)
        elif sess_idx == 1:
            _run_vl10g(get_sap_session(1), processed)
        elif sess_idx == 2:
            _run_zrma(get_sap_session(2), processed, variant='rlkr', date_mode='next_month')
        elif sess_idx == 3:
            _run_zrma(get_sap_session(3), processed, variant='q2', date_mode='year_end')
    except ConnectionError as e:
        logger.error(f"세션{sess_idx} 연결 실패: {e}")


LUNCH_START = (11, 20)  # 11:20
LUNCH_END   = (13, 20)  # 13:20


def _is_lunch_break():
    """평일 점심시간(11:20~13:20) 여부 확인."""
    now = datetime.now()
    if now.weekday() >= 5:  # 토/일 제외
        return False
    t = (now.hour, now.minute)
    return LUNCH_START <= t < LUNCH_END


def run_loop():
    logger.info(f"자동화 시작 (간격: {REFRESH_INTERVAL_MINUTES}분, 점심 휴식: {LUNCH_START[0]:02d}:{LUNCH_START[1]:02d}~{LUNCH_END[0]:02d}:{LUNCH_END[1]:02d})")
    logger.info(f"수동 실행: 다른 터미널에서 'python now.py' 실행")
    while True:
        if _is_lunch_break():
            logger.info("점심시간 — 실행 대기 중...")
            # 점심 끝날 때까지 5초마다 대기 (수동 트리거는 허용)
            while _is_lunch_break():
                time.sleep(5)
                if os.path.exists(TRIGGER_FILE):
                    os.remove(TRIGGER_FILE)
                    logger.info("▶ 점심 중 수동 전체 실행 트리거 감지 → 즉시 실행")
                    run_once()
                for sess_idx, flag_path in TRIGGER_FILES.items():
                    if os.path.exists(flag_path):
                        os.remove(flag_path)
                        logger.info(f"▶ 점심 중 세션{sess_idx} 단독 실행 트리거 감지")
                        _run_session(sess_idx)
                        break
            logger.info("점심시간 종료 → 재개")

        run_once()
        logger.info(f"{REFRESH_INTERVAL_MINUTES}분 후 다음 실행... (수동: now.py / now.py 0~3)")
        # 5초마다 trigger 파일 확인
        for _ in range(REFRESH_INTERVAL_MINUTES * 60 // 5):
            time.sleep(5)
            # 전체 실행 트리거
            if os.path.exists(TRIGGER_FILE):
                os.remove(TRIGGER_FILE)
                logger.info("▶ 수동 전체 실행 트리거 감지 → 즉시 실행")
                break
            # 개별 세션 트리거
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

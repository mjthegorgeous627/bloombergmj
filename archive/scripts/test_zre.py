"""
ZRE 66991011 테스트: VL10G에서 ZS 블록 해제 + S/N 수집 + Excel 기록.
"""
import sys
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

TARGET_ORDER = '66991011'

from sap_handler import get_sap_session
from vl10g_handler import (
    navigate_to_vl10g, is_on_vl10g_list, refresh_vl10g,
    get_all_rows_from_vl10g, process_zre_with_block_release,
)
from excel_handler import write_orders_to_excel

def main():
    logger.info(f"=== ZRE {TARGET_ORDER} 테스트 ===")

    # VL10G 세션 (세션1)
    try:
        sess = get_sap_session(1)
    except ConnectionError as e:
        logger.error(f"SAP 세션1 연결 실패: {e}")
        sys.exit(1)

    # VL10G 화면인지 확인, 아니면 진입
    if not is_on_vl10g_list(sess):
        logger.info("VL10G 진입 중...")
        navigate_to_vl10g(sess)
    else:
        logger.info("VL10G 새로고침...")
        refresh_vl10g(sess)

    # 전체 행 읽기
    rows = get_all_rows_from_vl10g(sess)
    logger.info(f"VL10G 행 수: {len(rows)}")

    # 66991011 찾기
    target_row = None
    for r in rows:
        if r['orig_doc'] == TARGET_ORDER:
            target_row = r
            break

    if not target_row:
        logger.error(f"VL10G에서 {TARGET_ORDER} 을 찾을 수 없음")
        logger.info(f"현재 VL10G 오더 목록: {[r['orig_doc'] for r in rows]}")
        sys.exit(1)

    logger.info(f"오더 발견: {target_row}")

    # process_zre_with_block_release 실행
    success, data = process_zre_with_block_release(
        sess,
        grid_idx=target_row['grid_idx'],
        orig_doc=target_row['orig_doc'],
        order_type=target_row['doc_type'],
        name1=target_row.get('name1', ''),
    )

    if not success:
        logger.error(f"처리 실패: {data}")
        sys.exit(1)

    logger.info(f"처리 성공 ({len(data)}행):")
    for i, row in enumerate(data):
        logger.info(f"  행{i+1}: {row.get('order_prefix')} | {row.get('description')} | S/N: {row.get('serial_number')}")

    # Excel 기록
    logger.info("Excel 기록 중...")
    write_orders_to_excel(data)
    logger.info("완료!")

if __name__ == '__main__':
    main()

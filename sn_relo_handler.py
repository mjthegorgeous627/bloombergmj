"""
회수 오더에 실제 회수된 S/N을 반영 시도하는 자동화 (2026-08-19).

배경: 오더 수집 시 회수 아이템에 SAP상 S/N이 없어서('X') 워크벤치에 X로
찍힌 경우, 사용자가 실제로 회수한 S/N을 워크벤치 대시보드에 수동 입력해
둔다. 그 S/N을 VA02로 오더를 열어 Technical Objects > Serial Numbers에
입력하고 저장을 시도하는 것이 이 스크립트의 역할 - 저장이 되면 그 S/N의
firm/cust와 오더의 firm/cust가 일치한다는 뜻이라 그대로 ZREC 진행 가능,
저장이 안 되면(에러) firm/cust 불일치라 paper relo 프로세스로 가야 한다.

실제 SAP 화면 조작(VA02 열기 → 회수 아이템 행 찾기 → Technical Objects →
기존 S/N 확인 → 없으면 입력 → 저장 → 상태바로 성공/실패 판정)은 전부
zrma_handler.input_sn_to_return_item()에 있음 - 이 파일은 그 함수를 위한
독립 SAP 세션을 열고/닫는 CLI 래퍼일 뿐이다 (zrec_handler.py와 같은
"subprocess 한 번 = 세션 하나 열고 쓰고 닫기" 패턴, workbench_app.py가
직접 SAP 세션을 들고 있지 않는 이 레포 전체의 아키텍처를 따름).

기존 VA02 세션을 찾아 재사용하지 않고 **항상 새 세션을 연다** - zrec_handler.
get_session()과 다른 점: 사용자가 수동으로 다른 오더를 VA02로 보고 있는
세션을 이 자동화가 실수로 가로채 엉뚱한 오더를 저장해버리는 사고를
피하기 위함. 세션은 작업 후 항상 닫는다(session count가 배치 처리 중에도
계속 늘어나지 않도록) - inspect_return_serial_popup.py가 이미 쓰던 안전한
close 패턴과 동일.

사용법:
  python sn_relo_handler.py fill --order 67078303 --serial 70631395
      (RESULT_JSON: 로 시작하는 한 줄로 결과 출력 - workbench_app.py용)
"""

import argparse
import json
import logging
import time

from sap_handler import get_scripting_engine
from zrma_handler import input_sn_to_return_item

logger = logging.getLogger(__name__)

MAX_SESSIONS = 6


def _connection():
    app = get_scripting_engine()
    return app.Children(0)


def _open_new_session():
    """새 SAP 세션을 연다. open_session.py/zrec_handler.py와 같은 이유로
    '마지막 child = 새 세션' 대신 SessionNumber diff로 진짜 새 세션만 잡는다
    (2026-08-18 실측된 세션 재사용 버그 회피 - 자세한 내용은
    zrec_handler._open_new_session()의 docstring 참고)."""
    conn = _connection()
    if conn.Children.Count >= MAX_SESSIONS:
        raise RuntimeError(
            f"SAP GUI 세션이 이미 최대({MAX_SESSIONS}개)입니다 - "
            "안 쓰는 세션을 먼저 닫으세요."
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
        raise RuntimeError("새 세션 생성 실패 (시간 초과)")
    return new_sess


def fill_sn_to_order(order_no, serial):
    """오더 1건에 S/N 1개 반영 시도. 결과는
    zrma_handler.input_sn_to_return_item()의 반환값 그대로
    (status: 'skipped'|'saved'|'error') + order/serial을 덧붙인 것."""
    session = _open_new_session()
    try:
        result = input_sn_to_return_item(session, order_no, serial)
        result["order"] = order_no
        result["serial"] = serial
        return result
    finally:
        try:
            session.findById("wnd[0]").close()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(
        description="회수 오더에 실제 회수된 S/N 반영 시도 (paper relo 필요 여부 판별)"
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_fill = sub.add_parser("fill", help="오더에 S/N 입력·저장 시도")
    p_fill.add_argument("--order", required=True)
    p_fill.add_argument("--serial", required=True)

    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    if args.cmd == "fill":
        try:
            result = fill_sn_to_order(args.order, args.serial)
            print("RESULT_JSON:" + json.dumps(result, ensure_ascii=False))
        except Exception as exc:
            logger.error(f"S/N 반영 실패: {exc}")
            print("RESULT_JSON:" + json.dumps(
                {"status": "error", "existing_sns": [], "message": str(exc),
                 "order": args.order, "serial": args.serial},
                ensure_ascii=False,
            ))
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

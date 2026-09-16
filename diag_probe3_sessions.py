"""
1번 기능(SAP 시작 & 자동루프) 중 '4개 세션 열고 화면 설정'까지만 검증.
Excel 열기 / catchup 실행 / 무한 루프는 하지 않는다.
"""

import sys
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)],
)

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from sap_handler import get_scripting_engine
from startup import setup_sap_sessions


def main():
    app = get_scripting_engine()
    conn = app.Children(0)

    print(f"시작 전 세션 수: {conn.Children.Count}")
    setup_sap_sessions(conn)

    print(f"\n최종 세션 수: {conn.Children.Count}")
    for i in range(conn.Children.Count):
        try:
            sess = conn.Children(i)
            title = sess.findById("wnd[0]").Text
            tcode = sess.Info.Transaction
            print(f"  세션[{i}] title='{title}' tcode={tcode}")
        except Exception as e:
            print(f"  세션[{i}]: 읽기 실패 ({e})")


if __name__ == "__main__":
    main()

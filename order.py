"""Simple command: python order.py <order_number>"""

import logging
import sys

from config import LOG_FILE
from manual_order_handler import run_manual_order

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
    ],
)


def main():
    if len(sys.argv) < 2:
        print("사용법: python order.py 오더번호")
        print("예시:   python order.py 67025693")
        return 1

    order_num = sys.argv[1].strip()
    ok = run_manual_order(order_num)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

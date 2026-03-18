"""수동 실행 트리거 - startup.py가 실행 중일 때 즉시 새 오더 확인."""
import os
flag = os.path.join(os.path.dirname(__file__), "run_now.flag")
open(flag, 'w').close()
print("▶ 즉시 실행 요청됨 (최대 5초 내 시작)")

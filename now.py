"""
수동 실행 트리거.

사용법:
  python now.py        → 전체 4개 세션 실행
  python now.py 0      → 세션0 VL06O만
  python now.py 1      → 세션1 VL10G만
  python now.py 2      → 세션2 ZRMA_Q RLKR만
  python now.py 3      → 세션3 ZRMA_Q Q2만
"""
import os, sys

SESSIONS = {
    '0': 'VL06O',
    '1': 'VL10G',
    '2': 'ZRMA RLKR',
    '3': 'ZRMA Q2',
}

base = os.path.dirname(__file__)
arg = sys.argv[1] if len(sys.argv) > 1 else None

if arg is None:
    flag = os.path.join(base, "run_now.flag")
    open(flag, 'w').close()
    print("▶ 전체 실행 요청됨 (최대 5초 내 시작)")
elif arg in SESSIONS:
    flag = os.path.join(base, f"run_now_{arg}.flag")
    open(flag, 'w').close()
    print(f"▶ 세션{arg} ({SESSIONS[arg]}) 단독 실행 요청됨 (최대 5초 내 시작)")
else:
    print(f"오류: 세션 번호는 0~3 중 하나여야 합니다. (입력값: {arg})")
    print("사용법: python now.py [0|1|2|3]")

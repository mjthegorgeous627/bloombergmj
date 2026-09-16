"""
오더 단위 프로세스 간 락.

문제: 자동 루프(main.py)가 오더를 수집~Excel 기록하는 도중, 사용자가 같은
오더에 "오더번호 반영"(order.py/수동 오더 진입)을 실행하면 두 프로세스가
동시에 같은 오더를 재수집하고 같은 라이브 Excel 파일(xlwings/COM)에 동시에
쓰게 된다. Excel COM은 여러 프로세스의 동시 조작을 안전하게 처리하지
못하므로(잠금/트랜잭션 없음) 어느 한쪽 쓰기가 몇 분씩 멈춰버리거나(사용자
눈에는 "에러도 없이 안 됨"으로 보임), 두 쓰기가 뒤섞여 같은 오더가 행이
중복/오염된 채로 들어갈 수 있다 (2026-08-14, ZRX 67076714 실제 사고 - 배송
1개/회수 1개여야 할 오더가 회수 2개로 중복 기록됨).

이 모듈은 오더번호별 파일 기반 락으로 "이 오더, 지금 다른 프로세스가 이미
수집~기록 중"을 감지해 두 번째 시도를 즉시(SAP 진입도 하기 전에) 막는다.
"""

import os
import time

_LOCK_DIR = os.path.join(os.path.dirname(__file__), "order_locks")
STALE_SECONDS = 30 * 60  # 30분 - 정상 처리 소요시간(수 분)보다 넉넉히 크게.
                          # 이보다 오래된 락은 죽은 프로세스가 남긴 것으로
                          # 간주하고 무시(탈취)한다.


def _lock_path(order_num):
    return os.path.join(_LOCK_DIR, f"{str(order_num).strip()}.lock")


def try_acquire(order_num):
    """
    이 오더에 대한 락을 시도한다.
    반환: True(획득 성공, release() 필요) / False(다른 프로세스가 보유 중).
    """
    os.makedirs(_LOCK_DIR, exist_ok=True)
    path = _lock_path(order_num)

    # 기존 락이 있으면 stale 여부만 확인 (내용은 참고용, 신뢰하지 않음).
    try:
        age = time.time() - os.path.getmtime(path)
        if age < STALE_SECONDS:
            return False
        # stale → 탈취 시도 (아래서 다시 원자적으로 생성)
        try:
            os.remove(path)
        except OSError:
            pass
    except FileNotFoundError:
        pass

    try:
        fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        with os.fdopen(fd, "w") as f:
            f.write(f"pid={os.getpid()} ts={time.time()}\n")
        return True
    except FileExistsError:
        return False


def release(order_num):
    try:
        os.remove(_lock_path(order_num))
    except OSError:
        pass


class OrderLock:
    """with OrderLock(order_num) as ok: ok가 False면 락을 못 잡은 것 - 처리하지 말 것."""

    def __init__(self, order_num):
        self.order_num = order_num
        self.acquired = False

    def __enter__(self):
        self.acquired = try_acquire(self.order_num)
        return self.acquired

    def __exit__(self, exc_type, exc, tb):
        if self.acquired:
            release(self.order_num)
        return False


def acquire_blocking(name, timeout_seconds=90, poll_seconds=0.3):
    """try_acquire()의 블로킹 버전. OrderLock은 "같은 오더면 이번 회차는
    건너뛰고 다음 20분 뒤 재시도"가 맞는 동작이라 non-blocking으로 충분했지만,
    이 함수를 쓰는 락(EXCEL_WRITE_LOCK_NAME)에서 그냥 건너뛰면 그 오더가
    Excel에 안 써진 채로 mark_processed()까지 불려서 다시는 재시도되지 않는다
    (main.py의 _write_and_notify는 write_orders_to_excel()이 예외 없이
    반환해야만 그 뒤 mark_processed를 부르므로, "건너뛴다"를 표현하려면 실제로
    예외를 던져야 한다) - 그래서 잠깐 기다렸다 재시도하고, 그래도 안 되면
    TimeoutError를 던져서 호출자 쪽 기존 except Exception 처리가 "이번
    사이클 실패, 다음 사이클에 재시도"로 다루게 한다."""
    deadline = time.time() + timeout_seconds
    while True:
        if try_acquire(name):
            return
        if time.time() >= deadline:
            raise TimeoutError(
                f"'{name}' 락 획득 실패 - {timeout_seconds}초 대기 후 타임아웃 "
                "(다른 프로세스가 오래 점유 중이거나 죽은 채 락을 남긴 상태일 수 있음)"
            )
        time.sleep(poll_seconds)


# 2026-08-24 구조 점검: OrderLock은 "같은 오더 번호"끼리만 서로 막아준다.
# 그런데 main.py 루프가 오더 A를 쓰는 도중 사용자가 order.py로 오더 B를
# 수동 입력하면 - 둘 다 오더 번호가 다르니 OrderLock은 통과시키고, 둘 다
# 그대로 같은 라이브 Excel COM 세션(xlwings로 붙잡은 같은 Excel.Application)에
# 동시에 Insert()/Merge()를 건다. 위 모듈 docstring이 원래 설명하는 위험
# ("Excel COM은 여러 프로세스의 동시 조작을 안전하게 처리하지 못해 몇 분씩
# 멈추거나 뒤섞인다")은 사실 오더 번호와 무관하게 "동시에 같은 파일을
# 건드리는" 모든 경우에 적용되는 얘기였는데, 실제 고쳐진 건 그중 "같은 오더"
# 케이스뿐이었다. 이 락은 오더 번호와 무관하게 write_orders_to_excel()/
# write_kakao_sent() 같은 실제 Excel COM 쓰기 구간 전체를 전역 직렬화한다 -
# SAP 스캔/조회 등 Excel을 안 건드리는 나머지 작업은 여전히 완전 병렬로 돈다.
EXCEL_WRITE_LOCK_NAME = "__excel_write__"


class ExcelWriteLock:
    """with ExcelWriteLock(): ... - 위 EXCEL_WRITE_LOCK_NAME 설명 참고.
    OrderLock과 달리 항상 락을 얻거나(블로킹) TimeoutError를 던진다 -
    "얻었는지 확인 후 건너뛰기" 패턴이 아니다."""

    def __enter__(self):
        acquire_blocking(EXCEL_WRITE_LOCK_NAME)
        return self

    def __exit__(self, exc_type, exc, tb):
        release(EXCEL_WRITE_LOCK_NAME)
        return False

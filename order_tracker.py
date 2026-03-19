import json
import os
from datetime import date, timedelta
from config import PROCESSED_ORDERS_FILE

PURGE_DAYS = 45  # 45일 이상 된 오더 자동 삭제


def _parse_date(date_str):
    try:
        return date.fromisoformat(date_str)
    except Exception:
        return date.today()


def _load_raw():
    """JSON 원본 로드. 구형(list) → dict 자동 변환."""
    if not os.path.exists(PROCESSED_ORDERS_FILE):
        return {}
    with open(PROCESSED_ORDERS_FILE, 'r') as f:
        data = json.load(f)
    if isinstance(data, list):
        # 구형 포맷: 날짜 모르므로 오늘 날짜로 변환 (90일간 유지)
        return {str(k).strip(): str(date.today()) for k in data}
    return {str(k).strip(): v for k, v in data.items()}


def load_processed():
    """처리된 오더 ID set 반환."""
    return set(_load_raw().keys())


def save_processed(order_ids):
    """
    현재 processed set을 저장.
    기존 항목은 날짜 보존, 새 항목은 오늘 날짜 부여.
    """
    raw = _load_raw()
    today_str = str(date.today())
    new_raw = {}
    for oid in order_ids:
        key = str(oid).strip()
        new_raw[key] = raw.get(key, today_str)
    with open(PROCESSED_ORDERS_FILE, 'w') as f:
        json.dump(new_raw, f, indent=2, ensure_ascii=False)


def mark_processed(order_id, processed_set):
    processed_set.add(str(order_id).strip())
    save_processed(processed_set)


def purge_old_processed(days=PURGE_DAYS):
    """N일 초과 오더 자동 삭제. 반환: 삭제된 수."""
    raw = _load_raw()
    cutoff = date.today() - timedelta(days=days)
    before = len(raw)
    raw = {k: v for k, v in raw.items() if _parse_date(v) >= cutoff}
    after = len(raw)
    if before != after:
        with open(PROCESSED_ORDERS_FILE, 'w') as f:
            json.dump(raw, f, indent=2, ensure_ascii=False)
    return before - after

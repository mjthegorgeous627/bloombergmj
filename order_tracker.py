import json
import os
from config import PROCESSED_ORDERS_FILE


def load_processed():
    if os.path.exists(PROCESSED_ORDERS_FILE):
        with open(PROCESSED_ORDERS_FILE, 'r') as f:
            return set(json.load(f))
    return set()


def save_processed(order_ids):
    with open(PROCESSED_ORDERS_FILE, 'w') as f:
        json.dump(list(order_ids), f, indent=2)


def mark_processed(order_id, processed_set):
    processed_set.add(str(order_id).strip())
    save_processed(processed_set)

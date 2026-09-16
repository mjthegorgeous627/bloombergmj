"""
카카오톡 핸들러 - REST API로 나에게 보내기.

최초 설정: python kakao_setup.py  (1회만)
이후 자동으로 토큰 갱신됨. 비용 없음.
"""

import json
import logging
import os
import re
import requests
from datetime import datetime

logger = logging.getLogger(__name__)

TOKENS_FILE = os.path.join(os.path.dirname(__file__), "kakao_tokens.json")
SEND_URL    = "https://kapi.kakao.com/v2/api/talk/memo/default/send"
REFRESH_URL = "https://kauth.kakao.com/oauth/token"

_ORDER_PREFIX_RE = re.compile(r'^(배송|회수)')

PRODUCT_MAP = [
    (['KEYBOARD', '5'],  '키보드5'),
    (['KEYBOARD'],       '일반키보드'),
    (['MONITOR'],        '모니터'),
    (['PC'],             'PC'),
    (['ROUTER'],         '라우터'),
    (['SERVER'],         '서버'),
]

US_KEYBOARD_MATERIALS = {'10045246'}


# ── 제품명 변환 ──────────────────────────────────────────────────────────────

def _get_product_name(description, material=''):
    desc = str(description).upper()
    material = str(material or '').strip()
    if material in US_KEYBOARD_MATERIALS or ('KEYBOARD' in desc and re.search(r'\bUS\b', desc)):
        return 'US일반키보드'
    for keywords, name in PRODUCT_MAP:
        if all(k in desc for k in keywords):
            return name
    return str(description).split()[0] if description else ''


# ── 메시지 포맷 ──────────────────────────────────────────────────────────────

def format_kakao_message(order_data_list):
    """
    SAP 오더 데이터 → 카카오톡 메시지 텍스트.

    형식:
      키보드 교체
      배송 키보드5 10063421
      배송 일반키보드 10045196 X 2개
      회수 일반키보드 10045196
      CHULMIN KANG / 02-2004-9908
      SHINYOUNG SECURITIES CO LTD 34-8 YOIDO-DONG

    같은 (구분/제품/자재번호) 조합은 줄바꿈 없이 합쳐서 "X n개"로 표시한다.
    """
    if not order_data_list:
        return ''

    prefixes = []
    products = []
    item_counts = {}  # (prefix, product, mn) -> count, 등장 순서 보존

    for item in order_data_list:
        prefix  = item.get('order_prefix', '')
        product = _get_product_name(item.get('description', ''), item.get('material', ''))
        mn      = item.get('material', '')

        prefixes.append(prefix)
        products.append(product)
        key = (prefix, product, mn)
        item_counts[key] = item_counts.get(key, 0) + 1

    item_lines = []
    for (prefix, product, mn), count in item_counts.items():
        line = f"{prefix} {product} {mn}"
        if count > 1:
            line += f" X {count}개"
        item_lines.append(line)

    # 제목
    unique_products = list(dict.fromkeys(products))
    prod_str = '/'.join(unique_products)
    prefix_set = set(prefixes)
    if '배송' in prefix_set and '회수' in prefix_set:
        action = '교체'
    elif '배송' in prefix_set:
        action = '배송'
    else:
        action = '회수'
    title = f"[{prod_str} {action}]"

    # 고객 정보 (첫 번째 행 기준)
    r0 = order_data_list[0]
    customer_line = f"{r0.get('customer', '')} / {r0.get('phone', '')}"

    # 주소 한 줄 조합
    parts = []
    if r0.get('company'):
        parts.append(r0['company'])
    if r0.get('street'):
        parts.append(r0['street'])
    if r0.get('street2'):
        parts.append(r0['street2'])
    address_line = ' '.join(parts)

    lines = [title] + item_lines + [customer_line, address_line]
    return '\n'.join(line for line in lines if line.strip())


# ── 토큰 관리 ────────────────────────────────────────────────────────────────

def _load_tokens():
    with open(TOKENS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def _save_tokens(tokens):
    with open(TOKENS_FILE, 'w', encoding='utf-8') as f:
        json.dump(tokens, f, ensure_ascii=False, indent=2)


def _refresh_access_token(tokens):
    data = {
        'grant_type':   'refresh_token',
        'client_id':    tokens['rest_api_key'],
        'refresh_token': tokens['refresh_token'],
    }
    if tokens.get('client_secret'):
        data['client_secret'] = tokens['client_secret']
    resp = requests.post(REFRESH_URL, data=data, timeout=10)
    resp.raise_for_status()
    result = resp.json()
    tokens['access_token'] = result['access_token']
    if 'refresh_token' in result:
        tokens['refresh_token'] = result['refresh_token']
    _save_tokens(tokens)
    logger.info("[카카오] 토큰 갱신 완료")
    return tokens


# ── 전송 ────────────────────────────────────────────────────────────────────

def send_kakao_message(text):
    """
    카카오 나에게 보내기. 토큰 만료 시 자동 갱신.
    반환: True(성공) / False(실패)
    """
    if not os.path.exists(TOKENS_FILE):
        logger.error("[카카오] kakao_tokens.json 없음 → python kakao_setup.py 를 먼저 실행하세요")
        return False

    tokens = _load_tokens()

    template = {
        "object_type": "text",
        "text": text,
        "link": {"web_url": "", "mobile_web_url": ""},
    }

    def _do_send(access_token):
        headers = {"Authorization": f"Bearer {access_token}"}
        data    = {"template_object": json.dumps(template, ensure_ascii=False)}
        return requests.post(SEND_URL, headers=headers, data=data, timeout=10)

    resp = _do_send(tokens['access_token'])

    # 401 = 토큰 만료 → 갱신 후 재시도
    if resp.status_code == 401:
        logger.warning("[카카오] 액세스 토큰 만료 → 갱신 중...")
        try:
            tokens = _refresh_access_token(tokens)
            resp = _do_send(tokens['access_token'])
        except Exception as e:
            logger.error(f"[카카오] 토큰 갱신 실패: {e}")
            return False

    if resp.status_code == 200 and resp.json().get('result_code') == 0:
        logger.info("[카카오] 전송 성공")
        return True
    else:
        logger.error(f"[카카오] 전송 실패: {resp.status_code} {resp.text}")
        return False


def send_kakao_order(order_data_list):
    """오더 데이터 → 포맷 → 카카오 전송."""
    message = format_kakao_message(order_data_list)
    if not message:
        logger.warning("[카카오] 메시지 생성 실패 (데이터 없음)")
        return False
    logger.info(f"[카카오] 메시지:\n{'-'*30}\n{message}\n{'-'*30}")
    return send_kakao_message(message)

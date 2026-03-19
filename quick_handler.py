"""
퀵 배송 문서 출력 핸들러.

사용법:
  python main.py --quick   → Excel 마지막 오더 기준으로 퀵.docx 작성 후 출력
"""

import os
import re
import shutil
import tempfile
import requests
import win32com.client
from docx import Document
from datetime import datetime

from excel_handler import get_workbook, get_today_sheet_name

TEMPLATE_PATH = r"C:\1\May\퀵.docx"
PRINTER_NAME  = "Samsung C56x Series"

_DATE_RE = re.compile(r'(\d{1,2})월\s*(\d{1,2})일')


# ── 카카오 주소/키워드 검색 ──────────────────────────────────────────────────

def _load_rest_api_key():
    tokens_file = os.path.join(os.path.dirname(__file__), "kakao_tokens.json")
    if not os.path.exists(tokens_file):
        return None
    import json
    with open(tokens_file, encoding='utf-8') as f:
        return json.load(f).get('rest_api_key')


def _kakao_address_search(query, api_key):
    """영문 주소 → 카카오 주소 검색 → 한국어 도로명주소 반환."""
    try:
        resp = requests.get(
            "https://dapi.kakao.com/v2/local/search/address.json",
            headers={"Authorization": f"KakaoAK {api_key}"},
            params={"query": query, "analyze_type": "similar"},
            timeout=5
        )
        docs = resp.json().get('documents', [])
        if docs:
            ra = docs[0].get('road_address') or docs[0].get('address')
            if ra:
                return ra.get('address_name', '')
    except Exception:
        pass
    return ''


def _kakao_keyword_search(query, api_key):
    """영문 회사명 → 카카오 키워드 검색 → 한국어 장소명 반환."""
    try:
        resp = requests.get(
            "https://dapi.kakao.com/v2/local/search/keyword.json",
            headers={"Authorization": f"KakaoAK {api_key}"},
            params={"query": query, "size": 1},
            timeout=5
        )
        docs = resp.json().get('documents', [])
        if docs:
            return docs[0].get('place_name', '')
    except Exception:
        pass
    return ''


# ── 주소 포맷 ────────────────────────────────────────────────────────────────

def _split_address_lines(korean_addr, street2=''):
    """
    카카오 반환 한국어 주소 → (줄1, 줄2) 분리.
    줄1: 시/도 제외한 구+도로명+번지
    줄2: 상세주소(동/층/호) if any
    """
    # '서울 영등포구 의사당대로 88' → '영등포구 의사당대로 88'
    addr = re.sub(r'^[가-힣]+시\s*|^서울\s*|^부산\s*|^대구\s*|^인천\s*|^광주\s*|^대전\s*|^울산\s*|^세종\s*', '', korean_addr).strip()
    line2 = street2.strip() if street2 else ''
    return addr, line2


def _ask_user(label, suggestion=''):
    """터미널에서 사용자 확인/수정."""
    if suggestion:
        user_input = input(f"  {label} [{suggestion}]: ").strip()
        return user_input if user_input else suggestion
    else:
        return input(f"  {label}: ").strip()


# ── Excel에서 오더 읽기 ───────────────────────────────────────────────────────

def _find_order_in_excel(order_num):
    """
    전체 시트에서 오더번호가 A열에 포함된 행 그룹 찾기.
    반환: dict or None
    """
    wb = get_workbook()

    for ws in wb.sheets:
        max_row = ws.used_range.last_cell.row
        group_start = None
        end_row = None

        for r in range(1, max_row + 1):
            a_val = str(ws.cells(r, 1).value or '')
            if order_num in a_val:
                # 이 행이 속한 그룹의 시작(E열에 값 있는 행) 탐색
                g_start = r
                for up in range(r, 0, -1):
                    if ws.cells(up, 5).value:
                        g_start = up
                        break
                    if up < r and not ws.cells(up, 1).value:
                        break

                # 그룹 끝: 같은 E값(또는 E=None)이 유지되는 마지막 행
                g_end = r
                for down in range(r + 1, max_row + 1):
                    if ws.cells(down, 1).value and ws.cells(down, 5).value:
                        break
                    if ws.cells(down, 1).value:
                        g_end = down
                    elif not any(ws.cells(down, c).value for c in range(1, 9)):
                        break

                group_start = g_start
                end_row = g_end
                break

        if group_start is None:
            continue

        customer   = str(ws.cells(group_start, 5).value or '')
        phone      = str(ws.cells(group_start, 6).value or '')
        addr_raw   = str(ws.cells(group_start, 7).value or '')
        addr_parts = addr_raw.split('\n')
        company_en = addr_parts[0].strip() if addr_parts else ''
        street_en  = addr_parts[1].strip() if len(addr_parts) > 1 else ''
        street2_en = addr_parts[2].strip() if len(addr_parts) > 2 else ''

        items = []
        for r in range(group_start, end_row + 1):
            a_val = ws.cells(r, 1).value
            if not a_val:
                continue
            prefix = '배송' if '배송' in str(a_val) else '회수'
            desc   = str(ws.cells(r, 2).value or '')
            items.append({'prefix': prefix, 'description': desc})

        return {
            'customer':   customer,
            'phone':      phone,
            'company_en': company_en,
            'street_en':  street_en,
            'street2_en': street2_en,
            'items':      items,
        }

    return None


# ── 아이템 설명 생성 ─────────────────────────────────────────────────────────

_PRODUCT_MAP = [
    (['KEYBOARD', '5'],  '키보드5'),
    (['KEYBOARD'],       '일반키보드'),
    (['MONITOR'],        '모니터'),
    (['PC'],             'PC'),
    (['ROUTER'],         '라우터'),
    (['SERVER'],         '서버'),
]


def _get_product_name(description):
    desc = str(description).upper()
    for keywords, name in _PRODUCT_MAP:
        if all(k in desc for k in keywords):
            return name
    return description.split()[0] if description else ''


def _build_item_line(items):
    """
    예) 배송 키보드5 1개
        회수 일반키보드 1개
    여러 개면 각 줄에.
    """
    from collections import Counter
    lines = []
    counter = Counter()
    prefix_map = {}
    for item in items:
        prod = _get_product_name(item['description'])
        counter[prod] += 1
        prefix_map[prod] = item['prefix']
    for prod, cnt in counter.items():
        lines.append(f"{prefix_map[prod]} {prod} {cnt}개")
    return '\n'.join(lines)


# ── docx 작성 ───────────────────────────────────────────────────────────────

def _fill_docx(addr_line1, addr_line2, company_kr, name_en, phone, item_line):
    """
    퀵.docx 템플릿 복사 후 내용 채워서 임시 파일 경로 반환.
    Para[0]: 주소1 \n 주소2(+회사)
    Para[3]: 회사명 \n 이름 님 \n 전화번호
    Para[5]: 블룸버그 \n 아이템
    """
    tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
    tmp.close()
    shutil.copy2(TEMPLATE_PATH, tmp.name)

    doc = Document(tmp.name)
    paras = doc.paragraphs

    # Para[0]: 주소 블록
    # run[0] = 주소1, run[1] = '\n', run[2] = 주소2+회사명
    if len(paras) > 0 and len(paras[0].runs) >= 3:
        paras[0].runs[0].text = addr_line1
        paras[0].runs[2].text = f"{addr_line2} {company_kr}".strip() if addr_line2 else company_kr

    # Para[3]: 회사명 \n 이름 님 \n 전화번호
    if len(paras) > 3 and len(paras[3].runs) >= 6:
        paras[3].runs[0].text = company_kr
        paras[3].runs[2].text = name_en
        paras[3].runs[3].text = ' 님'
        paras[3].runs[5].text = phone
        paras[3].runs[6].text = ''

    # Para[5]: 블룸버그 \n 아이템
    if len(paras) > 5 and len(paras[5].runs) >= 3:
        paras[5].runs[2].text = item_line

    doc.save(tmp.name)
    return tmp.name


# ── 출력 ────────────────────────────────────────────────────────────────────

def _print_docx(docx_path):
    """Word COM으로 흑백 1매 출력."""
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))

        # 프린터 설정
        prev_printer = word.ActivePrinter
        ports = ['', ' on Ne00:', ' on Ne01:', ' on Ne02:', ' on Ne03:']
        for port in ports:
            try:
                word.ActivePrinter = PRINTER_NAME + port
                break
            except Exception:
                continue

        # 흑백
        try:
            doc.PageSetup.BlackAndWhite = True
        except Exception:
            pass

        doc.PrintOut(Copies=1)
        print(f"[퀵] 출력 완료")

        doc.Close(False)
        word.ActivePrinter = prev_printer
    finally:
        word.Quit()


# ── 공개 함수 ────────────────────────────────────────────────────────────────

def run_quick(order_num):
    """지정한 오더번호 기준으로 퀵 배송 문서 작성 후 출력."""
    print(f"\n[퀵 배송 출력] 오더번호: {order_num}")

    data = _find_order_in_excel(order_num)
    if not data:
        print(f"[퀵] 오더번호 {order_num} 를 Excel에서 찾을 수 없음")
        return False

    api_key = _load_rest_api_key()

    print("\n주소/회사명 자동 검색 중...")

    # 주소 변환
    korean_addr = ''
    if api_key and data['street_en']:
        korean_addr = _kakao_address_search(data['street_en'], api_key)
    addr_line1_suggest, addr_line2_suggest = _split_address_lines(korean_addr, data['street2_en'])

    # 회사명 변환
    company_kr_suggest = ''
    if api_key and data['company_en']:
        company_kr_suggest = _kakao_keyword_search(data['company_en'], api_key)

    print("\n─ 아래 내용을 확인하세요. 맞으면 Enter, 수정하려면 입력 후 Enter ─")
    addr_line1  = _ask_user("주소 1줄", addr_line1_suggest)
    addr_line2  = _ask_user("주소 2줄 (층/동 등, 없으면 Enter)", addr_line2_suggest)
    company_kr  = _ask_user("회사명 (한국어)", company_kr_suggest)
    name_en     = _ask_user("이름 (영문)", data['customer'])
    phone       = _ask_user("전화번호", data['phone'])

    item_line = _build_item_line(data['items'])
    print(f"  아이템: {item_line}")

    print("\n문서 작성 중...")
    tmp_path = _fill_docx(addr_line1, addr_line2, company_kr, name_en, phone, item_line)

    try:
        _print_docx(tmp_path)
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

    return True

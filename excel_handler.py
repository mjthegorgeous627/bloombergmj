"""
Excel 핸들러 - xlwings 사용으로 Excel 열린 상태에서도 실시간 데이터 입력 가능.
"""

import xlwings as xw
import openpyxl
import re
from datetime import datetime
from config import EXCEL_PATH


def get_today_sheet_name():
    """시트명 형식: '3-17' (월-일)."""
    today = datetime.today()
    return f"{today.month}-{today.day}"


def get_workbook():
    """
    이미 열려있는 Excel 파일에 연결. 없으면 파일 열기.
    반환: xlwings Book 객체
    """
    filename = EXCEL_PATH.split("\\")[-1]
    # Excel이 실행 중이면 열려있는 파일에서 찾기
    try:
        for book in xw.books:
            if book.name == filename:
                return book
    except Exception:
        pass
    # Excel이 없거나 파일이 안 열려있으면 열기
    return xw.Book(EXCEL_PATH)


def find_today_section_end(ws):
    """
    오늘 날짜 섹션의 마지막 데이터 행 번호 반환.
    xlwings Sheet 객체를 받아 처리.
    """
    today = datetime.today()
    today_pattern = f"{today.month}월 {today.day}일"
    section_keywords = ["Scheduled", "Delayed", "Bloomberg", "Client"]

    # 사용 중인 전체 범위
    used = ws.used_range
    max_row = used.last_cell.row
    max_col = 8  # H열까지

    # 오늘 날짜 헤더 행 찾기
    date_header_row = None
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = ws.cells(r, c).value
            if val and isinstance(val, str) and today_pattern in val:
                date_header_row = r
                break
        if date_header_row:
            break

    if date_header_row is None:
        return max_row

    # 다음 섹션 시작 행 찾기
    next_section_row = None
    for r in range(date_header_row + 2, max_row + 1):
        for c in range(1, max_col + 1):
            val = ws.cells(r, c).value
            if val and isinstance(val, str):
                if "년 " in val and "월 " in val and "일" in val and today_pattern not in val:
                    next_section_row = r
                    break
                if any(kw in val for kw in section_keywords):
                    next_section_row = r
                    break
        if next_section_row:
            break

    if next_section_row is None:
        return max_row

    # 마지막 데이터 행 (다음 섹션 직전)
    last_data_row = date_header_row + 1
    for r in range(date_header_row + 2, next_section_row):
        row_has_data = any(ws.cells(r, c).value for c in range(1, max_col + 1))
        if row_has_data:
            last_data_row = r

    return last_data_row


def write_orders_to_excel(order_data_list):
    """
    order_data_list: 한 오더의 아이템들 리스트.
    Excel이 열려있어도 실시간으로 반영됨.
    """
    wb = get_workbook()

    # 시트 찾기
    sheet_name = get_today_sheet_name()
    sheet_names = [s.name for s in wb.sheets]

    if sheet_name in sheet_names:
        ws = wb.sheets[sheet_name]
    else:
        # 월-일 패턴으로 탐색
        today = datetime.today()
        matched = None
        for sname in sheet_names:
            if str(today.month) in sname and str(today.day) in sname:
                matched = sname
                break
        ws = wb.sheets[matched] if matched else wb.sheets.active

    # 삽입 위치 결정
    insert_after_row = find_today_section_end(ws)
    insert_at = insert_after_row + 1
    num_rows = len(order_data_list)

    # 행 삽입: 한 행씩 삽입 (안정적)
    for _ in range(num_rows):
        ws.range(f"A{insert_at}:H{insert_at}").api.EntireRow.Insert()

    # 데이터 입력
    for idx, item in enumerate(order_data_list):
        row_num = insert_at + idx

        # A열: Order #
        order_line = f"{item['order_prefix']} {item['order_type']} {item['order_num']}"
        if item.get('extra_orders'):
            for eo in item['extra_orders']:
                order_line += f"\n{eo}"
        cell_a = ws.cells(row_num, 1)
        cell_a.value = order_line
        cell_a.api.WrapText = True

        # B열: item
        ws.cells(row_num, 2).value = item.get('description', '')

        # C열: M/N (숫자형 방지 - 문자열로 저장)
        ws.cells(row_num, 3).value = str(item.get('material', ''))

        # D열: S/N (비워둠)
        ws.cells(row_num, 4).value = ''

        # E~H열: 첫 번째 아이템에만 입력 (나중에 병합)
        if item.get('is_first_item', False):
            ws.cells(row_num, 5).value = item.get('customer', '')
            ws.cells(row_num, 6).value = item.get('phone', '')

            # G열: ADDRESS
            address_parts = []
            if item.get('company'):
                address_parts.append(item['company'])
            street = item.get('street', '')
            street2 = item.get('street2', '')
            if street and street2:
                address_parts.append(f"{street}, {street2}")
            elif street:
                address_parts.append(street)
            elif street2:
                address_parts.append(street2)
            cell_g = ws.cells(row_num, 7)
            cell_g.value = '\n'.join(address_parts)
            cell_g.api.WrapText = True

            # H열: memo
            if item.get('memo'):
                cell_h = ws.cells(row_num, 8)
                cell_h.value = item['memo']
                cell_h.api.WrapText = True

    # E~H열 병합: 같은 오더 여러 행인 경우
    if num_rows > 1:
        for col in (5, 6, 7, 8):  # E, F, G, H
            merge_rng = ws.range(ws.cells(insert_at, col), ws.cells(insert_at + num_rows - 1, col))
            merge_rng.api.MergeCells = True

    # 저장
    wb.save()
    print(f"[Excel] {num_rows}개 행 삽입 완료 → {ws.name} 시트 행 {insert_at}~{insert_at + num_rows - 1}")
    return True


def get_existing_order_numbers():
    """
    Excel 파일의 A열(Order #)에서 이미 입력된 오더번호(7자리 이상 숫자) 추출.
    processed_orders.json 초기화용으로 사용.
    """
    found = set()
    num_re = re.compile(r'\b(\d{7,})\b')
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                val = row[0]
                if val and isinstance(val, str):
                    for m in num_re.finditer(val):
                        found.add(m.group(1))
        wb.close()
    except Exception as e:
        print(f"[Excel] 기존 오더 스캔 실패: {e}")
    return found

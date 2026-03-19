"""
Excel 핸들러 - xlwings 사용으로 Excel 열린 상태에서도 실시간 데이터 입력 가능.
"""

import xlwings as xw
import openpyxl
import re
from datetime import datetime
from config import EXCEL_PATH

_DATE_RE = re.compile(r'(\d{1,2})월\s*(\d{1,2})일')


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
    구조: 시트 맨 위가 오늘 블록 (아주 가끔 전날 블록이 위에 있을 수 있음).
    날짜 헤더 형식 무관: '3월 19일', '3월19일 목요일', '2026년 3월 19일' 모두 지원.
    """
    today = datetime.today()
    today_m = today.month
    today_d = today.day

    used = ws.used_range
    max_row = used.last_cell.row
    max_col = 8

    # G열(7열)에서 날짜 헤더 행 순서대로 수집
    date_headers = []  # [(row, month, day), ...]
    for r in range(1, max_row + 1):
        val = ws.cells(r, 7).value
        if val and isinstance(val, str):
            m = _DATE_RE.search(val)
            if m:
                date_headers.append((r, int(m.group(1)), int(m.group(2))))

    # 오늘 날짜 헤더 찾기
    today_header_row = None
    today_idx = None
    for i, (r, mo, d) in enumerate(date_headers):
        if mo == today_m and d == today_d:
            today_header_row = r
            today_idx = i
            break

    if today_header_row is None:
        return max_row

    # 다음 날짜 헤더 행 (없으면 파일 끝 다음)
    next_header_row = date_headers[today_idx + 1][0] if today_idx + 1 < len(date_headers) else max_row + 1

    # 오늘 섹션의 마지막 데이터 행
    last_data_row = today_header_row + 1
    for r in range(today_header_row + 2, next_header_row):
        if any(ws.cells(r, c).value for c in range(1, max_col + 1)):
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

        # A열: Order # (배송/회수 prefix + 오더번호 + 추가 오더번호들)
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

        # D열: S/N (회수 아이템은 SAP에서 수집, 없으면 X, 배송은 공백)
        ws.cells(row_num, 4).value = item.get('serial_number', '')

        # E열: 담당자
        ws.cells(row_num, 5).value = item.get('customer', '')

        # F열: 전화번호
        ws.cells(row_num, 6).value = item.get('phone', '')

        # G열: 주소 (모든 행에 기입)
        address_parts = []
        if item.get('company'):
            address_parts.append(item['company'])
        street = item.get('street', '')
        street2 = item.get('street2', '')
        if street and street2:
            address_parts.append(f"{street},\n{street2}")
        elif street:
            address_parts.append(street)
        elif street2:
            address_parts.append(street2)
        cell_g = ws.cells(row_num, 7)
        cell_g.value = '\n'.join(address_parts)
        cell_g.api.WrapText = True

        # H열: 메모 (모든 행에 기입, 상단에 타임스탬프 추가)
        memo = item.get('memo', '')
        timestamp = datetime.now().strftime('%y.%m.%d %H:%M')
        memo_with_ts = f"{timestamp}\n{memo}" if memo else timestamp
        cell_h = ws.cells(row_num, 8)
        cell_h.value = memo_with_ts
        cell_h.api.WrapText = True

    # E/F/G/H 열 병합 (같은 오더의 여러 행)
    if num_rows > 1:
        end_row = insert_at + num_rows - 1
        wb.app.display_alerts = False
        for col in [5, 6, 7, 8]:  # E=담당자, F=전화, G=주소, H=메모
            ws.range(ws.cells(insert_at, col), ws.cells(end_row, col)).api.Merge()
        wb.app.display_alerts = True

    # 모든 테두리 적용 (A~H, insert_at ~ insert_at+num_rows-1)
    border_range = ws.range(
        ws.cells(insert_at, 1),
        ws.cells(insert_at + num_rows - 1, 8)
    )
    for border_idx in [7, 8, 9, 10, 11, 12]:  # Left/Top/Bottom/Right/InsideV/InsideH
        border_range.api.Borders(border_idx).LineStyle = 1   # xlContinuous
        border_range.api.Borders(border_idx).Weight = 2      # xlThin

    # 저장
    wb.save()
    end_row = insert_at + num_rows - 1
    print(f"[Excel] {num_rows}개 행 삽입 완료 → {ws.name} 시트 행 {insert_at}~{end_row}")
    return ws, insert_at, end_row


def write_kakao_sent(ws, start_row, end_row):
    """
    J열에 카톡 전송 시각 기록.
    여러 행이면 병합해서 한 칸에 표시.
    프린트 범위(A~H)에 포함되지 않아 인쇄에 영향 없음.
    """
    timestamp = datetime.now().strftime('카톡 ✓ %y.%m.%d %H:%M')
    wb = ws.book
    if end_row > start_row:
        wb.app.display_alerts = False
        ws.range(ws.cells(start_row, 10), ws.cells(end_row, 10)).api.Merge()
        wb.app.display_alerts = True
    cell = ws.cells(start_row, 10)
    cell.value = timestamp
    cell.api.WrapText = True
    wb.save()
    print(f"[Excel] J열 카톡 전송 기록: {timestamp}")


def _is_recent_sheet(sheet_name, days=90):
    """시트명(M-D 형식)이 최근 N일 내인지 확인."""
    try:
        m, d = map(int, sheet_name.split('-'))
        today = datetime.today()
        sheet_date = datetime(today.year, m, d)
        if sheet_date > today:
            sheet_date = datetime(today.year - 1, m, d)
        return (today - sheet_date).days <= days
    except Exception:
        return False


def get_existing_order_numbers(days=45):
    """
    Excel 파일의 A열(Order #)에서 이미 입력된 오더번호(7자리 이상 숫자) 추출.
    최근 N일 시트만 스캔하여 성능 최적화.
    """
    found = set()
    num_re = re.compile(r'\b(\d{7,})\b')
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        for sname in wb.sheetnames:
            if not _is_recent_sheet(sname, days=days):
                continue
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

"""
프린트 핸들러 - 날짜 묶음 단위로 Excel 인쇄.

사용법 (main.py 통해):
  python main.py --print-afternoon   → 오늘 날짜 두 번째 묶음 인쇄 (오후)
  python main.py --print-tomorrow    → 내일 날짜 묶음 인쇄
"""

import re
from datetime import datetime, timedelta

from excel_handler import get_workbook, get_today_sheet_name

_DATE_RE = re.compile(r'(\d{1,2})월\s*(\d{1,2})일')
PRINTER_NAME = "Samsung C56x Series"


# ── 시트 탐색 ────────────────────────────────────────────────────────────────

def _get_worksheet(wb=None):
    """오늘 시트 반환."""
    if wb is None:
        wb = get_workbook()
    sheet_name = get_today_sheet_name()
    names = [s.name for s in wb.sheets]
    if sheet_name in names:
        return wb, wb.sheets[sheet_name]
    today = datetime.today()
    for sname in names:
        if str(today.month) in sname and str(today.day) in sname:
            return wb, wb.sheets[sname]
    return wb, wb.sheets.active


def _find_sheet_by_key(wb, sheet_key):
    """
    '4-2' 형식 시트 키로 시트 반환.
    정확히 일치 → 월/일 포함 유연 탐색 순으로 시도.
    """
    names = [s.name for s in wb.sheets]
    if sheet_key in names:
        return wb.sheets[sheet_key]
    try:
        month, day = map(int, sheet_key.split('-'))
        for sname in names:
            if str(month) in sname and str(day) in sname:
                return wb.sheets[sname]
    except Exception:
        pass
    return None


def _find_date_headers(ws):
    """
    G열 전체 스캔 → 날짜 헤더 행 목록 반환.
    반환: [(row, month, day), ...]  (행 순서대로)
    """
    max_row = ws.used_range.last_cell.row
    headers = []
    for r in range(1, max_row + 1):
        val = ws.cells(r, 7).value  # G열
        if val and isinstance(val, str):
            m = _DATE_RE.search(val)
            if m:
                headers.append((r, int(m.group(1)), int(m.group(2))))
    return headers


def _find_section_range(ws, header_row, all_headers):
    """
    header_row 기준 묶음의 (시작행, 끝행) 반환.
    - 시작: header_row (날짜 헤더 행 포함)
    - 끝: 빈 행(A~H 모두 비어있음) 직전 또는 다음 날짜 헤더 직전
    빈 행은 섹션 구분자로 처리하여 Scheduled/Delayed 등 다른 묶음 포함 방지.
    """
    max_row = ws.used_range.last_cell.row

    # 다음 날짜 헤더 위치
    next_header_row = None
    for r, mo, d in all_headers:
        if r > header_row:
            next_header_row = r
            break

    limit = (next_header_row - 1) if next_header_row else max_row

    end_row = header_row
    for r in range(header_row + 1, limit + 1):
        if any(ws.cells(r, c).value for c in range(1, 9)):
            end_row = r
        else:
            break  # 빈 행 = 섹션 구분자 → 중단

    return header_row, end_row


# ── 프린터 설정 ───────────────────────────────────────────────────────────────

def _try_set_printer(app_api):
    """
    Samsung C56x 프린터로 변경 시도.
    Excel ActivePrinter는 'NAME on PORT:' 형식이라 포트를 순서대로 시도.
    반환: 이전 프린터 이름 (복원용), 실패 시 None
    """
    saved = app_api.ActivePrinter
    # 이미 Samsung이면 그대로
    if PRINTER_NAME.lower() in saved.lower():
        return saved

    ports = ['', ' on Ne00:', ' on Ne01:', ' on Ne02:', ' on Ne03:',
             ' on USB001:', ' on USB002:', ' on USB003:']
    for port in ports:
        try:
            app_api.ActivePrinter = PRINTER_NAME + port
            print(f"[프린트] 프린터 설정: {PRINTER_NAME + port}")
            return saved
        except Exception:
            continue

    print(f"[프린트] Samsung 프린터 설정 실패 → 현재 프린터 사용: {saved}")
    return None  # 변경 실패 (복원 불필요)


# ── 인쇄 실행 ────────────────────────────────────────────────────────────────

def _print_section(wb, ws, start_row, end_row):
    """
    A{start_row}:H{end_row} 범위를 가로 A4, 흑백, 2매 인쇄.
    """
    app = wb.app.api
    ps = ws.api.PageSetup

    # 인쇄 영역
    ps.PrintArea = f"$A${start_row}:$H${end_row}"

    # 페이지 설정
    try:
        ps.Orientation = 2      # xlLandscape (가로)
    except Exception:
        pass
    try:
        ps.PaperSize = 9        # xlPaperA4
    except Exception:
        pass
    try:
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = 1
    except Exception:
        pass

    # 흑백 시도 (PageSetup.BlackAndWhite: 셀 색상 무시하고 흑백 출력)
    bw_set = False
    try:
        ps.BlackAndWhite = True
        bw_set = True
        print("[프린트] 흑백 설정 완료")
    except Exception:
        print("[프린트] 흑백 설정 불가 → 컬러 출력")

    # 프린터 변경 시도
    prev_printer = _try_set_printer(app)

    try:
        ws.api.PrintOut(Copies=2, Collate=True)
        print(f"[프린트] 완료: {ws.name} 시트 {start_row}행~{end_row}행 (A~H), 2매")
    finally:
        # 복원
        if prev_printer:
            try:
                app.ActivePrinter = prev_printer
            except Exception:
                pass
        ps.PrintArea = ""
        if bw_set:
            try:
                ps.BlackAndWhite = False
            except Exception:
                pass


# ── 공개 함수 ────────────────────────────────────────────────────────────────

def print_section(sheet_key, date_key=None, afternoon=False):
    """
    지정 시트의 지정 날짜 섹션 인쇄.

    sheet_key : '4-2' 형식 — 어느 시트에서 찾을지
    date_key  : '4-2' 또는 '4-3' 형식 — 어느 날짜 섹션을 인쇄할지.
                None이면 sheet_key와 동일 날짜 사용.
    afternoon : False=첫 번째(오전), True=두 번째(오후) 섹션
    """
    if date_key is None:
        date_key = sheet_key

    wb = get_workbook()
    ws = _find_sheet_by_key(wb, sheet_key)
    if ws is None:
        print(f"[프린트] 시트 '{sheet_key}' 없음")
        return False

    try:
        month, day = map(int, date_key.split('-'))
    except Exception:
        print(f"[프린트] 날짜 형식 오류: '{date_key}' (올바른 예: 4-2)")
        return False

    headers = _find_date_headers(ws)
    matched = [(r, mo, d) for r, mo, d in headers if mo == month and d == day]

    if not matched:
        print(f"[프린트] {date_key} 섹션 없음 (시트: {ws.name})")
        return False

    if afternoon:
        if len(matched) < 2:
            print(f"[프린트] {date_key} 오후 섹션 없음 — 섹션이 {len(matched)}개뿐")
            return False
        target_row = matched[1][0]
        label = f"{date_key} 오후"
    else:
        target_row = matched[0][0]
        label = f"{date_key} 오전" if len(matched) > 1 else date_key

    start_row, end_row = _find_section_range(ws, target_row, headers)
    print(f"[프린트] {label} 섹션: {ws.name} 시트 {start_row}행~{end_row}행")
    _print_section(wb, ws, start_row, end_row)
    return True


def print_afternoon():
    """
    오늘 날짜 묶음이 2개 이상이면 두 번째(오후) 묶음 인쇄.
    묶음 구분: G열에 같은 날짜가 두 번 등장.
    """
    wb, ws = _get_worksheet()
    today = datetime.today()
    headers = _find_date_headers(ws)

    today_sections = [
        (r, mo, d) for r, mo, d in headers
        if mo == today.month and d == today.day
    ]

    if len(today_sections) < 2:
        print(f"[프린트] 오후 묶음 없음 — 오늘({today.month}/{today.day}) 섹션이 {len(today_sections)}개")
        return False

    target_row = today_sections[1][0]  # 두 번째 섹션
    start_row, end_row = _find_section_range(ws, target_row, headers)
    print(f"[프린트] 오후 묶음: {start_row}행~{end_row}행")
    _print_section(wb, ws, start_row, end_row)
    return True


def print_tomorrow():
    """
    내일 날짜 묶음 인쇄.
    1) 오늘 시트에서 내일 날짜 섹션 탐색
    2) 없으면 내일 시트(예: '3-20')의 첫 번째 섹션 인쇄
    오늘 중 추가될 수 있으므로 실행 시점에 end_row를 동적으로 탐색.
    """
    wb, ws = _get_worksheet()
    tomorrow = datetime.today() + timedelta(days=1)

    # ── 1) 오늘 시트에서 내일 섹션 탐색 ────────────────────────
    headers = _find_date_headers(ws)
    tomorrow_sections = [
        (r, mo, d) for r, mo, d in headers
        if mo == tomorrow.month and d == tomorrow.day
    ]

    if tomorrow_sections:
        target_row = tomorrow_sections[0][0]
        start_row, end_row = _find_section_range(ws, target_row, headers)
        print(f"[프린트] 내일 묶음 (오늘 시트): {start_row}행~{end_row}행")
        _print_section(wb, ws, start_row, end_row)
        return True

    # ── 2) 내일 시트 탐색 ────────────────────────────────────────
    tomorrow_sheet_name = f"{tomorrow.month}-{tomorrow.day}"
    sheet_names = [s.name for s in wb.sheets]

    target_ws = None
    if tomorrow_sheet_name in sheet_names:
        target_ws = wb.sheets[tomorrow_sheet_name]
    else:
        # 월/일이 포함된 시트명 유연 탐색 (예: '3월20일', '20-3' 등 대비)
        for sname in sheet_names:
            if str(tomorrow.month) in sname and str(tomorrow.day) in sname:
                target_ws = wb.sheets[sname]
                break

    if target_ws is None:
        print(f"[프린트] 내일({tomorrow.month}/{tomorrow.day}) 묶음 없음 — 오늘 시트에도 내일 시트에도 없음")
        return False

    tmr_headers = _find_date_headers(target_ws)
    if not tmr_headers:
        print(f"[프린트] 내일 시트({target_ws.name})에 날짜 헤더 없음")
        return False

    # 내일 시트의 첫 번째 섹션
    target_row = tmr_headers[0][0]
    start_row, end_row = _find_section_range(target_ws, target_row, tmr_headers)
    print(f"[프린트] 내일 묶음 (내일 시트 '{target_ws.name}'): {start_row}행~{end_row}행")
    _print_section(wb, target_ws, start_row, end_row)
    return True

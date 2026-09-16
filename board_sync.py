"""
workbench.db의 현재 상태를 REALTIME_SYNC_PATH("2026 배송장_실시간 동기화.xlsx")에
맞추기 위한 임시 브릿지 스크립트 (2026-08-28, 사용자 요청).

배경: workbench 보드(workbench.db)가 진짜 메인이 되길 원하지만, workbench가
가끔 실시간 에러를 내다 보니 그동안 써온 배송장 엑셀을 실시간으로 손수 맞춰와서
- 1과 2를 동시에 손으로 맞추다 실수가 남. workbench.db가 이미 갖고 있는
전체 그림을 배송장 엑셀과 같은 구조(날짜별 배너 섹션이 쌓이는 한 시트)의 별도
파일 안에 자동으로 채워 넣어서 손으로 맞추는 일을 없애는 게 목적. workbench가
충분히 안정화돼서 이 파일 자체가 더 이상 필요 없어지면 스크립트째로 치우면 됨 -
사용자가 명시적으로 "임시방편"이라 부른 것.

2026-09-01 대상 파일 변경: 원래는 실제 공유 배송장 파일(EXCEL_PATH,
"C:\\1\\배송장\\2026 배송장.xlsx")의 오늘 시트에 직접 썼는데, 그 파일은 다른
사람들과 공유하는 파일이라 에러 없이 항상 깔끔해야 하고 workbench.db 기반의
"아직 확정 전" 상태가 섞이면 안 된다는 사용자 지적으로, 완전히 분리된
REALTIME_SYNC_PATH(config.py)로 쓰기 대상을 옮겼다. EXCEL_PATH는 이제
write_orders_to_excel() 등 매일 SAP 수집 경로만 그대로 쓴다 - 이 스크립트는
더 이상 EXCEL_PATH를 열거나 손대지 않는다. 전환 시점 이전에 이미 EXCEL_PATH에
반영됐던 내용(예: 9/1 섹션)은 사용자 결정으로 그대로 두고 되돌리지 않았음 -
REALTIME_SYNC_PATH는 전환 시점부터 빈 상태로 새로 시작한다.

배송장 엑셀의 실제 구조(2026-08-28 실측): 시트 이름은 "M-D"(예: "8-28")이고,
그 시트 하나 안에 오늘 날짜 배너부터 앞으로 예정된 여러 미래 날짜 배너까지
쭉 이어져 쌓여 있다(각 배너 아래 Order#/item/M-N/S-N/customer/phone/
ADDRESS/memo 헤더 + 데이터 행). 아침마다 어제 시트를 복사해서 오늘 이름으로
바꾸고 이미 지난 날짜 섹션만 지우는 방식으로 유지돼 왔음(사용자 확인).
REALTIME_SYNC_PATH가 아직 존재하지 않을 때(전환 첫 실행)는 이 구조의 씨앗이 될
첫 배너+헤더를 _seed_realtime_sync_file()이 openpyxl로 직접 그려 넣는다 -
find_or_create_section()은 복제할 기존 배너가 최소 하나는 있어야 동작하므로.

두 진입점:
  --new-day : 아침 루틴. 오늘 시트가 이미 있으면 아무 것도 안 함. 없으면
              (오늘 이전 날짜 중 가장 최근인) 시트를 복사해 오늘 이름으로
              만들고, 지난 날짜 섹션을 지운 뒤 sync()까지 이어서 실행해
              그 시점 workbench 전체 내용을 한 번에 채워 넣는다.
  --sync    : workbench.db에 있는데 아직 해당 시트에 없는 오더를, 그 오더의
              source_date에 맞는 날짜 배너 섹션 안에 끼워넣는다(그런 배너가
              없으면 새로 만듦). 오더 번호/식별자로 이미 있는지 확인하므로
              여러 번 실행해도 안전(추가만 하지 지우거나 덮어쓰지 않음) -
              workbench의 "기존 배송장 엑셀로 동기화" 버튼이 이 모드(sync())를
              그 자리에서 바로 호출함(서브프로세스 아님).

실제 Excel COM 쓰기는 excel_handler.py의 write_orders_to_excel()이 쓰는 것과
같은 xlwings 경로(_com_retry)를 그대로 재사용하되, 워크북을 여는 부분만
REALTIME_SYNC_PATH 전용으로 이 파일 안에 따로 둔다(_get_sync_workbook 참고) -
excel_handler.get_workbook()은 EXCEL_PATH 전용이라 그대로 못 씀. 파일이 열려
있어도 실시간 반영되고, 동시 쓰기 문제는 기존과 동일하게 ExcelWriteLock으로
막는다. write_orders_to_excel() 자체는 건드리지 않는다 - 그 함수는 지금도
main.py/manual_order_handler.py가 매일 의존하는 실제 운영 경로라, 이 신규 기능
때문에 손대는 위험을 지지 않기 위한 선택. 대신 그 함수 본문에서 행 삽입/서식
부분만 필요한 만큼 아래 _write_order_rows()에 별도로 옮겨 적었다(임의의
insert_at을 받을 수 있어야 해서 - write_orders_to_excel은 항상 "오늘 섹션"만
대상으로 하므로 그대로는 재사용 불가).

주의(라이브 검증 필요, 다른 Excel 자동화 변경들과 동일한 원칙):
이 파일은 실제 배송 문서를 건드리는 스크립트라 처음 실제로 켤 때는 반드시
사용자가 지켜보는 자리에서 한 번 확인해야 한다. 특히 새 날짜 배너를 만드는
경로(find_or_create_section의 배너 없음 분기)는 자주 타지 않는 드문 경로라
더 그렇다.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sqlite3
from datetime import date
from pathlib import Path

import openpyxl
import xlwings as xw
from openpyxl.styles import Alignment, Font, PatternFill

from config import REALTIME_SYNC_PATH
from excel_handler import _com_retry
from order_lock import ExcelWriteLock

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "workbench.db"

# --new-day는 Task Scheduler로 무인 실행되므로(콘솔 없는 pythonw) print()만으로는
# 실패해도 아무 기록이 안 남는다 - 다른 무인 스크립트들과 동일하게 파일 로그를 남긴다.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.FileHandler(BASE_DIR / "board_sync.log", encoding="utf-8")],
)
logger = logging.getLogger(__name__)

_DATE_RE = re.compile(r'(\d{1,2})월\s*(\d{1,2})일')
_WEEKDAYS = ['월', '화', '수', '목', '금', '토', '일']


# ── 공용 유틸 ──────────────────────────────────────────────────────────────

def _sheet_name_for(d):
    return f"{d.month}-{d.day}"


def _banner_text(d):
    return f"{d.year}년 {d.month}월 {d.day}일 {_WEEKDAYS[d.weekday()]}요일"


# ── REALTIME_SYNC_PATH 워크북 열기 (2026-09-01, EXCEL_PATH와 분리) ────────────

# 마감 내보내기(workbench_app.py의 _write_export_sheet, 새 시트 branch)와 같은
# 값 - 처음 만드는 파일의 첫 배너+헤더 서식을 그것과 통일하기 위해 그대로
# 옮겨왔다. openpyxl로 딱 한 번만 그리고 나면 그 뒤로는 전부 COM(xlwings)이
# 다루므로 이 상수들은 오직 _seed_realtime_sync_file()에서만 쓰인다.
_SEED_COL_WIDTHS = {1: 23.375, 2: 40.0, 3: 13.75, 4: 14.5, 5: 16.625, 6: 19.875, 7: 47.75, 8: 25.125}
_SEED_HEADER_FILL = PatternFill(fill_type="solid", fgColor="FBE3D6")
_SEED_HEADER = ["Order #", "item", "M/N", "S/N", "customer", "phone", "ADDRESS", "memo"]


def _seed_realtime_sync_file(path, first_date):
    """REALTIME_SYNC_PATH 파일이 아직 없을 때(전환 첫 실행) 딱 한 번만 호출된다.
    find_or_create_section()/_locate_section() 등 이 파일의 나머지 로직은
    전부 "시트 안에 배너가 최소 하나는 있다"고 가정하므로, 그 전제를 만족시킬
    첫 배너(first_date)+헤더만 openpyxl로 직접 그려 넣는다 - 이후 실제 데이터
    행 삽입/서식은 항상 COM(xlwings) 경로(_write_order_rows 등)가 담당한다."""
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _sheet_name_for(first_date)
    for col, width in _SEED_COL_WIDTHS.items():
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    banner_cell = ws.cell(row=1, column=7)
    banner_cell.value = _banner_text(first_date)
    banner_cell.font = Font(bold=True)
    banner_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, label in enumerate(_SEED_HEADER, start=1):
        cell = ws.cell(row=2, column=col)
        cell.value = label
        cell.fill = _SEED_HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    wb.save(path)
    wb.close()


def _get_sync_workbook(seed_date=None):
    """REALTIME_SYNC_PATH에 연결. 파일 자체가 없으면(전환 첫 실행) seed_date
    (보통 오늘) 기준으로 첫 배너를 먼저 만들어둔 뒤 연다.

    2026-09-02 사용자 지적: 이 파일은 사람이 볼 필요 없는 내부용 파일인데,
    (예전에 썼던) `xw.Book(path)`는 새로 열 때 기본적으로 화면에 보이는
    Excel 창을 띄운다 - main.py 20분 자동루프가 매번 이걸 열면서 아침마다
    화면에 떠 있는 걸 사용자가 발견함(진짜 배송장.xlsx까지 같이 열려있던
    건 별개 원인 - startup.py의 open_excel() 자동호출 제거로 따로 처리).
    그래서 직접 "이미 열린 인스턴스 찾기(전체 경로 기준) → 없으면 안 보이는
    새 Excel 인스턴스로 열기" 순서로 바꿨다. 한 번 안 보이게 열리면 그
    App이 살아있는 동안(main.py/workbench_app.py 프로세스 생애주기) COM
    ROT에 등록된 채로 남아 있어, 다음 호출부터는 아래 탐색 루프가 그 인스턴스를
    찾아 재사용한다(매번 새로 열지 않음)."""
    path = Path(REALTIME_SYNC_PATH)
    if not path.exists():
        _seed_realtime_sync_file(path, seed_date or date.today())

    target = os.path.normcase(os.path.abspath(str(path)))
    try:
        for app in xw.apps:
            for book in app.books:
                if os.path.normcase(os.path.abspath(book.fullname)) == target:
                    return book
    except Exception:
        pass

    app = xw.App(visible=False)
    return app.books.open(str(path))


def _find_banners(ws):
    """G열(7열)에서 날짜 배너 행을 문서 순서대로 전부 찾아
    [(row, month, day), ...] 로 반환. used_range 기준이라 시트가 비어있으면
    빈 리스트."""
    used = _com_retry(lambda: ws.used_range)
    max_row = used.last_cell.row
    banners = []
    for r in range(1, max_row + 1):
        val = ws.cells(r, 7).value
        if val and isinstance(val, str):
            m = _DATE_RE.search(val)
            if m:
                banners.append((r, int(m.group(1)), int(m.group(2))))
    return banners, max_row


def _section_end(ws, banner_row, next_banner_row, max_row):
    """banner_row(날짜 배너) 섹션의 마지막 실제 데이터 행. 데이터가 하나도
    없으면 헤더 행(banner_row+1) 자체를 반환(그 다음 행부터 삽입). 연속 빈
    행 3개 이상이면 그 섹션이 끝난 것으로 본다(excel_handler.py의
    find_today_section_end와 동일 규칙)."""
    boundary = next_banner_row if next_banner_row else (max_row + 1)
    last_data_row = banner_row + 1
    consecutive_empty = 0
    for r in range(banner_row + 2, boundary):
        if any(ws.cells(r, c).value for c in range(1, 9)):
            last_data_row = r
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty >= 3:
                break
    return last_data_row


def _locate_section(ws, target_date):
    """target_date에 해당하는 배너를 찾아 (banner_row, section_end_row)를
    반환. 없으면 None."""
    banners, max_row = _find_banners(ws)
    for i, (r, mo, d) in enumerate(banners):
        if mo == target_date.month and d == target_date.day:
            next_row = banners[i + 1][0] if i + 1 < len(banners) else None
            return r, _section_end(ws, r, next_row, max_row)
    return None


def _section_column_a_text(ws, banner_row, end_row):
    """그 섹션(배너+헤더 다음 행부터 end_row까지)의 A열 텍스트만 모은다 -
    전체 시트가 아니라 이 섹션 안에서만 "이미 있는지" 판단하기 위함(아래
    _already_in_sheet 사용처 설명 참고)."""
    parts = []
    for r in range(banner_row + 2, end_row + 1):
        v = ws.cells(r, 1).value
        if v:
            parts.append(str(v))
    return "\n".join(parts)


def _copy_cell_format(src_cell, dst_cell):
    """src_cell(xlwings Range, 단일 셀)의 서식만 dst_cell로 옮긴다 - 값은
    건드리지 않는다. Copy()/Paste() 클립보드 경로를 완전히 피하기 위한 대체.

    배경: 클립보드 기반 Paste는 대상 크기가 원본과 정확히 안 맞으면 Excel이
    "선택 영역을 채우도록 반복 복제"로 해석해버리는 위험이 있다(2026-08-28
    최초 사고 - A1:H2를 전체 행 폭에 붙여넣어 헤더가 XFD열까지 복제됨). 그
    직후 "대상 크기를 원본과 정확히 맞추면 안전하다"는 수정을 넣었는데,
    2026-08-31 재검증 중 그 수정만으로는 부족하다는 걸 확인했다 - 완료 안 된
    과거 오더를 한 번에 여러 날짜에 걸쳐 이월하는(=이 함수를 한 세션 안에서
    연달아 여러 번 호출하는) 정상적인 시나리오에서, 이전 호출의 클립보드
    상태가 다음 호출과 꼬이면서 같은 반복 복제 현상이 다시 재현됐다(라이브가
    아닌 사본에서 발견). 그래서 클립보드를 아예 쓰지 않고 셀 속성을 하나씩
    직접 읽어 그대로 설정하는 방식으로 바꿨다 - 느리지만(새 배너를 만드는
    드문 경로라 상관없음) 이 버그 클래스 자체가 원천적으로 발생할 수 없다."""
    src_api, dst_api = src_cell.api, dst_cell.api
    dst_api.Interior.Color = src_api.Interior.Color
    dst_api.Interior.Pattern = src_api.Interior.Pattern
    dst_api.Font.Bold = src_api.Font.Bold
    dst_api.Font.Size = src_api.Font.Size
    dst_api.Font.Name = src_api.Font.Name
    dst_api.Font.Color = src_api.Font.Color
    dst_api.HorizontalAlignment = src_api.HorizontalAlignment
    dst_api.VerticalAlignment = src_api.VerticalAlignment
    dst_api.WrapText = src_api.WrapText
    dst_api.NumberFormat = src_api.NumberFormat
    for idx in [7, 8, 9, 10, 11, 12]:  # xlEdgeLeft/Top/Bottom/Right, xlInsideVertical/Horizontal
        sb = src_api.Borders(idx)
        db = dst_api.Borders(idx)
        db.LineStyle = sb.LineStyle
        db.Weight = sb.Weight


def _clear_beyond_h(ws, row):
    """이 워크북의 배송 표는 A~H열(8열)만 쓰는데, 행을 삽입(EntireRow.Insert)할
    때 Excel이 인접 행 서식을 일부 물려받으면서 I열 이후(화면 훨씬 오른쪽 밖,
    최대 XFD열)까지 배경색 등이 번지는 경우가 있다(2026-08-31 실사용에서
    발견 - 헤더 배경색이 새로 삽입된 데이터 행에 번져 "주황색 줄"처럼
    보였음). 이 함수가 만들거나 손댄 행은 항상 I열 이후를 깨끗이 비워
    이 문제를 원천 차단한다."""
    _com_retry(lambda: ws.api.Range(ws.api.Cells(row, 9), ws.api.Cells(row, 16384)).Clear())


def _refit_row_heights(ws):
    """시트 전체(배너부터 마지막 행까지) 행 높이를 실제 줄바꿈 내용에 맞게
    다시 계산한다 (2026-09-01 사용자 발견 - 'backup_2\\...backup.xlsx'와
    줄간격을 맞춰달라는 요청).

    원인: Excel의 EntireRow.AutoFit()은 그 행에 세로 병합된 셀(E~H, 오더
    그룹의 customer/phone/address/memo)이 하나라도 걸려 있으면 실제 내용
    기준으로 다시 계산을 못 하는 잘 알려진 한계가 있다. 2026-08-31 이전
    버전의 find_or_create_section()은 새 배너를 만들 때 배너+헤더 행에
    (데이터 행이 아직 없는 상태의) 높이를 고정으로 넣었는데, 그 뒤로 그
    섹션 위에 실제 데이터가 병합된 채로 쌓이면서 AutoFit이 더는 못 고치는
    상태로 굳어버렸다. 그 섹션들이 매일 아침 "어제 시트를 복사해 오늘로"
    방식으로 그대로 전파돼, 라이브 파일 안에 배너/헤더/데이터가 전부 같은
    눌린 높이로 남은 옛 섹션이 여럿 쌓여 있었다(실측: 9/4, 9/7, 9/9 섹션).

    해결: 병합된 블록은 잠깐 풀고(UnMerge) 그 행 범위만 AutoFit한 뒤 같은
    범위로 다시 합친다 - 매 sync()마다 시트 전체에 대해 실행하므로, 오늘
    새로 쓴 행뿐 아니라 예전부터 남아있던 섹션도 다음 동기화 때 저절로
    교정되고, 그 뒤로 복사되는 날에는 더 이상 이 문제가 전파되지 않는다."""
    used = _com_retry(lambda: ws.used_range)
    max_row = used.last_cell.row
    r = 1
    while r <= max_row:
        cell_e = ws.api.Cells(r, 5)
        if cell_e.MergeCells:
            area = cell_e.MergeArea
            start, n = area.Row, area.Rows.Count
            end = start + n - 1
            if start == r and n > 1:
                row_range = f"{start}:{end}"
                for col in (5, 6, 7, 8):
                    _com_retry(lambda col=col: ws.api.Cells(start, col).MergeArea.UnMerge())
                _com_retry(lambda: ws.api.Rows(row_range).AutoFit())
                ws.book.app.display_alerts = False
                try:
                    for col in (5, 6, 7, 8):
                        _com_retry(lambda col=col: ws.api.Range(ws.api.Cells(start, col), ws.api.Cells(end, col)).Merge())
                finally:
                    ws.book.app.display_alerts = True
                r = end + 1
                continue
        _com_retry(lambda r=r: ws.api.Rows(r).AutoFit())
        r += 1


def find_or_create_section(ws, target_date):
    """target_date에 해당하는 배너 섹션에 새 데이터를 넣을 삽입 행(insert_at)을
    반환. 그런 배너가 없으면 날짜순으로 맞는 자리에 새로 만든다(기존 배너
    중 첫 번째 것의 배너+헤더 두 행의 서식/텍스트를 셀 단위로 그대로
    복제 - 클립보드는 쓰지 않음, _copy_cell_format 참고)."""
    located = _locate_section(ws, target_date)
    if located:
        _, end_row = located
        return end_row + 1

    banners, max_row = _find_banners(ws)
    if not banners:
        raise RuntimeError(
            f"'{ws.name}' 시트에 날짜 배너가 하나도 없어 새 섹션 서식을 복제할 수 없음"
        )

    # 날짜순으로 삽입할 위치 찾기 (target_date보다 나중인 첫 배너 앞에 삽입).
    # 같은 해 기준으로만 비교 - 몇 달 앞 미래 오더만 다루는 현재 운영 범위에서는
    # 충분(연말/연초 경계를 넘나드는 케이스는 이 스크립트 책임 밖으로 남겨둠).
    insert_before = None
    for r, mo, d in banners:
        try:
            bd = date(target_date.year, mo, d)
        except ValueError:
            continue
        if bd > target_date:
            insert_before = r
            break
    banner_row = insert_before if insert_before else (max_row + 1)
    template_row = banners[0][0]

    # 빈 행 삽입은 클립보드 없이 EntireRow.Insert()만으로(다른 모든 행 삽입과
    # 동일한, 안전이 검증된 방식).
    for _ in range(2):
        _com_retry(lambda: ws.range(f"A{banner_row}:H{banner_row}").api.EntireRow.Insert())

    # 2026-08-31 발견 버그: template_row는 삽입 "전"에 계산해뒀는데, 과거
    # 오더를 이월할 때는 항상 target_date가 현재 가장 이른 배너보다 앞이라
    # banner_row == template_row가 된다. 이 상태로 바로 위 EntireRow.Insert()를
    # 하면 template_row 위치의 실제 내용(배너+헤더)이 2행 아래로 밀려나는데,
    # template_row 변수는 갱신 안 된 옛 값 그대로라 "방금 삽입한 빈 행 자신"을
    # 템플릿으로 착각해서 그대로 복제해버린다 - 그 결과 헤더 라벨이 통째로
    # 빈 채 만들어지고(실사용자가 스크린샷으로 발견), 높이도 삽입된 빈 행의
    # 기본값에 고정돼 그 아래 데이터 행들까지 줄바꿈에 안 맞게 눌린다.
    # insert_before가 template_row보다 뒤(또는 없어 맨 끝에 추가)인 경우는
    # template_row가 삽입 지점보다 위라 안 밀리므로 이 보정이 필요 없다.
    if template_row >= banner_row:
        template_row += 2

    def _apply_template():
        for col in range(1, 9):
            src_banner = ws.cells(template_row, col)
            dst_banner = ws.cells(banner_row, col)
            _copy_cell_format(src_banner, dst_banner)

            src_header = ws.cells(template_row + 1, col)
            dst_header = ws.cells(banner_row + 1, col)
            _copy_cell_format(src_header, dst_header)
            # 헤더 라벨(Order #/item/M/N/...)은 서식과 함께 값도 원본 그대로.
            dst_header.value = src_header.value

        # 높이는 템플릿 값을 그대로 못박기보다 AutoFit으로 - 고정 높이를
        # 넣으면 이후 그 자리에 삽입되는 실제 데이터 행들이 그 고정 높이를
        # 물려받아 줄바꿈 내용에 안 맞게 눌리는 문제가 있었다(2026-08-31).
        ws.api.Rows(f"{banner_row}:{banner_row + 1}").AutoFit()

        # I열 이후는 이 서식에서 항상 비어있어야 한다 - 인접 행 서식을 일부
        # 물려받으며 화면 밖(I열~)까지 색이 번지는 걸 방지(2026-08-31 실사용
        # 중 실제로 발견: 헤더색이 I열 이후로 번져 "주황색 줄"처럼 보임).
        _clear_beyond_h(ws, banner_row)
        _clear_beyond_h(ws, banner_row + 1)

    _com_retry(_apply_template)

    ws.cells(banner_row, 7).value = _banner_text(target_date)
    return banner_row + 2


def _has_banner_for(ws, d):
    banners, _ = _find_banners(ws)
    return any(mo == d.month and dd == d.day for _, mo, dd in banners)


def _strip_past_sections(ws, today):
    """맨 앞부터 today보다 이전 날짜인 배너 섹션을 통째로(배너+헤더+데이터)
    지운다. 여러 개가 연달아 지난 경우(예: 주말을 건너뛰어 시트를 며칠 만에
    새로 만드는 경우)까지 반복 처리."""
    while True:
        banners, max_row = _find_banners(ws)
        if not banners:
            return
        r, mo, d = banners[0]
        try:
            bd = date(today.year, mo, d)
        except ValueError:
            return
        if bd >= today:
            return
        stale_end = banners[1][0] - 1 if len(banners) > 1 else max_row
        _com_retry(lambda r=r, stale_end=stale_end: ws.range(f"{r}:{stale_end}").api.EntireRow.Delete())


def _order_is_open(item_rows):
    """이 오더의 항목 중 아직 완료 처리(board_status='done'/'cancelled')되지
    않은 게 하나라도 있으면 True. orders 테이블 자체엔 완료 여부가 없고,
    workbench 보드가 쓰는 완료 마커는 order_items.board_status(항목별)뿐이라
    거기서 판단한다. 2026-08-31 발견: 이게 없어서 sync()가 "과거 날짜=완료된
    날짜"로 잘못 취급해, 아직 처리 중인 과거 오더까지 새 시트에서 통째로
    누락시키는 버그가 있었다(실사용자가 손으로 이월해온 것과 다른 동작)."""
    return any((it['board_status'] or '') not in ('done', 'cancelled') for it in item_rows)


def _already_in_sheet(section_text, order_row):
    """오더 번호가 있으면 그걸로, 없으면(수동 메모 행) order_type 문자열
    전체로 - 이미 그 오더의 날짜 섹션 안에 반영됐는지 확인. 반드시 그 오더의
    "날짜 섹션 안"의 텍스트만 대상으로 해야 한다(전체 시트가 아니라) -
    "SDSK 1333208374"처럼 같은 티켓 번호를 여러 날짜(9/15 Event, 9/28
    Event처럼)에 걸쳐 별도 오더로 반복 사용하는 경우가 실제로 있어서, 전체
    시트를 대상으로 하면 한쪽 날짜에 이미 써진 걸 보고 다른 날짜 것까지
    "이미 있음"으로 오판해 영원히 못 들어가는 버그가 있었음(2026-08-28
    사용자가 "동기화 안 된다"고 보고해서 실측 발견 - order_type이 "memo"/
    "SDSK"처럼 아주 짧은 경우도 전체 시트 기준이면 우연히 다른 데 있는
    텍스트와 매치될 위험이 있었음). 둘 다 없는 행은 식별할 단서가 없어
    매번 새로 추가될 수 있음(드문 경우, 알려진 한계)."""
    order_no = (order_row['order_no'] or '').strip()
    if order_no:
        return re.search(rf'\b{re.escape(order_no)}\b', section_text) is not None
    order_type = (order_row['order_type'] or '').strip()
    if order_type:
        return order_type in section_text
    return False


# ── 워크벤치 보드의 셀별 수동 강조(cell_highlights) 반영 ──────────────────────
# workbench_app.py의 set_cell_highlight()/HIGHLIGHT_FIELDS와 동일한 필드
# 이름/색상 - 우클릭으로 지정한 배경(노랑/핑크)/글자색(빨강)/Bold를 그대로
# 배송장 엑셀에도 반영한다(2026-08-31 사용자 요청 - 이전엔 전부 무시하고
# 흰 배경/기본 글자색으로만 썼음). 색상 값은 이미 있는 "마감" 내보내기
# (_EXPORT_HIGHLIGHT_FILL/_EXPORT_HIGHLIGHT_FONT_COLOR)와 통일했다 - 워크벤치
# 화면 자체의 CSS 색(#fde047 등)과는 약간 다르지만, 이 코드베이스 안에서 이미
# "강조 → 엑셀"에 쓰기로 정한 색이라 그걸 따르는 게 일관적이다. 참고:
# 브라우저 "배송장 인쇄"(Ctrl+P 스타일)는 별개 기능이고 거기는 의도적으로
# 흑백 유지 - 이 강조는 그것과 무관하게 EXCEL_PATH 파일에만 적용된다.
_HIGHLIGHT_FIELDS_BY_COL = {
    1: ('order', 'type'), 2: ('description',), 3: ('material',), 4: ('serial',),
    5: ('customer',), 6: ('phone',), 7: ('address',), 8: ('memo',),
}
_HIGHLIGHT_BG_RGB = {'yellow': (255, 255, 0), 'pink': (255, 192, 203)}
_HIGHLIGHT_FONT_RGB = {'red': (255, 0, 0)}
_DEFAULT_CELL_RGB = (255, 255, 255)

# order_items.board_status → 배경/글자색 (2026-09-01 사용자 요청 - workbench
# 화면 자체의 상태색과 통일). workbench_app.py 프론트(row-done/row-pending/
# row-cancelled CSS, 약 4168~4173행)와 정확히 같은 값: 완료(done)=파랑,
# 회수 대기(pending, "Ready" 버튼)=빨강, 취소(cancelled, "Cancel" 버튼)=회색+
# 취소선, 그 외(초기화/미지정)=흰색(_DEFAULT_CELL_RGB 그대로). 수동 셀 강조
# (cell_highlights)가 지정돼 있으면 그게 항상 우선(기존 동작 그대로) -
# 상태색은 강조가 없을 때의 기본값으로만 쓰인다.
_STATUS_BG_RGB = {
    'done': (191, 219, 254),      # #bfdbfe
    'pending': (254, 202, 202),   # #fecaca
    'cancelled': (209, 213, 219),  # #d1d5db
}
_STATUS_FONT_RGB = {'cancelled': (107, 114, 128)}  # #6b7280


def _fetch_cell_highlights(item_ids):
    """item_ids에 해당하는 cell_highlights 행을 {(item_id, field): style_dict}
    로 반환. style_dict는 workbench_app.py와 동일하게 {"bg":.., "color":..,
    "bold":..} 중 실제로 지정된 키만 담김."""
    item_ids = [i for i in item_ids if i]
    if not item_ids:
        return {}
    conn = _connect_db()
    try:
        placeholders = ','.join('?' * len(item_ids))
        cur = conn.cursor()
        cur.execute(
            f"SELECT item_id, field, style_json FROM cell_highlights WHERE item_id IN ({placeholders})",
            item_ids,
        )
        return {
            (r['item_id'], r['field']): (json.loads(r['style_json']) if r['style_json'] else {})
            for r in cur.fetchall()
        }
    finally:
        conn.close()


def _apply_cell_style(cell, style, default_bg=_DEFAULT_CELL_RGB, default_font=None, strike=False):
    """cell(xlwings Range 단일 셀)에 강조 스타일을 적용. style이 None/빈
    dict면 default_bg(호출부가 넘긴 board_status 색, 안 넘기면 흰 배경)로 -
    매번 명시적으로 지정해야 이전에 강조됐던 셀이 재사용될 때 색이 안
    지워지고 남는 걸 방지한다. 수동 강조(style)가 있으면 항상 그게 우선."""
    bg = (style or {}).get('bg')
    cell.color = _HIGHLIGHT_BG_RGB.get(bg, default_bg)
    if style and style.get('color') in _HIGHLIGHT_FONT_RGB:
        cell.font.color = _HIGHLIGHT_FONT_RGB[style['color']]
    elif default_font:
        cell.font.color = default_font
    if style and style.get('bold'):
        cell.font.bold = True
    cell.font.strikethrough = strike


# ── write_orders_to_excel()에서 필요한 부분만 떼어온 행 쓰기 ──────────────────

def _write_order_rows(ws, insert_at, order_data_list):
    """order_data_list(한 오더의 아이템들, excel_handler.write_orders_to_excel과
    동일한 dict 모양)를 insert_at부터 삽입해서 쓴다. excel_handler.py의
    _write_orders_to_excel_locked() 본문과 동일한 서식(줄바꿈/병합/테두리) -
    그 함수는 항상 "오늘 섹션"만 대상으로 해서 그대로 재사용할 수 없어 이
    파일 전용으로 옮겨 적었다."""
    num_rows = len(order_data_list)

    for _ in range(num_rows):
        _com_retry(lambda: ws.range(f"A{insert_at}:H{insert_at}").api.EntireRow.Insert())

    # I열 이후는 항상 비어있어야 하는데, 행 삽입 시 인접 행 서식을 일부
    # 물려받으며 화면 밖(I열~XFD열)까지 색이 번지는 경우가 있었다(2026-08-31
    # 실사용에서 발견 - 헤더 배경색이 새로 삽입된 데이터 행에 번져 눈에 띄는
    # "주황색 줄"처럼 보였음). 값을 쓰기 전에 미리 비워 원천 차단한다.
    for i in range(num_rows):
        _clear_beyond_h(ws, insert_at + i)

    # 워크벤치 보드의 셀별 수동 강조(cell_highlights)를 그대로 반영 -
    # customer/phone/address/memo는 오더 전체에 걸쳐 병합되는 컬럼이라
    # workbench 자체 규칙대로 "그룹의 첫 항목" 기준으로만 조회한다.
    item_ids = [item.get('item_id') for item in order_data_list]
    highlights = _fetch_cell_highlights(item_ids)
    anchor_item_id = order_data_list[0].get('item_id') if order_data_list else None
    # E~H(customer/phone/address/memo)는 그룹 전체가 한 셀로 병합되므로 색도
    # anchor(그룹 첫 항목)의 board_status 기준 하나로 통일한다.
    anchor_status = (order_data_list[0].get('board_status') or '') if order_data_list else ''
    anchor_bg = _STATUS_BG_RGB.get(anchor_status, _DEFAULT_CELL_RGB)
    anchor_font = _STATUS_FONT_RGB.get(anchor_status)
    anchor_strike = anchor_status == 'cancelled'

    def _style_for(item_id, fields):
        for f in fields:
            st = highlights.get((item_id, f))
            if st:
                return st
        return None

    for idx, item in enumerate(order_data_list):
        row_num = insert_at + idx
        item_id = item.get('item_id')
        # A~D(order#/item/M-N/S-N)는 항목별로 board_status가 다를 수 있어
        # (예: 같은 오더 안에서 배송은 완료, 회수는 대기) 그 항목 자신의
        # 상태를 쓴다 - workbench 화면도 행(항목) 단위로 색칠한다.
        status = item.get('board_status') or ''
        bg = _STATUS_BG_RGB.get(status, _DEFAULT_CELL_RGB)
        font_rgb = _STATUS_FONT_RGB.get(status)
        strike = status == 'cancelled'

        order_line = f"{item['order_prefix']} {item['order_type']} {item['order_num']}"
        if item.get('obd'):
            order_line += f"\nOBD {item['obd']}"
        if item.get('extra_orders'):
            for eo in item['extra_orders']:
                order_line += f"\n{eo}"
        cell_a = ws.cells(row_num, 1)
        cell_a.value = order_line
        cell_a.api.WrapText = True
        _apply_cell_style(cell_a, _style_for(item_id, _HIGHLIGHT_FIELDS_BY_COL[1]), bg, font_rgb, strike)

        description = item.get('description', '')
        try:
            qty_value = int(float(item.get('quantity') or 1))
        except (ValueError, TypeError):
            qty_value = 1
        if qty_value > 1:
            description = f"{description}\nQty: {qty_value}" if description else f"Qty: {qty_value}"
        cell_b = ws.cells(row_num, 2)
        cell_b.value = description
        cell_b.api.WrapText = True
        _apply_cell_style(cell_b, _style_for(item_id, _HIGHLIGHT_FIELDS_BY_COL[2]), bg, font_rgb, strike)

        cell_c = ws.cells(row_num, 3)
        cell_c.value = str(item.get('material', ''))
        _apply_cell_style(cell_c, _style_for(item_id, _HIGHLIGHT_FIELDS_BY_COL[3]), bg, font_rgb, strike)

        cell_d = ws.cells(row_num, 4)
        cell_d.value = item.get('serial_number', '')
        _apply_cell_style(cell_d, _style_for(item_id, _HIGHLIGHT_FIELDS_BY_COL[4]), bg, font_rgb, strike)

        cell_e = ws.cells(row_num, 5)
        cell_e.value = item.get('customer', '')
        _apply_cell_style(cell_e, _style_for(anchor_item_id, _HIGHLIGHT_FIELDS_BY_COL[5]), anchor_bg, anchor_font, anchor_strike)

        cell_f = ws.cells(row_num, 6)
        cell_f.value = item.get('phone', '')
        _apply_cell_style(cell_f, _style_for(anchor_item_id, _HIGHLIGHT_FIELDS_BY_COL[6]), anchor_bg, anchor_font, anchor_strike)

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
        if item.get('cust_no'):
            address_parts.append(f"(cust# {item['cust_no']})")
        cell_g = ws.cells(row_num, 7)
        cell_g.value = '\n'.join(address_parts)
        cell_g.api.WrapText = True
        _apply_cell_style(cell_g, _style_for(anchor_item_id, _HIGHLIGHT_FIELDS_BY_COL[7]), anchor_bg, anchor_font, anchor_strike)

        memo = item.get('memo', '')
        cell_h = ws.cells(row_num, 8)
        cell_h.value = memo
        cell_h.api.WrapText = True
        _apply_cell_style(cell_h, _style_for(anchor_item_id, _HIGHLIGHT_FIELDS_BY_COL[8]), anchor_bg, anchor_font, anchor_strike)

    if num_rows > 1:
        end_row = insert_at + num_rows - 1
        ws.book.app.display_alerts = False
        try:
            for col in [5, 6, 7, 8]:
                _com_retry(lambda col=col: ws.range(ws.cells(insert_at, col), ws.cells(end_row, col)).api.Merge())
        finally:
            ws.book.app.display_alerts = True

    border_range = ws.range(
        ws.cells(insert_at, 1),
        ws.cells(insert_at + num_rows - 1, 8)
    )
    for border_idx in [7, 8, 9, 10, 11, 12]:
        border_range.api.Borders(border_idx).LineStyle = 1
        border_range.api.Borders(border_idx).Weight = 2

    # 줄바꿈된 내용(주소/메모 등)에 맞게 높이를 다시 계산 - 삽입된 행이 인접
    # 행의 고정 높이를 물려받아 눌린 채로 남는 문제 방지(2026-08-31 실사용
    # 발견, find_or_create_section의 같은 조치 참고).
    ws.api.Rows(f"{insert_at}:{insert_at + num_rows - 1}").AutoFit()


# ── workbench.db 읽기 ────────────────────────────────────────────────────

def _connect_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _split_extra_codes(order_row):
    obd = (order_row['delivery_no'] or '').strip()
    extra = (order_row['extra_codes'] or '').strip()
    lines = [l.strip() for l in extra.splitlines() if l.strip()] if extra else []
    return obd, [l for l in lines if not obd or obd not in l]


def _build_order_data_list(order_row, item_rows):
    obd, extra_orders = _split_extra_codes(order_row)
    items = []
    for it in item_rows:
        items.append({
            'item_id': it['id'],
            'board_status': (it['board_status'] or '').strip(),
            'order_prefix': it['item_type'] or '',
            'order_type': order_row['order_type'] or '',
            'order_num': order_row['order_no'] or '',
            'obd': obd,
            'extra_orders': extra_orders,
            'description': it['description'] or '',
            'quantity': it['qty'] or 1,
            'material': it['material'] or '',
            'serial_number': it['serial'] or '',
            'customer': order_row['customer'] or '',
            'phone': order_row['phone'] or '',
            'company': '',
            'street': order_row['address'] or '',
            'street2': '',
            'cust_no': '',
            'memo': order_row['memo'] or '',
        })
    return items


def fetch_syncable_orders_all():
    """workbench.db 전체(지운 항목 제외)에서 (order_row, item_rows, target_date)
    를 돌려준다. 이미 시트에 있는지는 여기서 걸러내지 않는다(호출부가
    _already_in_sheet로 판단) - 그래야 여러 번 불러도 안전.

    항목은 order_items.item_date(워크벤치 보드에서 드래그로 다른 날짜로 옮긴
    경우 여기 저장됨)가 있으면 그 날짜로, 없으면 오더의 source_date로 묶어서
    반환한다 - 같은 오더 안에서도 항목별로 다른 날짜에 가 있을 수 있다.
    2026-08-31 실사용에서 발견: 이걸 무시하고 항상 오더 전체를 source_date
    하나로만 묶으면, 사용자가 워크벤치에서 이미 다른(주로 미래) 날짜로 옮겨둔
    오더가 옛날 source_date 자리에 중복으로 다시 만들어지는 버그가 있었다
    (실제 사용자 리포트로 발견 - 67055178은 8/22→9/7로, 7846435는
    8/24→9/9로 옮겼는데 8/22, 8/24 자리에도 또 생겼음)."""
    conn = _connect_db()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM orders ORDER BY id")
        out = []
        for o in cur.fetchall():
            if not o['source_date']:
                continue
            cur.execute(
                "SELECT * FROM order_items WHERE order_id=? AND deleted_at IS NULL ORDER BY line_no",
                (o['id'],),
            )
            items = cur.fetchall()
            if not items:
                continue

            groups = {}  # target_date -> [item_row, ...] (한 오더 안에서도 항목별로 날짜가 갈릴 수 있음)
            for it in items:
                eff_date_str = it['item_date'] or o['source_date']
                try:
                    eff_date = date.fromisoformat(eff_date_str)
                except ValueError:
                    continue
                groups.setdefault(eff_date, []).append(it)

            for target_date, group_items in groups.items():
                out.append((o, group_items, target_date))
        return out
    finally:
        conn.close()


def _rebuild_today_rows(ws, sheet_date):
    """sheet_date(보통 오늘)에 해당하는 섹션은 append 방식이 아니라 매번
    통째로 지우고 다시 쓴다. workbench 보드는 항목별로 board_pos(드래그 순서)/
    board_status/serial 등을 수시로 바꾸는데, 나머지 sync() 로직(이미 있으면
    건너뛰기)은 한 번 써진 행을 절대 갱신하지 않아서 시간이 지나면 엑셀이
    workbench 화면과 어긋나는 문제가 있었다(2026-08-31 실사용 발견: 순서가
    워크벤치의 board_pos와 다르고, 나중에 채운 시리얼 번호도 반영 안 됨).
    배너가 아직 없으면 find_or_create_section으로 새로 만든 뒤 채운다 -
    create_new_day_sheet()가 보통 미리 만들어두지만, sync()가 단독으로
    (버튼으로) 호출될 때도 항상 동작하게 하기 위함."""
    located = _locate_section(ws, sheet_date)
    if located:
        banner_row, end_row = located
        header_row = banner_row + 1
        if end_row > header_row:
            _com_retry(lambda: ws.range(f"{header_row + 1}:{end_row}").api.EntireRow.Delete())
    else:
        find_or_create_section(ws, sheet_date)
        banner_row, _ = _locate_section(ws, sheet_date)
        header_row = banner_row + 1

    conn = _connect_db()
    try:
        cur = conn.cursor()
        # 2026-09-02 사용자 요청: 이 파일(REALTIME_SYNC_PATH)은 "혹시 잃어버릴
        # 수 있는 데이터를 확인하는 안전망" 용도라, 마감 버튼으로 workbench.db
        # 에서 삭제된(soft-delete, 7일 복구 가능) 항목도 여기서는 사라지면 안
        # 된다는 지적 - 원래는 i.deleted_at IS NULL만 걸러서, 마감 직후 바로
        # 다음 sync()에서 그 항목이 오늘 섹션 재작성(rebuild) 때 통째로
        # 빠졌었다. 오늘 안에(date(deleted_at)=오늘) 삭제된 항목은 계속
        # 포함시켜 매번 다시 그려지게 한다 - 내일이 되면 이 날짜는 더 이상
        # "오늘"이 아니라 rebuild 대상에서 빠지므로(append 전용 섹션이 됨)
        # 그 시점 마지막으로 그려진 내용 그대로 영구히 남는다.
        cur.execute(
            """
            SELECT i.*, o.order_no, o.order_type, o.delivery_no, o.extra_codes,
                   o.customer, o.phone, o.address, o.memo AS order_memo
            FROM order_items i JOIN orders o ON o.id = i.order_id
            WHERE (i.deleted_at IS NULL OR date(i.deleted_at) = ?)
              AND (i.item_date = ? OR (i.item_date IS NULL AND o.source_date = ?))
            ORDER BY (i.board_pos IS NULL), i.board_pos, i.order_id, i.line_no
            """,
            (sheet_date.isoformat(), sheet_date.isoformat(), sheet_date.isoformat()),
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    # order_id별로 인접한 항목을 하나의 블록으로 묶는다(같은 오더는 병합
    # 서식이 함께 적용돼야 하므로) - board_pos로 정렬된 순서는 그대로 유지.
    blocks = []
    for r in rows:
        if blocks and blocks[-1][0] == r['order_id']:
            blocks[-1][1].append(r)
        else:
            blocks.append((r['order_id'], [r]))

    insert_at = header_row + 1
    for _order_id, item_group in blocks:
        first = item_group[0]
        order_row = {
            'order_no': first['order_no'], 'order_type': first['order_type'],
            'delivery_no': first['delivery_no'], 'extra_codes': first['extra_codes'],
            'customer': first['customer'], 'phone': first['phone'],
            'address': first['address'], 'memo': first['order_memo'],
        }
        order_data_list = _build_order_data_list(order_row, item_group)
        _write_order_rows(ws, insert_at, order_data_list)
        insert_at += len(order_data_list)

    return True


def _remove_empty_sections(ws, keep_date=None):
    """데이터가 하나도 없는(배너+헤더만 있는) 날짜 섹션은 통째로 지운다 -
    항목이 다른 날짜로 옮겨가면서 빈 배너만 남는 경우가 있다(2026-08-31
    실사용 발견: 9/15, 9/28 항목이 각각 9/18, 9/29로 옮겨간 뒤 빈 배너만
    남아있었음). keep_date(보통 시트 날짜)는 지금 비어있어도 지우지 않는다 -
    하루 시작 시점엔 당연히 비어있을 수 있는 자리라 지우면 오히려 헷갈림."""
    while True:
        banners, max_row = _find_banners(ws)
        removed = False
        for i, (r, mo, d) in enumerate(banners):
            if keep_date and mo == keep_date.month and d == keep_date.day:
                continue
            next_row = banners[i + 1][0] if i + 1 < len(banners) else None
            end_row = _section_end(ws, r, next_row, max_row)
            if end_row == r + 1:  # 헤더 행 자체가 반환됨 = 데이터 없음
                stale_end = next_row - 1 if next_row else max_row
                _com_retry(lambda r=r, stale_end=stale_end: ws.range(f"{r}:{stale_end}").api.EntireRow.Delete())
                removed = True
                break  # 행 번호가 다 바뀌었으니 배너 목록을 처음부터 다시 읽는다
        if not removed:
            return


# ── 진입점 ────────────────────────────────────────────────────────────────

def sync(sheet_date=None):
    """workbench.db에 있고 아직 해당 시트에 없는 오더를 채워 넣는다."""
    sheet_date = sheet_date or date.today()
    sheet_name = _sheet_name_for(sheet_date)
    with ExcelWriteLock():
        wb = _get_sync_workbook(sheet_date)
        if sheet_name not in [s.name for s in wb.sheets]:
            raise RuntimeError(f"'{sheet_name}' 시트가 없음 - 먼저 --new-day로 시트를 만드세요")
        ws = wb.sheets[sheet_name]

        # 오늘(시트 날짜) 섹션은 아래 append 루프가 아니라 항상 통째로 다시
        # 쓴다 - _rebuild_today_rows 독스트링 참고. 이미 처리했으니 아래
        # 루프에서는 오늘 날짜를 건너뛴다.
        _rebuild_today_rows(ws, sheet_date)

        added = 0
        for order_row, item_rows, target_date in fetch_syncable_orders_all():
            if target_date == sheet_date:
                continue
            # 시트 날짜보다 지난 날짜의 오더는, "완료된" 것만 건너뛴다(과거
            # 섹션은 create_new_day_sheet()의 _strip_past_sections()가 이미
            # 지운 뒤이니, 완료된 오더까지 이 체크 없이 되살리면 2026-08-28에
            # 실제로 재현된 "지운 지난 섹션이 되살아나는" 버그가 남). 반대로
            # 아직 안 끝난(board_status가 done/cancelled가 아닌 항목이 하나라도
            # 있는) 과거 오더는 여기서 건너뛰지 않고 그 오더의 원래 날짜로 새
            # 섹션을 만들어서라도 다시 넣는다 - 실사용자가 매일 손으로
            # "안 끝난 건 다음 날로 이월"해온 것과 같은 동작(2026-08-31 실제
            # 라이브 데이터로 재현 확인: 이게 없으면 아직 처리 중인 과거 오더
            # 10건이 새 시트에서 통째로 누락됨).
            if target_date < sheet_date and not _order_is_open(item_rows):
                continue
            # 이미 있는지는 그 오더의 "날짜 섹션 안"에서만 확인한다(전체
            # 시트 기준이면 같은 티켓 번호를 여러 날짜에 걸쳐 반복 사용하는
            # 오더가 영원히 안 들어가는 버그가 남 - _already_in_sheet
            # 독스트링 참고, 2026-08-28 실사용에서 발견).
            located = _locate_section(ws, target_date)
            if located:
                banner_row, end_row = located
                section_text = _section_column_a_text(ws, banner_row, end_row)
                if _already_in_sheet(section_text, order_row):
                    continue
                insert_at = end_row + 1
            else:
                insert_at = find_or_create_section(ws, target_date)
            order_data_list = _build_order_data_list(order_row, item_rows)
            _write_order_rows(ws, insert_at, order_data_list)
            added += 1

        _remove_empty_sections(ws, keep_date=sheet_date)

        # 개별 행마다 clear_beyond_h를 해도, 이 시트에 이미 남아있던 오래된
        # 헤더/데이터 행에 I열 이후로 번져있던 서식이 그 다음 EntireRow.Insert()
        # 시점에 다시 인접 행으로 옮겨붙는 경우가 있었다(2026-08-31 실사용
        # 발견 - 정확한 재현 조건은 못 찾았지만, "오래된(과거부터 있던) 섹션에
        # 처음으로 행을 추가할 때"에서 재현됨). 정확한 메커니즘을 다 밝히기보다,
        # 저장 직전에 이 시트 전체를 대상으로 I열 이후를 한 번 더 통째로 쓸어서
        # 지우는 게 더 확실하다 - A~H만 쓰는 표 구조상 그 바깥은 항상 비어있어야
        # 하므로 안전하다. _rebuild_today_rows도 행을 새로 쓰므로 added==0이어도
        # 항상 실행한다.
        used = ws.api.UsedRange
        last_row = used.Row + used.Rows.Count - 1
        _com_retry(lambda: ws.api.Range(ws.api.Cells(1, 9), ws.api.Cells(last_row, 16384)).Clear())

        # 시트 전체 줄간격을 실제 내용 기준으로 다시 계산(_refit_row_heights
        # 참고) - 오늘 새로 쓴 행뿐 아니라 예전부터 눌린 채 남아있던 섹션도
        # 매 sync()마다 교정된다.
        _refit_row_heights(ws)

        wb.save()
        result = {"sheet": sheet_name, "added": added}
        # workbench 버튼은 board_sync.sync()를 함수로 바로 호출하지 CLI를
        # 거치지 않아서(__main__ 블록의 로깅을 안 탐) 여기서 직접 남긴다 -
        # 안 그러면 "동기화 안 되는 것 같다"는 문의가 와도 뭘 했는지 흔적이
        # 하나도 안 남는다(2026-08-28 실제로 이래서 원인 파악에 애먹음).
        logger.info(f"sync() {result}")
        return result


def create_new_day_sheet(today=None):
    """오늘 시트가 이미 있으면 아무 것도 안 하고 끝낸다. 없으면 가장 최근
    (오늘 이전) 날짜 시트를 복사해서 오늘 이름으로 만들고, 지난 날짜 섹션을
    지운 뒤 sync()까지 이어서 실행한다."""
    today = today or date.today()
    sheet_name = _sheet_name_for(today)
    with ExcelWriteLock():
        wb = _get_sync_workbook(today)
        names = [s.name for s in wb.sheets]
        if sheet_name in names:
            wb.save()
            created = False
            src_name = None
        else:
            candidates = []
            for n in names:
                m = re.match(r'^(\d{1,2})-(\d{1,2})$', n)
                if m:
                    try:
                        d = date(today.year, int(m.group(1)), int(m.group(2)))
                    except ValueError:
                        continue
                    if d <= today:
                        candidates.append((d, n))
            if not candidates:
                raise RuntimeError("복사할 이전 날짜 시트를 찾지 못함 (오늘 이전 'M-D' 형식 시트가 없음)")
            candidates.sort()
            _, src_name = candidates[-1]
            src_ws = wb.sheets[src_name]
            src_index = [s.name for s in wb.sheets].index(src_name)

            _com_retry(lambda: src_ws.api.Copy(After=src_ws.api))
            new_ws = wb.sheets[src_index + 1]
            new_ws.name = sheet_name

            if not _has_banner_for(new_ws, today):
                find_or_create_section(new_ws, today)
            _strip_past_sections(new_ws, today)

            wb.save()
            created = True

    result = {"created": created, "sheet": sheet_name, "copied_from": src_name}
    sync_result = sync(today)  # sync() 자체도 자기 로그를 남김
    result["sync"] = sync_result
    logger.info(f"create_new_day_sheet() {result}")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="workbench.db → 배송장 엑셀 동기화")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--new-day", action="store_true", help="오늘 시트 생성 + 최초 전체 반영")
    group.add_argument("--sync", action="store_true", help="누락된 오더만 채워 넣기")
    args = parser.parse_args()

    try:
        if args.new_day:
            result = create_new_day_sheet()
        else:
            result = sync()
        logger.info(result)
        print(f"[board_sync] {result}")
    except Exception:
        logger.exception("board_sync 실패")
        raise

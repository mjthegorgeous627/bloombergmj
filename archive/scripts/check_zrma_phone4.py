"""
ZRMA Partners 탭 전화번호 완전 탐색.
1. Partners 테이블 모든 컬럼 확인
2. subSCREEN 내 모든 버튼/툴바 확인
3. 행 선택 후 btnTEL 재시도
"""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
navigate_to_zrma_order(session, group_zrma_by_order(rows)['66999156']['grid_idx'])
time.sleep(1)

# Partners 탭 진입
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
time.sleep(1.5)

_TABLE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
    "/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)
_SUB = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
)

table = session.findById(_TABLE)
sub   = session.findById(_SUB)

# ── 1. 컬럼 수 및 컬럼명 확인 ──────────────────────────────
print("=" * 60)
print("[Partners 테이블 컬럼 전체]")
print(f"RowCount={table.RowCount}  ColumnCount={table.ColumnCount}")

# 모든 컬럼 헤더
for c in range(table.ColumnCount):
    try:
        title = table.GetColumnTitles(c).ElementAt(0) if hasattr(table.GetColumnTitles(c), 'ElementAt') else ''
    except:
        title = ''
    # Ship-to 행(4) 값으로 컬럼 확인
    try:
        val = table.GetCell(4, c).Text.strip()
    except:
        val = '?'
    print(f"  Col[{c}] title='{title}' Ship-to_val='{val}'")

# ── 2. subSCREEN 내 모든 컨트롤 확인 ───────────────────────
print("\n" + "=" * 60)
print("[subSUBSCREEN_PARTNER_OVERVIEW 자식 목록]")
for i in range(sub.Children.Count):
    try:
        ch = sub.Children.ElementAt(i)
        short = ch.Id.split("/")[-1]
        etxt = getattr(ch, 'Text', '') or ''
        print(f"  [{ch.Type}] {short}" + (f" '{str(etxt)[:40]}'" if etxt else ""))
    except Exception as e:
        print(f"  [ERR {e}]")

# ── 3. 행 선택 방법 시도 ────────────────────────────────────
print("\n" + "=" * 60)
print("[행 선택 후 btnTEL 시도]")

_BTN_TEL = _SUB + "/btnTEL"

# 방법 A: SelectedRows 속성으로 행 선택
try:
    table.SelectedRows = "4"
    time.sleep(0.3)
    print(f"SelectedRows='4' 설정 성공")
    session.findById(_BTN_TEL).press()
    time.sleep(1.5)
    print(f"팝업: {session.findById('wnd[1]').Text}")
    # 팝업 내부
    def dump1(el, indent=0, depth=5):
        px="  "*indent
        try:
            eid=el.Id.split("/")[-1]; et=el.Type
            v=getattr(el,'Text','') or getattr(el,'Value','') or ''
            print(f"{px}[{et}] {eid}" + (f" '{str(v)[:60]}'" if v else ""))
        except: return
        if indent>=depth: return
        try:
            for i in range(el.Children.Count):
                try: dump1(el.Children.ElementAt(i),indent+1,depth)
                except: pass
        except: pass
    dump1(session.findById("wnd[1]"), depth=5)
    session.findById("wnd[1]").sendVKey(12)
    session.findById("wnd[0]").sendVKey(3)
    import sys; sys.exit(0)
except Exception as e:
    print(f"방법A 실패: {e}")

# 방법 B: GetCell(4,0).setFocus() + 일정 대기 후 btnTEL
try:
    table.GetCell(4, 0).setFocus()
    time.sleep(0.5)
    # 해당 행 click 시도
    table.GetCell(4, 1).setFocus()
    time.sleep(0.3)
    session.findById(_BTN_TEL).press()
    time.sleep(1.5)
    print(f"방법B 팝업: {session.findById('wnd[1]').Text}")
    dump1(session.findById("wnd[1]"), depth=5)
    session.findById("wnd[1]").sendVKey(12)
    session.findById("wnd[0]").sendVKey(3)
    import sys; sys.exit(0)
except Exception as e:
    print(f"방법B 실패: {e}")

# 방법 C: currentCellRow 직접 설정
try:
    table.currentCellRow = 4
    time.sleep(0.3)
    session.findById(_BTN_TEL).press()
    time.sleep(1.5)
    print(f"방법C 팝업: {session.findById('wnd[1]').Text}")
    dump1(session.findById("wnd[1]"), depth=5)
    session.findById("wnd[1]").sendVKey(12)
    session.findById("wnd[0]").sendVKey(3)
    import sys; sys.exit(0)
except Exception as e:
    print(f"방법C 실패: {e}")

session.findById("wnd[0]").sendVKey(3)

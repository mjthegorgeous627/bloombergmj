"""
ZRMA Partners 탭 전화번호 탐색 - GuiTableControl 방식
"""
import win32com.client, time, sys

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
order_map = group_zrma_by_order(rows)
print(f"오더 목록: {list(order_map.keys())[:5]}")

# 첫 번째 오더로 테스트
target = list(order_map.keys())[0]
navigate_to_zrma_order(session, order_map[target]['grid_idx'])
time.sleep(1)
print(f"오더 {target} 진입 완료")

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

# ── 1. 컬럼 수 (Columns.Count) ──────────────────────────────
try:
    col_count = table.Columns.Count
    print(f"RowCount={table.RowCount}  Columns.Count={col_count}")
except Exception as e:
    col_count = 15
    print(f"Columns.Count 실패: {e} → {col_count}으로 시도")

# ── 2. 각 행/컬럼 값 출력 ────────────────────────────────────
print("\n[Partners 테이블 내용]")
ship_to_row = -1
for r in range(table.RowCount):
    row_vals = []
    for c in range(col_count):
        try:
            v = table.GetCell(r, c).Text.strip()
            row_vals.append(f"[{c}]='{v}'")
        except:
            break
    line = "  ".join(row_vals)
    print(f"  행[{r}]: {line}")
    if 'Ship' in line:
        ship_to_row = r

print(f"\nShip-to 행: {ship_to_row}")

# ── 3. btnTEL 위치 확인 ──────────────────────────────────────
print("\n[subSCREEN 자식 목록]")
for i in range(sub.Children.Count):
    try:
        ch = sub.Children.ElementAt(i)
        short = ch.Id.split("/")[-1]
        etxt = getattr(ch, 'Text', '') or getattr(ch, 'Tooltip', '') or ''
        print(f"  [{ch.Type}] {short}" + (f" '{str(etxt)[:40]}'" if etxt else ""))
    except Exception as e:
        print(f"  [ERR {e}]")

# ── 4. Ship-to 행 선택 후 btnTEL 시도 ───────────────────────
if ship_to_row < 0:
    ship_to_row = 0  # fallback

_BTN_TEL = _SUB + "/btnTEL"

def try_btn_tel(label):
    try:
        btn = session.findById(_BTN_TEL)
        btn.press()
        time.sleep(1.5)
        popup = session.findById("wnd[1]")
        print(f"\n[{label}] 팝업 열림: '{popup.Text}'")
        # 팝업 내부 전체 덤프
        def dump(el, indent=0, depth=6):
            px = "  " * indent
            try:
                short = el.Id.split("/")[-1]
                v = getattr(el, 'Text', '') or getattr(el, 'Value', '') or ''
                print(f"{px}[{el.Type}] {short}" + (f" '{str(v)[:60]}'" if str(v).strip() else ""))
            except: return
            if indent >= depth: return
            try:
                for i in range(el.Children.Count):
                    try: dump(el.Children.ElementAt(i), indent+1, depth)
                    except: pass
            except: pass
        dump(popup)
        popup.sendVKey(12)
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"[{label}] 실패: {e}")
        try:
            session.findById("wnd[1]").sendVKey(12)
        except: pass
        return False

# 방법 A: SelectedRows
try:
    table.SelectedRows = str(ship_to_row)
    time.sleep(0.3)
    print(f"SelectedRows={ship_to_row} 설정 성공")
except Exception as e:
    print(f"SelectedRows 실패: {e}")

if try_btn_tel("방법A-SelectedRows"):
    sys.exit(0)

# 방법 B: currentCellRow
try:
    table.currentCellRow = ship_to_row
    time.sleep(0.3)
    print(f"currentCellRow={ship_to_row} 설정 성공")
except Exception as e:
    print(f"currentCellRow 실패: {e}")

if try_btn_tel("방법B-currentCellRow"):
    sys.exit(0)

# 방법 C: GetCell setFocus
try:
    table.GetCell(ship_to_row, 0).setFocus()
    time.sleep(0.5)
    print("setFocus(row,0) 성공")
except Exception as e:
    print(f"setFocus 실패: {e}")

if try_btn_tel("방법C-setFocus"):
    sys.exit(0)

print("\n모든 방법 실패 - Ship-to 행 더블클릭 시도")
# 방법 D: double-click → 팝업 주소 (SAPLSZA1)
try:
    table.GetCell(ship_to_row, 1).setFocus()
    time.sleep(0.3)
    table.doubleClickCurrentCell()
    time.sleep(1.5)
    popup = session.findById("wnd[1]")
    print(f"더블클릭 팝업: '{popup.Text}'")
    def dump2(el, indent=0, depth=6):
        px = "  " * indent
        try:
            short = el.Id.split("/")[-1]
            v = getattr(el, 'Text', '') or getattr(el, 'Value', '') or ''
            print(f"{px}[{el.Type}] {short}" + (f" '{str(v)[:60]}'" if str(v).strip() else ""))
        except: return
        if indent >= depth: return
        try:
            for i in range(el.Children.Count):
                try: dump2(el.Children.ElementAt(i), indent+1, depth)
                except: pass
        except: pass
    dump2(popup)
    popup.sendVKey(12)
except Exception as e:
    print(f"더블클릭 실패: {e}")

session.findById("wnd[0]").sendVKey(3)
print("\n완료")

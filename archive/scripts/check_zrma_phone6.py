"""
ZRMA 전화번호 - btnDETAIL 방식 + btnTEL 경로 직접 확인
"""
import win32com.client, time, sys

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
order_map = group_zrma_by_order(rows)
target = list(order_map.keys())[0]
navigate_to_zrma_order(session, order_map[target]['grid_idx'])
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

# Ship-to 행 찾기 (행[4])
ship_to_row = -1
for r in range(min(table.RowCount, 10)):
    try:
        parvw = table.GetCell(r, 0).Text.strip()
        if 'Ship' in parvw:
            ship_to_row = r
            break
    except: break

print(f"Ship-to 행: {ship_to_row}")

# ── 방법1: Ship-to 행 포커스 → btnDETAIL ──────────────────────
def dump_popup(label):
    try:
        popup = session.findById("wnd[1]")
        print(f"\n[{label}] 팝업 열림: '{popup.Text}'")
        def dump(el, indent=0, depth=7):
            px = "  " * indent
            try:
                short = el.Id.split("/")[-1]
                v = getattr(el, 'Text', '') or getattr(el, 'Value', '') or ''
                if str(v).strip():
                    print(f"{px}[{el.Type}] {short} '{str(v)[:80]}'")
                else:
                    print(f"{px}[{el.Type}] {short}")
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
        print(f"[{label}] 팝업 없음: {e}")
        return False

# Ship-to 행 포커스
if ship_to_row >= 0:
    table.GetCell(ship_to_row, 1).setFocus()
    time.sleep(0.5)

# btnDETAIL 실제 ID 확인
sub = session.findById(_SUB)
print(f"\n[btnDETAIL 실제 ID 확인]")
for i in range(sub.Children.Count):
    try:
        ch = sub.Children.ElementAt(i)
        if 'btn' in ch.Id.split('/')[-1].lower():
            print(f"  {ch.Id}")
    except: pass

# btnDETAIL 클릭
try:
    session.findById(_SUB + "/btnDETAIL").press()
    time.sleep(1.5)
    if dump_popup("btnDETAIL"):
        session.findById("wnd[0]").sendVKey(3)
        sys.exit(0)
except Exception as e:
    print(f"btnDETAIL 실패: {e}")

# ── 방법2: Contact person 행 포커스 → btnDETAIL ───────────────
print("\n[Contact person 행으로 시도]")
contact_row = -1
for r in range(min(table.RowCount, 10)):
    try:
        parvw = table.GetCell(r, 0).Text.strip()
        if 'Contact' in parvw:
            contact_row = r
            break
    except: break

if contact_row >= 0:
    table.GetCell(contact_row, 1).setFocus()
    time.sleep(0.5)
    try:
        session.findById(_SUB + "/btnDETAIL").press()
        time.sleep(1.5)
        dump_popup("Contact-btnDETAIL")
    except Exception as e:
        print(f"Contact btnDETAIL 실패: {e}")

# ── 방법3: btnTEL - 실제 ID로 재시도 ─────────────────────────
print("\n[btnTEL 실제 ID로 재시도]")
if ship_to_row >= 0:
    table.GetCell(ship_to_row, 1).setFocus()
    time.sleep(0.5)
# btnTEL 찾기
for i in range(sub.Children.Count):
    try:
        ch = sub.Children.ElementAt(i)
        if 'TEL' in ch.Id.split('/')[-1].upper():
            print(f"  btnTEL ID: {ch.Id}")
            ch.press()
            time.sleep(1.5)
            dump_popup("btnTEL-실제ID")
            break
    except Exception as e:
        print(f"  btnTEL press 실패: {e}")

session.findById("wnd[0]").sendVKey(3)
print("완료")

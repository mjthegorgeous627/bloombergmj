"""ZRMA 전화번호 경로 2가지 확인: Texts 탭, btnTEL."""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
navigate_to_zrma_order(session, group_zrma_by_order(rows)['66999156']['grid_idx'])
time.sleep(1)

def dump(element, indent=0, max_depth=5):
    prefix = "  " * indent
    try:
        eid = element.Id.split("wnd[0]/usr/")[-1]
        etype = element.Type
        etext = getattr(element, 'Text', '')
        print(f"{prefix}[{etype}] {eid}" + (f"  '{str(etext)[:60]}'" if etext else ""))
    except: return
    if indent >= max_depth: return
    try:
        for i in range(element.Children.Count):
            try: dump(element.Children.ElementAt(i), indent+1, max_depth)
            except: pass
    except: pass

# ── 1. Texts 탭 ──────────────────────────────────────
print("=" * 60)
print("[Goto > Header > Texts]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[10]").select()
time.sleep(1.5)
print(f"화면: {session.findById('wnd[0]').Text}")
dump(session.findById("wnd[0]/usr"), max_depth=6)
session.findById("wnd[0]").sendVKey(3)
time.sleep(1)

# ── 2. Partners > btnTEL ─────────────────────────────
print("\n" + "=" * 60)
print("[Goto > Header > Partners → btnTEL]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
time.sleep(1.5)

# btnTEL 클릭 (행 선택 없이 바로)
try:
    btn_tel_path = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
        "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
        "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnTEL"
    )
    session.findById(btn_tel_path).press()
    time.sleep(1.5)
    print(f"팝업 화면: {session.findById('wnd[1]').Text}")
    dump(session.findById("wnd[1]"), max_depth=5)
    session.findById("wnd[1]").sendVKey(12)
except Exception as e:
    print(f"btnTEL 실패: {e}")

session.findById("wnd[0]").sendVKey(3)

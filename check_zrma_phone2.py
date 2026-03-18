"""ZRMA Texts 탭 내부 + Partners btnTEL 정밀 확인."""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
navigate_to_zrma_order(session, group_zrma_by_order(rows)['66999156']['grid_idx'])
time.sleep(1)

# ── 1. Texts 탭 내부 깊이 탐색 ───────────────────────
print("=" * 60)
print("[Texts 탭 - SPLITTER_CONTAINER 내부]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[10]").select()
time.sleep(1.5)

splitter_cont = session.findById(
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4152"
    "/subSUBSCREEN_TEXT:SAPLV70T:2100"
    "/cntlSPLITTER_CONTAINER/shellcont"
)
print(f"shellcont 자식 수: {splitter_cont.Children.Count}")
for i in range(splitter_cont.Children.Count):
    child = splitter_cont.Children.ElementAt(i)
    print(f"\n  child[{i}]: type={child.Type}  id={child.Id}")
    for attr in ['Value', 'Text']:
        try:
            v = getattr(child, attr, '')
            if v:
                print(f"    {attr}: '{str(v)[:300]}'")
        except: pass
    # 한 단계 더
    try:
        for j in range(child.Children.Count):
            gc = child.Children.ElementAt(j)
            print(f"    grandchild[{j}]: type={gc.Type}  id={gc.Id.split('/')[-1]}")
            for attr in ['Value', 'Text']:
                try:
                    v = getattr(gc, attr, '')
                    if v:
                        print(f"      {attr}: '{str(v)[:300]}'")
                except: pass
    except: pass

session.findById("wnd[0]").sendVKey(3)
time.sleep(1)

# ── 2. Partners btnTEL - Partners 탭 먼저 클릭 후 시도 ──
print("\n" + "=" * 60)
print("[Partners 탭 직접 클릭 → btnTEL]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
time.sleep(1.5)

# 탭 직접 클릭으로 확실히 활성화
try:
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07").select()
    time.sleep(1)
except Exception as e:
    print(f"탭 select: {e}")

btn_tel_path = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnTEL"
)
try:
    session.findById(btn_tel_path).press()
    time.sleep(1.5)
    print(f"팝업: {session.findById('wnd[1]').Text}")
    # 팝업 내부 덤프
    def dump(el, indent=0, depth=4):
        px = "  "*indent
        try:
            eid = el.Id.split("/")[-1]; et = el.Type; etxt = getattr(el,'Text','')
            print(f"{px}[{et}] {eid}" + (f" '{str(etxt)[:60]}'" if etxt else ""))
            for attr in ['Value']:
                try:
                    v = getattr(el, attr, '')
                    if v: print(f"{px}  Value='{str(v)[:100]}'")
                except: pass
        except: return
        if indent>=depth: return
        try:
            for i in range(el.Children.Count):
                try: dump(el.Children.ElementAt(i), indent+1, depth)
                except: pass
        except: pass
    dump(session.findById("wnd[1]"), depth=5)
    session.findById("wnd[1]").sendVKey(12)
except Exception as e:
    print(f"btnTEL 실패: {e}")

session.findById("wnd[0]").sendVKey(3)

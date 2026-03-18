"""ZRMA_Q Partners 화면 element tree 확인."""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

# 오더 66999156 진입
from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
order_map = group_zrma_by_order(rows)
navigate_to_zrma_order(session, order_map['66999156']['grid_idx'])
time.sleep(1)

# Goto > Header > Partners
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
time.sleep(1.5)
print(f"현재 화면: {session.findById('wnd[0]').Text}")

def explore(element, indent=0, max_depth=6):
    prefix = "  " * indent
    try:
        eid = element.Id.split("wnd[0]/usr/")[-1]
        etype = element.Type
        etext = getattr(element, 'Text', '')
        print(f"{prefix}[{etype}] {eid}" + (f"  '{str(etext)[:50]}'" if etext else ""))
    except Exception as e:
        print(f"{prefix}[ERR: {e}]"); return
    if indent >= max_depth: return
    try:
        for i in range(element.Children.Count):
            try: explore(element.Children.ElementAt(i), indent+1, max_depth)
            except Exception: pass
    except Exception: pass

explore(session.findById("wnd[0]/usr"), max_depth=6)

# 복귀
session.findById("wnd[0]").sendVKey(3)

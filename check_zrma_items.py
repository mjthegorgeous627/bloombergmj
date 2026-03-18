"""ZRMA_Q 오더 상세 화면 - 아이템 테이블 ID 확인용."""
import win32com.client

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)  # 세션2

print(f"현재 화면: {session.findById('wnd[0]').Text}\n")

# Sales 탭의 SUBSCREEN_TC 컨테이너 깊이 탐색
def explore(element, indent=0, max_depth=8):
    prefix = "  " * indent
    try:
        eid = element.Id.split("wnd[0]/usr/")[-1]  # 짧게 표시
        etype = element.Type
        ename = getattr(element, 'Name', '')
        etext = getattr(element, 'Text', '')
        line = f"{prefix}[{etype}] {eid}"
        if etext:
            line += f"  text='{str(etext)[:50]}'"
        print(line)
    except Exception as e:
        print(f"{prefix}[ERR: {e}]")
        return

    if indent >= max_depth:
        return
    try:
        for i in range(element.Children.Count):
            try:
                explore(element.Children.ElementAt(i), indent + 1, max_depth)
            except Exception:
                pass
    except Exception:
        pass

# 1. Sales 탭 안의 SUBSCREEN_TC 탐색
print("=" * 60)
print("[Sales 탭 - SUBSCREEN_TC 내부]")
try:
    tc = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01"
        "/ssubSUBSCREEN_BODY:SAPMV45A:4400"
        "/subSUBSCREEN_TC:SAPMV45A:4900"
    )
    explore(tc, max_depth=5)
except Exception as e:
    print(f"SUBSCREEN_TC 접근 실패: {e}")

# 2. Item Overview 탭 탐색
print("\n" + "=" * 60)
print("[Item Overview 탭 내부]")
try:
    tab2 = session.findById(
        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02"
    )
    explore(tab2, max_depth=5)
except Exception as e:
    print(f"Item Overview 탭 접근 실패: {e}")

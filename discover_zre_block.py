"""
ZRE 오더 Delivery Block 필드 직접 접근 테스트.
세션1이 'Change RMA Order: Overview' 화면에 있어야 함.
"""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(1)
print(f"현재 화면: {session.findById('wnd[0]').Text}\n")

BLOCK_PATH = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK"

try:
    field = session.findById(BLOCK_PATH)
    print(f"[발견] 현재 값: '{field.Text.strip()}'")
    print(f"  key: '{field.key}'")
    print("→ 경로 OK")
except Exception as e:
    print(f"[실패] {e}")

    # Sales 탭 클릭 후 재시도
    print("\nSales 탭 클릭 후 재시도...")
    try:
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01").select()
        time.sleep(0.5)
        field = session.findById(BLOCK_PATH)
        print(f"[발견] 현재 값: '{field.Text.strip()}'")
        print("→ Sales 탭 select 후 접근 가능")
    except Exception as e2:
        print(f"[실패] {e2}")

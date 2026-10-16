"""
VL10G ZZ 블록 해제 경로 테스트.
7774732 ZOR 오더 진입 → Delivery Block 필드 찾기 → 실제 저장은 안 함.
"""
import win32com.client, time
from vl10g_handler import VL10G_GRID_PATH, COL_ORIG_DOC

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(1)

grid = session.findById(VL10G_GRID_PATH)

# 7774732 ZOR 행 찾기 (grid_idx=1)
TARGET_IDX = 1
TARGET_DOC = "7774732"

print(f"오더 {TARGET_DOC} 진입 (VBELV 클릭)...")
grid.setCurrentCell(TARGET_IDX, COL_ORIG_DOC)
grid.clickCurrentCell()
time.sleep(2)

title = session.findById("wnd[0]").Text
print(f"화면: '{title}'")

# Delivery Block 필드 탐색
block_candidates = [
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/cmbVBAK-LIFSK",
    "wnd[0]/usr/cmbVBAK-LIFSK",
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/cmbVBAK-LIFSP",
    "wnd[0]/usr/cmbVBAK-LIFSP",
]

found_fid = None
for fid in block_candidates:
    try:
        field = session.findById(fid)
        val = getattr(field, 'Value', '') or getattr(field, 'Text', '') or field.key
        print(f"[발견] {fid.split('/')[-1]} = '{val}'")
        found_fid = fid
        break
    except Exception as e:
        print(f"[없음] {fid.split('/')[-1]}: {e}")

if not found_fid:
    print("\n필드 못 찾음 - 화면 구조 덤프:")
    # 탭 T\\01 확인
    for tab_id in ["tabpT\\01", "tabpT\\02", "tabpT\\03"]:
        try:
            tab = session.findById(f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/{tab_id}")
            print(f"  탭 {tab_id} 존재")
        except:
            pass
    # LIFSK/LIFSP 직접 검색
    for sfx in ["LIFSK", "LIFSP"]:
        for pfx in ["cmb", "txt", "ctxt"]:
            fid2 = f"wnd[0]/usr/{pfx}VBAK-{sfx}"
            try:
                f2 = session.findById(fid2)
                print(f"  [발견] {fid2}")
            except:
                pass
else:
    print(f"\n→ Block 필드 확인됨: {found_fid.split('/')[-1]}")
    print("  (실제 해제는 하지 않음 - 테스트만)")

# F3 복귀
session.findById("wnd[0]").sendVKey(3)
time.sleep(1)
print(f"\n복귀 화면: {session.findById('wnd[0]').Text}")

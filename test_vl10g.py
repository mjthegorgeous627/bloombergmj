"""
VL10G 동작 테스트 - 행 읽기 + Background 버튼 흐름 확인
실제 배송 생성은 하지 않음 (행 선택 → Background 클릭 전 중단)
"""
import win32com.client, time
from vl10g_handler import get_all_rows_from_vl10g, VL10G_GRID_PATH, COL_ORIG_DOC

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(1)
print(f"화면: {session.findById('wnd[0]').Text}")

# ── 1. 행 읽기 ──────────────────────────────────────────────
rows = get_all_rows_from_vl10g(session)
print(f"\n전체 행 수: {len(rows)}")
for r in rows:
    print(f"  {r}")

if not rows:
    print("행 없음 - 종료")
    exit()

# ── 2. 첫 번째 행 선택 방식 테스트 ──────────────────────────
grid = session.findById(VL10G_GRID_PATH)
first = rows[0]
print(f"\n첫 번째 행 ({first['orig_doc']}) 선택 시도...")

# ALV Grid 행 선택: selectedRows 속성
try:
    grid.selectedRows = str(first['grid_idx'])
    time.sleep(0.3)
    print(f"selectedRows={first['grid_idx']} 설정 성공")
    sel = grid.selectedRows
    print(f"현재 selectedRows: {sel}")
except Exception as e:
    print(f"selectedRows 실패: {e}")

# ── 3. VBELV 셀 클릭 → VA02 진입 여부 확인 ──────────────────
print(f"\nVBELV 셀 클릭 테스트 (오더: {first['orig_doc']})")
try:
    grid.setCurrentCell(first['grid_idx'], COL_ORIG_DOC)
    time.sleep(0.3)
    grid.clickCurrentCell()
    time.sleep(1.5)
    title_after = session.findById("wnd[0]").Text
    print(f"클릭 후 화면: '{title_after}'")
    # 원래 화면이 아니면 F3로 복귀
    if "Activities" not in title_after:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)
        print("F3 복귀")
except Exception as e:
    print(f"셀 클릭 실패: {e}")

# ── 4. LIFSP 값 확인 ─────────────────────────────────────────
print("\n[LIFSP (Delivery Block) 값 확인]")
for r in rows:
    print(f"  오더 {r['orig_doc']} | 유형 {r['doc_type']} | Block='{r['deliv_block']}' | VBELN='{r['vbeln']}'")

print("\n완료 (Background 버튼은 클릭하지 않음)")

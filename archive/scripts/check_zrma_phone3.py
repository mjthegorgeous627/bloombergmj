"""
ZRMA 전화번호 2가지 접근법 심층 탐색.
1. Texts 탭 SPLITTER_CONTAINER 깊은 탐색 (shellcont 내부 GuiTextedit)
2. Partners 탭 Ship-to 행 더블클릭 → 팝업 주소 읽기
"""
import win32com.client, time

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)

from zrma_handler import get_all_rows_from_zrma, group_zrma_by_order, navigate_to_zrma_order
rows = get_all_rows_from_zrma(session)
navigate_to_zrma_order(session, group_zrma_by_order(rows)['66999156']['grid_idx'])
time.sleep(1)

# ── 1. Texts 탭 - SPLITTER_CONTAINER 깊이 탐색 ──────────────
print("=" * 60)
print("[Texts 탭 깊이 탐색 - shellcont/shell 하위]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[10]").select()
time.sleep(1.5)

def dump_deep(el, indent=0, max_depth=8, path=""):
    """Value/Text/Lines까지 출력하는 깊은 덤프."""
    px = "  " * indent
    try:
        eid = el.Id
        etype = el.Type
    except:
        return

    short_id = eid.split("/")[-1]
    print(f"{px}[{etype}] {short_id}")

    # 텍스트 내용 출력
    for attr in ['Value', 'Text']:
        try:
            v = getattr(el, attr, None)
            if v and str(v).strip():
                snippet = str(v)[:200].replace('\n', '\\n')
                print(f"{px}  ** {attr}: '{snippet}'")
        except:
            pass

    # Lines (GuiTextedit)
    try:
        lines_obj = el.Lines
        if lines_obj and lines_obj.Count > 0:
            print(f"{px}  ** Lines.Count: {lines_obj.Count}")
            for i in range(min(lines_obj.Count, 5)):
                try:
                    print(f"{px}     Line[{i}]: '{lines_obj.ElementAt(i)}'")
                except:
                    pass
    except:
        pass

    if indent >= max_depth:
        return

    try:
        cnt = el.Children.Count
        for i in range(cnt):
            try:
                dump_deep(el.Children.ElementAt(i), indent+1, max_depth)
            except:
                pass
    except:
        pass

# SPLITTER_CONTAINER 전체 구조
try:
    base = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08"
        "/ssubSUBSCREEN_BODY:SAPMV45A:4152"
        "/subSUBSCREEN_TEXT:SAPLV70T:2100"
    )
    sub = session.findById(base)
    print(f"subSUBSCREEN_TEXT 자식 수: {sub.Children.Count}")
    dump_deep(sub, max_depth=8)
except Exception as e:
    print(f"subSUBSCREEN_TEXT 접근 실패: {e}")

# F3 복귀
session.findById("wnd[0]").sendVKey(3)
time.sleep(1)

# ── 2. Partners 탭 - Ship-to 행 더블클릭 ──────────────────────
print("\n" + "=" * 60)
print("[Partners 탭 Ship-to 행 더블클릭]")
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[9]").select()
time.sleep(1.5)

_TABLE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
    "/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)

try:
    table = session.findById(_TABLE)
    print(f"Partners 테이블 행 수: {table.RowCount}")

    # Ship-to 행 찾기
    ship_row = -1
    for i in range(table.RowCount):
        try:
            parvw = table.GetCell(i, 0).Text.strip()
            name1 = table.GetCell(i, 4).Text.strip()
            print(f"  행[{i}] PARVW='{parvw}' NAME1='{name1}'")
            if 'Ship' in parvw:
                ship_row = i
        except:
            break

    if ship_row >= 0:
        print(f"\n>>> Ship-to 행 {ship_row} 더블클릭 시도")
        # 방법 1: GetCell setFocus 후 doubleClick
        try:
            cell = table.GetCell(ship_row, 4)
            cell.setFocus()
            time.sleep(0.3)
            # doubleClick - GuiTableControl 행 더블클릭
            table.doubleClickCurrentCell()
            time.sleep(1.5)
            # 팝업 확인
            try:
                popup_title = session.findById("wnd[1]").Text
                print(f"팝업 열림: '{popup_title}'")
                dump_deep(session.findById("wnd[1]"), max_depth=6)
                session.findById("wnd[1]").sendVKey(12)
            except Exception as e:
                print(f"팝업 없음: {e}")
        except Exception as e:
            print(f"더블클릭 실패: {e}")
            # 방법 2: btnTEL (행 선택 후)
            print("\n>>> btnTEL 재시도")
            try:
                table.GetCell(ship_row, 0).setFocus()
                time.sleep(0.5)
                btn_path = (
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
                    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
                    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/btnTEL"
                )
                session.findById(btn_path).press()
                time.sleep(1.5)
                popup_title = session.findById("wnd[1]").Text
                print(f"btnTEL 팝업: '{popup_title}'")
                dump_deep(session.findById("wnd[1]"), max_depth=6)
                session.findById("wnd[1]").sendVKey(12)
            except Exception as e2:
                print(f"btnTEL도 실패: {e2}")
    else:
        print("Ship-to 행 못 찾음")

except Exception as e:
    print(f"Partners 테이블 접근 실패: {e}")

session.findById("wnd[0]").sendVKey(3)

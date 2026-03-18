"""
SAP GUI 요소 탐색 도구.

SAP GUI가 특정 화면을 열고 있을 때 이 스크립트를 실행하면
현재 화면의 모든 UI 요소 ID를 출력합니다.
sap_handler.py의 element ID를 맞추는 데 사용하세요.

사용법:
  1. SAP GUI에서 원하는 화면을 열어둔 상태로
  2. python discover_sap.py 실행
  3. 출력된 ID 목록에서 필요한 요소를 찾아 sap_handler.py에 반영
"""

import win32com.client
import time


def print_element_tree(element, indent=0, max_depth=6):
    prefix = "  " * indent
    try:
        elem_id = element.Id
        elem_type = element.Type
        elem_name = getattr(element, 'Name', '')
        elem_text = getattr(element, 'Text', '')
        elem_value = ''
        try:
            elem_value = getattr(element, 'Value', '')
        except Exception:
            pass

        line = f"{prefix}[{elem_type}] id={elem_id}"
        if elem_name:
            line += f"  name='{elem_name}'"
        if elem_text:
            line += f"  text='{str(elem_text)[:60]}'"
        if elem_value:
            line += f"  value='{str(elem_value)[:60]}'"
        print(line)

    except Exception as e:
        print(f"{prefix}[ERROR reading element: {e}]")
        return

    if indent >= max_depth:
        return

    try:
        count = element.Children.Count
        for i in range(count):
            try:
                child = element.Children.ElementAt(i)
                print_element_tree(child, indent + 1, max_depth)
            except Exception:
                pass
    except Exception:
        pass


def print_grid_columns(session):
    """ALV 그리드의 컬럼 이름 출력 (오더 목록 화면에서 실행)."""
    grid_paths = [
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlGRID/shellcont/shell",
    ]
    for path in grid_paths:
        try:
            grid = session.findById(path)
            print(f"\n[그리드 발견] 경로: {path}")
            print(f"  행 수: {grid.RowCount}")
            cols = grid.ColumnOrder
            print(f"  컬럼 목록: {list(cols)}")
            # 첫 번째 행 데이터 샘플
            if grid.RowCount > 0:
                print("  [0행 데이터 샘플]")
                for col in list(cols)[:10]:
                    try:
                        val = grid.GetCellValue(0, col)
                        print(f"    {col} = '{val}'")
                    except Exception:
                        pass
            return grid
        except Exception:
            continue
    print("[그리드를 찾을 수 없음 - 현재 화면이 오더 목록 화면인지 확인]")
    return None


def main():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        app = sap_gui_auto.GetScriptingEngine
        conn = app.Children(0)

        # 열려있는 세션 목록 출력
        session_count = conn.Children.Count
        print(f"열려있는 SAP 세션 수: {session_count}")
        for i in range(session_count):
            try:
                s = conn.Children(i)
                print(f"  세션{i}: {s.findById('wnd[0]').Text}")
            except Exception:
                print(f"  세션{i}: (읽기 실패)")

        idx = input(f"사용할 세션 번호 (0~{session_count-1}): ").strip()
        session = conn.Children(int(idx))
        print(f"세션{idx} 연결 성공")
        print(f"현재 화면 제목: {session.findById('wnd[0]').Text}")
        print(f"현재 트랜잭션: {session.Info.Transaction}\n")
    except Exception as e:
        print(f"SAP 연결 실패: {e}")
        return

    print("=" * 70)
    print("옵션을 선택하세요:")
    print("  1. 현재 화면 전체 요소 트리 출력 (느릴 수 있음)")
    print("  2. ALV 그리드 컬럼명만 출력 (오더 목록 화면에서)")
    print("  3. 메뉴바 구조 출력")
    print("  4. 팝업(wnd[1]) 요소 트리 - 5초 카운트다운 후 스캔")
    print("=" * 70)
    choice = input("선택 (1/2/3/4): ").strip()

    if choice == "1":
        print("\n[현재 화면 요소 트리]")
        print_element_tree(session.findById("wnd[0]"), max_depth=5)

    elif choice == "2":
        print_grid_columns(session)

    elif choice == "3":
        print("\n[메뉴바 구조]")
        try:
            mbar = session.findById("wnd[0]/mbar")
            for i in range(mbar.Children.Count):
                menu = mbar.Children.ElementAt(i)
                print(f"  menu[{i}]: {menu.Text}")
                try:
                    for j in range(menu.Children.Count):
                        sub = menu.Children.ElementAt(j)
                        print(f"    menu[{i}]/menu[{j}]: {sub.Text}")
                        try:
                            for k in range(sub.Children.Count):
                                subsub = sub.Children.ElementAt(k)
                                print(f"      menu[{i}]/menu[{j}]/menu[{k}]: {subsub.Text}")
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception as e:
            print(f"메뉴 읽기 실패: {e}")

    elif choice == "4":
        print("\n[팝업 스캔] SAP에서 팝업을 열어두세요.")
        for i in range(5, 0, -1):
            print(f"  {i}초 후 스캔...", end="\r")
            time.sleep(1)
        print("\n[팝업(wnd[1]) 요소 트리]")
        try:
            popup = session.findById("wnd[1]")
            print_element_tree(popup, max_depth=6)
        except Exception as e:
            print(f"팝업(wnd[1]) 없음: {e}")

    print("\n완료.")


if __name__ == "__main__":
    main()

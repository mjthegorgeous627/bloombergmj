"""
Non-interactive SAP GUI Scripting probe.
SAP GUI 8.0 UI 변경 진단용 - 사용자 입력 없이 현재 연결 상태를 출력한다.

사용법:
  python diag_probe.py
"""

import sys


def main():
    try:
        import win32com.client
    except Exception as e:
        print(f"[FAIL] pywin32 import 실패: {e}")
        return 1

    sap_gui_auto = None
    for moniker in ("SAPGUI", "SAPGUISERVER"):
        try:
            sap_gui_auto = win32com.client.GetObject(moniker)
            print(f"[OK] GetObject('{moniker}') 성공")
            break
        except Exception as e:
            print(f"[FAIL] GetObject('{moniker}') 실패: {e}")

    if sap_gui_auto is None:
        return 1

    try:
        app = sap_gui_auto.GetScriptingEngine
    except Exception as e:
        print(f"[FAIL] GetScriptingEngine 실패: {e}")
        return 1

    print(f"[OK] Scripting Engine 연결 성공")

    try:
        conn_count = app.Children.Count
    except Exception as e:
        print(f"[FAIL] Connections 조회 실패: {e}")
        return 1

    print(f"연결(Connection) 수: {conn_count}")

    if conn_count == 0:
        print("[INFO] 열려있는 SAP 연결이 없습니다. SAP GUI 아이콘을 실행하고 로그인한 뒤 다시 실행하세요.")
        return 0

    for ci in range(conn_count):
        try:
            conn = app.Children(ci)
        except Exception as e:
            print(f"  연결[{ci}]: 조회 실패 ({e})")
            continue

        try:
            conn_desc = conn.Description
        except Exception:
            conn_desc = "?"
        print(f"\n연결[{ci}] Description={conn_desc}")

        try:
            sess_count = conn.Children.Count
        except Exception as e:
            print(f"  세션 수 조회 실패: {e}")
            continue

        print(f"  세션 수: {sess_count}")

        for si in range(sess_count):
            try:
                sess = conn.Children(si)
            except Exception as e:
                print(f"  세션[{si}]: 조회 실패 ({e})")
                continue

            try:
                title = sess.findById("wnd[0]").Text
            except Exception as e:
                title = f"(읽기 실패: {e})"

            try:
                tcode = sess.Info.Transaction
            except Exception:
                tcode = "?"

            try:
                program = sess.Info.Program
            except Exception:
                program = "?"

            try:
                screen = sess.Info.ScreenNumber
            except Exception:
                screen = "?"

            print(f"    세션[{si}] title='{title}' tcode={tcode} program={program} screen={screen}")

    return 0


if __name__ == "__main__":
    sys.exit(main())

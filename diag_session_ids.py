"""
Read-only diagnostic: dump session Id / Info.SessionNumber / title for every
open SAP session, to check whether SessionNumber is stable and distinct from
the session's current position (tab order) under NWBC.

Usage:
  python diag_session_ids.py
"""

import sys


def main():
    import win32com.client

    sap_gui_auto = None
    for moniker in ("SAPGUI", "SAPGUISERVER"):
        try:
            sap_gui_auto = win32com.client.GetObject(moniker)
            break
        except Exception:
            pass
    if sap_gui_auto is None:
        print("[FAIL] SAP GUI Scripting 연결 실패")
        return 1

    engine = sap_gui_auto.GetScriptingEngine
    if not hasattr(engine, "Children"):
        engine = engine()

    conn_count = engine.Children.Count
    print(f"연결 수: {conn_count}")

    for ci in range(conn_count):
        conn = engine.Children(ci)
        sess_count = conn.Children.Count
        print(f"\n연결[{ci}] 세션 수: {sess_count}")

        for si in range(sess_count):
            sess = conn.Children(si)
            try:
                title = sess.findById("wnd[0]").Text
            except Exception as e:
                title = f"(읽기 실패: {e})"
            try:
                sess_id = sess.Id
            except Exception as e:
                sess_id = f"(실패: {e})"
            try:
                sess_num = sess.Info.SessionNumber
            except Exception as e:
                sess_num = f"(실패: {e})"
            try:
                tcode = sess.Info.Transaction
            except Exception:
                tcode = "?"

            print(f"  위치[{si}] title='{title}' tcode={tcode} Id={sess_id} SessionNumber={sess_num}")

    return 0


if __name__ == "__main__":
    sys.exit(main())

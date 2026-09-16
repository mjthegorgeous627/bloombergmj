"""Inspect SAP return-item serial popup without saving changes."""

import sys
import time

from sap_handler import run_transaction, get_scripting_engine
from zrma_handler import EXTRAS_TECH_OBJ, ZRMA_ITEMS_TABLE_PATH, _read_serial_numbers_from_popup


def dump_control(ctrl, indent=0, max_depth=4):
    pad = "  " * indent
    try:
        text = getattr(ctrl, "Text", "")
    except Exception:
        text = ""
    try:
        name = getattr(ctrl, "Name", "")
    except Exception:
        name = ""
    try:
        ctype = getattr(ctrl, "Type", "")
    except Exception:
        ctype = ""
    print(f"{pad}{ctype} name={name} id={ctrl.Id} text={text!r}")
    if indent >= max_depth:
        return
    try:
        count = ctrl.Children.Count
    except Exception:
        return
    for i in range(count):
        try:
            dump_control(ctrl.Children.ElementAt(i), indent + 1, max_depth)
        except Exception as exc:
            print(f"{pad}  child {i}: {exc}")


def dump_menu(session):
    print("\nMENUBAR:")
    for mi in range(8):
        try:
            menu = session.findById(f"wnd[0]/mbar/menu[{mi}]")
        except Exception:
            continue
        try:
            text = menu.Text
        except Exception:
            text = ""
        print(f"menu[{mi}] text={text!r} id={menu.Id}")
        try:
            count = menu.Children.Count
        except Exception:
            continue
        for ci in range(count):
            try:
                child = menu.Children.ElementAt(ci)
                try:
                    ctext = child.Text
                except Exception:
                    ctext = ""
                print(f"  child[{ci}] type={child.Type} text={ctext!r} id={child.Id}")
            except Exception as exc:
                print(f"  child[{ci}] {repr(exc)}")


def open_new_session():
    # sap_handler.get_scripting_engine()가 NWBC(SAPGUISERVER)/classic(SAPGUI)
    # 모니커 차이를 이미 안전하게 처리함 - 여기서 win32com.client.GetObject를
    # 직접 부르면 안 됨 (memory: 저수준 pywin32 우회는 실제 SAP GUI 크래시를
    # 한 번 낸 적 있음. get_scripting_engine()을 그대로 쓸 것).
    app = get_scripting_engine()
    conn = app.Children(0)
    before = {conn.Children(i).Info.SessionNumber for i in range(conn.Children.Count)}
    conn.Children(0).createSession()
    deadline = time.time() + 10
    while time.time() < deadline:
        time.sleep(0.3)
        for i in range(conn.Children.Count):
            s = conn.Children(i)
            if s.Info.SessionNumber not in before:
                return s
    raise RuntimeError("SAP session create timeout")


def close_popup(session):
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.5)
    except Exception:
        try:
            session.findById("wnd[1]").sendVKey(0)
            time.sleep(0.5)
        except Exception:
            pass


def main():
    order = sys.argv[1] if len(sys.argv) > 1 else "67030748"
    session = open_new_session()
    try:
        run_transaction(session, "/nVA02")
        time.sleep(1)
        session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = order
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(2)
        print("TITLE:", session.findById("wnd[0]").Text)

        table = session.findById(ZRMA_ITEMS_TABLE_PATH)
        print("TABLE:", table.Id, "rows=", table.RowCount)
        dump_menu(session)
        for i in range(table.RowCount):
            try:
                item_no = table.GetCell(i, 0).Text.strip()
            except Exception:
                break
            if not item_no:
                continue
            mat = table.GetCell(i, 1).Text.strip()
            desc = table.GetCell(i, 2).Text.strip()
            qty = table.GetCell(i, 3).Text.strip()
            try:
                route = table.GetCell(i, 12).Text.strip()
            except Exception:
                route = ""
            print(f"ROW {i}: item={item_no} mat={mat} desc={desc} qty={qty} route={route}")

        target_rows = []
        for i in range(table.RowCount):
            try:
                route = table.GetCell(i, 12).Text.strip().upper()
                desc = table.GetCell(i, 2).Text.strip().upper()
                item_no = table.GetCell(i, 0).Text.strip()
            except Exception:
                break
            if item_no and route == "RETURN" and "BUNIT" not in desc:
                target_rows.append(i)

        print("RETURN_ROWS:", target_rows)
        for row_i in target_rows:
            print(f"\n--- inspect row {row_i} ---")
            table = session.findById(ZRMA_ITEMS_TABLE_PATH)
            for col in (0, 1, 2):
                try:
                    table.GetCell(row_i, col).setFocus()
                    time.sleep(0.2)
                    print(f"focus col {col}: ok")
                except Exception as exc:
                    print(f"focus col {col}: {exc}")
            try:
                table.rows.elementAt(row_i).selected = True
                print("row selected ok")
            except Exception as exc:
                print("row selected failed:", exc)
            session.findById(EXTRAS_TECH_OBJ).select()
            time.sleep(1.5)
            try:
                popup = session.findById("wnd[1]")
                popup_text = popup.Text
                print("POPUP:", popup_text)
                dump_control(popup)
                if "Serial" in popup_text:
                    print("READ_SN:", _read_serial_numbers_from_popup(session))
                close_popup(session)
                if popup_text == "Information":
                    print("RETRY TECHNICAL OBJECTS AFTER INFORMATION")
                    session.findById(EXTRAS_TECH_OBJ).select()
                    time.sleep(1.5)
                    popup = session.findById("wnd[1]")
                    print("POPUP RETRY:", popup.Text)
                    dump_control(popup)
                    if "Serial" in popup.Text:
                        print("READ_SN_RETRY:", _read_serial_numbers_from_popup(session))
                    close_popup(session)
            except Exception as exc:
                print("NO POPUP:", repr(exc))
    finally:
        try:
            session.findById("wnd[0]").close()
        except Exception:
            pass


if __name__ == "__main__":
    main()

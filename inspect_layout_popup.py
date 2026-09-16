import win32com.client

sap = win32com.client.GetObject("SAPGUI")
app = sap.GetScriptingEngine
conn = app.Children(0)

for sess_idx in range(conn.Children.Count):
    sess = conn.Children(sess_idx)
    try:
        title = sess.findById("wnd[0]").Text
    except Exception:
        title = ""
    print(f"\n=== session {sess_idx}: {title} ===")
    try:
        popup = sess.findById("wnd[1]")
        print("popup:", popup.Id, popup.Type, getattr(popup, "Text", ""))
    except Exception as e:
        print("no popup", e)
        continue

    def walk(obj, depth=0, limit=5):
        if depth > limit:
            return
        try:
            children = obj.Children
            count = children.Count
        except Exception:
            return
        for i in range(count):
            try:
                child = children(i)
                text = getattr(child, "Text", "")
                name = getattr(child, "Name", "")
                print("  " * depth + f"[{i}] id={child.Id} type={child.Type} name={name} text={text}")
                walk(child, depth + 1, limit)
            except Exception as e:
                print("  " * depth + f"[{i}] <err> {e}")
    walk(popup)

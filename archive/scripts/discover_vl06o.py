"""
VL06O 초기 화면 구조 탐색.
VL06O 진입 직후 화면(아직 아무것도 누르기 전)에서 실행.
"""
import win32com.client, time

sap  = win32com.client.GetObject("SAPGUI").GetScriptingEngine
sess = sap.Children(0).Children(0)  # 세션0

sess.findById("wnd[0]/tbar[0]/okcd").text = "/nVL06O"
sess.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("=== 화면 제목:", sess.findById("wnd[0]").Text)
print()

def dump(node, depth=0):
    try:
        ntype = node.Type
        nid   = node.Id.split("/")[-1]   # 마지막 부분만
        fullid = node.Id
        try:    text = repr(node.Text)
        except: text = ""
        try:    tooltip = repr(node.Tooltip)
        except: tooltip = ""

        if ntype in ("GuiTextField","GuiCTextField","GuiLabel"):
            print("  "*depth + f"[{ntype}] {fullid!r}  text={text}")
        elif ntype in ("GuiButton","GuiMenubar"):
            print("  "*depth + f"[{ntype}] {fullid!r}  tooltip={tooltip}  text={text}")

        for i in range(node.Children.Count):
            dump(node.Children(i), depth+1)
    except:
        pass

dump(sess.findById("wnd[0]"))

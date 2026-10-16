"""
ZRMA_Q 초기 화면 구조 탐색 + Shift+F5 (Get Variants) 팝업 구조.
ZRMA_Q 진입 직후 화면에서 실행.
"""
import win32com.client, time

sap  = win32com.client.GetObject("SAPGUI").GetScriptingEngine
sess = sap.Children(0).Children(2)  # 세션2

sess.findById("wnd[0]/tbar[0]/okcd").text = "/nZRMA_Q"
sess.findById("wnd[0]").sendVKey(0)
time.sleep(2)

print("=== 화면 제목:", sess.findById("wnd[0]").Text)
print()

def dump(node, depth=0):
    try:
        ntype = node.Type
        fullid = node.Id
        try:    text = repr(node.Text)
        except: text = ""
        try:    tooltip = repr(node.Tooltip)
        except: tooltip = ""

        if ntype in ("GuiTextField","GuiCTextField","GuiLabel"):
            print("  "*depth + f"[{ntype}] {fullid!r}  text={text}")
        elif ntype in ("GuiButton",):
            print("  "*depth + f"[{ntype}] {fullid!r}  tooltip={tooltip}  text={text}")

        for i in range(node.Children.Count):
            dump(node.Children(i), depth+1)
    except:
        pass

dump(sess.findById("wnd[0]"))

# Shift+F5 (vkey 41) 눌러서 Variants 팝업 열기
print("\n=== Shift+F5 (Get Variants) 팝업 열기...")
sess.findById("wnd[0]").sendVKey(41)
time.sleep(1.5)

try:
    wnd1 = sess.findById("wnd[1]")
    print("팝업 제목:", wnd1.Text)
    dump(wnd1)
except Exception as e:
    print("팝업 없음:", e)

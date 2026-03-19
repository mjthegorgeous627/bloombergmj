"""
VL10G 선택 화면의 필드 ID 탐색.
VL10G 선택 화면(F8 누르기 전)이 열려 있는 상태에서 실행.
"""
import win32com.client

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
sess = sap.Children(0).Children(1)  # 세션1 (VL10G)

# 현재 화면으로 VL10G 이동
sess.findById("wnd[0]/tbar[0]/okcd").text = "/nVL10G"
sess.findById("wnd[0]").sendVKey(0)

import time
time.sleep(2)

print("=== 현재 화면:", sess.findById("wnd[0]").Text)

def dump(node, depth=0):
    try:
        ctype = node.Type
        cid   = node.Id
        try:
            ctext = node.Text
        except:
            ctext = ""
        if any(k in ctype for k in ("txt", "ctxt", "Field", "Edit")):
            print("  " * depth + f"[{ctype}] {cid!r}  text={ctext!r}")
        for i in range(node.Children.Count):
            dump(node.Children(i), depth + 1)
    except:
        pass

dump(sess.findById("wnd[0]/usr"))

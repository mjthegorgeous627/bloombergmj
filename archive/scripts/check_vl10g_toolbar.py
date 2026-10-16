import win32com.client
sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(1)
tb = session.findById("wnd[0]/tbar[1]")
for i in range(tb.Children.Count):
    try:
        b = tb.Children.ElementAt(i)
        print(f'[{i}] id={b.Id.split("/")[-1]} text="{getattr(b,"Text","")}" tip="{getattr(b,"Tooltip","")}"')
    except:
        pass

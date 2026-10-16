import win32com.client

GRID = "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell"
sap = win32com.client.GetObject("SAPGUI")
app = sap.GetScriptingEngine
conn = app.Children(0)

for sess_idx in [2, 3]:
    sess = conn.Children(sess_idx)
    print(f"\n=== session {sess_idx} ===")
    try:
        grid = sess.findById(GRID)
    except Exception as e:
        print("grid not found", e)
        continue
    print("id", grid.Id, "type", grid.Type, "rows", getattr(grid, "RowCount", None))
    cols = []
    for i in range(30):
        try:
            col = grid.ColumnOrder(i)
        except Exception:
            try:
                col = grid.ColumnOrder[i]
            except Exception:
                break
        cols.append(col)
    print("cols", cols)
    for r in range(min(10, grid.RowCount)):
        vals = []
        for c in cols or ["LAYOUT", "TEXT", "VARIANT", "DESCRIPT"]:
            try:
                vals.append(f"{c}={grid.GetCellValue(r, c)!r}")
            except Exception:
                pass
        print(r, " | ".join(vals))

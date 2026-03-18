"""ZRMA_Q 그리드 컬럼명 및 샘플 데이터 확인용 일회성 스크립트."""
import win32com.client

sap = win32com.client.GetObject("SAPGUI").GetScriptingEngine
session = sap.Children(0).Children(2)  # 세션2 (RMA List)

grid = session.findById("wnd[0]/usr/cntlCUST_CONT/shellcont/shell")
print(f"행 수: {grid.RowCount}")
cols = list(grid.ColumnOrder)
print(f"컬럼 목록: {cols}")

if grid.RowCount > 0:
    print("\n[0행 데이터 샘플]")
    for col in cols:
        try:
            val = grid.GetCellValue(0, col)
            print(f"  {col} = '{val}'")
        except Exception:
            pass

import os as _os
from datetime import date as _date


def _resolve_excel_path(folder=r"C:\1\배송장"):
    """올해 배송장 파일 경로. 올해 파일이 아직 없으면(1월 초 등) 가장 최근 연도 파일로 대체."""
    for y in range(_date.today().year, _date.today().year - 4, -1):
        p = _os.path.join(folder, f"{y} 배송장.xlsx")
        if _os.path.exists(p):
            return p
    return _os.path.join(folder, f"{_date.today().year} 배송장.xlsx")


EXCEL_PATH = _resolve_excel_path()
# 2026-10-08: 별도 "마감 파일"(FINALIZE_EXPORT_PATH)을 폐지하고 EXCEL_PATH 하나로 통합했다.
# 마감 버튼도, Dashboard/이력 검색/퀵 금액 정리도 모두 이 파일을 쓴다.
PROCESSED_ORDERS_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\processed_orders.json"
LOG_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\automation.log"
SESSION_MAP_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\session_map.json"
SAP_VARIANT = "KSCPs1"
REFRESH_INTERVAL_MINUTES = 20

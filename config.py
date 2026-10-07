EXCEL_PATH = r"C:\1\배송장\2026 배송장.xlsx"
# workbench의 "마감 (내보내기 + 삭제)" 전용 내보내기 대상 - 매일 SAP 수집이
# 쓰는 EXCEL_PATH와는 별개의 파일. 날짜별 시트로 쌓는 방식은 동일, 단 실제
# 배송장 파일을 건드리지 않는다 (사용자 요청, 2026-08-14).
# 2026-09-01: 폴더가 backup_2→backup로, 파일명이 ..._backup.xlsx→..._마감.xlsx로
# 사용자에 의해 정리됨 - 경로를 맞춰 갱신(안 그러면 마감 버튼이 존재하지 않는
# 옛 backup_2 경로를 찾다 실패함).
FINALIZE_EXPORT_PATH = r"C:\1\배송장\backup\2026 배송장_마감.xlsx"
PROCESSED_ORDERS_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\processed_orders.json"
LOG_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\automation.log"
SESSION_MAP_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\session_map.json"
SAP_VARIANT = "KSCPs1"
REFRESH_INTERVAL_MINUTES = 20

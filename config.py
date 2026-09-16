EXCEL_PATH = r"C:\1\배송장\2026 배송장.xlsx"
# workbench의 "마감 (내보내기 + 삭제)" 전용 내보내기 대상 - 매일 SAP 수집이
# 쓰는 EXCEL_PATH와는 별개의 파일. 날짜별 시트로 쌓는 방식은 동일, 단 실제
# 배송장 파일을 건드리지 않는다 (사용자 요청, 2026-08-14).
# 2026-09-01: 폴더가 backup_2→backup로, 파일명이 ..._backup.xlsx→..._마감.xlsx로
# 사용자에 의해 정리됨 - 경로를 맞춰 갱신(안 그러면 마감 버튼이 존재하지 않는
# 옛 backup_2 경로를 찾다 실패함).
FINALIZE_EXPORT_PATH = r"C:\1\배송장\backup\2026 배송장_마감.xlsx"
# board_sync.py("동기화" 버튼)의 실제 쓰기 대상 - EXCEL_PATH(다른 사람들과
# 공유하는 실제 배송장 파일)는 에러 없이 깔끔해야 하는데, workbench.db 기반
# 동기화는 아직 확정 전 상태를 담을 수 있어 거기 섞으면 안 된다는 사용자 지적
# (2026-09-01)으로 EXCEL_PATH와 완전히 분리된 별도 파일로 이전. 구조(날짜별
# 배너가 쌓이는 한 시트)는 EXCEL_PATH와 동일 - board_sync.py만 이 경로를 쓰고,
# write_orders_to_excel() 등 매일 SAP 수집 경로는 여전히 EXCEL_PATH 그대로 씀.
REALTIME_SYNC_PATH = r"C:\1\배송장\backup\2026 배송장_실시간 동기화.xlsx"
PROCESSED_ORDERS_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\processed_orders.json"
LOG_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\automation.log"
SESSION_MAP_FILE = r"C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\session_map.json"
SAP_VARIANT = "KSCPs1"
REFRESH_INTERVAL_MINUTES = 20

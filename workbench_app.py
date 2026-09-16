"""Local operations workbench for SAP/Bloomberg delivery automation.

This app moves day-to-day work state out of the shipping Excel file and into a
small SQLite database. Excel is treated as an import source, not the runtime
source of truth for portal execution.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import win32api
import win32event
import winerror

from db import *
from excel_sync import *
from items import *
from sap_ops import *
from sap_ops import _sap_loop_watchdog_thread
from portal_ops import *
def render_notes_page():
    """2026-08-31 사용자 요청: 오더/날짜와 무관한 잡메모(링크, 전화번호, 계정정보
    등)를 관리할 곳이 없어서 신설. 처음엔 항목별 카드+카테고리 UI로 만들었는데,
    실제로 옮기려는 내용(고객센터 번호/SAP 계정/배송기한 규칙 등 엑셀에 박스로
    적어둔 15개+ 잡다한 블록)이 "카테고리 하나 + 한 줄" 같은 정형 항목이 아니라
    사용자가 힘들다고 함 - 항목 단위 add/카테고리 입력 자체가 안 맞는 모델이었음.
    그래서 카드 모델을 버리고 탭(페이지) 여러 개 + 탭마다 큰 textarea 하나인
    메모장 방식으로 교체 - 엑셀 내용을 그대로 복사/붙여넣기만 하면 되게."""
    pages_json = json.dumps(list_pages(), ensure_ascii=False)
    return NOTES_PAGE_TEMPLATE.replace("__PAGES__", pages_json)


def render_mobile_page():
    """2026-08-26 사용자 요청: 회사 밖에서 폰으로 workbench를 볼 때(원격
    데스크톱 화면을 확대해서 보는 게 아니라) 터치에 맞는 화면이 필요해서 추가.
    데스크톱용 PAGE_TEMPLATE(복잡한 표/컨트롤 패널 전부)을 반응형으로 늘리는
    대신 완전히 별도의 가벼운 페이지로 분리했다 - 기존 데스크톱 화면을 건드릴
    위험 없이, SAP 루프 상태 확인 + 세션별 즉시 조회 버튼 + 오늘/이후 오더를
    카드로 보는 것까지만 딱 필요한 만큼만 담는다. 실제 SAP 화면 스크린샷은
    안 보여준다(사용자 확인: workbench가 긁어온 오더 데이터면 충분) - 이미
    board_data()가 보여주는 게 SAP에서 방금 긁어온 그 데이터 그대로다."""
    data_json = json.dumps(board_data(), ensure_ascii=False)
    return MOBILE_PAGE_TEMPLATE.replace("__BOARD_DATA__", data_json)


def render_dashboard_page():
    data_json = json.dumps(board_data(), ensure_ascii=False)
    custom_types_json = json.dumps(list_custom_types(), ensure_ascii=False)
    nonworking_json = json.dumps(nonworking_days_payload(), ensure_ascii=False)
    return (
        PAGE_TEMPLATE
        .replace("__BOARD_DATA__", data_json)
        .replace("__CUSTOM_TYPES__", custom_types_json)
        .replace("__NONWORKING_DAYS__", nonworking_json)
        .replace("__DELETED_ITEM_RETENTION_DAYS__", str(DELETED_ITEM_RETENTION_DAYS))
    )


MOBILE_PAGE_TEMPLATE = (TEMPLATES_DIR / "mobile.html").read_text(encoding="utf-8")


NOTES_PAGE_TEMPLATE = (TEMPLATES_DIR / "notes.html").read_text(encoding="utf-8")



PAGE_TEMPLATE = (TEMPLATES_DIR / "dashboard.html").read_text(encoding="utf-8")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        with WORKBENCH_ACCESS_LOG.open("a", encoding="utf-8") as fh:
            fh.write(now_text() + " [Workbench] " + (fmt % args) + "\n")

    def send_json(self, data, status=200):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def read_json(self):
        length = int(self.headers.get("Content-Length", "0") or 0)
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_GET(self):
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/":
                payload = render_dashboard_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/mobile":
                payload = render_mobile_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/notes":
                payload = render_notes_page().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
                return
            if parsed.path == "/api/pages":
                self.send_json({"pages": list_pages()})
                return
            if parsed.path == "/api/orders":
                qs = parse_qs(parsed.query)
                self.send_json({"orders": list_orders(qs.get("q", [""])[0], qs.get("status", [""])[0])})
                return
            if parsed.path == "/api/board_data":
                # 새 오더가 SAP 수집으로 workbench.db에 들어와도 이미 열려있는
                # 브라우저 탭은 새로고침 전엔 모른다 - 클라이언트가 이 엔드포인트를
                # 주기적으로 폴링해서 새 itemId만 감지해 추가한다(사용자 요청,
                # 2026-08-14). board_data()는 page-load 때 __BOARD_DATA__로 쓰는
                # 것과 완전히 같은 함수라 모양이 항상 일치함.
                self.send_json(board_data())
                return
            if parsed.path == "/api/board_version":
                # board_data()의 가벼운 버전 - 클라이언트가 이걸 자주(1~2초)
                # 찔러보고 바뀐 걸 감지했을 때만 진짜 board_data()를 부른다.
                # 2026-09-01, see board_version()'s own docstring.
                self.send_json({"version": board_version()})
                return
            if parsed.path == "/api/sap/status":
                self.send_json(sap_status())
                return
            if parsed.path == "/api/items/deleted":
                self.send_json({"items": list_deleted_items()})
                return
            if parsed.path == "/api/nonworking_days":
                self.send_json(nonworking_days_payload())
                return
            m = re.match(r"^/api/orders/(\d+)$", parsed.path)
            if m:
                data = get_order(int(m.group(1)))
                if not data:
                    self.send_json({"error": "not found"}, 404)
                else:
                    self.send_json(data)
                return
            m = re.match(r"^/api/process/(\d+)/running$", parsed.path)
            if m:
                self.send_json({"running": pid_running(int(m.group(1)))})
                return
            self.send_json({"error": "not found"}, 404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)

    def do_POST(self):
        try:
            parsed = urlparse(self.path)
            if parsed.path == "/api/import/today":
                self.send_json(load_today_from_excel())
                return
            if parsed.path == "/api/import/range":
                body = self.read_json()
                self.send_json(import_excel_range(
                    body.get("sheet") or "", body.get("rowStart") or 1, body.get("rowEnd") or None,
                ))
                return
            if parsed.path == "/api/export":
                body = self.read_json()
                self.send_json(export_dates_to_excel(body.get("dates") or {}))
                return
            if parsed.path == "/api/sync_delivery_excel":
                self.send_json(sync_delivery_excel())
                return
            if parsed.path == "/api/items/delete":
                body = self.read_json()
                self.send_json(delete_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/items/restore":
                body = self.read_json()
                self.send_json(restore_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/nonworking_days/set":
                body = self.read_json()
                self.send_json(set_manual_holiday(body.get("date") or "", bool(body.get("on"))))
                return
            if parsed.path == "/api/items/move_date":
                body = self.read_json()
                self.send_json(move_items_date(body.get("itemIds") or [], body.get("date") or ""))
                return
            if parsed.path == "/api/items/status":
                body = self.read_json()
                self.send_json(set_items_board_status(body.get("itemIds") or [], body.get("status") or ""))
                return
            if parsed.path == "/api/orders/add_manual":
                body = self.read_json()
                self.send_json(add_manual_note(
                    body.get("cells") or [], body.get("date") or "", body.get("itemType") or "",
                    body.get("orderLabel") or "", body.get("cellCount"),
                ))
                return
            if parsed.path == "/api/items/edit_field":
                body = self.read_json()
                self.send_json(update_item_field(body.get("itemId"), body.get("field") or "", body.get("value") or ""))
                return
            if parsed.path == "/api/orders/edit_field":
                body = self.read_json()
                self.send_json(update_order_field(body.get("itemIds") or [], body.get("field") or "", body.get("value") or ""))
                return
            if parsed.path == "/api/orders/edit_delivery_codes":
                body = self.read_json()
                self.send_json(update_delivery_codes(body.get("itemIds") or [], body.get("value") or ""))
                return
            if parsed.path == "/api/items/highlight":
                body = self.read_json()
                self.send_json(set_cell_highlight(body.get("itemId"), body.get("field") or "", body.get("style") or {}))
                return
            if parsed.path == "/api/items/reorder":
                body = self.read_json()
                self.send_json(reorder_items(body.get("itemIds") or []))
                return
            if parsed.path == "/api/items/merge":
                body = self.read_json()
                self.send_json(merge_orders(body.get("itemIds") or [], body.get("fields") or []))
                return
            if parsed.path == "/api/items/unmerge":
                body = self.read_json()
                self.send_json(unmerge_orders(body.get("itemIds") or [], body.get("fields") or []))
                return
            if parsed.path == "/api/undo":
                self.send_json(undo_last_action())
                return
            if parsed.path == "/api/sap/start_loop":
                self.send_json(sap_start_loop())
                return
            if parsed.path == "/api/sap/stop_loop":
                self.send_json(sap_stop_loop())
                return
            if parsed.path == "/api/workbench/restart":
                self.send_json(workbench_force_restart())
                return
            if parsed.path == "/api/custom_types/add":
                body = self.read_json()
                self.send_json(add_custom_type(body.get("name") or ""))
                return
            if parsed.path == "/api/pages/add":
                body = self.read_json()
                self.send_json(add_page(body.get("title") or ""))
                return
            if parsed.path == "/api/pages/edit":
                body = self.read_json()
                self.send_json(update_page(body.get("id"), body.get("field") or "", body.get("value") if body.get("value") is not None else ""))
                return
            if parsed.path == "/api/pages/delete":
                body = self.read_json()
                self.send_json(delete_page(body.get("id")))
                return
            if parsed.path == "/api/sap/morning_routine":
                self.send_json(sap_run_morning_routine())
                return
            if parsed.path == "/api/sap/run_all":
                self.send_json(sap_run_all())
                return
            if parsed.path == "/api/sap/run_session":
                body = self.read_json()
                self.send_json(sap_run_session(int(body.get("session", -1))))
                return
            if parsed.path == "/api/sap/run_order":
                body = self.read_json()
                self.send_json(sap_run_order(body.get("orderNo") or ""))
                return
            if parsed.path == "/api/sap/open_session":
                body = self.read_json()
                self.send_json(sap_open_session(body.get("tcode") or ""))
                return
            if parsed.path == "/api/sap/open_order":
                body = self.read_json()
                self.send_json(sap_open_order(body.get("orderNo") or ""))
                return
            if parsed.path == "/api/portal/login_manual":
                self.send_json(portal_login_manual())
                return
            if parsed.path == "/api/portal/print_latest_zpl":
                self.send_json(print_latest_zpl())
                return
            m = re.match(r"^/api/orders/(\d+)/run_portal$", parsed.path)
            if m:
                body = self.read_json()
                self.send_json(run_portal(int(m.group(1)), skip_qr=bool(body.get("skipQr"))))
                return
            if parsed.path == "/api/sn_relo/fill_one":
                body = self.read_json()
                self.send_json(sn_relo_one(body.get("itemId")))
                return
            if parsed.path == "/api/zrec/lookup_batch":
                body = self.read_json()
                self.send_json(zrec_lookup_batch([int(i) for i in (body.get("itemIds") or [])]))
                return
            if parsed.path == "/api/zrec/commit_one":
                body = self.read_json()
                self.send_json(zrec_commit_one(body.get("itemId"), body.get("orderNo") or ""))
                return
            if parsed.path == "/api/zrec/verify_batch":
                body = self.read_json()
                self.send_json(zrec_verify_batch([int(i) for i in (body.get("itemIds") or [])]))
                return
            if parsed.path == "/api/pod/preview":
                body = self.read_json()
                self.send_json(pod_preview(
                    [int(i) for i in (body.get("orderIds") or [])],
                    body.get("signedBy") or "", body.get("datetime") or "",
                    body.get("remarks") or "", bool(body.get("pickupDone")),
                    body.get("itemIds") or [],
                ))
                return
            if parsed.path == "/api/pod/update":
                body = self.read_json()
                self.send_json(run_pod_update(
                    [int(i) for i in (body.get("orderIds") or [])],
                    body.get("signedBy") or "", body.get("datetime") or "",
                    body.get("remarks") or "", bool(body.get("pickupDone")),
                    body.get("itemIds") or [],
                ))
                return
            m = re.match(r"^/api/orders/(\d+)/status$", parsed.path)
            if m:
                body = self.read_json()
                status = body.get("status", "new")
                with connect() as con:
                    con.execute("UPDATE orders SET status=?, updated_at=? WHERE id=?", (status, now_text(), int(m.group(1))))
                    add_event(con, int(m.group(1)), f"Status changed to {status}")
                self.send_json({"ok": True})
                return
            self.send_json({"error": "not found"}, 404)
        except Exception as exc:
            self.send_json({"error": str(exc)}, 500)


# 2026-08-28: workbench_app.py 서버 자체는(SAP 루프와 달리, sap_start_loop()
# 참고) 여태 중복 실행을 전혀 막지 않았다 - workbench.bat을 두 번 실행하면
# 그냥 두 서버가 동시에 뜬다. 실측으로 확인한 결과 Python HTTPServer의
# allow_reuse_address=True 기본값 + Windows의 느슨한 SO_REUSEADDR 처리
# 때문에 두 번째 프로세스도 같은 포트(8765)에 조용히 bind까지 성공해서
# "에러도 없이 어느 쪽이 응답하는지 알 수 없는" 상태가 된다. 실제로
# 사용자가 "workbench가 안 보인다"며 workbench.bat을 손으로 한 번 더
# 실행해서 이 상태가 실제로 재현됨(자세한 경위는 세션 기록 참고).
# 포트/프로세스 목록 확인은 그 자체로 확인-후-실행 사이 경합(race)이
# 남으므로, 원자적인 Win32 named mutex로 "이 프로세스가 유일한 서버인지"를
# 판정한다.
_SINGLE_INSTANCE_MUTEX_NAME = "Global\\MJSuh_Workbench_App_SingleInstance"
_single_instance_mutex_handle = None  # GC/커널 핸들 해제로 락이 풀리지 않도록 계속 참조 유지


def _acquire_single_instance_lock():
    """뮤텍스 획득에 성공하면 True(이 프로세스가 유일한 서버), 이미 다른
    workbench_app.py 서버가 뮤텍스를 쥐고 있으면 False."""
    global _single_instance_mutex_handle
    handle = win32event.CreateMutex(None, False, _SINGLE_INSTANCE_MUTEX_NAME)
    already_running = (win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS)
    if already_running:
        return False
    _single_instance_mutex_handle = handle  # 참조 유지 (프로세스 종료 시 자동 해제됨)
    return True


def serve(host, port, open_browser=True):
    url = f"http://{host}:{port}/"
    if not _acquire_single_instance_lock():
        # 이미 다른 workbench_app.py 서버가 떠 있음 - 새 서버를 띄우는 대신
        # 기존 서버의 브라우저 탭만 열어준다(수동 실행/워치독/아침 루틴
        # 어느 경로로 여기 왔든 동일하게 안전한 동작).
        logger.info(f"workbench_app.py 이미 실행 중인 인스턴스가 있음 - 새 서버 없이 브라우저만 엶: {url}")
        if open_browser:
            webbrowser.open(url)
        return
    # 위 뮤텍스가 최종 방어선이지만, 혹시 뮤텍스가 뚫리는 경우에도(예: 다른
    # 사용자 세션) 포트가 이미 쓰이고 있으면 조용히 이중 bind되지 않고
    # 시끄럽게(예외로) 실패하도록 재사용 허용을 꺼둔다.
    ThreadingHTTPServer.allow_reuse_address = False
    ensure_db()
    threading.Thread(target=_sap_loop_watchdog_thread, daemon=True).start()
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Workbench running: {url}", flush=True)
    logger.info(f"Workbench 서버 시작: {url} (PID {os.getpid()})")
    # 2026-08-25: workbench.bat/워치독 재시작 모두 pythonw(콘솔 창 없음)로
    # 바뀌면서, "떠 있나"를 눈으로 확인할 창이 사라졌다 - 그 대신 매 시작마다
    # 토스트 한 번씩 띄워서 최소한의 확인 수단을 남겨둔다. 죽었다가 워치독이
    # 재시작하는 비정상 케이스는 workbench_watchdog_check.py가 재시작 '직전'에
    # 별도의 더 눈에 띄는 토스트를 먼저 띄우므로, 이건 항상 뒤이어 오는 평범한
    # "정상 기동" 확인용 - 실패해도(윈도우 알림 API 자체 오류 등) 서버 구동을
    # 막으면 안 되므로 예외는 무조건 삼킨다.
    try:
        from win_notify import send_windows_notification
        send_windows_notification("Workbench 시작됨", url, duration="short")
    except Exception:
        logger.warning("시작 토스트 알림 실패(무시)", exc_info=True)
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except Exception:
        logger.exception("serve_forever()가 처리되지 않은 예외로 종료됨 - 이게 workbench가 죽은 실제 원인")
        raise
    finally:
        # serve_forever()는 정상 상황에서 절대 반환되지 않는다 - 여기 도달했다는
        # 것 자체가 곧 workbench 프로세스 종료라는 뜻이므로 항상 남긴다.
        logger.warning("Workbench 서버 종료(serve_forever 반환/예외)")


def main():
    parser = argparse.ArgumentParser(description="MJSuh local operations workbench")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-open", action="store_true")
    # Manual/emergency-only from here on - see load_today_from_excel()'s own
    # docstring for why. The routine SAP->workbench path is --import-rows-json.
    parser.add_argument("--import-today", action="store_true")
    parser.add_argument(
        "--import-excel-range", nargs=3, metavar=("SHEET", "START_ROW", "END_ROW"),
        help="한 시트의 특정 행 범위만 1회성으로 불러오기 (비상용) - 예: --import-excel-range 8-6 10 25",
    )
    # The routine path: main.py/manual_order_handler.py call this right after
    # writing a freshly-collected order to Excel, passing the exact same rows
    # as a JSON file (avoids a huge argv string) - see import_sap_rows().
    parser.add_argument("--import-rows-json", metavar="PATH")
    args = parser.parse_args()
    ensure_db()
    if args.import_today:
        print(json.dumps(load_today_from_excel(), ensure_ascii=False, indent=2))
        return
    if args.import_excel_range:
        sheet, start_row, end_row = args.import_excel_range
        print(json.dumps(import_excel_range(sheet, start_row, end_row), ensure_ascii=False, indent=2))
        return
    if args.import_rows_json:
        with open(args.import_rows_json, "r", encoding="utf-8") as fh:
            rows = json.load(fh)
        print(json.dumps(import_sap_rows(rows), ensure_ascii=False, indent=2))
        return
    serve(args.host, args.port, open_browser=not args.no_open)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except BaseException:
        logger.exception("workbench_app.py 최상위에서 처리되지 않은 예외로 프로세스 종료")
        raise

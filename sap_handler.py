"""
SAP GUI Scripting - VL06O 배송 오더 데이터 추출 모듈.
확인된 element ID 기준으로 작성됨.
"""

import win32com.client
import subprocess
import threading
import time
import re
import json
import logging

from config import SESSION_MAP_FILE

logger = logging.getLogger(__name__)

# 오더번호 패턴: ZOR/SOR/ZINX/ZINP/ZINT/SDSK/ZRX/ZRE/ORD + 숫자 (앞의 00 제거)
ORDER_PATTERN = re.compile(
    r'\b(ZOR|SOR|ZINX|ZINP|ZINT|SDSK|ZRX|ZRE|ORD#?)\s*:?\s*(00)?(\d{6,})\b',
    re.IGNORECASE
)

# 확인된 element ID
GRID_PATH        = "wnd[0]/usr/cntlGRID1/shellcont/shell"
ADDR_BTN         = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/btnBT_WADR_T"
MENU_TEXTS       = "wnd[0]/mbar/menu[2]/menu[1]/menu[9]"   # Goto > Header > Texts (VL06O)
MENU_ENV_SHIP    = "wnd[0]/mbar/menu[4]/menu[0]"           # Environment > Ship-To Party

# ZRMA_Q (VA03 스타일) 메뉴 - VL06O와 인덱스 다름
ZRMA_MENU_PARTNERS = "wnd[0]/mbar/menu[2]/menu[1]/menu[9]"   # Goto > Header > Partners
ZRMA_MENU_TEXTS    = "wnd[0]/mbar/menu[2]/menu[1]/menu[10]"  # Goto > Header > Texts

# ZRMA Texts 탭 텍스트 에디터 경로 (GuiSplitterShell 하위 shellcont[1])
ZRMA_TEXT_EDITOR = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4152"
    "/subSUBSCREEN_TEXT:SAPLV70T:2100"
    "/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
)
ITEM_TABLE       = "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER"


def get_sap_gui_auto():
    """
    SAP GUI Scripting COM 객체 반환.
    - 클래식 SAP Logon 실행 방식: "SAPGUI" 모니커
    - NWBC(NetWeaver Business Client) 실행 방식: "SAPGUISERVER" 모니커
    두 방식 모두 시도해 어느 쪽으로 SAP를 띄우든 동작하게 한다.
    """
    last_exc = None
    for moniker in ("SAPGUI", "SAPGUISERVER"):
        try:
            return win32com.client.GetObject(moniker)
        except Exception as e:
            last_exc = e
    raise ConnectionError(f"SAP GUI Scripting 연결 실패 (SAPGUI/SAPGUISERVER 모두 실패): {last_exc}")


_cached_engine = None


def get_scripting_engine(force_refresh=False):
    """
    SAP GUI Scripting Engine(app) 객체 반환.

    NWBC(SAPGUISERVER)는 같은 프로세스에서 GetObject()를 두 번째로 호출하면
    반환되는 객체가 깨지는(속성이 함수 취급되는) 문제가 있어, 프로세스당 한 번만
    연결하고 이후에는 캐시된 엔진 객체를 재사용한다.

    추가로 실측 확인된 문제: `sap_gui_auto.GetScriptingEngine`이 property로
    자동 호출되지 않고 바인딩된 python method 객체로 반환되는 경우가 있다
    (`'function' object has no attribute 'Children'` 에러의 실제 원인).
    이 경우 명시적으로 () 호출해야 진짜 엔진 객체를 얻는다.
    """
    global _cached_engine
    if force_refresh or _cached_engine is None:
        engine = get_sap_gui_auto().GetScriptingEngine
        if not hasattr(engine, "Children"):
            engine = engine()
        _cached_engine = engine
    return _cached_engine


def graceful_close_all_sap_connections():
    """
    2026-09-14: evening_shutdown.py/startup.py가 SAP 프론트엔드가 응답 없을 때
    NWBC/SapGuiServer 프로세스를 Stop-Process -Force로 죽여왔는데, 이건 그냥
    프로세스를 죽이는 것뿐 SAP 백엔드 로그오프가 아니라서 서버 쪽엔 로그온이
    그대로 남는다 - 다음 로그인 시도가 그 유령 세션과 충돌해
    handle_multi_logon_popup()이 처리해야 하는 상황을 반복해서 만드는 근본
    원인이었다(실제로 2026-09-13 저녁 강제종료 이후 남은 유령 세션이
    2026-09-14 낮에 새 로그인과 충돌하는 것을 실측 확인).

    강제종료 직전에 한 번 시도: 스크립팅 엔진에 붙어있는 모든 connection을
    CloseConnection()으로 정상 종료(실제 SAP 로그오프)한다. 이미 죽어서
    스크립팅 엔진 자체가 응답 없는 경우(원래 강제종료가 필요했던 그 상황)엔
    여기서도 예외/블록이 날 수 있으므로, 호출하는 쪽에서 반드시 타임아웃으로
    감싸고 실패해도 그냥 이어서 강제종료하면 된다 - 이 함수는 "되면 좋고
    안 되면 마는" 예방 조치지 강제종료를 대체하지 않는다."""
    try:
        app = get_sap_gui_auto()
        engine = app.GetScriptingEngine
        if not hasattr(engine, "Children"):
            engine = engine()
    except Exception as e:
        logger.info(f"정상 로그오프 생략 (스크립팅 엔진 연결 안 됨): {e}")
        return False

    closed_any = False
    try:
        count = engine.Children.Count
    except Exception as e:
        logger.info(f"정상 로그오프 생략 (연결 목록 조회 실패): {e}")
        return False

    for i in range(count - 1, -1, -1):
        try:
            conn = engine.Children(i)
            for j in range(conn.Children.Count - 1, -1, -1):
                try:
                    handle_multi_logon_popup(conn.Children(j))
                except Exception:
                    pass
            conn.CloseConnection()
            closed_any = True
            logger.info(f"SAP connection[{i}] 정상 로그오프 완료")
        except Exception as e:
            logger.info(f"SAP connection[{i}] 정상 로그오프 실패 (무시): {e}")
    return closed_any


def run_with_timeout(fn, timeout_sec):
    """fn()을 데몬 스레드에서 실행하고 timeout_sec 안에 안 끝나면 (False, None).
    SAP GUI Scripting COM 호출은 자체 타임아웃이 없어서, 원격 NWBC/SapGuiServer가
    죽어있으면(예: 주말 내내 idle이던 세션이 백엔드 타임아웃/VPN 끊김으로
    응답불능이 되는 경우) 그냥 영원히 블록된다 - 예외를 던지는 게 아니라
    "응답 없음"이라 try/except로는 절대 못 잡는다.

    2026-08-30 실사고: 이 실패 유형이 정확히 이 함수가 감싸는 "기존 세션
    재사용" 체크(startup.py launch_sap_and_connect() 1단계) 안에서 터졌다.
    워치독이 30분마다 python 프로세스를 강제 재시작해도, 새로 뜬 startup.py가
    이 체크에서 또 그 죽은 세션을 발견하고 다시 블록되는 걸 반복 - 그래서
    일요일 오전부터 월요일 아침까지 20번 넘게 재시작해도 하나도 안 풀렸다.
    타임아웃으로 감싸서 "N초 안에 응답 없음"을 "죽음"으로 취급해야 그
    자리에서 실제로 죽은 프로세스를 찾아 죽이고 새로 시작할 수 있다.

    타임아웃이 나도 이 스레드 자체를 강제로 죽일 방법은 없다(daemon=True라
    최소한 프로세스 종료는 안 막음) - 대부분은 이 함수를 부른 직후
    kill_stuck_sap_gui_processes()가 그 죽은 NWBC/SapGuiServer를 실제로
    죽여버리므로, 블록돼 있던 COM 호출도 곧 에러로 풀려나며 스레드가
    알아서 끝난다.

    2026-09-17: startup.py 전용이었던 걸 sap_handler.py로 옮김 - workbench의
    자체 워치독(sap_ops._watchdog_restart_loop, "automation.log 30분 이상
    무갱신" 감지 시 startup.py를 재시작하는 그 경로)이 python 프로세스만
    taskkill하고 진짜 멈춰있던 NWBC/SapGuiServer는 그대로 남겨둬서, 재시작된
    startup.py의 새 로그인 시도가 그 좀비와 충돌해 "License Information for
    Multiple Logons" 팝업에 다시 걸리는 게 반복됐다(2026-09-17 실측 스크린샷).
    startup.py만 쓰던 이 정리 로직을 워치독도 같이 쓸 수 있게 공용 모듈로
    옮겼다."""
    result = {}

    def runner():
        try:
            result["value"] = fn()
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=runner, daemon=True)
    t.start()
    t.join(timeout_sec)
    if t.is_alive():
        return False, None
    if "error" in result:
        raise result["error"]
    return True, result.get("value")


def kill_stuck_sap_gui_processes():
    """실제 SAP 프론트엔드 프로세스(NWBC/NwbcProcessAgent/SapGuiServer,
    구형 saplogon 포함)를 강제 종료. python.exe는 절대 안 건드림 - 워치독의
    기존 taskkill(main.py/startup.py 대상)은 이 프로세스들을 전혀 안
    건드려서, 진짜 죽은 SAP GUI는 그대로 남아있고 재시작마다 그 좀비를
    다시 붙잡는 게 2026-08-30 사고의 근본 원인이었다. 죽일 게 없어도
    안전하게 아무 일도 안 함.

    2026-09-14: 강제종료(Stop-Process) 직전에 정상 로그오프를 먼저 시도한다
    (타임아웃 8초로 감싸서, 이미 완전히 죽어 응답 없는 경우엔 그냥 넘어가고
    바로 강제종료로 진행 - graceful_close_all_sap_connections() 자체가
    막혀있는 스크립팅 엔진을 부를 수 있어서 run_with_timeout 없이 직접
    부르면 여기서도 영원히 블록될 수 있음). 로그오프 없이 그냥 죽이면 SAP
    백엔드에 로그온이 남아 다음 로그인 시도가 "License Information for
    Multiple Logons" 팝업과 충돌하는 게 근본 원인이었다."""
    try:
        run_with_timeout(graceful_close_all_sap_connections, timeout_sec=8)
    except Exception as e:
        logger.info(f"정상 로그오프 시도 중 예외 (무시하고 강제종료로 진행): {e}")

    ps_cmd = (
        "Get-Process -Name NWBC,NwbcProcessAgent,SapGuiServer,saplogon "
        "-ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=15,
        )
        logger.info("응답 없는 SAP GUI 프로세스(NWBC/SapGuiServer 등) 강제 종료")
    except Exception as e:
        logger.warning(f"SAP GUI 프로세스 강제 종료 실패 (무시하고 계속 진행): {e}")
    time.sleep(2)


def handle_multi_logon_popup(sess, end_other_logons=True):
    """
    2026-09-14 실측: 전날 저녁 세션을 로그오프 없이 강제종료(evening_shutdown.py
    등의 Stop-Process -Force)하면 SAP 백엔드에는 그 유저의 로그온이 그대로
    남아있고, 다음 새 로그온 시도가 "License Information for Multiple Logons"
    팝업(wnd[1])과 충돌한다. 이 팝업은 일반 GuiMessageBox가 아니라 라디오버튼
    2개짜리 모달이라 F12로 안 닫히고, 아무도 안 눌러주면 그 세션이 로그인
    화면에 영구히 멈춰선 채로 새 SAP 창만 계속 늘어나는 형태로 사용자에게
    보였다(중복 에러창).

    end_other_logons=True(기본)면 "Continue with this logon and end any
    other logons in the system"를 선택해 오래된 유령 세션을 정리하고 이번
    로그온을 계속 진행한다 - 무인 자동화 맥락에서 유령 세션을 남에게 쓰이는
    실제 작업으로 오인해 방치하는 것보다 안전하다고 판단(2026-09-14 실측
    당시 문제의 유령 세션도 실제로는 아무도 안 쓰던 좀비였음). False면
    "Terminate this logon"으로 이번 새 로그온만 취소한다.

    라디오버튼/확인버튼을 하드코딩 ID(radMULTI_LOGON_OPT1 등) 대신 화면
    문구로 찾는다 - SAP 버전에 따라 ID가 바뀔 수 있어도 문구는 안정적.
    팝업이 없거나 이 팝업이 아니면 조용히 False.
    """
    try:
        wnd1 = sess.findById("wnd[1]")
    except Exception:
        return False

    try:
        if "Multiple Logon" not in (wnd1.Text or ""):
            return False
    except Exception:
        return False

    logger.warning(f"다중 로그온(License Information for Multiple Logons) 팝업 감지 "
                    f"→ {'기존 세션 종료 후 계속' if end_other_logons else '이번 로그온 취소'}")

    target_phrase = "end any other logons" if end_other_logons else "terminate this logon"

    def _find_radio(el):
        try:
            kids = el.Children
        except Exception:
            return None
        for k in range(kids.Count):
            c = kids(k)
            try:
                if c.Type == "GuiRadioButton" and target_phrase in (c.Text or "").lower():
                    return c
            except Exception:
                pass
            found = _find_radio(c)
            if found is not None:
                return found
        return None

    try:
        radio = _find_radio(wnd1)
        if radio is not None:
            radio.Selected = True
        else:
            logger.warning("다중 로그온 팝업 - 라디오버튼을 문구로 못 찾음, 기본 선택값으로 진행")
        wnd1.sendVKey(0)  # Enter = 확인(초록 체크)
        time.sleep(1)
        return True
    except Exception as e:
        logger.warning(f"다중 로그온 팝업 처리 실패: {e}")
        return False


def save_session_map(mapping):
    """
    역할(0~3) → SAP가 부여한 SessionNumber 매핑을 파일에 저장.
    setup_sap_sessions() 직후, 즉 위치[0..3]가 확실히 VL06O/VL10G/ZRMA RLKR/ZRMA
    Q2 순서인 시점에 호출한다. NWBC 탭이 나중에 중간에 삽입되어 위치가 밀려도
    이 매핑으로 원래 세션을 다시 찾을 수 있다.
    """
    try:
        with open(SESSION_MAP_FILE, "w", encoding="utf-8") as f:
            json.dump(mapping, f)
        logger.info(f"세션 매핑 저장: {mapping}")
    except Exception as e:
        logger.warning(f"세션 매핑 저장 실패: {e}")


def _load_session_map():
    try:
        with open(SESSION_MAP_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _find_session_by_map(conn, session_idx):
    """저장된 SessionNumber로 세션을 찾는다. 매핑이 없거나 못 찾으면 None."""
    mapping = _load_session_map()
    if not mapping:
        return None
    target_num = mapping.get(str(session_idx))
    if target_num is None:
        return None
    try:
        count = conn.Children.Count
    except Exception:
        return None
    for i in range(count):
        try:
            sess = conn.Children(i)
            if sess.Info.SessionNumber == target_num:
                return sess
        except Exception:
            continue
    return None


def get_sap_session(session_idx=0, retries=1):
    """
    실행 중인 SAP GUI 세션에 연결.
    session_idx: 0=VL06O, 1=VL10G, 2=ZRMA_Q RLKR, 3=ZRMA_Q Q2

    session_map.json에 기록된 SessionNumber로 먼저 찾는다 (NWBC 탭이 중간에
    삽입/삭제되어 위치가 밀려도 안전). 매핑이 없거나 실패하면 기존처럼
    위치(conn.Children(session_idx))로 폴백한다.

    실패 시 짧게 대기 후, 캐시된 엔진을 버리고 새로 연결해 한 번 재시도한다.
    """
    last_exc = None
    for attempt in range(retries + 1):
        try:
            app = get_scripting_engine(force_refresh=(attempt > 0))
            conn = app.Children(0)
            session = _find_session_by_map(conn, session_idx)
            if session is None:
                session = conn.Children(session_idx)
            logger.info(f"SAP 연결 성공 [세션{session_idx}]: {session.findById('wnd[0]').Text}")
            return session
        except Exception as e:
            last_exc = e
            if attempt < retries:
                time.sleep(1.5)
    raise ConnectionError(f"SAP GUI 연결 실패 (세션 {session_idx}): {last_exc}")


def run_transaction(session, tcode):
    """T-code 실행."""
    session.findById("wnd[0]/tbar[0]/okcd").text = tcode
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)


def is_on_list_screen(session):
    """현재 VL06O 오더 목록 또는 선택 화면인지 확인."""
    try:
        title = session.findById("wnd[0]").Text
        return ("List of Outbound Deliveries" in title
                or "General Delivery List" in title)
    except Exception:
        return False


def refresh_sap_list(session, grid_path, label="SAP", try_f5=True, toolbar_button_id="&REFRESH"):
    """
    현재 목록 화면을 유지한 채 새로고침한다.
    ALV Grid 툴바 Refresh를 우선 사용하고, 실패하면 F5로 fallback.

    NWBC 8.00에서는 ALV 그리드가 "&REFRESH" 버튼을 스크립팅에 노출하지 않는 경우가
    있다. VL06O/VL10G 그리드는 툴바 버튼이 아예 0개로 노출된다 (실측:
    grid.ToolbarButtonCount == 0). ZRMA_Q 그리드는 버튼은 있지만 새로고침 버튼의
    실제 id가 "&REFRESH"가 아니라 "REF"다 (grid.GetToolbarButtonId()로 실측 확인,
    tooltip="Refresh"). toolbar_button_id로 화면별 실제 id를 넘긴다.
    버튼이 0개인 경우 pressToolbarButton()은 항상 COM 예외를 던지므로, 먼저 버튼
    수를 읽어 0이면 시도 자체를 건너뛴다 (매 사이클 반복되는 이 예외가
    SapGuiServer.exe를 불안정하게 만드는 것으로 의심되어, 확실히 실패할 호출은
    최소화한다).
    """
    grid = session.findById(grid_path)

    has_toolbar_buttons = True
    try:
        has_toolbar_buttons = grid.ToolbarButtonCount > 0
    except Exception:
        pass  # 버튼 수를 알 수 없으면 기존처럼 시도

    if has_toolbar_buttons:
        try:
            grid.pressToolbarButton(toolbar_button_id)
            time.sleep(2)
            logger.info(f"{label} 목록 새로고침 완료 (ALV Refresh)")
            return True
        except Exception as e:
            logger.warning(f"{label} ALV Refresh 실패 → F5 시도: {e}")
    else:
        logger.info(f"{label} ALV 툴바에 버튼 없음(NWBC) → ALV Refresh 생략")

    if not try_f5:
        return False

    try:
        session.findById("wnd[0]").sendVKey(5)  # F5 = Refresh
        time.sleep(2)
        logger.info(f"{label} 목록 새로고침 완료 (F5)")
        return True
    except Exception as e:
        logger.warning(f"{label} F5 새로고침 실패: {e}")
        return False


def navigate_to_vl06o_list(session):
    """
    VL06O → Variant KSCPs1 → F8 실행 → 오더 목록 화면 진입.
    이미 목록 화면이면 F5(새로고침)만 실행.
    """
    if is_on_list_screen(session):
        title = session.findById("wnd[0]").Text
        if "General Delivery List" in title:
            # 선택화면 → F8로 실행
            logger.info("선택화면 → F8 실행")
            session.findById("wnd[0]").sendVKey(8)
            time.sleep(2)
        else:
            logger.info("VL06O list screen; refreshing current list")
            if refresh_sap_list(session, GRID_PATH, "VL06O"):
                return
            logger.warning("VL06O refresh failed; reopening selection")

    logger.info("VL06O 진입 중...")
    run_transaction(session, "/nVL06O")
    time.sleep(1.5)

    # Variant 필드가 숨겨져 있으면 "Display Variants" 버튼으로 표시
    try:
        session.findById("wnd[0]/usr/ctxtLF_SVAR6")
    except Exception:
        session.findById("wnd[0]/tbar[1]/btn[31]").press()
        time.sleep(0.5)

    # "List Outbound Deliveries" 행의 Variant 필드에 KSCPs1 입력 후 버튼 클릭
    session.findById("wnd[0]/usr/ctxtLF_SVAR6").text = "KSCPs1"
    session.findById("wnd[0]/usr/btnBUTTON6").press()
    time.sleep(2)

    # 선택화면(General Delivery List)에서 F8로 실행
    session.findById("wnd[0]").sendVKey(8)
    time.sleep(3)
    logger.info("목록 화면 진입 완료")


def get_all_rows_from_list(session):
    """
    VL06O 목록 그리드에서 전체 행 데이터 읽기.
    반환: [{'ebeln': '7763549', 'vbeln': '92086300', 'matnr': '10045196',
             'arktx': 'KEYBOARD...', 'lfimg': '2', 'name_we': 'KOREA UNIV'}, ...]
    """
    rows = []
    try:
        grid = session.findById(GRID_PATH)
        row_count = grid.RowCount
        logger.info(f"그리드 행 수: {row_count}")

        for i in range(row_count):
            try:
                row = {
                    'grid_idx': i,                                       # 실제 그리드 행 번호 (더블클릭용)
                    'ebeln':   grid.GetCellValue(i, "EBELN").strip(),   # Pur.Doc (오더번호)
                    'vbeln':   grid.GetCellValue(i, "VBELN").strip(),   # 배송 문서번호
                    'matnr':   grid.GetCellValue(i, "MATNR").strip(),   # Material
                    'arktx':   grid.GetCellValue(i, "ARKTX").strip(),   # Description
                    'lfimg':   grid.GetCellValue(i, "LFIMG").strip(),   # Delivery Qty
                    'name_we': grid.GetCellValue(i, "NAME_WE").strip(), # Ship-to 회사명
                }
                if row['ebeln']:
                    rows.append(row)
            except Exception as e:
                logger.warning(f"행 {i} 읽기 오류: {e}")
                continue

    except Exception as e:
        logger.error(f"그리드 읽기 실패: {e}")

    return rows


def group_by_order(rows):
    """
    그리드 행을 EBELN(오더번호)별로 그룹화.
    row_idx: 그리드에서 해당 오더의 첫 번째 행 인덱스 (더블클릭 진입에 사용)
    반환: {'7763549': {'vbeln': '92086300', 'row_idx': 0, 'items': [...], 'name_we': '...'}, ...}
    """
    orders = {}
    for row in rows:
        ebeln = row['ebeln']
        if ebeln not in orders:
            orders[ebeln] = {
                'vbeln':   row['vbeln'],
                'name_we': row['name_we'],
                'row_idx': row['grid_idx'],  # 실제 그리드 행 인덱스 (enumerate 아님)
                'items':   [],
            }
        orders[ebeln]['items'].append({
            'vbeln': row['vbeln'],
            'matnr': row['matnr'],
            'arktx': row['arktx'],
            'lfimg': row['lfimg'],
        })
    return orders


def navigate_to_delivery(session, row_idx, vbeln):
    """
    VL06O 목록 그리드에서 해당 행을 더블클릭하여 배송 상세 화면 진입.
    VL03N을 별도로 열지 않음.
    """
    logger.info(f"배송 문서 {vbeln} 진입 (행 {row_idx} 더블클릭)...")
    try:
        grid = session.findById(GRID_PATH)
        grid.setCurrentCell(row_idx, "VBELN")
        grid.doubleClickCurrentCell()
        time.sleep(2)

        title = session.findById("wnd[0]").Text
        delivery_titles = ("Outbound delivery", "Outbound Delivery", "Replenishment Dlv.")
        if any(token in title for token in delivery_titles) and str(vbeln) in title:
            logger.info(f"상세 화면 진입 성공: {title}")
            return True
        else:
            logger.error(f"예상치 못한 화면: {title}")
            try:
                session.findById("wnd[0]").sendVKey(3)
                time.sleep(1)
            except Exception:
                pass
            return False
    except Exception as e:
        logger.error(f"더블클릭 진입 실패: {e}")
        return False


def get_text_content(session, menu_id=None):
    """
    Goto → Header → Texts 진입 후 텍스트 내용 추출.
    menu_id: None이면 MENU_TEXTS(VL06O) 사용. ZRMA는 ZRMA_MENU_TEXTS 전달.
    반환: 텍스트 문자열 (없으면 '')
    """
    text_content = ""
    try:
        session.findById(menu_id or MENU_TEXTS).select()
        time.sleep(1)

        # 텍스트 화면에서 내용 읽기 시도 (여러 방식)
        editor_ids = [
            "wnd[0]/usr/cntlSFTXT_0100_EDITOR/shellcont/shell",
            "wnd[0]/usr/txtTDLINES-TDLINE",
        ]
        if menu_id == ZRMA_MENU_TEXTS:
            editor_ids = [ZRMA_TEXT_EDITOR] + editor_ids

        for editor_id in editor_ids:
            try:
                elem = session.findById(editor_id)
                val = getattr(elem, 'Value', '') or getattr(elem, 'Text', '')
                if val and str(val).strip():
                    # \r\n → \n 정규화 (SAP GuiShell Value에 \r 포함됨)
                    text_content = str(val).replace('\r\n', '\n').replace('\r', '\n').strip()
                    break
            except Exception:
                continue

        # 텍스트 화면에서 테이블 방식으로 읽기 (fallback)
        if not text_content:
            try:
                table = session.findById("wnd[0]/usr/tblSAPDOCU_LINES")
                lines = []
                for i in range(table.RowCount):
                    try:
                        line = table.GetCell(i, 0).Text.strip()
                        if line:
                            lines.append(line)
                    except Exception:
                        continue
                text_content = '\n'.join(lines)
            except Exception:
                pass

        # 뒤로가기 (F3)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)

    except Exception as e:
        logger.warning(f"Text 접근 실패 (내용 없을 수 있음): {e}")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass

    return text_content.strip()


def parse_extra_orders(text):
    """
    텍스트에서 오더번호 패턴 추출.
    반환: ['SDSK 001320768583', 'ZOR 7763549', ...]  중복 없이, 앞의 00 제거
    """
    found = []
    seen = set()
    for match in ORDER_PATTERN.finditer(text):
        prefix = match.group(1).upper()
        num    = match.group(3)  # 앞의 00은 group(2)에서 이미 분리됨
        entry  = f"{prefix} {num}"
        if entry not in seen:
            seen.add(entry)
            found.append(entry)
    return found


def add_extra_order_dedup(extra_orders, label, number):
    """
    extra_orders(오더번호 텍스트 스캔으로 opportunistically 얻은 리스트, 예:
    parse_extra_orders() 결과)에 collect_order_extra_fields()로 직접 읽은
    ORD/SDSK 번호를 추가한다. 같은 숫자가 이미 스캔 결과에 들어있으면
    (라벨이 다르더라도) 추가하지 않는다 - 두 출처가 같은 번호를 가리키는
    실제 사고가 있었음 (2026-08-14, ZRX 67076714의 Order# 셀에 "SDSK
    1332695640"이 두 줄로 중복 기록됨: 기존 자유 텍스트 스캔이 이미 잡아낸
    번호를, VA02 PO Number 필드에서 직접 읽은 값이 다시 붙였기 때문).
    """
    number = str(number or '').strip()
    if not number:
        return list(extra_orders)
    digits = number.lstrip('0') or '0'
    for existing in extra_orders:
        existing_digits = re.sub(r'\D', '', existing).lstrip('0') or '0'
        if existing_digits == digits:
            return list(extra_orders)
    return list(extra_orders) + [f"{label} {number}"]


def normalize_phone(raw):
    """
    전화번호 정규화.
    - 82로 시작 → 한국 번호 형식 변환
    - 81로 시작 → [해외-일본] 표기
    - 기타 국제번호 → [해외] 표기
    - 이미 한국 번호 형식이면 그대로
    """
    if not raw:
        return ''
    digits = re.sub(r'\D', '', raw)
    if not digits:
        return raw.strip()

    if digits.startswith('82'):
        local = '0' + digits[2:]
        return _format_korean(local)
    elif digits.startswith('81'):
        return f'[해외-일본] {raw.strip()}'
    elif len(digits) >= 10 and not digits.startswith('0'):
        # 기타 국제번호 (0으로 시작하지 않는 긴 번호)
        return f'[해외] {raw.strip()}'
    else:
        return _format_korean(digits)


def _format_korean(digits):
    """한국 번호 형식으로 변환 (앞에 0 포함된 숫자열)."""
    if digits.startswith('010') and len(digits) == 11:
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'
    elif digits.startswith('02'):
        rest = digits[2:]
        if len(rest) == 7:
            return f'02-{rest[:3]}-{rest[3:]}'
        elif len(rest) == 8:
            return f'02-{rest[:4]}-{rest[4:]}'
        else:
            return f'02-{rest}'
    elif len(digits) == 11:
        return f'{digits[:3]}-{digits[3:7]}-{digits[7:]}'
    elif len(digits) == 10:
        return f'{digits[:3]}-{digits[3:6]}-{digits[6:]}'
    return digits


def parse_memo_for_display(text):
    """
    SAP Text → H열 메모용 정보 추출.
    - 오더번호 참조 줄 제거 (ZRX/SDSK/ZOR 등으로 시작하는 줄)
    - Caller Phone, Caller E-Mail 제거 (이미 별도 컬럼에 있음)
    - Delivery Note, Caller Name, 날짜/시간/배송 지시사항 등 유지
    """
    if not text:
        return ''

    # [Text] 접두사 제거
    text = re.sub(r'^\[Text\]\s*', '', text.strip())

    ORDER_REF_RE = re.compile(
        r'^(ZRX|ZRE|ZOR|SOR|ZINX|ZINP|ZINT|SDSK|ORD#?)\s+\d+',
        re.IGNORECASE
    )
    SKIP_RE = re.compile(
        r'^(Caller\s+Phone|Caller\s+E-?Mail)\s*[:\-]',
        re.IGNORECASE
    )

    result = []
    prev_blank = False
    for line in text.split('\n'):
        s = line.strip()
        if not s:
            if result and not prev_blank:
                result.append('')
            prev_blank = True
            continue
        prev_blank = False
        if ORDER_REF_RE.match(s):
            continue
        if SKIP_RE.match(s):
            continue
        result.append(s)

    # 앞뒤 빈 줄 제거
    while result and result[0] == '':
        result.pop(0)
    while result and result[-1] == '':
        result.pop()

    return '\n'.join(result)


def parse_contact_from_text(text):
    """
    텍스트에서 전화번호 등 연락처 추출.
    'Caller Phone:'/'Contact Number:' 등 명시 라벨 우선, 일반 번호 패턴 fallback.
    반환: {'phone': '', 'found': False}
    """
    result = {'phone': '', 'found': False}

    # 1차: "Caller Phone: +82-2-767-5886" / "Contact Number: +82232714865" 등
    # 명시 라벨 패턴 (오더 67043332에서 확인: ZRE/ZRX 텍스트는 "Contact Number:"
    # 라벨을 쓰는데 "Caller Phone:"만 인식하다 보니 2차 fallback으로 새서
    # 아래의 일반 번호 패턴이 텍스트 속 오더번호(0으로 패딩된 참조번호)를
    # 전화번호로 잘못 집어내는 사고가 있었음 - 라벨 매칭 폭 확대로 해결)
    caller_re = re.compile(
        r'(?:Caller\s+Phone|Contact\s+(?:Number|Phone))\s*[:\-]\s*([+\d][\d\-.\s]{6,})',
        re.IGNORECASE
    )
    m = caller_re.search(text)
    if m:
        result['phone'] = normalize_phone(m.group(1).strip())
        result['found'] = True
        return result

    # 2차: 일반 번호 패턴 (+82-2-xxx-xxxx, 010-xxxx-xxxx 등) - 단, 오더번호
    # 참조 줄(ZRX/ZRE/ZOR/SDSK 등, 0-패딩되어 전화번호처럼 보일 수 있음)은
    # 먼저 제거하고 검색해 위 사고 유형의 오탐을 줄인다. 그래도 라벨 없이
    # 우연히 숫자 패턴만 맞는 다른 참조번호가 남아있을 수 있어, 실제
    # 한국 지역/휴대폰 국번으로 시작하는 후보만 채택한다 (오더 67043332
    # 사고: 국번이 존재하지 않는 "067..."이 전화번호로 잘못 채택됐었음).
    scan_text = ORDER_PATTERN.sub(' ', text)
    phone_re = re.compile(
        r'(\+?82[-.\s]?|0)[1-9]\d?[-.\s]?\d{3,4}[-.\s]?\d{4}'
    )
    for m in phone_re.finditer(scan_text):
        candidate = m.group().strip()
        digits = re.sub(r'\D', '', candidate)
        if _valid_kr_phone_digits(digits):
            result['phone'] = normalize_phone(candidate)
            result['found'] = True
            break
    return result


_KR_PHONE_PREFIXES = (
    "02",
    "031", "032", "033",
    "041", "042", "043", "044",
    "051", "052", "053", "054", "055",
    "061", "062", "063", "064", "070",
    "010", "011", "016", "017", "018", "019",
)


def _valid_kr_phone_digits(digits):
    """digits(숫자만)가 실제 한국 지역/휴대폰 국번으로 시작하고 길이가
    그럴듯한 경우에만 True. 라벨 없는 일반 번호 패턴 매칭이 오더/문서
    번호 같은 임의의 숫자열을 전화번호로 오인하지 않도록 하는 필터."""
    if not digits:
        return False
    local = digits
    if local.startswith('82'):
        local = '0' + local[2:]
    if not local.startswith('0') or len(local) not in (9, 10, 11):
        return False
    return local.startswith(_KR_PHONE_PREFIXES)


def get_ship_to_address(session):
    """
    Ship-to Party 주소 팝업 읽기.
    1차: ADDR_BTN 버튼 클릭 → 팝업
    2차: Goto > Header > Partners → WE 행 더블클릭 → 팝업
    3차: 헤더 한줄 주소 (txtRV50A-TXTWE) fallback
    반환: {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}
    """
    result = {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}

    def read_field(*ids):
        for fid in ids:
            try:
                val = session.findById(fid).text.strip()
                if val:
                    return val
            except Exception:
                continue
        return ''

    def read_addr_popup():
        """팝업(wnd[1])에서 SAPLSZA1 주소 필드 읽기. (discover_sap.py로 확인된 실제 ID)"""
        base = "wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301"
        phone_raw = read_field(
            f"{base}/txtSZA1_D0100-TEL_NUMBER",
            f"{base}/txtSZA1_D0100-MOB_NUMBER",
        )
        r = {
            'company':  read_field(f"{base}/txtADDR1_DATA-NAME1"),
            'customer': read_field(f"{base}/txtADDR1_DATA-NAME2"),
            'street':   read_field(f"{base}/txtADDR1_DATA-STR_SUPPL1"),
            'street2':  read_field(f"{base}/txtADDR1_DATA-STR_SUPPL2"),
            'phone':    normalize_phone(phone_raw),
        }
        logger.info(f"팝업 주소 읽기 결과: {r}")
        return r

    def close_popup():
        try:
            session.findById("wnd[1]").sendVKey(12)  # F12 닫기
            time.sleep(0.5)
        except Exception:
            pass

    # 1차 시도: Overview 화면 헤더의 Ship-to Party 주소 버튼 (ADDR_BTN)
    try:
        session.findById(ADDR_BTN).press()
        time.sleep(1.5)
        session.findById("wnd[1]")  # 팝업 존재 확인
        r = read_addr_popup()
        close_popup()
        if any(r.values()):
            return r
        logger.warning("ADDR_BTN 팝업 필드 모두 빈값")
    except Exception as e:
        logger.warning(f"ADDR_BTN 시도 실패: {e}")

    # 2차 시도: Goto > Header > Partners → WE(Ship-to) 행 더블클릭
    try:
        session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[8]").select()
        time.sleep(1)

        # Partners 화면: 테이블에서 WE(Ship-to Party) 행 찾아 더블클릭
        partner_table_ids = [
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\08/ssubSUBSCREEN_BODY:SAPMV50A:1204/tblSAPMV50ATC_VBPA_HEAD_OVER",
            "wnd[0]/usr/tblSAPMV50ATC_VBPA_HEAD_OVER",
        ]
        table = None
        for tid in partner_table_ids:
            try:
                table = session.findById(tid)
                break
            except Exception:
                continue

        we_row = -1
        if table:
            for row_i in range(table.RowCount):
                try:
                    parvw = table.getCellValue(row_i, "PARVW").strip()
                    if parvw == "WE":
                        we_row = row_i
                        break
                except Exception:
                    # 컬럼 이름이 다를 수 있으므로 셀 텍스트로 시도
                    try:
                        cell_val = table.GetCell(row_i, 0).Text.strip()
                        if cell_val == "WE":
                            we_row = row_i
                            break
                    except Exception:
                        continue

        if we_row >= 0 and table:
            try:
                table.setCurrentCell(we_row, "KUNNR")
                table.doubleClickCurrentCell()
                time.sleep(1.5)
                session.findById("wnd[1]")  # 팝업 열렸는지 확인
                r = read_addr_popup()
                close_popup()
                # Partners 화면에서 Overview로 복귀
                session.findById("wnd[0]").sendVKey(3)
                time.sleep(0.5)
                if any(r.values()):
                    return r
            except Exception as e:
                logger.warning(f"Partners WE 더블클릭 실패: {e}")
                try:
                    session.findById("wnd[0]").sendVKey(3)
                    time.sleep(0.5)
                except Exception:
                    pass
        else:
            logger.warning(f"Partners 테이블에서 WE 행 못 찾음 (table={table}, we_row={we_row})")
            try:
                session.findById("wnd[0]").sendVKey(3)
                time.sleep(0.5)
            except Exception:
                pass

    except Exception as e:
        logger.warning(f"Goto>Header>Partners 시도 실패: {e}")
        try:
            session.findById("wnd[0]").sendVKey(3)
        except Exception:
            pass

    # 3차 fallback: 헤더의 한줄 주소 텍스트 (불완전하지만 없는 것보다 나음)
    try:
        oneline = session.findById(
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV50A:1502/txtRV50A-TXTWE"
        ).text.strip()
        if oneline:
            result['company'] = oneline
            logger.warning(f"주소 fallback(한줄): {oneline}")
    except Exception:
        pass

    return result


_ZRMA_PARTNER_TABLE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
    "/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)
_ZRMA_PARTNER_SUB = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\07"
    "/ssubSUBSCREEN_BODY:SAPMV45A:4352"
    "/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000"
)


def get_ship_to_address_zrma(session, retries=1):
    """
    ZRMA_Q 오더에서 Ship-to Party 주소 읽기.
    Goto > Header > Partners → Ship-to 행 포커스 → btnDETAIL 클릭
    → SAPLSZA1 팝업에서 company/customer/street/street2/phone 읽기.

    2026-08-07(오더 67043332)에서 확인된 문제: btnDETAIL을 눌러도 SAP 응답이
    느리면 주소 팝업(wnd[1])이 고정된 1.5초 대기 안에 뜨지 않아, 이후 모든
    필드 읽기가 rf()의 except로 조용히 빈 문자열이 되고 "ZRMA 주소 읽기:
    {all blank}"로 마치 정상 조회인 것처럼 로그가 남았다. wnd[1]이 실제로
    뜰 때까지 짧게 폴링하고, 그래도 못 뜨거나 결과가 전부 빈 값이면 전체
    시퀀스(메뉴 선택 → 행 포커스 → Detail 클릭)를 한 번 더 재시도한다.
    """
    result = {'company': '', 'customer': '', 'street': '', 'street2': '', 'phone': ''}

    for attempt in range(retries + 1):
        try:
            session.findById(ZRMA_MENU_PARTNERS).select()
            time.sleep(1)

            table = session.findById(_ZRMA_PARTNER_TABLE)

            # Ship-to 행 찾기
            ship_to_row = -1
            for row_i in range(min(table.RowCount, 10)):
                try:
                    parvw = table.GetCell(row_i, 0).Text.strip()
                    if 'Ship' in parvw:
                        ship_to_row = row_i
                        break
                except Exception:
                    break

            if ship_to_row < 0:
                logger.warning("ZRMA: Ship-to 행 못 찾음")
                session.findById("wnd[0]").sendVKey(3)
                return result

            # Ship-to 행 포커스 → btnDETAIL → SAPLSZA1 팝업
            table.GetCell(ship_to_row, 1).setFocus()
            time.sleep(0.5)
            session.findById(_ZRMA_PARTNER_SUB + "/btnDETAIL").press()

            # 팝업(wnd[1])이 실제로 뜰 때까지 최대 3초 폴링 (고정 sleep 대신)
            popup_ready = False
            for _ in range(6):
                time.sleep(0.5)
                try:
                    session.findById("wnd[1]")
                    popup_ready = True
                    break
                except Exception:
                    continue

            if not popup_ready:
                raise RuntimeError("주소 팝업(wnd[1])이 열리지 않음")

            base = "wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301"

            def rf(*ids):
                for fid in ids:
                    try:
                        v = session.findById(fid).text.strip()
                        if v:
                            return v
                    except Exception:
                        continue
                return ''

            phone_raw = rf(
                f"{base}/txtSZA1_D0100-TEL_NUMBER",
                f"{base}/txtSZA1_D0100-MOB_NUMBER",
            )
            result = {
                'company':  rf(f"{base}/txtADDR1_DATA-NAME1"),
                'customer': rf(f"{base}/txtADDR1_DATA-NAME2"),
                'street':   rf(f"{base}/txtADDR1_DATA-STR_SUPPL1"),
                'street2':  rf(f"{base}/txtADDR1_DATA-STR_SUPPL2"),
                'phone':    normalize_phone(phone_raw),
            }
            logger.info(f"ZRMA 주소 읽기: {result}")

            session.findById("wnd[1]").sendVKey(12)
            time.sleep(0.5)

            if any(result.values()) or attempt >= retries:
                break

            logger.warning(f"ZRMA 주소 전부 빈 값 → 재시도 ({attempt + 1}/{retries})")
            time.sleep(1.0)

        except Exception as e:
            logger.warning(f"ZRMA Partners 접근 실패 (시도 {attempt + 1}/{retries + 1}): {e}")
            try:
                session.findById("wnd[1]").sendVKey(12)
            except Exception:
                pass
            if attempt < retries:
                time.sleep(1.0)
                continue

    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)
    except Exception:
        pass

    return result


def _delivery_order_type(ebeln, extra_orders):
    target = str(ebeln or '').lstrip('0')
    for ref in extra_orders or []:
        text = str(ref or '').upper()
        digits = ''.join(ch for ch in text if ch.isdigit()).lstrip('0')
        if target and digits == target:
            match = re.search(r'\b(ZOR|SOR|ZINX|ZINP|ZINT|SDSK|ZRX|ZRE|ORD#?)\b', text)
            if match:
                return match.group(1).replace('#', '')
    if str(ebeln).startswith(('4', '5')):
        return 'ZINP'
    return 'ZOR'


def _va03_new_session():
    """Opens a brand-new SAP session (never reuses/touches the 4 sessions
    the main automation loop owns) - same technique
    manual_order_handler._open_new_sap_session() already uses for its own
    single-order lookups."""
    app = get_scripting_engine()
    conn = app.Children(0)
    base = conn.Children(0)
    before = conn.Children.Count
    base.createSession()
    deadline = time.time() + 10
    while time.time() < deadline:
        try:
            if conn.Children.Count > before:
                sess = conn.Children(conn.Children.Count - 1)
                return sess
        except Exception:
            pass
        time.sleep(0.5)
    raise RuntimeError("새 SAP 세션 생성 실패 (VA03 추가필드 조회용)")


def _va03_close_session(session):
    try:
        session.findById("wnd[0]").close()
    except Exception:
        pass


def _strip_leading_zeros(value):
    value = (value or "").strip()
    stripped = value.lstrip("0")
    return stripped if stripped else value


def read_order_extra_fields(session):
    """Reads PO Number(VBKD-BSTKD)/Ship-to Party number(KUWEV-KUNNR)/
    Req.Deliv.Date(RV45A-KETDAT) from a VA02/VA03 order screen the caller
    has already navigated to (the initial Overview screen - all 3 sit there,
    no menu navigation needed). Field IDs confirmed live 2026-08-06 against
    real ZOR and ZRX orders - see project_sap_order_collection_gaps memory.
    Read-only lookups (never writes), safe in either display or change mode.
    Any field that isn't found/blank on this particular order/item type just
    comes back '' - never raises, since a missing "nice to have" field must
    not break the actual collection pipeline."""
    result = {"po_number": "", "cust_no": "", "delivery_date": ""}
    try:
        result["po_number"] = _strip_leading_zeros(
            session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text
        )
    except Exception:
        pass
    try:
        result["cust_no"] = session.findById(
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR"
        ).text.strip()
    except Exception:
        pass
    try:
        result["delivery_date"] = session.findById(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4400"
            "/ssubHEADER_FRAME:SAPMV45A:4440/ctxtRV45A-KETDAT"
        ).text.strip()
    except Exception:
        pass
    return result


def collect_order_extra_fields(order_num, order_type, existing_session=None):
    """High-level wrapper: PO Number is labelled ORD for ZOR, SDSK for every
    other order type (user confirmed 2026-08-06 - ZRE/ZRX/ZINX/ZINP/ZINT all
    share the same "PO Number field = SDSK" rule). If existing_session is
    given (the ZRMA_Q path, which already has the order open for other
    reasons), reuses it - no extra session needed. Otherwise (the VL06O/ZOR
    path, which never opens VA03 on its own) opens a throwaway 5th-ish
    session via VA03, reads, and closes it again - deliberately never
    touches the 4 sessions the main loop owns. Never raises - swallows any
    failure and returns empty fields, so a lookup problem here can't break
    the real Excel/kakao pipeline the caller is in the middle of."""
    po_label = "ORD" if order_type == "ZOR" else "SDSK"
    empty = {"po_label": po_label, "po_number": "", "cust_no": "", "delivery_date": ""}
    session = existing_session
    opened_new = False
    try:
        if session is None:
            session = _va03_new_session()
            opened_new = True
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = str(order_num)
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)
        fields = read_order_extra_fields(session)
        return {"po_label": po_label, **fields}
    except Exception as exc:
        logger.warning(f"오더 {order_num} 추가필드(PO Number/Ship-to/Delivery Date) 조회 실패: {exc}")
        return empty
    finally:
        if opened_new and session is not None:
            _va03_close_session(session)


def _compact_delivery_row_qty(qty, item):
    """Large no-serial internal shipments stay as one Excel row with a quantity marker."""
    try:
        qty_int = max(1, int(float(qty)))
    except (ValueError, TypeError):
        qty_int = 1
    if qty_int > 10:
        return [(qty_int, '')]
    return [(1, '') for _ in range(qty_int)]


def build_excel_rows(ebeln, order_info, address, extra_orders, memo):
    """
    하나의 오더에서 Excel 입력용 행 리스트 생성.
    아이템 1개 × 수량 = 행 수
    """
    rows = []
    is_first = True
    order_type = _delivery_order_type(ebeln, extra_orders)

    for item in order_info['items']:
        try:
            qty = max(1, int(float(item['lfimg'])))
        except (ValueError, TypeError):
            qty = 1

        for row_qty, serial_number in _compact_delivery_row_qty(qty, item):
            rows.append({
                'order_prefix': '배송',
                'order_type':   order_type,
                'order_num':    ebeln,
                'obd':          item.get('vbeln') or order_info.get('vbeln', ''),
                'extra_orders': extra_orders,
                'material':     item['matnr'],
                'description':  item['arktx'],
                'quantity':     row_qty,
                'serial_number': serial_number,
                'customer':     address.get('customer', ''),
                'phone':        address.get('phone', ''),
                'company':      address.get('company', ''),
                'street':       address.get('street', ''),
                'street2':      address.get('street2', ''),
                # Ship-to Party's own SAP number (KUWEV-KUNNR, distinct from
                # its address) - rendered as a trailing "(cust# ...)" note at
                # the end of the address text, not a separate column (user's
                # explicit call, 2026-08-06 - see collect_order_extra_fields()).
                'cust_no':      address.get('cust_no', ''),
                'memo':         memo,
                'is_first_item': is_first,
            })
            is_first = False

    return rows


def process_new_orders(session, new_ebelns, order_map):
    """
    새 오더 목록을 처리하여 Excel 입력용 데이터 반환.
    new_ebelns: ['7763549', ...] (7-prefix, 미처리)
    order_map:  group_by_order() 결과
    반환: {ebeln: [row, row, ...], ...}
    """
    results = {}

    for ebeln in new_ebelns:
        logger.info(f"=== 오더 {ebeln} 처리 중 ===")
        order_info = order_map.get(ebeln)
        if not order_info:
            continue

        vbeln   = order_info['vbeln']
        row_idx = order_info['row_idx']

        # 상세 화면 진입 (그리드 행 더블클릭)
        if not navigate_to_delivery(session, row_idx, vbeln):
            logger.error(f"오더 {ebeln} 진입 실패 - 건너뜀")
            continue

        # 1. 주소 먼저 읽기 (팝업 - 화면 이동 없음)
        address = get_ship_to_address(session)

        # 2. Text 읽기 (다른 화면으로 이동했다가 복귀)
        text = get_text_content(session)
        extra_orders = parse_extra_orders(text) if text else []
        text_contact = parse_contact_from_text(text) if text else {'found': False}
        memo = parse_memo_for_display(text) if text else ""

        # text에서 전화번호 찾은 경우 주소에 없으면 보완
        if text_contact['found'] and not address.get('phone'):
            address['phone'] = text_contact['phone']

        # 3. ORD 번호 / Ship-to Party 번호(Cust#) / Delivery Date - VL06O 자체엔
        # 없는 정보라 VA03을 별도 세션으로 열어서 조회 (사용자 확인, 2026-08-06:
        # "ord 넘버 수집하는거 중요하다, va03을 열어서라도 수집해야함"). 4개 운영
        # 세션은 안 건드리고 새 세션 하나 열었다가 바로 닫는다 - 실패해도 나머지
        # 파이프라인엔 영향 없음(collect_order_extra_fields 자체가 절대 raise 안 함).
        order_type_for_lookup = _delivery_order_type(ebeln, extra_orders)
        extra_fields = collect_order_extra_fields(ebeln, order_type_for_lookup)
        if extra_fields.get('cust_no'):
            address['cust_no'] = extra_fields['cust_no']
        if extra_fields.get('po_number'):
            extra_orders = list(extra_orders) + [f"{extra_fields['po_label']} {extra_fields['po_number']}"]
        if extra_fields.get('delivery_date'):
            note = f"(delivery date: {extra_fields['delivery_date']})"
            memo = f"{memo}\n{note}" if memo else note

        # Excel 행 생성
        rows = build_excel_rows(ebeln, order_info, address, extra_orders, memo)
        results[ebeln] = rows

        logger.info(f"오더 {ebeln}: {len(rows)}행 생성")

        # 목록 화면으로 복귀 (F3 Back)
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(1)

    return results

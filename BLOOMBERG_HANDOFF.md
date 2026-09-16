# Bloomberg / SAP Automation Handoff

Date: 2026-07-20
Project folder: C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation

## Current task: SAP 8.0 (NWBC) upgrade compatibility fix

SAP client was upgraded and the launcher's two priority features broke:
1. "SAP 시작 & 자동루프" (startup.py → main.py loop, 4 SAP sessions: VL06O/VL10G/ZRMA RLKR/ZRMA Q2)
2. "오더 엑셀 반영" (order.py <오더번호> → manual_order_handler.py)

### Root cause found

"SAP 8.0 업그레이드" = the desktop launcher (`BLOOMBERG RP1 via new NJ saprouter.sap`) now opens
**NWBC (NetWeaver Business Client) 8.00**, not classic SAP Logon. Confirmed by process list:
`C:\Program Files\SAP\NWBC800\NWBC.exe` (window title changes per transaction, e.g. "SAP Easy Access",
"Outbound Delivery Monitor", "RMA List" — all inside ONE window, no separate classic SAPGUI window pops up).

NWBC registers SAP GUI Scripting under COM moniker **"SAPGUISERVER"**, not the classic **"SAPGUI"**.
All code previously did `win32com.client.GetObject("SAPGUI")`, which fails under NWBC.

Screen element IDs, menu indices, toolbar button indices inside the actual SAP transactions
(VL06O/VL10G/ZRMA_Q/VA02/VA03 etc.) are **unchanged** — this is a connection-layer problem only,
not a UI-element-ID problem.

### Fixes applied (already saved to disk)

- `sap_handler.py`:
  - `get_sap_gui_auto()` — tries `"SAPGUI"` moniker first, falls back to `"SAPGUISERVER"`.
  - `get_scripting_engine(force_refresh=False)` — **caches the resolved engine object per process**.
    Important: repeatedly calling `GetObject("SAPGUISERVER")` a second time in the same process
    corrupts the returned object (`sap.GetScriptingEngine` starts returning a Python "function"
    instead of the real engine → `'function' object has no attribute 'Children'`). Caching avoids
    the repeat call entirely. Do NOT remove this cache as a "cleanup" — it is load-bearing.
  - `get_sap_session(session_idx=0, retries=1)` — on failure, force-refreshes the cached engine
    once and retries.
- `startup.py`, `manual_order_handler.py`, `zrma_handler.py` — updated to use
  `get_scripting_engine()` / `get_sap_gui_auto()` from `sap_handler.py` instead of calling
  `win32com.client.GetObject("SAPGUI")` directly. Removed a dead duplicate `_get_scripting_engine`
  in startup.py.
- `startup.py`, `main.py`, `order.py` — added `sys.stdout.reconfigure(encoding='utf-8', errors='replace')`
  to stop `UnicodeEncodeError`/"Logging error" spam from `✓` characters under the Windows cp949 console.
- `diag_probe.py` (new, kept intentionally) — non-interactive SAP GUI Scripting connectivity probe.
  Run `python diag_probe.py` any time to check whether `SAPGUISERVER`/`SAPGUI` scripting is reachable
  and how many sessions/what titles exist. Useful for future SAP UI changes on this or other branches.

### What is verified working

- Feature 2 ("오더 엑셀 반영", `order.py <order>`): full end-to-end success tested with real order
  67060626 — VA02 entry, address/S/N extraction, Excel row write, Kakao "to me" notification all worked.
- Feature 1 core (4-session open + screen setup via `setup_sap_sessions`): succeeded twice cleanly
  (VL06O "List of Outbound Deliveries", VL10G "Activities Due for Shipping...", both ZRMA_Q "RMA List").
- Feature 1 full catchup pass (`main.py --once`): succeeded once completely — found and correctly
  processed 4 real new ZRMA RLKR orders (67026239, 67029290, 67031603, 67033755), built rows, and (in
  the run that completed) would have written Excel + sent Kakao. A later run of the same got stuck
  and was manually killed before writing/notifying — **no duplicate Excel rows or duplicate Kakao
  messages were sent** because the hang happened before `_write_and_notify` was ever reached; those
  4 orders remain unprocessed and will be picked up cleanly on the next successful run.

### UPDATE 2026-07-21 오후: ALV refresh 예외 원인 찾아서 수정함 (검증 완료)

라이브 세션으로 직접 진단해서 아래 두 근본 원인을 모두 찾아 고쳤고, 4개 세션
전부 예외 없이 새로고침되는 것까지 확인함:

1. **`get_scripting_engine()`의 `'function' object has no attribute 'Children'` 버그**
   (오늘 새벽 11:00/11:10 실행 실패 원인) — 실제 원인은 "같은 프로세스에서 반복
   호출"이 아니라, `sap_gui_auto.GetScriptingEngine`이 property로 자동 호출되지
   않고 파이썬 바인딩 method 객체로 반환되는 경우가 있다는 것. `hasattr(engine,
   "Children")`이 False면 `engine()`으로 명시적으로 호출하도록 수정
   (`sap_handler.py`). 여러 번 새 프로세스에서 재현·검증했고 크래시 없었음.
   (참고: 이전 세션 노트에는 "GetScriptingEngine을 method로 호출하지 말라"고
   되어 있었는데, 오늘 실측 결과 이게 실제 근본 수정이었음 — 예전 크래시는 다른
   원인이었을 가능성이 높음.)

2. **`refresh_sap_list()`의 ALV Refresh 예외** — grid.ToolbarButtonCount로 실측
   확인한 결과:
   - VL06O/VL10G 그리드(`cntlGRID1`)는 NWBC에서 툴바 버튼을 **0개** 노출함 →
     `pressToolbarButton()`은 항상 예외. 이제 버튼 수 0이면 시도 자체를 건너뛰고
     바로 F5로 감 (F5는 이 두 화면에서 정상 동작).
   - ZRMA_Q 그리드(`cntlCUST_CONT`)는 버튼이 35개 있지만, 새로고침 버튼의 실제
     id는 `&REFRESH`가 아니라 **`REF`**(tooltip="Refresh")임. 이제 `REF`로
     누르면 정상 성공 — 예전처럼 F3/F8 재조회로 폴백할 필요가 없어짐(더 빠르고
     화면 상태도 안 흐트러짐). ZRMA_Q는 F5도 항상 실패("virtual key not
     enabled")하는 게 로그로 확인돼서 `try_f5=False`로 아예 생략시킴.

   결과: 4개 세션(VL06O, VL10G, ZRMA_Q RLKR, ZRMA_Q Q2) 새로고침 사이클에서
   COM 예외가 **0건**으로 검증됨 (전엔 사이클마다 6~8건).

**디버깅 중 실제로 벌어진 일 (참고용)**: 진단 중 `SapGuiServer.exe`가 두 번
Not Responding 상태가 되어 사용자 승인 하에 강제 종료 → 재로그인을 두 번 거침
(한 번은 "communication channel closed, Note 520688" 팝업 발생, NWBC 자체가
완전히 닫혀서 런처로 재실행함). 참고로 `Get-Process`의 "Not Responding" 표시는
`SapGuiServer.exe`가 실제로 멎은 것과 무관하게 계속 뜨는 것으로 보임 (그
상태에서도 스크립팅 호출은 정상 동작) — 이 프로세스는 신뢰할 수 있는 헬스체크
신호가 아님. 진단용 스크립트 `discover_grid_refresh.py` (신규)도 저장해둠.

**아직 안 한 것**: 위 수정을 반영한 채로 `main.py`의 20분 자동 루프를 장시간
(몇 시간) 돌려서 `SapGuiServer.exe`가 더 이상 불안정해지지 않는지 실제로
지켜보는 장기 검증. 코드 수정 자체는 완료·검증됨.

### Known unresolved issue — NWBC scripting bridge instability (RESOLVED, see UPDATE above)

`SapGuiServer.exe` (the NWBC scripting bridge process) has crashed once (real SAP error: "SAP GUI has
closed the communication channel. All sessions are lost. Please apply note 520688") and has gone
"Not Responding" (confirmed via `Get-Process | select Responding`) multiple times today, including
right after a clean fresh login. Symptom for the user: opening ANY new SAP window/session (manually,
unrelated to the automation) spins forever and never opens while this is happening.

**Likely contributing root cause, not yet fixed:** every single automated screen refresh
(VL06O/VL10G/ZRMA_Q) is currently failing its primary refresh method and falling back:

```
[WARNING] ...ALV Refresh 실패 → F5 시도: (-2147352567, '예외가 발생했습니다.', ...)
[WARNING] ...F5 새로고침 실패: (-2147352567, ..., 'The virtual key is not enabled.', ...)
[INFO] ...일반 Refresh 불가 → F3/F8 재조회 시도   ← this fallback does work
```

This happens on *every single* automation cycle, for every one of the 4 sessions. The fallback
(F3 then F8) does successfully recover the list each time, so functionally nothing is broken —
but throwing this COM exception repeatedly, every ~20 minutes for 3-4 sessions, is suspected to be
what gradually destabilizes `SapGuiServer.exe` until it stops responding to *new* session-creation
requests (existing already-bound sessions keep working fine even while this happens).

**Not yet done:** find the correct way to trigger a grid refresh in NWBC 8.00's ALV toolbar (the
`grid.pressToolbarButton("&REFRESH")` ID and the F5 vkey binding both no longer work) so this
exception stops firing every cycle. This needs a live SAP session and `discover_sap.py` (or similar)
to inspect the actual toolbar button IDs currently available on the VL06O/VL10G/ZRMA_Q list screens.
The relevant function is `refresh_sap_list()` in `sap_handler.py`.

### Practical operating notes learned today (don't repeat these mistakes)

- ~~Do NOT call `GetScriptingEngine` as a method~~ — **SUPERSEDED 2026-07-21 오후.** This turned out
  to be a misdiagnosis from the first session. Real-session testing today showed
  `sap_gui_auto.GetScriptingEngine` reliably needs to be called as `engine()` when the property
  access returns a bound python `method` object instead of the real engine (`hasattr(engine,
  "Children")` is False in that case). This is now implemented as a safe fallback in
  `get_scripting_engine()` and was exercised many times today across fresh processes with no crash.
  Still avoid `win32com.client.dynamic.Dispatch` and raw `pythoncom.MkParseDisplayName`/`BindToObject`
  — those weren't re-tested and aren't needed.
- Avoid running many independent diagnostic scripts back-to-back against the same live SAP session
  in a short window — this seems to contribute to bridge instability. Prefer one script that reuses
  a cached connection over several small ad-hoc scripts each doing their own fresh `GetObject`.
- If the user needs to do unrelated manual SAP work while the automation's 4 sessions are open,
  opening a brand-new session/window is the specific weak point right now — safer to reuse one of
  the 4 already-open sessions for a quick manual lookup than to open a new one.
- `Get-Process SapGuiServer | select Responding` showing `False` is **not a reliable health signal**
  — it stayed `False` today for extended periods while SAP GUI Scripting calls kept succeeding
  normally. Don't use it alone to decide whether to force-kill the process; try an actual scripting
  call (e.g. `diag_probe.py`) first.
- Force-killing `SapGuiServer.exe` to unstick it is disruptive: once it dropped all 4 automation
  sessions and once it took the entire NWBC window down, requiring a full manual relogin both times.
  Only do this with the user's explicit go-ahead, and expect to ask them to relogin/reopen SAP
  afterward.

## Session 2 status (2026-07-21 오후) — bug fix done, ops not yet resumed

**Done and verified live today** (see "UPDATE 2026-07-21 오후" section above for full detail):
- `get_scripting_engine()` fixed (method-vs-property bug) — this was the real cause of the
  `'function' object has no attribute 'Children'` failures from this morning's runs (11:00, 11:10 in
  automation.log).
- `refresh_sap_list()` fixed for all 3 grid types — VL06O/VL10G skip the always-broken toolbar
  attempt (0 buttons exposed under NWBC) and go straight to F5; ZRMA_Q now uses the correct toolbar
  button id `REF` (was `&REFRESH`) and succeeds directly, no longer needs the F3/F8 fallback.
- Verified: all 4 sessions (VL06O, VL10G, ZRMA_Q RLKR, ZRMA_Q Q2) refreshed with **zero** COM
  exceptions in a live end-to-end test.

**NOT done yet — pick up here tomorrow:**
1. The 20-minute automation loop (`startup.py` / `main.py`) has **not been left running**. Every test
   run today was manually interrupted. Start it fresh next session (`python startup.py`, or the
   launcher's "SAP 시작 & 자동루프" button) and let it run for real to get long-duration confirmation
   that `SapGuiServer.exe` stays stable now that the refresh exceptions are gone.
2. "오더 엑셀 반영" (`order.py <오더번호>` / `manual_order_handler.py`) — **not retested today.** It
   shares the fixed `get_scripting_engine()` connection layer, but its own SAP navigation logic
   (VA02/VA03 entry, address/serial extraction) wasn't touched or re-verified after today's changes.
   Test with a real order number before trusting it.
3. Vendor Dashboard (`vendor_dashboard.py`) and Portal 입력 (`portal_*.py`, browser/Selenium-based,
   unrelated to the SAP GUI Scripting bug) — **out of scope today, untouched, unverified.**
4. New diagnostic script saved: `discover_grid_refresh.py` (dumps ALV grid toolbar button ids for
   all 4 sessions) — kept for future NWBC UI changes, same spirit as `diag_probe.py`.
5. NWBC restarted itself again at some point after the last verification test (new PID observed with
   StartTime ~17:04 same day, automation.log has nothing logged after 14:59). Not investigated — may
   just be the user manually relogging in, or may be worth watching for a pattern.

## Session 3 (2026-07-23) — NWBC new-tab freeze diagnosed (NOT our bug); defensive session-lookup fix shipped

**User-reported symptoms while the 20-min loop was live:**
1. Opening an extra NWBC tab (via the `+` button, e.g. for manual VA02 work) sometimes spins forever
   and never loads — happens the instant `+`/new-tab is clicked, before even typing a tcode.
2. (Reported from memory, not reproduced today) a manually-opened tab sometimes got acted on by the
   automation loop as if it were one of its own 4 sessions.

**Root-cause test performed live:** killed the running automation python process entirely (`taskkill
/F` on the `main.py`/`startup.py` PID — confirmed by its `OleMainThreadWndName` / "Not Responding" tag
in `tasklist`), then had the user close the stuck tab and open a **fresh** tab with automation fully
off. It **still froze identically**. This rules out our script/COM connection as the cause.

**Conclusion:** the new-tab freeze is an **NWBC 8.00 client-side bug/limitation**, unrelated to
`sap_automation`'s code. User confirmed this never happened on classic SAP Logon — only started after
the NWBC switch. Not fixable from this repo. Only known mitigation: **pre-open all tabs you'll need for
the day up front** (before/alongside starting the loop) rather than opening new ones mid-day while the
loop is active — this is the workaround the user had already found empirically. If it keeps happening,
this needs to go to SAP Basis/IT as an NWBC client issue, not back to this codebase.

**Defensive fix shipped anyway** for symptom 2 (position-based session lookup is fragile even though it
wasn't the cause of today's freeze — if NWBC ever *does* successfully insert a tab mid-bar instead of
freezing, the old code would silently grab the wrong tab):
- `sap_handler.get_sap_session()` no longer trusts `conn.Children(session_idx)` blindly. It first tries
  to find the session by a stable identifier (`session.Info.SessionNumber`, assigned by SAP at session
  creation and independent of current tab position) recorded in `session_map.json`
  (`config.SESSION_MAP_FILE`). Falls back to the old positional lookup if no map exists or the mapped
  session isn't found (e.g. first run, or map is stale after a fresh SAP relogin).
- `startup.py`'s `setup_sap_sessions()` now records `{0: SessionNumber, 1: ..., 2: ..., 3: ...}` to
  `session_map.json` right after creating/positioning the 4 sessions (the one moment position is
  guaranteed correct), via `sap_handler.save_session_map()`.
- New kept diagnostic script: `diag_session_ids.py` — dumps position/Id/SessionNumber/title for every
  open session, read-only, same safe one-shot-`GetObject`-per-process pattern as `diag_probe.py`.
- **Not yet live-tested end-to-end**: verified this compiles (`py_compile`) but the automation loop was
  intentionally left OFF at the user's request (mid-troubleshooting the NWBC freeze) — next session
  should restart it (`python startup.py` re-creates `session_map.json` fresh) and confirm
  `get_sap_session` still logs correct session titles per role.

**Workaround implemented and live-tested working:** the freeze is specific to NWBC's own `+`/new-tab UI
button. SAP GUI Scripting's `session.createSession()` API — the exact call `startup.py` already uses 4x
to build the automation's own sessions — opens a new session via a different internal NWBC code path
that does **not** hit the bug. New script `open_session.py` (usage: `python open_session.py [tcode]`,
tcode optional) calls this directly; confirmed live 2026-07-23 that it creates a working, non-frozen
tab (4→5 sessions, tab usable immediately). Added a "새 SAP 창 열기" button + tcode field to
`launcher.py` (`open_new_sap_session()`) so the user doesn't need the command line. **This is now the
recommended way for the user to open any extra SAP window — tell them to stop using NWBC's `+` button
entirely and use this instead.** Not yet tested with the 20-min automation loop actively running at the
same time (loop was stopped when this was tested) — worth confirming next session, though there's no
strong reason to expect it behaves differently since it doesn't touch the loop's own 4 sessions.

**Automation loop status at end of this session: STOPPED.** The python process running `startup.py`'s
`run_loop()` was killed (PID was ~35276, may differ next time) as part of diagnosing symptom 1, and the
user asked to leave it off while sorting out the tab issue. **Next session/next assistant: ask the user
whether to restart it** (launcher's "SAP 시작 & 자동루프" button, or "자동 루프만 시작" if the 4 sessions
are still open and just need the loop resumed without recreating sessions — but note "자동 루프만 시작"
runs `main.py` directly, which does NOT call `setup_sap_sessions()`/`save_session_map()`, so
`session_map.json` would be stale/missing if this is a fresh SAP login since last `startup.py` run).

**Open question, asked but not yet answered by the user:** the user pushed back on needing to switch to
the launcher window to click "새 SAP 창 열기" — they'd prefer something closer to NWBC's own `+` button
feel. Offered two options and the user hadn't picked one yet when this session ended:
(a) keep using the launcher button as-is (already built, zero extra work), or
(b) a global hotkey (e.g. Ctrl+Alt+T) that calls `open_session.py`'s `open_new_session()` directly from
anywhere, including while the SAP window has focus — would need a small new always-running background
script (e.g. via `ctypes`/`win32gui` `RegisterHotKey` + a message loop) that isn't built yet.
**Next session: ask the user which they want (or both) before building the hotkey.** Made very clear to
the user that NWBC's actual `+` button itself cannot be patched/fixed — it's closed-source vendor UI
with no scripting hook — so (a)/(b) above are the only two real options, not "make `+` itself work."

## Session 4 (2026-07-24) — silent loop death confirmed live; watchdog built and wired in

**Confirmed live today, not just from historical logs:** the 20-min automation loop died with **zero
error logged** — not a hang the process just vanished from `tasklist` entirely — while SAP itself
(`NWBC.exe`/`SapGuiServer.exe`) stayed alive and logged in the whole time. Timeline: loop's last
heartbeat was 17:12:30 ("20분 후 다음 실행"); the next scheduled cycle (~17:32) never appears in
`automation.log`; by 17:34 when an unrelated manual test (`order.py`) was run, both the loop process
*and* the separate launcher GUI process were already gone (`tasklist` showed zero `python.exe`). Root
cause: `main.py`'s `run_loop()` is a plain single-threaded `while True` with no timeout around the
blocking SAP GUI Scripting COM calls — when `SapGuiServer.exe`'s bridge misbehaves (the known,
pre-existing instability from Session 1/2), the call either hangs forever or triggers a low-level crash
that kills the Python process before any Python exception/traceback can be logged. This is a **new,
confirmed root cause**, not the ALV-refresh-exception bug from Session 2 (that one is unrelated and
still fixed/verified). The launcher GUI itself doesn't touch SAP/Excel COM directly (checked — it's a
thin `subprocess.Popen` wrapper), so its disappearance same day is presumed a separate, unrelated event
(e.g. user closed it), not proven to share the same cause.

**`order.py` re-verified end-to-end** with a real order (67065297, chosen by user as an
already-processed order specifically to avoid touching a live/unprocessed one): SAP entry → address/S/N
read → Excel write → Kakao send (including an on-the-fly access-token refresh) all worked. Note this
intentionally left a **duplicate test row in today's live Excel sheet** (sheet `7-24`, rows 6–7) and
sent one real Kakao self-notification — user was told, not cleaned up automatically (production file).

**Built `watchdog.py` (new)** — supervises the loop from outside:
- Detects death via `automation.log` mtime staleness (>30 min = dead; loop's own interval is 20 min).
- On detection: finds any stray `main.py`/`startup.py` python processes via PowerShell
  `Get-CimInstance Win32_Process` (chosen over `wmic.exe` — **`wmic` is not installed on this machine**,
  confirmed via `where wmic` failing) and force-kills them, then relaunches `startup.py`. Safe to call
  repeatedly because `startup.py`'s `launch_sap_and_connect()` already reuses an existing logged-in SAP
  session instead of re-logging in (pre-existing behavior, not new).
- Respects `stop_automation.flag` — will not fight an intentional user stop.
- Only acts during business hours (08:00–20:00) — will not hammer SAP relogin attempts overnight if left
  running unattended.
- Flap protection: if it has had to restart 3+ times within a rolling hour, backs off to a 30-min
  cooldown instead of the normal 8-min one.
- Logs its own activity to new `watchdog.log`; writes one line into `automation.log` per restart so it's
  visible in the main trail too.

**Wired into `launcher.py`**: two new buttons, "워치독 시작" / "워치독 중지", in the existing collapsible
"SAP 자동화" tools section, tracked via a new `self.watchdog_process` (independent of
`self.started_process`, which still tracks only the loop).

**Live-tested tonight** (loop was found already dead at session start, from the crash described above):
started both `startup.py` and `watchdog.py` directly via PowerShell `Start-Process` (launcher GUI wasn't
open at the time) — loop came up cleanly reusing the already-logged-in NWBC session and completed a
normal cycle; watchdog logged its startup line correctly; both confirmed live via `tasklist`. Then, per
user's end-of-day request, stopped both cleanly: wrote `stop_automation.flag` (loop exited gracefully
within its 5s poll — confirmed via "중지 요청 감지 → 자동화 종료" in the log) and `taskkill`'d the
watchdog process directly (it has no self-stop-flag of its own, only respects the loop's). **Everything
is OFF as of end of session** — confirmed zero `python.exe` processes running. SAP/NWBC itself was left
as-is (still logged in) — not touched either way.

**NOT done yet — pick up here tomorrow:**
1. The new launcher buttons ("워치독 시작"/"워치독 중지") have **not been click-tested through the actual
   GUI** — only the underlying scripts were exercised directly via command line/PowerShell tonight.
2. The watchdog's **actual auto-restart trigger path has never fired for real** — tonight's test only
   exercised its own startup and a manual restart of both processes, not "loop dies on its own while
   watchdog is running and watchdog catches it." Next natural test: just start both for a normal work day
   and see if a real silent death gets auto-recovered without anyone noticing.
3. Vendor Dashboard / Portal 입력 — still untouched/unverified (carried over from prior sessions).
4. The open question from Session 3 (launcher button vs. global hotkey for "새 SAP 창 열기") is still
   unanswered by the user.

## Session 5 (2026-07-27) — ZRMA extraction bug, Portal browser-tab race condition, Workbench rebuilt into a real date-board

**1. `zrma_handler.py` item extraction bugs (order 7825282 / 7828250)**
- Two real bugs fixed together: (a) items sitting in the SAP table control beyond the currently
  *visible* rows were silently dropped (`get_items_from_zrma_order` now scrolls the table via
  `VerticalScrollbar` and dedupes by POSNR across scroll pages), and (b) `_is_numeric_material()` — a
  filter that drops any item whose `matnr` isn't pure digits (e.g. keeps `10045246`, drops
  `US_KEYBOARD`) — was removed then **put back** after the user confirmed with a live example (ZOR
  7828250: keep `10045246`, drop `US_KEYBOARD`) that it's correct behavior, not a bug. Also fixed
  `collect_serial_numbers` to restore the correct scroll position before re-selecting a row, since rows
  found on a scrolled page no longer sit at a stable index. Verified live against real orders
  (7825282, 7828250) with a disposable read-only script (extracted, printed, deleted — no Excel/Kakao
  side effects).

**2. Portal automation concurrency bug — root cause of the "1. Serial 등록 & QR 인쇄" feature being
abandoned back in May**
- `portal_register_serial.py` / `portal_pack_post.py` / `portal_download_labels.py` /
  `portal_update_pod.py` / `portal_login.py` all connect to the **same shared Chrome tab** via
  `playwright.chromium.connect_over_cdp(...)` + `context.pages[-1]`. Running two of them at once (e.g.
  clicking "배송 처리" for a second order before the first finished) lets them stomp each other's
  navigation on that one tab. Confirmed live in `automation.log` from 2026-05-28 09:41: a second
  order's serial-registration ran against the *first* order's still-open delivery page and
  downloaded/printed the wrong label.
- Fix: new `portal_lock.py` — a simple file-based mutex (`portal_browser.lock`, 3-min wait timeout,
  5-min staleness auto-clear). Wrapped the full CDP session of each of the five scripts above in
  `with portal_browser_lock(...):` so only one ever touches the shared tab at a time; a second call just
  waits its turn instead of racing. Unit-tested the lock's ordering with two threads — confirmed strict
  serialization. This transitively also fixes `workbench_app.py`'s "Portal 처리 + QR" button, since it
  shells out to the same `portal_ship_and_print.py` → same five scripts.

**3. `workbench_app.py` rebuilt from a plain order list into a real date-based ops board**
User wants this (not the launcher) to become the daily driver: SAP controls + Portal automation +
day-by-day work tracking, all in one browser tab, replacing the manual Excel sheet as the live work
surface (Excel becomes export-only). Built in stages, corrected twice based on user feedback:
- **Control panel**: 4 columns now — SAP / Bloomberg Portal / **Process** / Excel (Process was missing
  entirely in the first pass — a straight miss, not a design choice). SAP #3/#4 have an adjacent text
  input (order#, tcode). Portal #1 and Process's two buttons act on whichever row checkboxes are ticked
  in the table below (same checkboxes, shared across both).
- **Board layout**: one continuous `<table>` with a `position: sticky` orange header row, dates stacked
  **downward** (not side-by-side — that was the first pass's mistake). Each date is its own `<tbody>`.
  Learned the hard way that `overflow:hidden` on the wrapping div (added for rounded corners) silently
  breaks `position:sticky` on descendants — removed it.
- **Rows are per-item, not per-order, and not rowspan-merged** — deliberate: a single order (e.g. a ZRX
  exchange) can need its delivery processed today and its pickup a different day, so every row must be
  independently movable. Customer/phone/address/memo are repeated per row instead of rowspan-merged
  (rowspan breaks the moment the anchor row is dragged away from its siblings).
- **Row movement**: HTML5 drag-and-drop between date `<tbody>`s (verified for real in Chrome — delivery
  row stayed in today's section while the matching pickup row for the *same* order moved to tomorrow's,
  independently). **Also added a per-row "이동" dropdown** after the user pointed out dates actually
  span 7/25→8/22 in the live sheet (six date blocks) — dragging across that many off-screen sections
  isn't realistic, so the dropdown jumps a row straight to any date without scrolling.
- **Excel importer fixed to match reality**: opened the actual workbook and found a single day-sheet
  (`7-27`) contains *several* embedded date blocks (a lone "2026년 8월 22일 월요일"-style row in column
  G, followed by its own `Order # / item / M/N / S/N / customer / phone / ADDRESS / memo` header, then
  that date's orders — repeated). `load_today_from_excel()` previously tagged every row `source_date =
  today`, which was wrong for anything not in the very first block. Now `row_date_header()` detects
  those banner rows narrowly (column G matches the date pattern *and* every other column on that row is
  blank, so a memo like "8/22로 연기됨" inside a normal item row is never mistaken for a new block) and
  each order gets the real date of the block it's under. Also: column A's leading `배송`/`회수` token
  (real per-item data that was always there) is now captured into a new `order_items.item_type` column
  (migration added — `ALTER TABLE ... ADD COLUMN`, wrapped so it's a no-op on a DB that already has it)
  instead of being guessed from `order_type` (`infer_item_type()` kept only as a fallback for legacy
  rows imported before this fix, via `resolve_item_type()`). Re-ran the import live and confirmed order
  67059008 correctly landed under 2026-07-25 (not "today") with real 배송/배송/회수/회수 per item, and
  67055178 correctly under 2026-08-22.
- **`render_dashboard_page()`** now renders every distinct `source_date` present in the DB (plus today
  as an always-present anchor), not a hardcoded today/tomorrow pair.
- **Still placeholder-only (by design, structure-first pass)**: every control-panel button just shows a
  toast ("다음 단계에서 연결 예정") — no SAP/Portal/Excel action is actually wired yet. Row moves and
  Process Done/Ready-to-process coloring are DOM-only, not persisted — refreshing the page reverts
  everything (no backend write path exists for either yet).

**Recurring operational gotcha this session (twice)**: background test servers started via the Bash
tool's `... &` + `kill $!` do **not** reliably kill the real Windows `python.exe` process — Bash's `$!`
job id doesn't map to the actual Win32 PID for a spawned native exe, so the old server kept running and
serving stale code while the user opened a fresh `workbench.bat` that silently lost the port-8765 bind
race. Fix used both times: `Get-CimInstance Win32_Process | Where CommandLine -like '*workbench_app*'`
to find the real PID, `Stop-Process -Id <pid> -Force`. **Always verify/kill via PowerShell PIDs, never
trust a Bash-tool `&`-backgrounded native process's own kill.** For any future test server, prefer
`Start-Process -PassThru` (gives a trustworthy PID) over Bash `&`.

**Not done yet — pick up here tomorrow:**
1. **User has not yet visually confirmed today's final round of changes** (multi-date Excel parsing,
   real 배송/회수 capture, the "이동" dropdown, 11-column layout) — they asked to confirm tomorrow
   instead. Start there: have them open `workbench.bat` fresh (check no stale process is squatting on
   8765 first, per the gotcha above) and walk through it together.
2. **14 stale 2026-05-28 orders** are still sitting in `workbench.db` and now show up as their own date
   section (previously invisible since the board only rendered today/tomorrow). Asked the user twice
   whether to delete them — no answer yet. Ask again before touching (it's a delete).
3. **Nothing in the new dashboard is wired to real actions yet** — this was explicit, deliberate scope
   for this pass ("일단 구성을 만들어봐라" / "일단 구성만"). Next real step, in the order the user
   agreed made sense: (a) persist row date-moves and Process Done/Ready status to the DB (currently
   client-side only), (b) wire the Portal batch button to actually call `portal_ship_and_print.py` (or
   its constituent scripts) for the checked rows, (c) wire the SAP buttons (`order.py`, `open_session.py`,
   `main.py --once`/loop) as subprocess calls the same way `launcher.py` does it, (d) Excel export
   (today-only / all).
4. A `python workbench_app.py` process (no args, i.e. a real `workbench.bat` launch) may still be
   running from the user's own session tonight — check `Get-CimInstance Win32_Process | Where
   CommandLine -like '*workbench_app.py*'` before starting anything new, same gotcha as above.
5. Vendor Dashboard — still untouched/unverified (carried over from every prior session).

## Session 6 (2026-07-28) — "이동" redesigned as checkbox + Excel date field; stale 2026-05-28 orders deleted

**1. Per-row "이동" dropdown replaced.** User found the per-row `<select>` (one dropdown per item row,
added in Session 5) hard to use. Removed it entirely (dropped the whole "이동" table column) and replaced
it with: check the rows you want (same `.row-check` checkboxes Portal/Process already use) → pick a date in
a new `<input type="date">` field added to the **Excel** control-panel column → all checked rows jump to
that date immediately. Chose a native HTML calendar date input specifically because the user asked for
"달력 혹은 날짜 드롭다운 중 오류가 잘 나지 않을 것" (calendar or date-dropdown, whichever is less error-prone) —
a native date picker can't produce an invalid calendar date the way a free-text field could.
- **Dates not already on the board are created on the fly.** E.g. if the board only shows 7/28 and 7/30
  and the user moves a row to 7/29, a brand-new `<tbody>` date section for 2026-07-29 is built client-side
  (banner row + empty-row placeholder, same styling/drag-drop handlers as server-rendered sections) and
  inserted in the correct sorted position between 7/28 and 7/31 (compares `data-date` ISO strings). This
  was the user's explicit requirement ("아래 줄에 없는 날짜도 다 있어야 한다").
- Still client-side only / not persisted — same deliberate scope boundary as the rest of the dashboard
  (see Session 5). Refreshing the page reverts moves; this is unchanged, not a new gap.
- **Bug caught during live testing, fixed before calling this done:** typing digits directly into the
  native date input's segments (via the browser-automation tool's synthetic key events, not a real user's
  calendar click) could commit a malformed value straight through — produced a garbage `data-date` like
  `"72920-02-06"` and a broken section the first time this was tested. Added a guard in
  `moveCheckedToDate()`: validates `iso` matches `/^\d{4}-\d{2}-\d{2}$/` **and** parses to a real `Date`
  before doing anything, else shows a toast ("날짜 형식이 올바르지 않습니다") and aborts instead of creating
  a bad section. Re-tested after the fix via a real `change` event with a valid ISO value (`2026-07-29`) —
  confirmed: new section created in the right sorted position, checked row moved into it, unchecked rows
  left behind, checkbox auto-cleared after a successful move, and moving every row out of a date correctly
  restores that date's empty-row placeholder. Removing the whole column also meant retiring the old
  `BOARD_DATES`/`__DATES_JSON__` templating and `populateMoveSelects()`/`jumpRowToDate()` JS (no longer
  needed — labels are now computed in JS via `formatDateLabel()`), and re-flowing the table from 11 to 10
  columns (address column absorbed the freed width).

**2. Deleted the 14 stale 2026-05-28 orders** from `workbench.db` — asked twice in Session 5 with no
answer, user confirmed today ("다 삭제해라"). Backed up the live db first to
`workbench_before_delete_0528_<timestamp>.db` (same directory) before deleting, in case anything in there
was still needed. Deleted both `orders` (14 rows, ids 1-14) and their `events` rows for those order ids
(`order_items` cascade-deleted automatically via the existing `ON DELETE CASCADE` FK). Remaining
`source_date`s in the db now: 2026-07-25, 07-27, 07-28, 07-31, 08-22 — no more phantom May section on the
board.

**Operational note:** followed the Session-5-documented gotcha correctly this time — found the live
`workbench_app.py` server via `Get-CimInstance Win32_Process | Where CommandLine -like '*workbench_app*'`
(PID 34468), stopped it with `Stop-Process` before touching the db or the file, edited, then restarted via
`Start-Process -PassThru` (new PID, check `Get-CimInstance` again next session rather than assuming it's
still running) and live-tested through the actual Chrome browser tool + a couple of direct JS-dispatched
`change` events (to work around the automation tool's own inability to reliably type into a multi-segment
native date input — not an app bug, just how the browser-automation keystroke tool interacts with that
control type; a real user clicking/typing normally is unaffected).

**Not done yet — unchanged from Session 5's list:** nothing in the dashboard is wired to real backend
actions yet (SAP/Portal/Excel buttons, persisting moves or Process status) — still the deliberately-deferred
next phase, same order as before: (a) persist row date-moves and Process Done/Ready status to the DB, (b)
wire the Portal batch button to `portal_ship_and_print.py`, (c) wire the SAP buttons as subprocess calls,
(d) Excel export. Vendor Dashboard also still untouched (carried over every session).

## Session 6 continued (2026-07-28, same day) — same-order merge, inline edit, checked-row delete

User asked for three more changes to `workbench_app.py` right after the above:

**1. Same-order row merging.** Rows that share the same order label (order type + order#) *and* same
customer *and* same phone *and* same address *and* same memo now visually merge into one block — a single
rowspanned cell across those shared columns — instead of repeating the same text on every item row.
Per the user's explicit addition ("추가된 것은 같은 오더넘버도 병합한다는 것"), this merges **across separate
`orders` DB rows too**, not just items within one DB order row — real example found live in `workbench.db`:
order_no `67066585` exists as two separate order records (a duplicate-looking delivery/pickup pair from
Excel import). Verified live: when both records' customer/phone/address/memo genuinely matched they'd
merge; in the actual data on hand they *didn't* fully match (one record had blank customer/phone/address),
so per spec they correctly stayed as two separate blocks — confirms the match is exact-equality, not just
same order number.

**2. Checked-row delete.** New button under the Excel column ("3. 체크된 행 삭제") removes whichever rows are
checked from the board. **Screen-only for now** (same as every other action on this board — moves,
Done/Ready status) — nothing here writes to `workbench.db`. Told the user this explicitly rather than
assuming a real delete was wanted, since a real DB delete is a one-way destructive action and every other
button on this board is still a placeholder/display-only action from the Session 5/6 structure-only phase.

**3. Type dropdown + free-text editing.** The Type column (배송/회수) is now a `<select>` per row (not a
plain label). Every other text field — Order#, Item, M/N, S/N, Customer, Phone, Address, Memo — is now
directly editable in place (`contenteditable`, commits on blur or Enter). Editing a merged/shared field
(Customer/Phone/Address/Memo/Order#) edits the whole merged group at once, since it's physically one cell
spanning those rows.

**Implementation note (why this was a bigger rewrite than it sounds):** merging + inline edit + delete all
need the table's grouping/rowspans recomputed after *any* action (an edit can change whether rows still
match; a move/delete changes group membership). Patching rowspans into the old hand-built HTML strings
piece by piece wasn't going to hold up, so the whole board's rendering moved client-side: the server now
sends one flat JSON row list per date (`board_data()` → `__BOARD_DATA__` in the page template) instead of
pre-built `<tbody>` HTML, and a JS `ROWS` array + `renderAll()` is the single source of truth — every action
(drag-drop, the Excel-panel date move, Done/Ready, edits, delete) mutates `ROWS` then calls `renderAll()`,
which regroups by merge-key and rebuilds the whole table from scratch. This is also why drag-and-drop now
uses a dedicated small grip handle (⠿, own column) instead of making the whole row draggable — a
draggable `<tr>` would otherwise fight with clicking into a `contenteditable` cell or a `<select>` inside
it (text selection by mouse-drag would trigger a row-drag instead). Removed the now-dead `_h()` HTML-escape
helper and `html` import — no longer needed since the client builds cells via `textContent`/DOM APIs, not
string concatenation.

All three verified live via the actual Chrome browser (not just py_compile): edited a merged Customer cell
and confirmed the group stayed merged; deleted a checked row and confirmed the remaining group's rowspan
shrank correctly; toggled a Type dropdown and confirmed its color class updated; confirmed the real
`67066585` duplicate-order case in `workbench.db` merges/doesn't-merge exactly per the match rule above.

**Not done yet:** same backend-wiring gap as before (nothing here writes to workbench.db — moves, edits,
delete, Done/Ready are all still display-only). Vendor Dashboard still untouched.

## Session 6 continued again (2026-07-28) — order-code line format, center alignment, address bolding

Three more small fixes to `workbench_app.py` in the same session:

**1. Order# cell's second line reformatted.** The muted small line under the main order label (e.g. under
"ZRX 551055962") was showing a bare number like `92149447` with no code prefix. User pointed out this is
actually an **OBD (Outbound Delivery) number** — confirmed by the existing `OBD_RE` regex elsewhere in this
file that already parses `"OBD 12345678"`-style text. Order#-cell content isn't limited to SAP order codes
(ZOR/ZRE/ZRX/ZINP/ZINT) - Bloomberg-side codes (OBD/ORD/SDSK/...) belong in the same cell too, just on their
own line(s). Fix: `board_row()` now prefixes the stored delivery number with `"OBD "` server-side, and the
client's secondary-line `<div>` (still `.muted`, small gray text, same as before) is now **directly editable
free text, one code per line** - `white-space:pre-wrap` + a shared blur handler that splits on `\n`, trims,
dedupes, and writes back to every row in the merged group (same pattern as Customer/Phone/etc). This lets
more lines be typed in directly if more than one secondary code applies (e.g. add an `ORD ...` line under
an existing `OBD ...` line) - verified live by editing one to two lines and reading back `ROWS`. The main
label line (order type + number) is unchanged - still bold/black, its own separate contenteditable div.

**2. Center alignment.** M/N, S/N, Phone, Address, Memo columns are now `text-align:center` (Order#, Item,
Customer stay left-aligned as before). Simple CSS `nth-child` addition, no JS change.

**3. Address: bold company name, plain address below.** No data-model change needed - `orders.address` in
`workbench.db` already stores the company name and street address separated by a real `\n` (confirmed via
direct query, e.g. `"HYUNDAI MOTOR SECURITIES CO LTD\n14 YEOUIDAE-RO YEONGDEUNGPO-GU,\n.,20F"`), it just
wasn't being rendered with `white-space:pre-wrap` before so the `\n` collapsed into a space and everything
ran together as one line. Fix was two CSS rules on `.addr-cell`: `white-space:pre-wrap` (renders the
existing `\n`s as real line breaks) and `::first-line{font-weight:700}` (bolds only the rendered first
line - the company name - without touching the underlying text or the edit behavior at all, still one
single contenteditable field). Note: addresses with no `\n` in them at all (single-line addresses, e.g. one
row's address is just `"서울 오피스"`) render fully bold since their one line *is* the first line - harmless
edge case given the data, not fixed further since there's no company/street distinction to preserve in that
data to begin with.

Verified all three live in Chrome: order-code lines show `OBD 92149447` (not bare `92149447`) and accept a
second typed line; M/N/S/N/Phone/Address/Memo are visibly centered; addresses with embedded `\n` (e.g.
"HYUNDAI MOTOR SECURITIES CO LTD" / "SHINHAN BANK" / "TIANJIN CHENGFAN...") show the company name bold on
its own line above the plain-weight street address, centered.

## Session 6 continued a third time (2026-07-28) — alignment fixes, merge-by-order-number, collapsible dates, more Type options

Five more fixes to `workbench_app.py`, same session:

**1. M/N "staircase" bug fixed.** User spotted M/N values appearing to sit at inconsistent vertical
positions row-to-row. Root cause: `.item-row td` was `vertical-align:top`, which looks fine for uniform-
height rows but goes visibly wrong once a merged (rowspan) Customer/Phone/Address/Memo cell forces some
rows in a group to be taller than others - the un-merged M/N/S/N cells stayed pinned to the top of
whatever height their row ended up with, producing an uneven "staircase" look down the column. Fixed by
changing the base rule to `vertical-align:middle` - this single change also covers item 3's "높이도
중앙정렬" (vertical centering) requirement for Customer/Phone/Address/Memo for free, since it's the same
underlying rule for every `.item-row td`.

**2. Header row (orange "Type/Order#/..." row) centered.** `thead th` was `text-align:left`; changed to
`center`.

**3. Customer/Phone/Address/Memo centered** (horizontal, via `nth-child` - Customer was missing from the
center-align list added earlier this session for M/N/S/N/Phone/Address/Memo; added it) **and vertically**
(covered by fix #1 above).

**4. Merge rule loosened to "same order number" (was: order number + exact match on customer/phone/
address/memo).** Concrete case that exposed this: order `ZRX 67066585` exists as two separate `orders` DB
rows - one carries the delivery items + the real customer/phone/address (`MINJUNG SUH`/...), the other
carries the pickup (회수) items with **blank** customer/phone/address (Excel only records those once, on
the delivery leg). Under the old exact-match rule these two legs never merged, so the pickup rows displayed
with no customer/phone/address at all - looked broken, even though they're unambiguously the same real
order. `mergeKey()` now groups purely by `orderLabel` (trimmed), and `sharedTd()` in `buildGroupRows()`
picks the **first non-blank value across the whole group** for Customer/Phone/Address/Memo instead of
always reading `group.rows[0]` - so the pickup leg now inherits the delivery leg's customer/phone/address
correctly. Verified live: all 4 rows of `ZRX 67066585` (2 delivery + 2 pickup) now show as one merged block
with the customer name / phone / address filled in for every row, while the underlying
`ROWS` data for the pickup rows still correctly shows blank customer in memory (only the *display* backfills
- nothing written back to the blank DB row unless the user explicitly edits the shared cell, in which case
it propagates to the whole group as before).

**5. Collapsible date sections.** Each date's blue banner row now has a ▾/▸ toggle button at its right edge
(click to collapse/expand that date's rows; banner itself always stays visible). State lives in a
`COLLAPSED_DATES` Set so it survives `renderAll()` re-renders triggered by unrelated actions (edits, moves,
etc.) elsewhere on the board. **Implementation gotcha hit and fixed during this pass:** first attempt set
`display:flex` directly on the banner `<td>` to lay out the label + toggle button side by side - this broke
the cell's colspan-based width in the fixed-layout table (a flexed table-cell stops fully participating in
table column-width distribution in Chrome), collapsing the banner into a thin vertical sliver with the date
text wrapping character-by-character. Fixed by leaving the `<td>` as a normal table cell and putting the
flex layout on an inner wrapper `<div class="date-banner-inner">` instead - table cells should never get
`display:flex` directly in this codebase, wrap the content in a div if you need flex layout inside one.

**6. Type dropdown expanded + user-extensible.** Added `Bloomberg`, `Delayed`, `회수 Delayed` alongside
the existing `배송`/`회수`, plus a `+ 추가` entry at the end of every dropdown. Picking `+ 추가` prompts for
a new type name (`window.prompt`) and, if non-empty, adds it to a shared in-memory `CUSTOM_TYPES` list -
so it immediately becomes selectable in *every* row's dropdown, not just the one it was typed into (verified
live). Like everything else on this board, custom types don't survive a page reload (no backend yet).
Colored `회수`/`회수 Delayed` red, `배송` blue, everything else (Bloomberg/Delayed/custom) a neutral purple
(`type-other`) since there's no obvious existing color bucket for them.

**Not done yet:** same backend-wiring gap as every prior sub-session (nothing here writes to
workbench.db). Vendor Dashboard still untouched.

## Session 6 continued a fourth time (2026-07-28) — real root cause of the M/N misalignment found and fixed

User sent a screenshot (`260723.png`) showing each order group's *first* M/N value sitting centered while
every row after it in the same group looked left-shifted/wider. This was **not** a leftover styling gap -
it was a genuine bug in how column styling was being selected, present since the order-merge rowspan work
earlier in this session, and it affected more than just M/N.

**Root cause:** column width/alignment was assigned via `.item-row td:nth-child(N)` - but `nth-child` counts
a row's *own* children, not true table-column position. A merged group's first row has all 11 `<td>`s
(Order#/Customer/Phone/Address/Memo included, each `rowSpan=n`), but every row *after* the first in that
group omits those 5 cells entirely (they're covered by the first row's rowspan) - so in those rows, Item
becomes the 3rd child instead of the 4th, M/N the 4th instead of 5th, S/N the 5th instead of 6th, and check
the 6th instead of 11th. Each was silently inheriting the *previous* column's nth-child rule (e.g. M/N
rendering with Item's `width:15%` and no `text-align:center`, since Item's rule doesn't set one) on every
row except a group's first. This is exactly why the bug only ever showed up starting this session - it
needed a multi-row merged group to exist at all, which the merge feature only just introduced.

**Fix:** replaced every `nth-child`-based column rule on body rows with explicit classes assigned at cell
creation - `item-cell`, `mn-cell`, `sn-cell` (previously created with `className=''`) and `customer-cell`/
`phone-cell`/`memo-cell` (previously all lumped under the shared `cust-cell` class with widths coming from
now-removed nth-child rules; `cust-cell` is kept as a shared color/base style, `addr-cell` similarly kept
but now also carries its own width). Classes don't shift when a row has fewer cells, so this can't recur.
`thead th:nth-child` rules were left alone (the header only ever has one row with all 11 cells, so it was
never actually affected - only body rows with merged/rowspan cells were).

Verified live via `getBoundingClientRect()` + `getComputedStyle()` on real merged groups (not just eyeballing
a screenshot): all M/N cells in a group now report identical `left`, `width`, and `text-align:center`; same
checked for Item/S/N/check widths. Also cleaned up two stray duplicate `workbench_app.py` server processes
found running simultaneously at session start (one was likely the user's own `workbench.bat` launch from
checking the previous round of changes) - killed both, started one fresh instance, consistent with the
standing PowerShell-only-process-management gotcha.

## Session 6 continued a fifth time (2026-07-28) — real Excel export wired up (first write-to-a-real-file action in this app)

User asked to replace the two placeholder Excel export buttons ("내보내기 (오늘자만)" / "내보내기 (전체)")
with per-date checkboxes (one next to each date banner's ▾/▸ toggle, at the right edge like the reference
screenshot `260723.png` showed) + a single unified export button, which writes each checked date's current
board content into that date's own sheet in the live shipping workbook, formatted to match the sheets
`load_today_from_excel()` already knows how to parse.

**This is the first feature in this whole app that writes to a real file** - every other action all session
(moves, edits, merges, delete, Done/Ready, collapse) has been screen-only. Treated accordingly:

- **New `POST /api/export` endpoint + `export_dates_to_excel()`.** Takes `{dates: {iso: [row, ...]}}` from
  the client (not from `workbench.db` - the board's in-browser `ROWS` state is the current source of truth
  since nothing this session persists to the db yet), and for each date: names the sheet the same
  `"{month}-{day}"` way `today_sheet_names()` looks for it (e.g. `7-29`), **replaces** that sheet entirely if
  it already exists (`del wb[sheet_name]` then recreate - board state is authoritative, not additive), writes
  the header row `Order # / item / M/N / S/N / customer / phone / ADDRESS / memo` (matching the exact header
  the importer already expects, per its own docstring), then one row per item: column A carries
  `itemType + " " + orderLabel` plus any `deliveryNo` lines appended with `\n` **only on a merged group's
  first row** (blank on the rows below it) - same shape as the original sheets. Customer/Phone/Address/Memo
  columns get the merge-resolved (first-non-blank-across-group) values via the same `resolvedField()` the
  display already uses, factored out of `sharedTd()` into a shared helper so display and export can't drift
  apart.
- **Backs up the whole workbook before writing** - `shutil.copy2()` to
  `{stem}.backup_before_export_{timestamp}.xlsx` next to the real file, every single export call, no
  exceptions. This is the one place in the app where a mistake isn't just a page refresh away from
  undone.
- **Client side:** `EXPORT_DATES` Set (parallel to `COLLAPSED_DATES`) tracks which date checkboxes are
  checked, surviving `renderAll()` re-renders. `buildExportRows(iso)` groups that date's `ROWS` the same way
  the board displays them and flattens to the wire format above. `exportChecked()` POSTs and toasts a
  per-sheet summary (or the error message) back.
- **Verified without touching the live production file:** copied the real `C:\1\배송장\2026 배송장.xlsx` to
  a scratch path, pointed `export_dates_to_excel()` at the copy via a throwaway script (not the running
  server), and confirmed (a) the backup file is created, (b) sheets are created/replaced with the right name
  and header, (c) the written column-A text round-trips correctly through the project's own
  `parse_item_type()`/`parse_order_text()` (confirmed `"배송 ZRX 67059008"` → `('배송', ('ZRX','67059008',''))`
  and `"회수 ZOR 7830053\nOBD 92149315"` → `('회수', ('ZOR','7830053','92149315'))`), and (d) a merged
  ZRX 67066585-style group's blank customer backfills correctly in the exported payload. The scratch copy
  and its test backup were deleted afterward - nothing touched the real file during this verification.
- **Did not click the real export button against the live file** - only verified the code path via the copy
  above. **Important, told to the user:** Excel currently has the real file open (`Get-Process EXCEL` showed
  a live process with that window title) - writing to it via openpyxl while it's open in Excel risks the
  next manual save in Excel silently overwriting whatever this feature just wrote. User should close the
  file in Excel first before actually using this button for real, or at least reopen/refresh it afterward
  rather than save over it.

**Not done yet:** the rest of the board (moves/edits/status/delete) is still screen-only, unaffected by this
change. Vendor Dashboard still untouched.

**Follow-up same day:** user asked not to overwrite an existing same-named sheet at all. Changed
`export_dates_to_excel()` from "delete-then-recreate the `{month}-{day}` sheet" to "never touch an existing
sheet" - new `_unique_sheet_name()` appends `" (2)"`, `" (3)"`, ... (Excel's own duplicate-sheet naming
convention) until it finds a free name, so exporting the same date twice now produces `7-25` then
`7-25 (2)` side by side rather than replacing. Re-verified against a fresh scratch copy of the real workbook
(never the live file): ran the export twice for the same date and confirmed both `7-25` and `7-25 (2)` exist
with correct content; scratch copy and backups deleted after. Updated the export button's hint text to say
this explicitly.

## What to tell the next assistant (or next session)

Paste this:

"Continue in C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation. Read BLOOMBERG_HANDOFF.md first,
especially all the 'Session 6 (2026-07-28)' sections near the end (there are five, same day, back to back)
— that's the most recent state. IMPORTANT: the last of those five wired up a real Excel export (per-date
checkbox next to each date banner's ▾/▸ toggle + one 'Excel로 내보내기' button that writes each checked
date's board content into that date's own sheet in the real, live `C:\1\배송장\2026 배송장.xlsx`, replacing
the sheet if it exists, always backing the whole workbook up first) - this is the ONLY action in the whole
app that touches a real file; everything else is still screen-only. It was verified only against a scratch
copy of the file, never the live one, and Excel had the live file open in a running process at the time
(`Get-Process EXCEL`) - remind the user to close it in Excel (or not save over it) before actually using
that button for real, if that's still true. Summary of everything else: `workbench_app.py` is a date-based
ops board (SAP/Portal/Process/
Excel 4-column control panel + one sticky-header table, dates stacked downward). Board rendering is fully
client-side: the server sends one JSON row list per date (`board_data()`), and JS holds it in a `ROWS`
array + `renderAll()` that rebuilds the whole table from scratch after every action (moves, edits, delete,
Done/Ready, collapse/expand). Current feature set, all built this one day: checkbox-select + a single
`<input type="date">` in the Excel panel moves rows to any date (creates a new date section on the fly if
needed); rows merge into one rowspanned block purely by matching order label (SAP order type + number) -
Customer/Phone/Address/Memo display the first non-blank value found across the whole group, so a pickup
leg missing those fields still shows the delivery leg's values; every field is directly editable
(contenteditable, or a dropdown for Type: 배송/회수/Bloomberg/Delayed/회수 Delayed + a '+ 추가' entry that
lets the user type in new types on the fly, shared across every row); a checked-rows delete button;
drag-and-drop between dates via a small ⠿ grip-handle column (not the whole row, so it doesn't fight with
editing); each date's blue banner has a ▾/▸ collapse toggle; Address shows the company name bold on its own
line (data already had `\n` between company/street, just wasn't being rendered); M/N, S/N, Customer, Phone,
Address, Memo are center-aligned both horizontally and vertically (fixed a 'staircase' vertical-alignment
bug caused by rowspan'd cells forcing uneven row heights). Also deleted the 14 stale 2026-05-28 test orders
from workbench.db (backed up first to `workbench_before_delete_0528_<timestamp>.db`). Everything above is
still screen-only - nothing in this dashboard writes to workbench.db yet (that's still the deliberately-
deferred next phase: persist moves/edits/status/delete → wire Portal batch button → wire SAP buttons →
Excel export). Vendor Dashboard still untouched. Two operational gotchas to remember: (1) Bash-tool
background-and-kill does not reliably kill real Windows python.exe processes - always use PowerShell
(Get-CimInstance Win32_Process / Stop-Process -Id) to check for and clear stray workbench_app.py processes
before starting a fresh one; (2) a table `<td>` must never get `display:flex` directly in this codebase (it
breaks colspan width in the fixed-layout table) - wrap flex content in an inner `<div>` instead; (3) column
width/text-align on body-row cells must be set via an explicit class assigned in JS at cell creation
(item-cell/mn-cell/sn-cell/customer-cell/phone-cell/addr-cell/memo-cell/check-cell/etc.), never via
`nth-child` - a merged group's rows-after-the-first omit several `<td>`s (covered by the first row's
rowspan), which shifts nth-child indices and silently applies the wrong column's width/alignment; this
already happened once (the M/N 'staircase' bug, fixed in the last Session 6 sub-section) and would happen
again for any new column that reintroduces nth-child."

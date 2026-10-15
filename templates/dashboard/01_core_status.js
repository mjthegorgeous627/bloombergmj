function tick(){document.getElementById('clock').textContent=new Date().toLocaleString('ko-KR',{weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'})}
setInterval(tick,1000);tick();

// 콘솔 창이 없어진(2026-08-25) 뒤로 SAP 자동루프가 실제로 돌고 있는지 화면에서
// 확인할 방법이 없어져서 추가 - 시계 옆에 점+문구로 표시. 백엔드 /api/sap/status는
// 이미 있었는데(2026-08-25) 프론트에서 아무도 안 부르고 있었음.
async function updateLoopStatus(){
  const el=document.getElementById('sapLoopStatus');
  try{
    const res=await fetch('/api/sap/status');
    const data=await res.json();
    el.classList.remove('running','stopped');
    if(data.loopRunning){
      el.classList.add('running');
      el.querySelector('.txt').textContent='SAP 루프 실행 중 (PID '+data.pids.join(', ')+')';
    }else{
      el.classList.add('stopped');
      el.querySelector('.txt').textContent='SAP 루프 중지됨';
    }
  }catch(err){
    el.classList.remove('running','stopped');
    el.querySelector('.txt').textContent='SAP 루프 상태 확인 실패';
  }
}
updateLoopStatus();
setInterval(updateLoopStatus,5000);

function toggleCtrlCol(id){
  const col=document.getElementById(id);
  const collapsed=col.classList.toggle('collapsed');
  col.querySelector('.ctrl-col-toggle').textContent=collapsed?'▸':'▾';
  syncStickyTopHeight();
}

// Keeps the board's sticky column-header row (already position:sticky) glued
// directly beneath the sticky-top control bar instead of overlapping it.
// Re-measured whenever the control bar's own height can change (column
// collapse/expand, window resize) since that height isn't fixed.
function syncStickyTopHeight(){
  const el=document.querySelector('.sticky-top');
  if(el) document.documentElement.style.setProperty('--sticky-top-h', el.offsetHeight+'px');
}
window.addEventListener('resize', syncStickyTopHeight);
syncStickyTopHeight();

function toast(msg){
  const t=document.createElement('div');
  t.className='toast';
  t.textContent=msg;
  document.body.appendChild(t);
  requestAnimationFrame(()=>t.classList.add('show'));
  setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},1900);
}

// -- Control panel: structure-only placeholders, wired to real actions next --
async function runMorningRoutine(){
  try{
    const res=await fetch('/api/sap/morning_routine', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('아침 루틴 강제 실행됨 (PID '+data.pid+') - 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('아침 루틴 실행 실패: '+err.message);
  }
}
// 2026-10-15: 티켓 확인 - 서버가 숨김 프로세스로 Claude를 띄워 수신함을 읽고 "보드 반영 제안"을 만든다(보드는 수정 안 함).
let _ticketPoll=null;
function _ticketBtn(text, busy){
  const b=document.getElementById('ticketCheckBtn');
  if(!b) return;
  b.textContent=text; b.disabled=!!busy;
}
function _closeTicketReport(){ const o=document.getElementById('ticketReportOverlay'); if(o) o.remove(); }
function _openTicketReport(text){
  _closeTicketReport();
  const o=document.createElement('div');
  o.id='ticketReportOverlay';
  o.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center';
  o.onclick=(e)=>{ if(e.target===o) _closeTicketReport(); };
  const box=document.createElement('div');
  box.style.cssText='background:#fff;color:#111;max-width:900px;width:92%;max-height:86vh;overflow:auto;border-radius:10px;padding:18px 22px;box-shadow:0 8px 30px rgba(0,0,0,.3)';
  const head=document.createElement('div');
  head.style.cssText='display:flex;justify-content:space-between;align-items:center;margin-bottom:10px';
  head.innerHTML='<b style="font-size:16px">티켓 확인 결과 (제안만 - 보드는 바뀌지 않았습니다)</b>';
  const x=document.createElement('button'); x.textContent='닫기'; x.className='btn'; x.onclick=_closeTicketReport;
  head.appendChild(x);
  const pre=document.createElement('pre');
  pre.style.cssText='white-space:pre-wrap;word-break:break-word;font:13px/1.55 inherit;margin:0';
  pre.textContent=text;
  box.appendChild(head); box.appendChild(pre); o.appendChild(box); document.body.appendChild(o);
}
async function showTicketReport(){
  try{
    const st=await (await fetch('/api/tickets/status')).json();
    if(st.report) _openTicketReport(st.report);
    else toast('아직 보여줄 티켓 확인 결과가 없습니다 ('+(st.state||'idle')+')');
  }catch(err){ toast('결과를 불러오지 못했습니다: '+err.message); }
}
async function _pollTicketCheck(){
  try{
    const st=await (await fetch('/api/tickets/status')).json();
    if(st.state==='waiting'){ _ticketBtn('티켓 확인 (SAP 사이클 대기중…)', true); return; }
    if(st.state==='running'){ _ticketBtn('티켓 확인 중… (PC 만지지 마세요)', true); return; }
    clearInterval(_ticketPoll); _ticketPoll=null; _ticketBtn('티켓 확인', false);
    if(st.state==='done'){ toast('티켓 확인 완료 - 결과를 엽니다'); if(st.report) _openTicketReport(st.report); }
    else if(st.state==='error'){ toast('티켓 확인 실패: '+(st.message||'')); }
  }catch(err){ /* 서버 재시작 중 등 - 다음 폴링에서 재시도 */ }
}
async function runTicketCheck(){
  if(!window.confirm('블룸버그 터미널에서 새 티켓을 읽습니다 (보통 2~5분).\n\n- 읽는 동안 마우스·키보드를 건드리지 마세요.\n- SAP 자동루프는 이 동안 잠깐 대기합니다.\n- 보드는 바뀌지 않고, 반영 제안만 만들어집니다.\n\n시작할까요?')) return;
  try{
    const res=await fetch('/api/tickets/check', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'시작 실패');
    if(!data.started){ toast(data.message); return; }
    _ticketBtn('티켓 확인 (시작 중…)', true);
    if(_ticketPoll) clearInterval(_ticketPoll);
    _ticketPoll=setInterval(_pollTicketCheck, 3000);
  }catch(err){ toast('티켓 확인 시작 실패: '+err.message); }
}
async function runSap(kind){
  try{
    if(kind==='start'){
      const res=await fetch('/api/sap/start_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'시작 실패');
      toast(data.started ? 'SAP 자동루프 시작됨 (PID '+data.pid+')' : data.message);
      updateLoopStatus();
    } else if(kind==='stop'){
      if(!window.confirm('SAP 자동루프를 지금 중지할까요? (워치독도 다시 켤 때까지 재시작하지 않습니다)')) return;
      const res=await fetch('/api/sap/stop_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'중지 실패');
      toast(data.message);
      updateLoopStatus();
    } else if(kind==='all'){
      const res=await fetch('/api/sap/run_all', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
      toast(data.message);
    }
  }catch(err){
    toast('SAP '+kind+' 실패: '+err.message);
  }
}
// 2026-08-28: 트레이 아이콘 우클릭 메뉴와 같은 조작을 시계 옆 상태 점에서도
// 할 수 있게 - 클릭하면 열리는 드롭다운. 문서 아무 곳이나 클릭하면 닫힘.
function toggleSapLoopMenu(ev){
  ev.stopPropagation();
  document.getElementById('sapLoopMenu').classList.toggle('open');
}
function closeSapLoopMenu(){
  document.getElementById('sapLoopMenu').classList.remove('open');
}
document.addEventListener('click', closeSapLoopMenu);

// 2026-09-13 사용자 요청: 시계 옆 📅 버튼 - 주말/한국 공휴일(자동화 비운영일)
// 을 달력 숫자로 빨갛게 보여주고, 클릭으로 수동 지정/해제도 가능하게. 색
// 판단은 이미 있는 nonworkingInfo()/MANUAL_HOLIDAYS를 그대로 재사용 -
// 서버의 /api/nonworking_days/set에 저장되는 것도 동일.
let CAL_VIEW = (() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; })();

function _calIso(y, m, day){
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function toggleCalendarPanel(ev){
  ev.stopPropagation();
  const panel = document.getElementById('calPanel');
  const opening = !panel.classList.contains('open');
  panel.classList.toggle('open');
  if(opening){
    const today = new Date();
    CAL_VIEW = { y: today.getFullYear(), m: today.getMonth() };
    renderCalendarPanel();
  }
}
function closeCalendarPanel(){
  document.getElementById('calPanel').classList.remove('open');
}
document.addEventListener('click', closeCalendarPanel);

function calShiftMonth(delta, ev){
  ev.stopPropagation();
  CAL_VIEW.m += delta;
  if(CAL_VIEW.m < 0){ CAL_VIEW.m = 11; CAL_VIEW.y -= 1; }
  if(CAL_VIEW.m > 11){ CAL_VIEW.m = 0; CAL_VIEW.y += 1; }
  renderCalendarPanel();
}

function calToggleDay(iso, ev){
  ev.stopPropagation();
  const willAdd = !MANUAL_HOLIDAYS.has(iso);
  if(willAdd) MANUAL_HOLIDAYS.add(iso); else MANUAL_HOLIDAYS.delete(iso);
  fetch('/api/nonworking_days/set', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({date: iso, on: willAdd}),
  }).catch(() => {});
  renderCalendarPanel();
}

function renderCalendarPanel(){
  const {y, m} = CAL_VIEW;
  const panel = document.getElementById('calPanel');
  const startOffset = new Date(y, m, 1).getDay();  // 0=일요일
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const now = new Date();
  const todayIso = _calIso(now.getFullYear(), now.getMonth(), now.getDate());

  let html = `<div class="cal-head">`
    + `<button type="button" onclick="calShiftMonth(-1,event)">◂</button>`
    + `<span class="cal-title">${y}. ${String(m + 1).padStart(2, '0')}</span>`
    + `<button type="button" onclick="calShiftMonth(1,event)">▸</button>`
    + `</div><div class="cal-grid">`;
  ['일', '월', '화', '수', '목', '금', '토'].forEach(w => { html += `<div class="cal-wd">${w}</div>`; });
  for(let i = 0; i < startOffset; i++) html += `<div></div>`;
  for(let day = 1; day <= daysInMonth; day++){
    const iso = _calIso(y, m, day);
    const info = nonworkingInfo(iso);
    const cls = ['cal-day'];
    if(info.non) cls.push('nonworking');
    if(iso === todayIso) cls.push('today');
    const title = info.non ? info.label : '클릭: 이 날짜를 수동으로 휴무일(빨간색) 지정';
    html += `<div class="${cls.join(' ')}" title="${title}" onclick="calToggleDay('${iso}',event)">${day}</div>`;
  }
  html += `</div><div class="cal-legend">빨간 숫자 = 그날 SAP 자동화 비운영일(주말/공휴일/수동 지정). 숫자를 클릭하면 수동으로 지정·해제할 수 있습니다.</div>`;
  panel.innerHTML = html;
}
async function restartWorkbench(){
  if(!window.confirm('Workbench를 강제로 재시작할까요?\n1~2초간 서버가 끊겼다가 새로 뜨고, 이 페이지는 자동으로 새로고침됩니다.')) return;
  try{
    const res=await fetch('/api/workbench/restart', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'재시작 트리거 실패');
    toast(data.message);
    setTimeout(()=>location.reload(), 3000);
  }catch(err){
    toast('Workbench 재시작 실패: '+err.message);
  }
}
async function runSapSession(idx){
  try{
    const res=await fetch('/api/sap/run_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({session:idx})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast(data.message);
  }catch(err){
    toast('세션 '+idx+' 조회 실패: '+err.message);
  }
}
async function runOrderNoAction(){
  const input=document.getElementById('orderNoInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/run_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    toast('오더번호 반영 시작됨: '+v+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더번호 반영 실패: '+err.message);
  }
}
async function runOrderJumpAction(){
  const input=document.getElementById('orderJumpInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/open_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    const tcode = v.startsWith('6') ? 'VA02' : 'VA03';
    toast('오더 '+v+' ('+tcode+') 여는 중 (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더 바로가기 실패: '+err.message);
  }
}
async function runSerialLookupAction(){
  const input=document.getElementById('serialLookupInput');
  const v=input.value.trim();
  if(!v){toast('시리얼을 입력하세요');return}
  try{
    const res=await fetch('/api/sap/open_serial', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({serial:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    toast('시리얼 '+v+' ZIH08 조회 중 (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('시리얼 조회 실패: '+err.message);
  }
}
async function runNewSessionAction(){
  const input=document.getElementById('tcodeInput');
  const v=input.value.trim();
  try{
    const res=await fetch('/api/sap/open_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({tcode:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('새 SAP 세션 여는 중'+(v?' ('+v+')':'')+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('새 세션 열기 실패: '+err.message);
  }
}
async function openPortalLoginManual(){
  try{
    const res=await fetch('/api/portal/login_manual', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('Chrome 열림 (PID '+data.pid+') - 직접 로그인하세요. 이 창을 켜둔 채로 계속 재사용하면 됨');
  }catch(err){
    toast('Portal 로그인 창 열기 실패: '+err.message);
  }
}
// -- Bloomberg Portal: real automation, not a stub - runs
// portal_ship_and_print.py per distinct order among the checked rows, ONE
// AT A TIME (waits for each to fully finish via pollOrderDone() before
// starting the next), and stops on the first failure/timeout rather than
// plowing through the rest - if one order's run failed, something is likely
// systemically wrong (not logged into Portal, Chrome not running, a bad
// scrape) and the same failure would probably just repeat. Confirms the
// exact order list up front, same safety habit launcher.py's own
// messagebox.askyesno already has for this same action. --
async function pollOrderDone(orderId, timeoutMs){
  const start=Date.now();
  while(Date.now()-start < timeoutMs){
    await new Promise(r=>setTimeout(r, 3000));
    const res=await fetch('/api/orders/'+orderId);
    const data=await res.json();
    if(data.status && data.status!=='working'){
      // events[0] is the most recent event (newest-first) - on failure this
      // is run_portal()'s own add_event() call, whose first line is always
      // a clean one-sentence summary of what actually went wrong (see
      // run_portal's watcher()), not just a bare exit code.
      const detail=(data.events && data.events[0] && data.events[0].message || '').split('\n')[0];
      return {status: data.status, detail};
    }
  }
  return {status: 'timeout', detail: ''};
}
async function runPortalBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const orderIds=[...new Set(checkedRows.map(r=>r.orderId))];
  const labelFor_=id=>{ const r=checkedRows.find(r=>r.orderId===id); return r ? r.orderLabel : ('오더#'+id); };
  const summary=orderIds.map(labelFor_).join(', ');
  if(!window.confirm(
    `다음 ${orderIds.length}개 오더에 대해 "Serial 등록 & QR 인쇄"를 순서대로 실행합니다:\n\n${summary}\n\n진행할까요?`
  )) return;

  for(const orderId of orderIds){
    const label=labelFor_(orderId);
    let data;
    try{
      const res=await fetch('/api/orders/'+orderId+'/run_portal', {method:'POST'});
      data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    }catch(err){
      toast(label+' 실행 실패: '+err.message+' - 나머지 오더는 중단합니다');
      return;
    }
    if(!data.started){
      toast(label+': '+(data.message||'건너뜀'));
      continue;
    }
    toast(label+' 진행 중 (PID '+data.pid+') - 완료될 때까지 대기...');
    const result=await pollOrderDone(orderId, 360000);
    if(result.status==='printed'){
      toast(label+' 완료');
    } else if(result.status==='timeout'){
      toast(label+' 시간 초과 (6분) - 상태를 직접 확인하세요. 나머지 오더는 중단합니다');
      return;
    } else {
      toast(label+' 실패: '+(result.detail||('상태: '+result.status))+' - 나머지 오더는 중단합니다');
      return;
    }
  }
  toast('선택한 오더 처리 완료');
}
async function runLatestZpl(){
  try{
    const res=await fetch('/api/portal/print_latest_zpl', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('ZPL 인쇄 시작됨 (PID '+data.pid+') - Downloads 폴더 최신 라벨을 프린터로 전송, 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('ZPL 인쇄 실패: '+err.message);
  }
}

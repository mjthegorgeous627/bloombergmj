// 2026-10-13 추가 기능 (A에 더하는 것, 기존 코드는 그대로):
//  - 오른쪽 열: 달력(날짜별 건수/상태, 클릭하면 체크한 행 이동 또는 그 날짜로 스크롤) + 빠른 메모
//  - 메모 칸 맨 위 라벨(Delayed / 확인필요)과 모아보기 필터
//  - 오더번호/SDSK·ORD 옆 작은 이동·복사 아이콘
// 주의: 이 파일은 마지막에 합쳐지지만 renderAll()은 05번에서 이미 한 번 실행되므로,
// 여기의 상태 변수는 let/const 대신 var로 선언한다 (TDZ 방지 - 2026-08-27 사고 교훈).

var SIDE_OPEN = true;
try { SIDE_OPEN = localStorage.getItem('wbSideOpen') !== '0'; } catch (e) {}
var SIDE_CAL = null;
var SIDE_MEMOS = [];
var SIDE_MEMO_Q = '';

function applySideState(){
  const col = document.getElementById('sideCol');
  if(!col) return;
  col.classList.toggle('closed', !SIDE_OPEN);
  document.body.classList.toggle('side-open', SIDE_OPEN);
  document.body.classList.toggle('side-closed', !SIDE_OPEN);
  const t = document.getElementById('sideToggle');
  if(t) t.textContent = SIDE_OPEN ? '▸' : '◂';
}
function toggleSideCol(){
  SIDE_OPEN = !SIDE_OPEN;
  try { localStorage.setItem('wbSideOpen', SIDE_OPEN ? '1' : '0'); } catch (e) {}
  applySideState();
}
function openSideCol(){
  if(!SIDE_OPEN){ SIDE_OPEN = true; try { localStorage.setItem('wbSideOpen', '1'); } catch (e) {} applySideState(); }
}

// ---------- 날짜별 집계 / 달력 ----------
// (SIDE_FIN / SIDE_FIN_LOADED 는 첫 renderAll() 이전에 있어야 해서 03_zrec.js 쪽에 선언됨)
function loadSideFin(y, m, force){
  const key = y + '-' + String(m + 1).padStart(2, '0');
  if(SIDE_FIN_LOADED[key] && !force) return;
  SIDE_FIN_LOADED[key] = 1;
  const from = key + '-01', to = key + '-' + String(new Date(y, m + 1, 0).getDate()).padStart(2, '0');
  fetch('/api/calendar/counts?from=' + from + '&to=' + to).then(r => r.json()).then(d => {
    const c = (d && d.counts) || {};
    Object.keys(c).forEach(iso => { SIDE_FIN[iso] = Math.max(SIDE_FIN[iso] || 0, c[iso] || 0); });
    renderSideCalendar();
  }).catch(() => { SIDE_FIN_LOADED[key] = 0; });
}
// 마감 직후: 방금 마감한 행들의 날짜별 오더 수를 서버에 누적하고(이력 재색인 전 공백 보완) 화면에도 바로 반영
function recordFinalizedCounts(rows){
  const per = {};
  rows.forEach(r => { const k = (typeof mergeKey === 'function' ? mergeKey(r) : '') || r.rowKey; (per[r.date] = per[r.date] || {})[k] = 1; });
  const counts = {};
  Object.keys(per).forEach(iso => { counts[iso] = Object.keys(per[iso]).length; SIDE_FIN[iso] = (SIDE_FIN[iso] || 0) + counts[iso]; });
  try{ fetch('/api/calendar/finalized', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ counts: counts }) }).catch(() => {}); }catch(e){}
}
function sideDateStat(iso){
  const rs = ROWS.filter(r => r.date === iso);
  const keys = {};
  rs.forEach(r => { keys[(typeof mergeKey === 'function' ? mergeKey(r) : '') || r.rowKey] = 1; });
  let left = 0, pend = 0;
  rs.forEach(r => {
    if(!['done', 'cancelled', 'billed'].includes(r.uiStatus)) left++;
    if(r.uiStatus === 'pending') pend++;
  });
  const fin = SIDE_FIN[iso] || 0;   // 마감(내보내기+삭제)으로 보드에서 사라진 오더 수 - 달력 숫자가 마감 후에도 남게 함
  return { ord: Object.keys(keys).length, fin: fin, tot: Object.keys(keys).length + fin, it: rs.length, left: left, pend: pend };
}
function _sideIso(d){ return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }

function renderSideCalendar(){
  const host = document.getElementById('sideCal');
  if(!host) return;
  const today = localTodayIso();
  if(!SIDE_CAL){ const t = new Date(); SIDE_CAL = { y: t.getFullYear(), m: t.getMonth() }; }
  const y = SIDE_CAL.y, m = SIDE_CAL.m;
  const first = new Date(y, m, 1).getDay(), dim = new Date(y, m + 1, 0).getDate();
  const td = sideDateStat(today);
  let wk = 0;
  for(let i = 0; i < 7; i++){ const d = new Date(today + 'T00:00:00'); d.setDate(d.getDate() + i); wk += sideDateStat(_sideIso(d)).tot; }
  loadSideFin(y, m);
  let h = '<div class="sc-sum"><span class="sc-acc">오늘 ' + td.tot + '건</span><span>남음 ' + td.left + '</span><span>7일간 ' + wk + '건</span></div>';
  h += '<div class="sc-nav"><button type="button" onclick="sideCalNav(-1)" aria-label="이전 달">‹</button><span>' + y + '. ' + String(m + 1).padStart(2, '0') + '</span><button type="button" onclick="sideCalNav(1)" aria-label="다음 달">›</button></div>';
  h += '<div class="sc-grid">';
  ['일', '월', '화', '수', '목', '금', '토'].forEach(w => { h += '<span class="sc-wd">' + w + '</span>'; });
  for(let i = 0; i < first; i++) h += '<span></span>';
  for(let d = 1; d <= dim; d++){
    const iso = _calIso(y, m, d), info = nonworkingInfo(iso), s = sideDateStat(iso);
    let cls = 'sc-day';
    if(info.non) cls += ' non';
    if(iso === today) cls += ' today';
    if(s.tot) cls += s.ord ? (s.pend ? ' pend' : (s.left ? ' left' : ' done')) : ' done fin';
    const tip = (info.non ? info.label + ' · ' : '') + (s.tot ? '오더 ' + s.tot + (s.fin ? ' (마감 ' + s.fin + ')' : '') + (s.ord ? ' · 품목 ' + s.it + ' · 남음 ' + s.left + (s.pend ? ' · 회수 대기 ' + s.pend : '') : '') : '확정된 건 없음');
    h += '<button type="button" class="' + cls + '" title="' + tip + '" onclick="sideCalClick(\'' + iso + '\',event)" oncontextmenu="sideCalHoliday(\'' + iso + '\',event)"><span class="n">' + d + '</span>' + (s.tot ? '<span class="c">' + s.tot + '</span>' : '') + '</button>';
  }
  h += '</div>';
  h += '<div class="sc-leg">우클릭 = 휴무일 지정/해제 · <a href="#" onclick="openMiscSheets();return false" style="color:inherit">메모 시트</a></div>';
  host.innerHTML = h;
}
function sideCalNav(delta){
  SIDE_CAL.m += delta;
  if(SIDE_CAL.m < 0){ SIDE_CAL.m = 11; SIDE_CAL.y--; }
  if(SIDE_CAL.m > 11){ SIDE_CAL.m = 0; SIDE_CAL.y++; }
  renderSideCalendar();
}
function sideCalHoliday(iso, ev){
  ev.preventDefault();
  calToggleDay(iso, ev);       // 기존 비운영일 수동 지정/해제 (우클릭 또는 Ctrl+클릭)
  renderSideCalendar();
  renderAll();
}
function sideCalClick(iso, ev){
  if(ev.ctrlKey || ev.metaKey){ sideCalHoliday(iso, ev); return; }
  if(document.querySelectorAll('.row-check:checked').length){ moveCheckedToDate(iso); return; }
  // 보드에 이미 없는 과거(마감된) 날짜 -> 이력 팝업 (수정하면 원본 배송장 엑셀에 셀 단위로 반영)
  if(iso <= localTodayIso() && !document.querySelector('tbody.date-group[data-date="' + iso + '"]')){ openDayHistory(iso); return; }
  jumpToDate(iso);
}
function jumpToDate(iso){
  if(COLLAPSED_DATES.has(iso)){ COLLAPSED_DATES.delete(iso); renderAll(); }
  const tb = document.querySelector('tbody.date-group[data-date="' + iso + '"]');
  if(!tb) return;
  const ban = tb.querySelector('.date-banner-row');
  (ban || tb).scrollIntoView({ block: 'start', behavior: 'smooth' });
  tb.classList.add('jump-flash');
  setTimeout(() => tb.classList.remove('jump-flash'), 1200);
}

// ---------- 빠른 메모 ----------
function loadSideMemos(){
  fetch('/api/memos').then(r => r.json()).then(d => { SIDE_MEMOS = d.memos || []; renderSideMemos(); }).catch(() => {});
}
function _memoDay(m){ return String(m.created_at || '').slice(0, 10); }
function renderSideMemos(){
  const host = document.getElementById('sideMemoList');
  if(!host) return;
  const q = SIDE_MEMO_Q.trim().toLowerCase();
  const list = SIDE_MEMOS.filter(m => !q || String(m.body).toLowerCase().includes(q));
  list.sort((a, b) => (b.pinned - a.pinned) || (a.created_at < b.created_at ? 1 : -1));
  const today = localTodayIso(), yd = new Date(today + 'T00:00:00'); yd.setDate(yd.getDate() - 1);
  const yIso = _sideIso(yd);
  let out = '', last = '';
  list.forEach(m => {
    const day = _memoDay(m);
    const g = m.pinned ? '고정' : day === today ? '오늘' : day === yIso ? '어제' : day.slice(5).replace('-', '/');
    if(g !== last){ out += '<div class="sm-grp">' + g + '</div>'; last = g; }
    let tx = String(m.body).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
    if(q) tx = tx.replace(new RegExp('(' + q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/[&<>]/g, '') + ')', 'ig'), '<mark>$1</mark>');
    out += '<div class="sm-item' + (m.pinned ? ' pin' : '') + (m.done ? ' dn' : '') + '" data-id="' + m.id + '">'
      + '<input type="checkbox" data-sm="done" ' + (m.done ? 'checked' : '') + ' aria-label="완료">'
      + '<div class="sm-tx" contenteditable="true" data-sm="text" spellcheck="false">' + tx + '</div>'
      + '<span class="sm-act"><button type="button" data-sm="pin" title="고정 메모 탭으로 옮기기">📌</button><button type="button" data-sm="del" title="삭제">×</button></span>'
      + '<span class="sm-tm">' + String(m.created_at || '').slice(5, 16) + '</span></div>';
  });
  host.innerHTML = out;
  const c = document.getElementById('sideMemoCount');
  if(c) c.textContent = SIDE_MEMOS.length ? SIDE_MEMOS.length : '';
}
function _memoPost(path, body){
  return fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })
    .then(r => r.json()).catch(() => ({}));
}
function addSideMemo(text){
  const v = text.trim();
  if(!v) return;
  const n = new Date(), p = z => String(z).padStart(2, '0');
  const m = { id: 'tmp' + Date.now(), body: v, pinned: 0, done: 0, created_at: localTodayIso() + ' ' + p(n.getHours()) + ':' + p(n.getMinutes()) };
  SIDE_MEMOS.push(m);
  renderSideMemos();
  _memoPost('/api/memos/add', { body: v }).then(d => { if(d && d.id){ m.id = d.id; renderSideMemos(); } });
}
function sideMemoById(id){ return SIDE_MEMOS.find(m => String(m.id) === String(id)); }
function openSideMemo(){
  openSideCol();
  const t = document.getElementById('sideMemoNew');
  if(t){ t.focus(); }
}
document.addEventListener('keydown', e => {
  if((e.ctrlKey || e.metaKey) && !e.altKey && String(e.key).toLowerCase() === 'm'){ e.preventDefault(); openSideMemo(); return; }
  if(e.target && e.target.id === 'sideMemoNew' && e.key === 'Enter' && !e.shiftKey){
    e.preventDefault(); addSideMemo(e.target.value); e.target.value = '';
  }
});
document.addEventListener('input', e => {
  if(e.target && e.target.id === 'sideMemoQ'){ SIDE_MEMO_Q = e.target.value; renderSideMemos(); }
});
document.addEventListener('click', e => {
  const b = e.target.closest('[data-sm]');
  if(!b) return;
  const item = b.closest('.sm-item');
  if(!item) return;
  const m = sideMemoById(item.dataset.id);
  if(!m) return;
  const k = b.dataset.sm;
  if(k === 'done'){ m.done = b.checked ? 1 : 0; renderSideMemos(); _memoPost('/api/memos/update', { id: m.id, done: m.done }); }
  if(k === 'pin'){ moveMemoToPin(m); }
  if(k === 'del'){ SIDE_MEMOS = SIDE_MEMOS.filter(z => z !== m); renderSideMemos(); _memoPost('/api/memos/delete', { id: m.id }); }
});
document.addEventListener('focusout', e => {
  const t = e.target;
  if(!t || !t.dataset || t.dataset.sm !== 'text') return;
  const item = t.closest('.sm-item'), m = item && sideMemoById(item.dataset.id);
  const v = t.textContent.trim();
  if(m && v && v !== m.body){ m.body = v; _memoPost('/api/memos/update', { id: m.id, body: v }); }
});

// ---------- 메모 칸 라벨 (Delayed / 확인필요) + 모아보기 ----------
// 함수로 둔 이유: 05번의 첫 renderAll()이 이 파일보다 먼저 실행되므로 var 값 대입에 의존하면 안 된다
function flagDefs(){ return { delayed: { label: 'Delayed', cls: 'flag-delayed' }, check: { label: '확인필요', cls: 'flag-check' } }; }
function parseFlags(s){ return String(s || '').split(',').filter(k => flagDefs()[k]); }
function flagMatch(r){
  if(typeof BOARD_FILTER === 'undefined' || BOARD_FILTER === 'all') return true;
  return parseFlags(r.flags).includes(BOARD_FILTER);
}
function editableText(td){
  const c = td.cloneNode(true);
  c.querySelectorAll('[data-noedit]').forEach(n => n.remove());
  return c.textContent.trim();
}
function setFlags(rows, arr){
  const value = arr.join(',');
  const ids = rows.map(r => r.itemId);
  rows.forEach(r => { r.flags = value; });
  persistOrderField(ids, 'flags', value).catch(err => toast('라벨 저장 실패 (DB 반영 안 됨): ' + err.message));
  renderAll();
}
function decorateMemoTd(td, rows){
  td.style.position = 'relative';
  const cur = parseFlags(rows[0].flags);
  if(cur.length){
    const box = document.createElement('div');
    box.className = 'memo-flags';
    box.contentEditable = 'false';
    box.dataset.noedit = '1';
    cur.forEach(k => { const s = document.createElement('span'); s.className = 'flag-label ' + flagDefs()[k].cls; s.textContent = flagDefs()[k].label; box.appendChild(s); });
    td.insertBefore(box, td.firstChild);
  }
  // 빈 메모칸: 안에 편집 불가 요소(⚑ 버튼/라벨)만 있으면 브라우저가 글자 커서를 못 넣어 타이핑이 안 된다
  // (2026-10-14 ZRX 67110807 사례) -> 커서가 들어갈 빈 줄(<br>)을 하나 둔다. textContent에는 영향 없음.
  if(!editableText(td)) td.appendChild(document.createElement('br'));
  const btn = document.createElement('button');
  btn.type = 'button'; btn.className = 'flag-btn'; btn.textContent = '⚑'; btn.title = '라벨 지정 (Delayed / 확인필요)';
  btn.contentEditable = 'false'; btn.dataset.noedit = '1';
  btn.addEventListener('mousedown', e => e.preventDefault());
  btn.addEventListener('click', e => { e.stopPropagation(); openFlagMenu(e.clientX, e.clientY, rows); });
  td.appendChild(btn);
}
function openFlagMenu(x, y, rows){
  closeHlMenu();
  const cur = parseFlags(rows[0].flags);
  const menu = document.createElement('div');
  menu.className = 'hl-menu';
  menu.style.left = x + 'px'; menu.style.top = y + 'px';
  const mk = (label, cls, fn) => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'hl-menu-btn' + (cls ? ' ' + cls : ''); b.textContent = label;
    b.addEventListener('click', e => { e.stopPropagation(); fn(); closeHlMenu(); });
    menu.appendChild(b);
  };
  Object.keys(flagDefs()).forEach(k => {
    const on = cur.includes(k);
    mk(flagDefs()[k].label + (on ? ' 끄기' : ' 켜기'), flagDefs()[k].cls + '-btn', () => setFlags(rows, on ? cur.filter(z => z !== k) : cur.concat([k])));
  });
  mk('라벨 모두 해제', '', () => setFlags(rows, []));
  document.body.appendChild(menu);
  hlMenuEl = menu;
  setTimeout(() => document.addEventListener('click', closeHlMenu), 0);
}
function setBoardFilter(k){ BOARD_FILTER = k; renderAll(); }
function updateFlagFilterBar(){
  const bar = document.getElementById('ffChips');
  if(!bar) return;
  const seen = {}, cnt = { delayed: 0, check: 0 };
  let total = 0;
  ROWS.forEach(r => {
    const key = r.date + '|' + (typeof mergeKey === 'function' ? mergeKey(r) : r.rowKey);
    if(seen[key]) return;
    seen[key] = 1; total++;
    parseFlags(r.flags).forEach(k => { cnt[k]++; });
  });
  const chip = (k, label, n, cls) => '<button type="button" class="ff-chip ' + cls + (BOARD_FILTER === k ? ' on' : '') + '" onclick="setBoardFilter(\'' + k + '\')">' + label + '<i>' + n + '</i></button>';
  bar.innerHTML = '<span class="ff-title">모아보기</span>' + chip('all', '전체', total, '') + chip('delayed', 'Delayed', cnt.delayed, 'flag-delayed') + chip('check', '확인필요', cnt.check, 'flag-check');
}

// ---------- 오더번호 / SDSK·ORD 옆 아이콘 ----------
async function openOrderFromBoard(no){
  try{
    const res = await fetch('/api/sap/open_order', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ orderNo: no }) });
    const data = await res.json();
    if(!res.ok || data.error) throw new Error(data.error || '실행 실패');
    if(!data.started){ toast(data.message); return; }
    toast('오더 ' + no + ' (' + (String(no).startsWith('6') ? 'VA02' : 'VA03') + ') 여는 중');
  }catch(err){ toast('오더 바로가기 실패: ' + err.message); }
}
async function openTicketInTerminal(code){
  try{
    const res = await fetch('/api/terminal/open_ticket', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ code: code }) });
    const data = await res.json();
    if(!res.ok || data.error) throw new Error(data.error || '실행 실패');
    toast('터미널에서 ' + code + ' 여는 중');
  }catch(err){ toast('터미널 이동 실패: ' + err.message); }
}
function copyText(v){
  const done = () => toast('복사됨: ' + v);
  if(navigator.clipboard && navigator.clipboard.writeText){ navigator.clipboard.writeText(v).then(done).catch(done); } else done();
}
function _iconBtn(txt, title, fn){
  const b = document.createElement('button');
  b.type = 'button'; b.className = 'ord-ico'; b.textContent = txt; b.title = title;
  b.addEventListener('mousedown', e => e.preventDefault());
  b.addEventListener('click', e => { e.stopPropagation(); fn(); });
  return b;
}
async function persistOrderNo(itemIds, value){
  try{
    const res = await fetch('/api/orders/rename_order_no', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ itemIds: itemIds, value: value }) });
    const data = await res.json();
    if(!res.ok || data.error){ throw new Error(data.error || '저장 실패'); }
    toast('오더번호 변경 저장됨 (Ctrl+Z로 되돌릴 수 있음)');
  }catch(err){
    toast('오더번호 저장 안 됨: ' + err.message + ' - 화면을 새로고침합니다');
    setTimeout(() => location.reload(), 2200);
  }
}
function _iconSlot(){ const e = document.createElement('span'); e.className = 'ord-slot'; return e; }
async function openObdFromBoard(line){
  const no = String(line).replace(/\D/g, '');
  if(!no){ toast('OBD 번호를 찾지 못했습니다'); return; }
  try{
    const res = await fetch('/api/sap/open_obd', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ deliveryNo: no }) });
    const data = await res.json();
    if(!res.ok || data.error) throw new Error(data.error || '실행 실패');
    toast('OBD ' + no + ' 시리얼 입력 화면 여는 중');
  }catch(err){ toast('OBD 열기 실패: ' + err.message); }
}
function makeOrderIcons(orderNo, label){
  const s = document.createElement('span');
  s.className = 'ord-icons'; s.contentEditable = 'false';
  if(orderNo) s.appendChild(_iconBtn('↗', 'SAP에서 열기 (' + (String(orderNo).startsWith('6') ? 'VA02' : 'VA03') + ')', () => openOrderFromBoard(orderNo)));
  else s.appendChild(_iconSlot());
  s.appendChild(_iconBtn('⧉', '복사', () => copyText(orderNo || label)));
  return s;
}
function makeCodeIcons(codeLines){
  const col = document.createElement('div');
  col.className = 'code-icons'; col.contentEditable = 'false';
  codeLines.forEach(line => {
    const row = document.createElement('div');
    row.className = 'code-icon-row';
    if(/^OBD/i.test(line)) row.appendChild(_iconBtn('↗', 'SAP에서 열기 - 시리얼 입력 화면 (VL02N)', () => openObdFromBoard(line)));
    else if(/^(SDSK|ORD)/i.test(line)) row.appendChild(_iconBtn('↗', '터미널에서 열기', () => openTicketInTerminal(line)));
    else row.appendChild(_iconSlot());
    row.appendChild(_iconBtn('⧉', '복사', () => copyText(line)));
    col.appendChild(row);
  });
  return col;
}

// ---------- 갱신 훅 (renderAll 끝에서 호출) ----------
function syncFilterBarHeight(){
  const el = document.getElementById('flagFilterBar');
  if(el) document.documentElement.style.setProperty('--ff-h', el.offsetHeight + 'px');
}
window.addEventListener('resize', syncFilterBarHeight);
function updateSideViews(){
  updateFlagFilterBar();
  syncFilterBarHeight();
  renderSideCalendar();
  decorateRowTools();
}

// ---------- 리스트에서 바로 처리: 줄 아래 ＋ / 이동칸(⠿) 우클릭 메뉴 ----------
function addRowAt(iso){
  addManualRow();
  const d = document.getElementById('manualRowDate');
  if(d && iso) d.value = iso;
}
function decorateRowTools(){
  document.querySelectorAll('tbody.date-group').forEach(tb => {
    if(tb.querySelector('.add-row-tr')) return;
    const tr = document.createElement('tr');
    tr.className = 'add-row-tr';
    const td = document.createElement('td');
    td.colSpan = 11;
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'add-row-btn'; b.textContent = '＋';
    b.title = '이 날짜에 새 행 추가';
    b.addEventListener('click', () => addRowAt(tb.dataset.date));
    td.appendChild(b); tr.appendChild(td); tb.appendChild(tr);
  });
}
let rowMenuEl = null;
function closeRowMenu(){
  const home = document.getElementById('mergeFieldHome');
  const mf = document.querySelector('.merge-field-row');
  if(rowMenuEl){
    if(mf && home && !home.contains(mf)) home.appendChild(mf);
    rowMenuEl.remove(); rowMenuEl = null;
  }
  document.removeEventListener('click', closeRowMenuOnClick, true);
}
function closeRowMenuOnClick(e){
  if(rowMenuEl && rowMenuEl.contains(e.target)) return;
  closeRowMenu();
}
function openRowMenu(x, y){
  closeHlMenu(); closeRowMenu();
  const n = document.querySelectorAll('.row-check:checked').length;
  const menu = document.createElement('div');
  menu.className = 'hl-menu row-menu';
  menu.style.left = x + 'px'; menu.style.top = y + 'px';
  const mk = (label, fn, cls) => {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'hl-menu-btn' + (cls ? ' ' + cls : ''); b.textContent = label;
    b.addEventListener('click', e => { e.stopPropagation(); closeRowMenu(); fn(); });
    menu.appendChild(b);
  };
  const head = document.createElement('div');
  head.className = 'row-menu-head'; head.textContent = '선택 ' + n + '행';
  menu.appendChild(head);
  mk('복사', copyChecked);
  mk('삭제', deleteChecked, 'row-menu-del');
  const sep = document.createElement('div'); sep.className = 'row-menu-sep'; menu.appendChild(sep);
  const mf = document.querySelector('.merge-field-row');
  if(mf) menu.appendChild(mf);
  mk('병합', mergeChecked);
  mk('병합 해제', unmergeChecked);
  document.body.appendChild(menu);
  const r = menu.getBoundingClientRect();
  if(r.right > window.innerWidth) menu.style.left = Math.max(4, window.innerWidth - r.width - 6) + 'px';
  if(r.bottom > window.innerHeight) menu.style.top = Math.max(4, window.innerHeight - r.height - 6) + 'px';
  rowMenuEl = menu;
  setTimeout(() => document.addEventListener('click', closeRowMenuOnClick, true), 0);
}
document.getElementById('boardTable').addEventListener('contextmenu', e => {
  const cell = e.target.closest('.handle-cell');
  if(!cell) return;
  e.preventDefault();
  const tr = cell.closest('tr.item-row');
  const cb = tr && tr.querySelector('.row-check');
  if(cb && !cb.checked){
    // 엑셀처럼: 선택 밖의 줄을 우클릭하면 그 줄만 선택
    document.querySelectorAll('.row-check:checked').forEach(c => { c.checked = false; c.dispatchEvent(new Event('change', { bubbles: true })); });
    cb.checked = true; cb.dispatchEvent(new Event('change', { bubbles: true }));
  }
  openRowMenu(e.clientX, e.clientY);
});

// ---------- 고정 메모 (탭별 기록 = 기존 📌 메모 페이지의 notepad_pages) ----------
let PIN_PAGES = [], PIN_ACTIVE = null, PIN_SAVE_T = null, PIN_EL = null;
function _esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function pinLoad(){
  return fetch('/api/pages').then(r => r.json()).then(d => { PIN_PAGES = d.pages || []; if(!PIN_PAGES.some(p => p.id === PIN_ACTIVE)) PIN_ACTIVE = PIN_PAGES.length ? PIN_PAGES[0].id : null; }).catch(() => {});
}
function pinCur(){ return PIN_PAGES.find(p => p.id === PIN_ACTIVE); }
function pinFlush(){
  clearTimeout(PIN_SAVE_T);
  const ed = document.getElementById('pinEditor'), pg = pinCur();
  if(!ed || !pg || ed.value === pg.body) return;
  pg.body = ed.value;
  _memoPost('/api/pages/edit', { id: pg.id, field: 'body', value: ed.value });
}
function pinRender(){
  if(!PIN_EL) return;
  const tabs = PIN_EL.querySelector('.pin-tabs');
  tabs.innerHTML = '';
  PIN_PAGES.forEach(p => {
    const t = document.createElement('div');
    t.className = 'pin-tab' + (p.id === PIN_ACTIVE ? ' on' : '');
    t.innerHTML = '<span class="pt-title">' + _esc(p.title) + '</span><span class="pt-x" title="삭제">×</span>';
    t.addEventListener('click', e => {
      if(e.target.classList.contains('pt-x')){
        if(!t.classList.contains('arm')){ t.classList.add('arm'); e.target.textContent = '삭제?'; setTimeout(() => { t.classList.remove('arm'); e.target.textContent = '×'; }, 2500); return; }
        PIN_PAGES = PIN_PAGES.filter(z => z !== p);
        if(PIN_ACTIVE === p.id) PIN_ACTIVE = PIN_PAGES.length ? PIN_PAGES[0].id : null;
        _memoPost('/api/pages/delete', { id: p.id }); pinRender(); return;
      }
      if(p.id !== PIN_ACTIVE){ pinFlush(); PIN_ACTIVE = p.id; pinRender(); }
    });
    t.querySelector('.pt-title').addEventListener('dblclick', e => {
      const el = e.target; el.contentEditable = 'true'; el.focus();
      const done = () => {
        el.contentEditable = 'false';
        const v = el.textContent.trim();
        if(v && v !== p.title){ p.title = v; _memoPost('/api/pages/edit', { id: p.id, field: 'title', value: v }); } else el.textContent = p.title;
      };
      el.addEventListener('blur', done, { once: true });
      el.addEventListener('keydown', ev => { if(ev.key === 'Enter'){ ev.preventDefault(); el.blur(); } });
    });
    tabs.appendChild(t);
  });
  const add = document.createElement('button');
  add.type = 'button'; add.className = 'pin-add'; add.textContent = '+'; add.title = '새 탭';
  add.addEventListener('click', async () => {
    pinFlush();
    const r = await _memoPost('/api/pages/add', { title: '새 메모장' });
    if(r && r.id){ PIN_PAGES.push({ id: r.id, title: r.title || '새 메모장', body: '' }); PIN_ACTIVE = r.id; pinRender(); }
  });
  tabs.appendChild(add);
  const ed = PIN_EL.querySelector('#pinEditor');
  const pg = pinCur();
  ed.disabled = !pg; ed.value = pg ? pg.body : '';
}
async function openPinPanel(){
  if(!PIN_EL){
    PIN_EL = document.createElement('div');
    PIN_EL.className = 'pin-panel';
    PIN_EL.innerHTML = '<div class="pin-head"><b>고정 메모</b><a href="/notes" target="_blank" title="전체 화면으로 열기">↗</a><button type="button" class="pin-close" title="닫기">×</button></div>'
      + '<div class="pin-tabs"></div><textarea id="pinEditor" spellcheck="false"></textarea>';
    document.body.appendChild(PIN_EL);
    PIN_EL.querySelector('.pin-close').addEventListener('click', closePinPanel);
    PIN_EL.querySelector('#pinEditor').addEventListener('input', () => { clearTimeout(PIN_SAVE_T); PIN_SAVE_T = setTimeout(pinFlush, 700); });
  }
  PIN_EL.style.display = 'flex';
  await pinLoad(); pinRender();
}
function closePinPanel(){ pinFlush(); if(PIN_EL) PIN_EL.style.display = 'none'; }
function togglePinPanel(){ if(PIN_EL && PIN_EL.style.display !== 'none') closePinPanel(); else openPinPanel(); }
async function moveMemoToPin(m){
  const body = String(m.body || '');
  const title = (body.split('\n')[0] || '메모').trim().slice(0, 18) || '메모';
  const r = await _memoPost('/api/pages/add', { title });
  if(!r || !r.id){ toast('고정 메모로 옮기지 못했습니다'); return; }
  await _memoPost('/api/pages/edit', { id: r.id, field: 'body', value: body });
  SIDE_MEMOS = SIDE_MEMOS.filter(z => z !== m); renderSideMemos();
  _memoPost('/api/memos/delete', { id: m.id });
  PIN_ACTIVE = r.id;
  await openPinPanel(); PIN_ACTIVE = r.id; pinRender();
  toast('고정 메모 탭으로 옮겼습니다');
}

// ---------- 이력 검색 (billed 옆) ----------
let HIST_T = null, HIST_SEQ = 0, HIST_CUR = [], HIST_EL = null;
const HIST_LABELS = ['오더/구분', '자재', 'M/N', 'S/N', '고객', '연락처', '주소', '메모'];
function histTerms(){ return document.getElementById('histQ').value.trim().split(/\s+/).filter(Boolean); }
function histHl(text, terms){
  let h = _esc(text);
  terms.forEach(t => { const re = new RegExp(_esc(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'); h = h.replace(re, m => '<mark>' + m + '</mark>'); });
  return h;
}
function histEl(){
  if(!HIST_EL){
    HIST_EL = document.createElement('div');
    HIST_EL.className = 'hist-pop';
    document.body.appendChild(HIST_EL);
    HIST_EL.addEventListener('click', histClick);
  }
  return HIST_EL;
}
function histPlace(){
  const inp = document.getElementById('histQ'), r = inp.getBoundingClientRect(), el = histEl();
  el.style.top = (r.bottom + 6) + 'px';
  const room = window.innerWidth - (document.body.classList.contains('side-open') ? 316 : 56);
  const w = Math.min(960, room - 20);
  el.style.width = w + 'px';
  el.style.left = Math.max(10, Math.min(r.left, room - w - 10)) + 'px';
}
function closeHist(){ if(HIST_EL) HIST_EL.style.display = 'none'; }
async function histSearch(){
  const q = document.getElementById('histQ').value.trim();
  const el = histEl();
  if(!q){ closeHist(); return; }
  const my = ++HIST_SEQ;
  let r;
  try { r = await fetch('/api/history/search?q=' + encodeURIComponent(q)).then(x => x.json()); } catch (e) { return; }
  if(my !== HIST_SEQ) return;
  HIST_CUR = r.results || [];
  const t = histTerms();
  histPlace(); el.style.display = 'flex';
  const head = '<div class="hist-head"><b>' + (r.total || 0) + '건</b>' + (r.total > HIST_CUR.length ? '<span>(최근 ' + HIST_CUR.length + '건 표시)</span>' : '')
    + '<a href="/history" target="_blank" title="이력 페이지에서 열기 (되돌리기 포함)">이력 페이지 ↗</a><button type="button" class="hist-x" title="닫기">×</button></div>';
  if(!HIST_CUR.length){ el.innerHTML = head + '<div class="hist-empty">결과 없음</div>'; return; }
  el.innerHTML = head + '<div class="hist-scroll"><table class="hist-tbl"><tr><th>날짜</th><th>구분</th><th>오더</th><th>자재</th><th>M/N</th><th>S/N</th><th>고객</th><th>메모</th></tr>'
    + HIST_CUR.map((x, i) => { const c = x.cells;
      return '<tr class="hist-row" data-i="' + i + '"><td>' + _esc(x.sheetDate || x.sheetName) + '</td><td class="k-' + _esc(x.kind) + '">' + _esc(x.kind) + '</td><td>' + histHl((c[0] || '').replace(/^(배송|회수)\s*/, ''), t) + '</td><td>' + histHl(c[1] || '', t) + '</td><td>' + histHl(c[2] || '', t) + '</td><td>' + histHl(c[3] || '', t) + '</td><td>' + histHl(c[4] || '', t) + '</td><td>' + histHl(c[7] || '', t) + '</td></tr>';
    }).join('') + '</table></div>';
}
async function histClick(e){
  if(e.target.classList.contains('hist-x')){ closeHist(); return; }
  const tr = e.target.closest('tr.hist-row'); if(!tr) return;
  const nx = tr.nextElementSibling;
  if(nx && nx.classList.contains('hist-detail')){ nx.remove(); tr.classList.remove('sel'); return; }
  const x = HIST_CUR[+tr.dataset.i];
  const d = await fetch('/api/history/group?sheet_id=' + x.sheetId + '&group_no=' + x.groupNo).then(r => r.json()).catch(() => null);
  if(!d || !d.rows) return;
  const t = histTerms();
  const det = document.createElement('tr'); det.className = 'hist-detail';
  det.innerHTML = '<td colspan="8"><div class="hd-src">' + _esc(d.sourceFile) + ' · 시트 ' + _esc(d.sheetName) + ' · 엑셀 ' + d.rows.map(r => r.excelRow).join(', ') + '행</div>'
    + d.rows.map(r => '<div class="hd-row"><div class="hd-cap">' + _esc(r.kind || '') + '</div>' + r.cells.map((v, i) => v ? '<div class="hd-f"><b>' + _esc(HIST_LABELS[i] || ('열' + (i + 1))) + '</b><span>' + histHl(v, t) + '</span></div>' : '').join('') + '</div>').join('') + '</td>';
  tr.classList.add('sel'); tr.after(det);
}
(function(){
  const inp = document.getElementById('histQ');
  if(!inp) return;
  inp.addEventListener('input', () => { clearTimeout(HIST_T); HIST_T = setTimeout(histSearch, 300); });
  inp.addEventListener('focus', () => { if(inp.value.trim() && HIST_CUR.length){ histPlace(); histEl().style.display = 'flex'; } });
  inp.addEventListener('keydown', e => { if(e.key === 'Escape'){ closeHist(); inp.blur(); } if(e.key === 'Enter'){ clearTimeout(HIST_T); histSearch(); } });
  document.addEventListener('mousedown', e => { if(HIST_EL && HIST_EL.style.display !== 'none' && !HIST_EL.contains(e.target) && e.target !== inp) closeHist(); });
})();

applySideState();
loadSideMemos();

// ---------- 퀵 기사님 안내문 인쇄 (2026-10-14) ----------
// 체크한 배송/회수 행 -> 팝업에서 내용 확인/수정(미리보기) -> 인쇄. Word 파일 없이 A4 세로로 바로 인쇄한다.
const QUICK_PRODUCTS = [[/KEYBOARD/i, '키보드'], [/MONITOR|DISPLAY/i, '모니터'], [/MOUSE/i, '마우스'], [/MACBOOK|LAPTOP|NOTEBOOK/i, '노트북'], [/\bPC\b|DESKTOP/i, 'PC']];
function quickProductName(desc){
  const d = String(desc || '');
  for(const p of QUICK_PRODUCTS){ if(p[0].test(d)) return p[1]; }
  return d.replace(/\s+/g, ' ').trim().slice(0, 24) || '물품';
}
function quickDefaults(rows){
  const counts = new Map();
  rows.forEach(r => {
    const t = (r.itemType || '').replace(/ Delayed$/, '') || '배송';
    const k = quickProductName(r.description) + '|' + t;
    counts.set(k, (counts.get(k) || 0) + 1);
  });
  const title = [...counts.entries()].map(([k, n]) => { const [p, t] = k.split('|'); return p + ' ' + n + '개 ' + t; }).join(', ');
  const r0 = rows[0];
  const lines = String(r0.address || '').split('\n').map(s => s.trim()).filter(Boolean);
  const keep = lines.map(l => l.replace(/\(?\s*cust#[^)]*\)?/gi, '').trim()).filter(Boolean);   // (cust# ...) 부분만 지우고 나머지(예: 2F)는 유지
  return {
    title: title,
    company: keep[0] || '',
    address: keep.slice(1).join('\n'),
    name: String(r0.customer || '').split('\n')[0].trim(),
    phone: String(r0.phone || '').split('\n')[0].trim(),
    note: ''
  };
}
function buildQuickSheet(d){
  const el = document.createElement('div');
  el.className = 'qs-sheet';
  const add = (cls, text) => { if(!String(text || '').trim()) return; const e = document.createElement('div'); e.className = cls; e.textContent = text; el.appendChild(e); };
  add('qs-title', d.title);
  add('qs-company', d.company);
  add('qs-addr', d.address);
  add('qs-name', d.name ? d.name + ' 님' : '');
  if(String(d.phone || '').trim()){
    const p = document.createElement('div'); p.className = 'qs-phone';
    const i = document.createElement('span'); i.className = 'qs-tel'; i.textContent = '☎';
    p.appendChild(i); p.appendChild(document.createTextNode(' ' + d.phone)); el.appendChild(p);
  }
  if(String(d.note || '').trim()){ const n = document.createElement('div'); n.className = 'qs-note'; n.textContent = d.note; el.appendChild(n); }
  const logo = document.createElement('img'); logo.className = 'qs-logo'; logo.alt = 'Bloomberg'; logo.src = window.__QUICK_LOGO__ || '/api/quick/logo';
  el.insertBefore(logo, el.firstChild);   // 로고는 제목 위(맨 위)
  return el;
}
let QUICK_ROWS = [];
function openQuickPrint(){
  const keys = [...document.querySelectorAll('.row-check:checked')].map(cb => cb.dataset.rowKey);
  QUICK_ROWS = ROWS.filter(r => keys.includes(r.rowKey));
  if(!QUICK_ROWS.length){ toast('퀵으로 보낼 행을 먼저 체크하세요'); return; }
  const d = quickDefaults(QUICK_ROWS);
  let ov = document.getElementById('quickOverlay');
  if(ov) ov.remove();
  ov = document.createElement('div');
  ov.id = 'quickOverlay'; ov.className = 'trash-overlay show';
  ov.innerHTML = '<div class="quick-modal"><div class="quick-head"><b>퀵 기사님 안내문</b><span>내용을 확인·수정하고 인쇄하세요 (A4 세로)</span><button type="button" class="quick-x" onclick="closeQuickPrint()">×</button></div>'
    + '<div class="quick-body"><div class="quick-form">'
    + '<label>내용 <i>(예: 키보드 1개 배송 / 키보드 1개 회수 · 줄바꿈 가능)</i><textarea id="qk_title" rows="2"></textarea></label>'
    + '<div class="quick-ref"><b>보드 원문 (영문 주소 · 참고용)</b><pre id="qk_ref"></pre></div>'
    + '<label>회사명 <i>(한글로 입력)</i><input id="qk_company"></label>'
    + '<label>주소<textarea id="qk_address" rows="3"></textarea></label>'
    + '<div id="qk_sugg" class="quick-sugg"></div>'
    + '<div class="quick-two"><label>고객명<input id="qk_name"></label><label>전화번호<input id="qk_phone"></label></div>'
    + '<label>요청 사항 <i>(예: 5시 전까지 배송해주세요)</i><textarea id="qk_note" rows="2"></textarea></label>'
    + '<div class="quick-actions"><button type="button" class="btn primary" onclick="printQuick()">인쇄</button><button type="button" class="btn" onclick="closeQuickPrint()">닫기</button></div>'
    + '</div><div class="quick-prev"><div class="quick-paper" id="qk_paper"></div></div></div></div>';
  document.body.appendChild(ov);
  ['title', 'company', 'address', 'name', 'phone', 'note'].forEach(k => {
    const e = document.getElementById('qk_' + k); e.value = d[k] || '';
    e.addEventListener('input', renderQuickPreview);
  });
  ov.addEventListener('mousedown', e => { if(e.target === ov) closeQuickPrint(); });
  document.getElementById('qk_ref').textContent = String(QUICK_ROWS[0].address || '').trim() || '(주소 없음)';
  renderQuickPreview();
  const en = (QUICK_ROWS[0].address || '').split('\n')[0].trim();
  if(en && /[A-Za-z]/.test(en)){
    fetch('/api/quick/suggest?company=' + encodeURIComponent(en)).then(r => r.json()).then(s => {
      const box = document.getElementById('qk_sugg'); if(!box) return;
      (s.suggestions || []).forEach(x => {
        const b = document.createElement('button'); b.type = 'button'; b.className = 'quick-chip';
        b.textContent = '이전 주소: ' + x.text + (x.date ? ' (' + x.date + ')' : ''); b.title = '클릭하면 주소 칸에 채웁니다';
        b.addEventListener('click', () => { document.getElementById('qk_address').value = x.text; renderQuickPreview(); });
        box.appendChild(b);
      });
    }).catch(() => {});
  }
}
function quickFormData(){
  const g = k => document.getElementById('qk_' + k).value;
  return { title: g('title'), company: g('company'), address: g('address'), name: g('name'), phone: g('phone'), note: g('note') };
}
function renderQuickPreview(){
  const paper = document.getElementById('qk_paper'); if(!paper) return;
  paper.textContent = '';
  paper.appendChild(buildQuickSheet(quickFormData()));
}
function closeQuickPrint(){ const ov = document.getElementById('quickOverlay'); if(ov) ov.remove(); }
function printQuick(){
  const area = document.getElementById('printArea');
  area.textContent = '';
  area.appendChild(buildQuickSheet(quickFormData()));
  const st = document.createElement('style');
  st.id = 'quickPrintStyle';
  st.textContent = '@page{size:A4 portrait;margin:0}';
  document.head.appendChild(st);
  const done = () => { st.remove(); area.textContent = ''; window.removeEventListener('afterprint', done); };
  window.addEventListener('afterprint', done);
  window.print();
}

// ---------- 서명 용지 POD / POC 인쇄 (2026-10-14) ----------
// POD = Proof of Delivery (ZOR 배송만), POC = Proof of Collection (ZRE 회수). 교체(ZRX)/내부배송(ZINT·ZINP 등)은 대상 아님.
// 체크박스(pod/poc)는 모아보기 줄에 있고 브라우저에 기억된다.
//  - [인쇄]: 체크한 오더행 중 pod 칸이 켜져 있으면 ZOR 오더의 POD, poc 칸이 켜져 있으면 ZRE 오더의 POC를 인쇄
//  - pod 칸이 켜져 있으면 "Serial 등록 & QR 인쇄"가 오더 하나 끝날 때마다 그 오더의 POD도 자동 인쇄(ZOR만)
function formFlag(k){ try{ return localStorage.getItem('wbForm_' + k) === '1'; }catch(e){ return false; } }
function setFormFlag(k, on){ try{ localStorage.setItem('wbForm_' + k, on ? '1' : '0'); }catch(e){} }
function initFormBoxes(){
  ['pod', 'poc'].forEach(k => {
    const el = document.getElementById('chk_' + k);
    if(!el) return;
    el.checked = formFlag(k);
    el.addEventListener('change', () => setFormFlag(k, el.checked));
  });
}
function formPrefix(label){ return String(label || '').trim().split(/\s+/)[0].toUpperCase(); }
function buildFormOrder(kind, orderId){
  const same = ROWS.filter(r => r.orderId === orderId);
  if(!same.length) return null;
  const want = kind === 'pod' ? /^배송/ : /^회수/;
  let rows = same.filter(r => want.test(r.itemType || ''));
  if(!rows.length) rows = same;
  const r0 = rows[0];
  const prefix = formPrefix(r0.orderLabel);
  const lines = String(r0.address || '').split('\n').map(s => s.trim()).filter(Boolean)
    .map(l => l.replace(/\(?\s*cust#[^)]*\)?/gi, '').trim()).filter(Boolean);
  const num = String(r0.orderNo || '').trim() || (String(r0.orderLabel || '').match(/\d+/) || [''])[0];
  return {
    prefix: prefix, number: num, date: r0.date || '',
    company: lines[0] || '', address: lines.slice(1).join('\n'),
    attn: String(r0.customer || '').split('\n')[0].trim(),
    items: rows.map(r => ({ material: r.material || '', description: r.description || '', qty: 1 }))
  };
}
async function sendFormPrint(kind, order){
  const res = await fetch('/api/forms/print', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ kind: kind, order: order }) });
  const data = await res.json();
  if(!res.ok || data.error) throw new Error(data.error || '인쇄 요청 실패');
  return data;
}
function formPages(kind, order){ const per = kind === 'pod' ? 5 : 7; const uniq = new Set(order.items.map(i => i.material + '|' + i.description)); return Math.max(1, Math.ceil(uniq.size / per)); }
async function printSignForms(){
  const wantPod = document.getElementById('chk_pod').checked;
  const wantPoc = document.getElementById('chk_poc').checked;
  if(!wantPod && !wantPoc){ toast('pod 또는 poc 칸을 먼저 체크하세요'); return; }
  const keys = [...document.querySelectorAll('.row-check:checked')].map(cb => cb.dataset.rowKey);
  const checked = ROWS.filter(r => keys.includes(r.rowKey));
  if(!checked.length){ toast('인쇄할 오더 행을 먼저 체크하세요'); return; }
  const ids = [...new Set(checked.map(r => r.orderId))];
  const jobs = [];
  ids.forEach(id => {
    const prefix = formPrefix((checked.find(r => r.orderId === id) || {}).orderLabel);
    if(wantPod && prefix === 'ZOR'){ const o = buildFormOrder('pod', id); if(o) jobs.push(['pod', o]); }
    if(wantPoc && prefix === 'ZRE'){ const o = buildFormOrder('poc', id); if(o) jobs.push(['poc', o]); }
  });
  if(!jobs.length){ toast('선택한 오더 중 인쇄 대상이 없습니다 (POD는 ZOR, POC는 ZRE 오더만)'); return; }
  const summary = jobs.map(([k, o]) => k.toUpperCase() + ' · ' + o.prefix + ' ' + o.number + ' (' + o.items.length + '품목, ' + formPages(k, o) + '장)').join('\n');
  if(!window.confirm('다음 서명 용지를 인쇄합니다:\n\n' + summary + '\n\n진행할까요?')) return;
  let ok = 0;
  for(const [k, o] of jobs){
    try{ const d = await sendFormPrint(k, o); if(d.started) ok++; else toast(k.toUpperCase() + ' ' + o.number + ': ' + (d.message || '건너뜀')); }
    catch(err){ toast(k.toUpperCase() + ' ' + o.number + ' 인쇄 실패: ' + err.message); }
  }
  if(ok) toast('서명 용지 ' + ok + '건 인쇄 시작 (Excel이 백그라운드에서 인쇄합니다)');
}
// "Serial 등록 & QR 인쇄"가 한 오더를 끝낼 때 호출 - pod 칸이 켜져 있고 ZOR 오더면 POD 자동 인쇄(확인창 없음)
async function autoPrintPod(orderId){
  const el = document.getElementById('chk_pod');
  if(!el || !el.checked) return;
  const row = ROWS.find(r => r.orderId === orderId);
  if(!row || formPrefix(row.orderLabel) !== 'ZOR') return;
  const o = buildFormOrder('pod', orderId);
  if(!o) return;
  try{
    const d = await sendFormPrint('pod', o);
    toast(d.started ? ('POD 서명 용지 인쇄 시작: ' + o.prefix + ' ' + o.number) : ('POD 인쇄 건너뜀: ' + (d.message || '')));
  }catch(err){ toast('POD 인쇄 실패: ' + err.message); }
}
initFormBoxes();

// ---------- 실행취소 / 다시실행 (2026-10-14) ----------
// 같은 이름으로 다시 선언해 v1의 undoLastAction을 대체한다: 되돌린 내용을 새로고침 뒤에도 알림으로 보여주고,
// "순서 변경"을 되돌린 경우엔 상태 변경이 그대로 남아 있다는 설명을 덧붙인다.
function toastLong(msg, ms){
  const t = document.createElement('div');
  t.className = 'toast'; t.textContent = msg;
  document.body.appendChild(t);
  requestAnimationFrame(() => t.classList.add('show'));
  setTimeout(() => { t.classList.remove('show'); setTimeout(() => t.remove(), 300); }, ms || 7000);
}
function showAfterReload(msg){ try{ sessionStorage.setItem('wbUndoMsg', msg); }catch(e){} }
async function undoLastAction(){
  try{
    const res = await fetch('/api/undo', { method: 'POST' });
    const data = await res.json();
    if(!res.ok || data.error) throw new Error(data.error || '실행취소 실패');
    if(!data.undone){ toast(data.message || '되돌릴 작업이 없습니다'); return; }
    let msg = '실행취소: ' + data.summary + ' (' + data.affected + '건)';
    if(/순서 변경/.test(data.summary)) msg += ' - 상태 변경은 그대로 남아 있습니다. 한 번 더 누르면 상태 변경도 되돌립니다';
    showAfterReload(msg + ' · ↷ 으로 다시 실행');
    location.reload();
  }catch(err){ toast('실행취소 실패: ' + err.message); }
}
async function redoLastAction(){
  try{
    const res = await fetch('/api/redo', { method: 'POST' });
    const data = await res.json();
    if(!res.ok || data.error) throw new Error(data.error || '다시실행 실패');
    if(!data.redone){ toast(data.message || '다시 실행할 작업이 없습니다'); return; }
    showAfterReload('다시실행: ' + data.summary + ' (' + data.affected + '건)');
    location.reload();
  }catch(err){ toast('다시실행 실패: ' + err.message); }
}
// Ctrl+Shift+Z / Ctrl+Y = 다시실행 (v1의 Ctrl+Z 처리보다 먼저 가로챔 - Shift+Z가 실행취소로 처리되지 않게)
window.addEventListener('keydown', e => {
  if(!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if(!((k === 'z' && e.shiftKey) || k === 'y')) return;
  const el = document.activeElement;
  if(el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)) return;
  e.preventDefault(); e.stopImmediatePropagation();
  redoLastAction();
}, true);
try{
  const m = sessionStorage.getItem('wbUndoMsg');
  if(m){ sessionStorage.removeItem('wbUndoMsg'); setTimeout(() => toastLong(m, 8000), 500); }
}catch(e){}

// ---------- 마감된 과거 날짜 이력 보기/수정 (2026-10-15) ----------
// 달력에서 보드에 없는 과거 날짜를 누르면 history.db의 그 날짜 시트를 보여준다(엑셀의 칸 색/병합도 표시).
// 칸을 고치면 "고친 셀 하나만" 원본 배송장 엑셀에 쓴다(시트 전체를 다시 쓰지 않음).
let DAYH_EL = null, DAYH_PENDING = {}, DAYH_T = null;
const DAYH_LABELS = ['오더/구분', '자재', 'M/N', 'S/N', '고객', '연락처', '주소', '메모'];
const DAYH_COLOR_NAMES = { '00B0F0': '완료(파랑)', 'D9D9D9': '취소(회색)', '92D050': 'billed(초록)', 'FFFF00': '노랑 표시', 'FFC0CB': '분홍 표시', 'FBE3D6': '헤더색' };
function dayhEl(){
  if(!DAYH_EL){
    DAYH_EL = document.createElement('div');
    DAYH_EL.className = 'dayh-back';
    DAYH_EL.addEventListener('mousedown', e => { if(e.target === DAYH_EL) closeDayHistory(); });
    document.body.appendChild(DAYH_EL);
    document.addEventListener('keydown', e => { if(e.key === 'Escape' && DAYH_EL && DAYH_EL.style.display !== 'none') closeDayHistory(); });
  }
  return DAYH_EL;
}
function closeDayHistory(){
  if(Object.keys(DAYH_PENDING).length) dayhFlush();
  if(DAYH_EL) DAYH_EL.style.display = 'none';
}
function _dayhLetter(i){ return String.fromCharCode(65 + i); }
async function openDayHistory(iso, sheetId){
  const el = dayhEl();
  el.style.display = 'flex';
  el.innerHTML = '<div class="dayh-box"><div class="dayh-head"><b>' + _esc(iso || '시트') + '</b><span>불러오는 중…</span></div></div>';
  let d;
  try { d = await fetch('/api/history/day?' + (sheetId ? 'sheet_id=' + sheetId : 'date=' + encodeURIComponent(iso))).then(r => r.json()); } catch (e) { d = null; }
  if(!d || !d.sheets){ el.innerHTML = '<div class="dayh-box"><div class="dayh-head"><b>' + _esc(iso || '시트') + '</b><span>불러오기 실패</span><button type="button" class="dayh-x">×</button></div></div>'; el.querySelector('.dayh-x').onclick = closeDayHistory; return; }
  const title = iso ? _esc(iso) + ' (' + ['일', '월', '화', '수', '목', '금', '토'][new Date(iso + 'T00:00:00').getDay()] + ')' : _esc((d.sheets[0] || {}).sheetName || '시트');
  let h = '<div class="dayh-box"><div class="dayh-head"><b>' + title + '</b>'
    + '<span>칸을 고치면 해당 칸만 배송장 엑셀에 바로 반영됩니다 · 칸 색은 엑셀의 색(파랑=완료, 회색=취소, 초록=billed)</span><i id="dayhState"></i><button type="button" class="dayh-x" title="닫기 (Esc)">×</button></div><div class="dayh-scroll">';
  if(!d.sheets.length) h += '<div class="hist-empty">이 날짜의 이력이 색인에 없습니다</div>';
  d.sheets.forEach(sh => { h += dayhSheetHtml(sh); });
  h += '</div></div>';
  el.innerHTML = h;
  el.querySelector('.dayh-x').onclick = closeDayHistory;
  el.querySelectorAll('.dayh-c').forEach(td => {
    td.addEventListener('blur', () => dayhCommit(td));
    td.addEventListener('keydown', e => { if(e.key === 'Enter' && !e.shiftKey){ e.preventDefault(); td.blur(); } });
  });
}
function dayhSheetHtml(sh){
  let ncol = sh.noHeader ? 1 : DAYH_LABELS.length;
  sh.rows.forEach(r => { ncol = Math.max(ncol, Math.min(14, r.cells.length)); });
  const merges = sh.merges || [];
  const rowNums = new Set(sh.rows.map(r => r.excelRow));
  const anchor = {}, covered = new Set();
  merges.forEach(m => {
    const r1 = m[0], c1 = m[1], r2 = m[2], c2 = m[3];
    if(!rowNums.has(r1)) return;                       // 시작 행이 표시되지 않으면 병합 무시(칸이 사라지는 것 방지)
    let rs = 0; sh.rows.forEach(r => { if(r.excelRow >= r1 && r.excelRow <= r2) rs++; });
    anchor[r1 + ':' + c1] = { rs: rs, cs: c2 - c1 + 1 };
    sh.rows.forEach(r => { if(r.excelRow < r1 || r.excelRow > r2) return; for(let c = c1; c <= c2; c++){ if(!(r.excelRow === r1 && c === c1)) covered.add(r.excelRow + ':' + c); } });
  });
  let h = '<div class="dayh-src">' + _esc(sh.sourceFile) + ' · 시트 ' + _esc(sh.sheetName) + ' · ' + sh.rows.length + '행' + (sh.noHeader ? ' · 머리글 없는 시트' : '') + '</div><div class="dayh-wrap"><table class="dayh-tbl"><tr><th></th>';
  for(let c = 0; c < ncol; c++) h += '<th>' + _esc(sh.noHeader ? _dayhLetter(c) : (DAYH_LABELS[c] || _dayhLetter(c))) + '</th>';
  h += '</tr>';
  let last = null;
  sh.rows.forEach(r => {
    const brk = last !== null && r.groupNo !== last ? ' dayh-brk' : '';
    last = r.groupNo;
    const rc = r.fills && r.fills['0'];
    const kStyle = rc ? ' style="border-left:6px solid #' + rc + '"' : '';
    const kTip = rc ? (DAYH_COLOR_NAMES[rc] || ('엑셀 색 #' + rc)) : '';
    h += '<tr class="' + brk + '"><td class="dayh-k k-' + _esc(r.kind) + '"' + kStyle + (kTip ? ' title="' + _esc(kTip) + '"' : '') + '>' + _esc(r.kind || '') + '</td>';
    for(let c = 0; c < ncol; c++){
      const key = r.excelRow + ':' + (c + 1);
      if(covered.has(key)) continue;
      const a = anchor[key];
      const v = r.cells[c] || '';
      const f = r.fills && r.fills[String(c)];
      const bg = f && f !== rc ? ' style="background:#' + f + '55"' : (f ? ' style="background:#' + f + '22"' : '');
      h += '<td class="dayh-c" contenteditable="plaintext-only" spellcheck="false"' + bg + (a && a.rs > 1 ? ' rowspan="' + a.rs + '"' : '') + (a && a.cs > 1 ? ' colspan="' + a.cs + '"' : '')
        + ' data-row="' + r.rowId + '" data-col="' + c + '" data-orig="' + _esc(v).replace(/"/g, '&quot;') + '">' + _esc(v) + '</td>';
    }
    h += '</tr>';
  });
  return h + '</table></div>';
}
// 날짜를 못 읽은 시트(메모/이벤트/주소록 등) 목록
async function openMiscSheets(){
  const el = dayhEl();
  el.style.display = 'flex';
  el.innerHTML = '<div class="dayh-box"><div class="dayh-head"><b>날짜 없는 시트 (메모 · 이벤트 · 주소록 등)</b><button type="button" class="dayh-x" title="닫기 (Esc)">×</button></div><div class="dayh-scroll" id="dayhMisc">불러오는 중…</div></div>';
  el.querySelector('.dayh-x').onclick = closeDayHistory;
  let d; try { d = await fetch('/api/history/misc').then(r => r.json()); } catch (e) { d = null; }
  const host = document.getElementById('dayhMisc'); if(!host) return;
  if(!d || !d.sheets){ host.textContent = '불러오기 실패'; return; }
  host.innerHTML = '<div class="dayh-misc">' + d.sheets.map(s => '<button type="button" onclick="openDayHistory(\'\',' + s.sheetId + ')"><b>' + _esc(s.sheetName.trim() || '(이름 없음)') + '</b><span>' + _esc(s.sourceFile.replace(/\s*배송장.*$/, '')) + ' · ' + s.rows + '행</span></button>').join('') + '</div>';
}
function dayhCommit(td){
  const val = td.innerText.replace(/ /g, ' ').replace(/\n+$/, '').trim();
  if(val === td.dataset.orig) return;
  const key = td.dataset.row + ':' + td.dataset.col;
  DAYH_PENDING[key] = { rowId: +td.dataset.row, col: +td.dataset.col, value: val, td: td };
  td.classList.add('dayh-saving');
  clearTimeout(DAYH_T); DAYH_T = setTimeout(dayhFlush, 500);
}
async function dayhFlush(){
  clearTimeout(DAYH_T);
  const batch = DAYH_PENDING; DAYH_PENDING = {};
  const items = Object.values(batch);
  if(!items.length) return;
  const st = document.getElementById('dayhState');
  if(st) st.textContent = '엑셀에 저장 중… (큰 파일은 몇 초 걸릴 수 있음)';
  let res;
  try {
    res = await fetch('/api/history/edit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ edits: items.map(i => ({ rowId: i.rowId, col: i.col, value: i.value })) }) }).then(r => r.json());
  } catch (e) { res = { results: [] }; }
  let bad = 0;
  items.forEach(i => {
    const r = (res.results || []).find(x => x.rowId === i.rowId && x.col === i.col);
    i.td.classList.remove('dayh-saving');
    if(r && r.ok){ i.td.dataset.orig = i.value; i.td.classList.remove('dayh-err'); i.td.classList.add('dayh-ok'); setTimeout(() => i.td.classList.remove('dayh-ok'), 1500); i.td.title = ''; }
    else { bad++; i.td.classList.add('dayh-err'); i.td.title = (r && r.error) || '저장 실패'; }
  });
  if(st) st.textContent = bad ? ('⚠ ' + bad + '칸 저장 실패 (빨간 칸에 마우스를 올리면 이유가 보입니다)') : ('✓ ' + items.length + '칸 저장됨');
}

// 2026-10-16: 상단 [📱 모바일] - 폰 크기 /m 창을 오른쪽 모니터에 띄운다(mobile_window.py)
async function openMobileWindow(){
  try{
    const res=await fetch('/api/mobile/open',{method:'POST'});
    const data=await res.json();
    if(!res.ok||data.error) throw new Error(data.error||'실패');
    toast('모바일 창을 오른쪽 모니터에 띄웁니다 (몇 초 걸립니다)');
  }catch(err){ toast('모바일 창 열기 실패: '+err.message); }
}

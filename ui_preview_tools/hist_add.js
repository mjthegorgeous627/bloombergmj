
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
  const w = Math.min(960, window.innerWidth - 20);
  el.style.width = w + 'px';
  el.style.left = Math.max(10, Math.min(r.left, window.innerWidth - w - 10)) + 'px';
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

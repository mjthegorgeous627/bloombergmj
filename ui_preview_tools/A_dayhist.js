
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

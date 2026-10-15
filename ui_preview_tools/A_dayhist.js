
// ---------- 마감된 과거 날짜 이력 보기/수정 (2026-10-15) ----------
// 달력에서 보드에 없는 과거 날짜를 누르면 history.db의 그 날짜 시트를 보여준다.
// 칸을 고치면 "고친 셀 하나만" 원본 배송장 엑셀에 쓴다(시트 전체를 다시 쓰지 않음).
let DAYH_EL = null, DAYH_PENDING = {}, DAYH_T = null;
const DAYH_LABELS = ['오더/구분', '자재', 'M/N', 'S/N', '고객', '연락처', '주소', '메모'];
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
async function openDayHistory(iso){
  const el = dayhEl();
  el.style.display = 'flex';
  el.innerHTML = '<div class="dayh-box"><div class="dayh-head"><b>' + _esc(iso) + '</b><span>불러오는 중…</span></div></div>';
  let d;
  try { d = await fetch('/api/history/day?date=' + encodeURIComponent(iso)).then(r => r.json()); } catch (e) { d = null; }
  if(!d || !d.sheets){ el.innerHTML = '<div class="dayh-box"><div class="dayh-head"><b>' + _esc(iso) + '</b><span>불러오기 실패</span><button type="button" class="dayh-x">×</button></div></div>'; el.querySelector('.dayh-x').onclick = closeDayHistory; return; }
  const wd = ['일', '월', '화', '수', '목', '금', '토'][new Date(iso + 'T00:00:00').getDay()];
  let h = '<div class="dayh-box"><div class="dayh-head"><b>' + _esc(iso) + ' (' + wd + ')</b>'
    + '<span>칸을 고치면 해당 칸만 배송장 엑셀에 바로 반영됩니다</span><i id="dayhState"></i><button type="button" class="dayh-x" title="닫기 (Esc)">×</button></div><div class="dayh-scroll">';
  if(!d.sheets.length) h += '<div class="hist-empty">이 날짜의 이력이 색인에 없습니다</div>';
  d.sheets.forEach(sh => {
    h += '<div class="dayh-src">' + _esc(sh.sourceFile) + ' · 시트 ' + _esc(sh.sheetName) + ' · ' + sh.rows.length + '행</div><table class="dayh-tbl"><tr><th></th>'
      + DAYH_LABELS.map(l => '<th>' + l + '</th>').join('') + '</tr>';
    let last = null;
    sh.rows.forEach(r => {
      const brk = last !== null && r.groupNo !== last ? ' dayh-brk' : '';
      last = r.groupNo;
      h += '<tr class="' + brk + '"><td class="k-' + _esc(r.kind) + '">' + _esc(r.kind || '') + '</td>';
      for(let c = 0; c < DAYH_LABELS.length; c++){
        const v = r.cells[c] || '';
        h += '<td class="dayh-c" contenteditable="plaintext-only" spellcheck="false" data-row="' + r.rowId + '" data-col="' + c + '" data-orig="' + _esc(v).replace(/"/g, '&quot;') + '">' + _esc(v) + '</td>';
      }
      h += '</tr>';
    });
    h += '</table>';
  });
  h += '</div></div>';
  el.innerHTML = h;
  el.querySelector('.dayh-x').onclick = closeDayHistory;
  el.querySelectorAll('.dayh-c').forEach(td => {
    td.addEventListener('blur', () => dayhCommit(td));
    td.addEventListener('keydown', e => { if(e.key === 'Enter' && !e.shiftKey){ e.preventDefault(); td.blur(); } });
  });
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

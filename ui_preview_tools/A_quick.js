
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

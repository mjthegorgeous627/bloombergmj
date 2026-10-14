
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

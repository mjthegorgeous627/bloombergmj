
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

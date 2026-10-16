// -- Item-level persistence helpers, shared by status/move actions below.
// Each posts to workbench.db first and only touches ROWS/renders on success,
// same pattern as deleteChecked() - never leave the screen out of sync with
// what actually got saved. --
async function persistItemStatus(itemIds, status){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/status', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, status})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'상태 저장 실패');
  return true;
}
// Done/Ready/Cancel/초기화 버튼은 사용자가 그 행을 직접 확인/처리했다는
// 신호이므로, set_items_board_status()가 서버(orders/order_items 테이블)에서
// 이미 지운 portal_error/pod_error/sn_relo_error를 화면에도 그 자리에서 함께
// 지운다 - 안 그러면 최대 1.5초짜리 board_version 폴링(5301행)이 돌 때까지
// 주황(row-error)이 그대로 남아 "버튼 눌러도 에러 표시가 안 없어진다"는 오해를
// 만든다(2026-09-02 사용자 지적: 수동으로 문제를 고치고 이 버튼들을 눌러도
// 자동화가 스스로 재시도해서 성공할 때만 지워지는 것처럼 보였음). portalError/
// podError는 orders 테이블 값이라 같은 오더의 다른 행에도 동일하게 붙어있으므로
// (board_row()의 해당 주석 참고) 체크된 항목뿐 아니라 같은 orderId를 가진 모든
// 행에서 지워야 서버 상태와 어긋나지 않는다; snReloError는 항목 단위라 체크된
// 항목 자신만 지운다.
function clearRowErrorFlags(itemIds){
  const idSet=new Set(itemIds);
  const orderIds=new Set(ROWS.filter(r=>idSet.has(r.itemId)).map(r=>r.orderId));
  ROWS.forEach(r=>{
    if(orderIds.has(r.orderId)){ r.portalError=false; r.portalErrorMsg=''; r.podError=false; r.podErrorMsg=''; }
    if(idSet.has(r.itemId)){ r.snReloError=false; r.snReloErrorMsg=''; }
  });
}
async function persistItemDate(itemIds, iso){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/move_date', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, date: iso})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'날짜 이동 저장 실패');
  return true;
}
async function persistItemReorder(itemIds){
  if(!itemIds.length)return true;
  const res=await fetch('/api/items/reorder', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'순서 저장 실패');
  return true;
}
// -- Inline-edit persistence: description/material/serial (per item) and
// customer/phone/address/memo (per order, resolved from the edited item ids -
// a shared cell can span multiple orders once merged). These cells used to
// be screen-only (ROWS-only, lost on reload/restart) - exactly the gap that
// silently ate a typed-in serial number. --
async function persistItemField(itemId, field, value){
  const res=await fetch('/api/items/edit_field', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemId, field, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}
async function persistOrderField(itemIds, field, value){
  const res=await fetch('/api/orders/edit_field', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, field, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  if(field==='memo' && String(value||'').includes('퀵')) autoQuickType(itemIds);
  return true;
}
// 2026-10-16: 메모에 '퀵'이 있으면 배송→배송_퀵, 회수→회수_퀵 (다른 Type은 그대로). 직접 다시 고를 수도 있음.
function autoQuickType(itemIds){
  let changed=false;
  ROWS.filter(r=>itemIds.includes(r.itemId)).forEach(r=>{
    const next=r.itemType==='배송'?'배송_퀵':r.itemType==='회수'?'회수_퀵':'';
    if(!next) return;
    r.itemType=next; changed=true;
    persistItemField(r.itemId,'item_type',next).catch(err=>toast('Type 저장 실패: '+err.message));
  });
  if(changed) setTimeout(()=>{ if(typeof renderAll==='function') renderAll(); },0);
}
async function persistCustomType(name){
  const res=await fetch('/api/custom_types/add', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}
async function persistDeliveryCodes(itemIds, value){
  const res=await fetch('/api/orders/edit_delivery_codes', {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, value})
  });
  const data=await res.json();
  if(!res.ok || data.error) throw new Error(data.error||'저장 실패');
  return true;
}

// -- Undo: reverts the most recent delete/move/status batch in workbench.db
// (server-side history, capped to the last 20 actions). Reloads the page
// afterward instead of patching ROWS in place - undo can touch items that
// aren't even on the currently-loaded board (e.g. undoing a delete brings an
// item's date section back), so a full reload is the simple, always-correct
// way to pick up whatever state the server ends up in. --
async function undoLastAction(){
  try{
    const res=await fetch('/api/undo', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행취소 실패');
    if(!data.undone){ toast(data.message||'되돌릴 작업이 없습니다'); return }
    toast('실행취소: '+data.summary);
    location.reload();
  }catch(err){
    toast('실행취소 실패: '+err.message);
  }
}
document.addEventListener('keydown', e=>{
  if((e.ctrlKey || e.metaKey) && e.key.toLowerCase()==='z'){
    const el=document.activeElement;
    const isEditing = el && (el.tagName==='INPUT' || el.tagName==='TEXTAREA' || el.isContentEditable);
    if(isEditing) return; // let the browser handle undo inside a text field being edited
    e.preventDefault();
    undoLastAction();
  }
});

// -- Trash: browse/restore soft-deleted items outside Ctrl+Z's
// UNDO_HISTORY_LIMIT-action reach - anything the server still returns here
// is, by definition, within DELETED_ITEM_RETENTION_DAYS (not yet hard-purged).
// Built with createElement/textContent throughout, not innerHTML, same as
// every other cell in this app - order/customer/memo text can contain
// anything typed into SAP or the sheet, so it's never treated as HTML. --
async function openTrash(){
  document.getElementById('trashOverlay').classList.add('show');
  await refreshTrashList();
}
function closeTrash(){
  document.getElementById('trashOverlay').classList.remove('show');
}
async function refreshTrashList(){
  const container=document.getElementById('trashList');
  container.textContent='불러오는 중...';
  try{
    const res=await fetch('/api/items/deleted');
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
    renderTrashList(data.items||[]);
  }catch(err){
    container.textContent='조회 실패: '+err.message;
  }
}
function daysLeftText(deletedAt){
  const deleted=new Date(String(deletedAt||'').replace(' ','T'));
  if(isNaN(deleted.getTime())) return '';
  const daysLeft=Math.max(0, DELETED_ITEM_RETENTION_DAYS - Math.floor((Date.now()-deleted.getTime())/86400000));
  return daysLeft+'일 남음';
}
function renderTrashList(items){
  const container=document.getElementById('trashList');
  container.textContent='';
  if(!items.length){
    const empty=document.createElement('div');
    empty.className='trash-empty';
    empty.textContent='삭제된 항목이 없습니다.';
    container.appendChild(empty);
    return;
  }
  const table=document.createElement('table');
  table.className='trash-table';
  const thead=document.createElement('thead');
  const headRow=document.createElement('tr');
  ['', 'Order#', 'Customer', 'Item', 'M/N', 'S/N', '삭제 일시', '남은 기간'].forEach(label=>{
    const th=document.createElement('th'); th.textContent=label; headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);
  const tbody=document.createElement('tbody');
  items.forEach(item=>{
    const tr=document.createElement('tr');
    const checkTd=document.createElement('td');
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.className='trash-check'; cb.value=item.id;
    checkTd.appendChild(cb);
    tr.appendChild(checkTd);
    [
      `${item.order_type||''} ${item.order_no||''}`.trim(),
      item.customer||'',
      item.description||'',
      item.material||'',
      item.serial||'',
      item.deleted_at||'',
      daysLeftText(item.deleted_at),
    ].forEach(text=>{
      const td=document.createElement('td'); td.textContent=text; tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  container.appendChild(table);
}
async function restoreCheckedTrash(){
  const itemIds=[...document.querySelectorAll('.trash-check:checked')].map(cb=>parseInt(cb.value, 10));
  if(!itemIds.length){toast('복구할 항목을 선택하세요');return}
  try{
    const res=await fetch('/api/items/restore', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'복구 실패');
    toast(data.restored+'건 복구됨');
  }catch(err){toast('복구 실패: '+err.message);return}
  location.reload();
}

// 2026-10-13: Done(파랑)/Ready(빨강)/Cancel(회색)/billed(초록) 처리된 행은 "일단락"이라
// 해당 날짜 섹션의 맨 위로 처리한 순서대로 쌓는다 - 먼저 처리한 게 맨 위,
// 나중에 처리한 게 그 아래. 이미 쌓여있던 행은 그대로 두고, 이번에 새로 처리된
// 행을 그 뒤에 붙인 뒤 나머지(미처리) 행이 이어짐. 순서는 reorder API로 저장.
const SETTLED_STATUSES=new Set(['done','pending','cancelled','billed']);
async function stackSettledToTop(newItemIds){
  const newSet=new Set(newItemIds);
  const dates=[...new Set(ROWS.filter(r=>newSet.has(r.itemId)).map(r=>r.date))];
  for(const iso of dates){
    const dayRows=ROWS.filter(r=>r.date===iso);
    const isSettled=r=>SETTLED_STATUSES.has(r.uiStatus);
    const prev=dayRows.filter(r=>isSettled(r) && !newSet.has(r.itemId));
    const fresh=dayRows.filter(r=>newSet.has(r.itemId) && isSettled(r));
    const rest=dayRows.filter(r=>!prev.includes(r) && !fresh.includes(r));
    const ordered=[...prev, ...fresh, ...rest];
    const firstIdx=ROWS.findIndex(r=>r.date===iso);
    ROWS=ROWS.filter(r=>r.date!==iso);
    ROWS.splice(firstIdx, 0, ...ordered);
    try{ await persistItemReorder(ordered.map(r=>r.itemId)); }
    catch(err){ toast('순서 저장 실패: '+err.message); }
  }
}
// -- Process: acts on whichever rows are checked, same checkboxes Portal uses --
async function processDone(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, 'done');
  }catch(err){toast('완료 처리 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='done'});
  clearRowErrorFlags(itemIds);
  await stackSettledToTop(itemIds);
  renderAll();
  toast(keys.length+'건 완료 처리(파랑) - DB에 반영됨');
}
async function processPickupReady(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>keys.includes(r.rowKey));
  const pendingIds=checkedRows.filter(r=>(r.itemType==='회수'||r.itemType==='회수_퀵')).map(r=>r.itemId);
  const doneIds=checkedRows.filter(r=>!(r.itemType==='회수'||r.itemType==='회수_퀵')).map(r=>r.itemId);
  try{
    await Promise.all([persistItemStatus(pendingIds,'pending'), persistItemStatus(doneIds,'done')]);
  }catch(err){toast('상태 저장 실패: '+err.message);return}
  checkedRows.forEach(row=>{row.uiStatus = (row.itemType==='회수'||row.itemType==='회수_퀵') ? 'pending' : 'done'});
  clearRowErrorFlags([...pendingIds, ...doneIds]);
  await stackSettledToTop([...pendingIds, ...doneIds]);
  renderAll();
  toast(keys.length+'건 처리: 회수는 대기(빨강), 나머지는 완료(파랑) - DB에 반영됨');
}
// 2026-08-07: 취소된 행 표시 (회색, 줄 전체) - 우클릭 셀 강조가 아니라
// Done/Ready to process와 같은 행 단위 상태로 만들어달라는 사용자 요청.
// board_status에 'cancelled'를 저장 - 삭제는 아니므로 행은 그대로 board에
// 남고(복구는 Done/Ready to process를 다시 눌러 원래 상태로 되돌리면 됨),
// undo(Ctrl+Z)도 다른 status 변경과 동일하게 적용됨.
async function processCancel(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, 'cancelled');
  }catch(err){toast('취소 처리 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='cancelled'});
  clearRowErrorFlags(itemIds);
  await stackSettledToTop(itemIds);
  renderAll();
  toast(keys.length+'건 취소 처리(회색) - DB에 반영됨');
}
// 2026-09-29: "billed" - 회수 오더가 안 돼서 고객에게 비용을 청구한 상태(초록).
// 자주 안 쓰는 버튼이라 작게 배치. 마감 파일엔 들어가지만 엑셀에서 파랑이
// 아닌 초록으로 찍혀 vendor_dashboard.py 수량 집계에서 자동 제외된다.
async function processBilled(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, 'billed');
  }catch(err){toast('billed 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus='billed'});
  clearRowErrorFlags(itemIds);
  await stackSettledToTop(itemIds);
  renderAll();
  toast(keys.length+'건 billed 처리(초록) - 대시보드 수량 제외');
}
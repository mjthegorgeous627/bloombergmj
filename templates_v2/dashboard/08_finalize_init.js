// -- Finalize ("마감"): the actual end-of-day close-out - export the checked
// ROWS (not the whole date - 2026-08-27 user request) to Excel, and ONLY if
// that succeeds, also clear those specific rows out of workbench.db (same
// soft-delete every other delete goes through - 7 day recovery window,
// undoable). Doing this as one action means there's no way to end up
// half-done: delete-before-export-succeeded, or export-then-forget-to-delete
// leaving stale "done" rows cluttering the board.
//
// Granularity is per-row, not per-date: a date's banner checkbox checks
// every row under it by default (see renderAll's exportCb listener), but any
// row can be unchecked first to leave it out of THIS 마감 - e.g. one order
// that isn't actually resolved yet while the rest of that day's orders are.
// That row just stays on the board, unchecked, date unchanged. Whenever it's
// later resolved, checking just that row and hitting 마감 again exports it
// alone - export_dates_to_excel() targets that row's own `date` field, so it
// lands appended to the bottom of the SAME sheet the rest of that date
// already went to, not a new sheet. --
// -- 월초 작업 1) 퀵 금액 정리 (2026-10-01): quick_settlement 저장소의 스크립트를
// draft(사본 미리보기) -> 확인 -> apply(원본 반영) 순서로 호출한다.
async function runQuickSettlement(){
  const now=new Date();
  const prev=new Date(now.getFullYear(), now.getMonth()-1, 1);
  const def=prev.getFullYear()+'-'+String(prev.getMonth()+1).padStart(2,'0');
  const month=(prompt('퀵 금액을 정리할 달 (YYYY-MM)', def)||'').trim();
  if(!month) return;
  const call=async(mode)=>{
    const res=await fetch('/api/monthly/quick_settlement',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({month,mode})});
    return await res.json();
  };
  toast(month+' 퀵 금액 정리 미리보기 만드는 중 (1~2분 걸릴 수 있음)...');
  let r;
  try{ r=await call('draft'); }catch(err){ toast('퀵 정리 실패: '+err.message); return; }
  if(!r.ok){ alert('퀵 금액 정리 실패\n\n'+(r.error||'알 수 없는 오류')); return; }
  const t=r.totals||{}, n=r.rows||{};
  const fmt=v=>v==null?'-':Number(v).toLocaleString();
  const same=Math.abs((t.block1||0)-(t.statement_payable||0))<=1;
  let msg=month+' 미리보기 (원본은 아직 안 건드림)\n\n'
    +'퀵회사 '+n.quick_company+'행 / 퀵배송·Taxi '+n.quick_other_taxi+'행 / 우체국 '+n.post_office+'행\n'
    +'합계  블록1 '+fmt(t.block1)+'  블록2 '+fmt(t.block2)+'  블록3 '+fmt(t.block3)+'\n'
    +'내역서 결제금액 '+fmt(t.statement_payable)+(same?'  (블록1과 일치)':'  (블록1과 불일치!)')+'\n'
    +'미리보기 파일: '+r.target;
  if((r.review||[]).length) msg+='\n\n[REVIEW - 해결 전에는 반영 불가]\n- '+r.review.join('\n- ');
  if((r.warnings||[]).length) msg+='\n\n[경고]\n- '+r.warnings.join('\n- ');
  if((r.review||[]).length){ alert(msg); return; }
  if(!confirm(msg+'\n\n원본에 반영할까요? (원본은 _backup 폴더에 자동 백업, 퀵 정리 엑셀은 닫혀 있어야 함)')) return;
  toast('원본에 반영 중...');
  try{ r=await call('apply'); }catch(err){ toast('반영 실패: '+err.message); return; }
  if(!r.ok){ alert('원본 반영 실패\n\n'+(r.error||'알 수 없는 오류')); return; }
  toast(month+' 퀵 금액 정리를 원본에 반영했습니다 ('+r.sheet+' 시트)');
}

// 2026-10-06 실사고(ZRX 67105149): 배송(10/6 완료)과 회수(10/7 예정, 미완료)를
// 날짜 나눠 저장해뒀는데 shift-클릭 범위선택 등으로 회수까지 체크된 채 마감이
// 눌려, 아직 안 한 회수가 마감 파일 10-7 시트로 나가고 보드에서도 사라졌다.
// 그래서 (1) 완료(파랑)/취소/billed가 아닌 행은 체크돼 있어도 마감에서 강제로
// 뺀다(체크도 해제), (2) 마감이 진행 중인 동안 재클릭을 막는다(응답이 느릴 때
// 다시 눌러 같은 행이 두 번 내보내진 사고). 서버도 같은 검증을 한 번 더 한다
// (items.check_finalizable).
function localTodayIso(){
  return new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,10);
}
let FINALIZING=false;
const FINALIZABLE_UI_STATUSES=['done','cancelled','billed'];
async function finalizeChecked(){
  if(FINALIZING){toast('이미 마감 처리 중입니다 - 끝날 때까지 기다리세요 (다시 누르지 마세요)');return}
  const checkedRowKeys=new Set([...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey));
  const allChecked=ROWS.filter(r=>checkedRowKeys.has(r.rowKey));
  // 2026-10-06 ZRX 67103429: 파랑(done)인데 날짜가 아직 안 온(예: 10/8) 회수가 10/2에 마감돼
  // 사라졌다. 오늘 뒤 날짜의 done은 마감 대상이 아니다(취소/billed는 예외, 서버도 동일 검증).
  const todayIso=localTodayIso();
  const notReady=allChecked.filter(r=>!FINALIZABLE_UI_STATUSES.includes(r.uiStatus)
    || (r.uiStatus==='done' && r.date>todayIso));
  if(notReady.length){
    notReady.forEach(r=>checkedRowKeys.delete(r.rowKey));
    document.querySelectorAll('.row-check:checked').forEach(cb=>{ if(!checkedRowKeys.has(cb.dataset.rowKey)) cb.checked=false; });
    window.alert('마감할 수 없는 행 '+notReady.length+'건은 제외했습니다 (보드에 그대로 남음) - 완료(파랑)/취소/billed가 아니거나, 아직 오지 않은 날짜의 완료 행입니다:\n\n'
      +notReady.map(r=>'- '+r.date+' '+(r.orderLabel||'(메모)')+' '+(r.itemType||'')).join('\n'));
  }
  const checkedRows=allChecked.filter(r=>checkedRowKeys.has(r.rowKey));
  if(!checkedRows.length){toast('마감할 행을 먼저 체크하세요 (완료(파랑)/취소/billed 행만 마감할 수 있습니다)');return}
  const isos=[...new Set(checkedRows.map(r=>r.date))];
  const payload={};
  isos.forEach(iso=>{payload[iso]=buildExportRows(iso, checkedRowKeys)});
  FINALIZING=true;
  const finBtn=document.querySelector('button.finish');
  if(finBtn) finBtn.disabled=true;
  try{
    await finalizeCheckedRun(checkedRows, checkedRowKeys, isos, payload);
  } finally {
    FINALIZING=false;
    if(finBtn) finBtn.disabled=false;
  }
}

async function finalizeCheckedRun(checkedRows, checkedRowKeys, isos, payload){
  toast('마감 처리 중 (엑셀로 내보내는 중)...');
  let exportData;
  try{
    const res=await fetch('/api/export', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({dates:payload, itemIds:checkedRows.map(r=>r.itemId)})
    });
    exportData=await res.json();
    if(!res.ok || exportData.error){throw new Error(exportData.error||'내보내기 실패')}
  }catch(err){
    window.alert('마감 실패 (배송장에 저장하지 못해 workbench.db는 그대로 둠)\n\n'+err.message+'\n\n배송장에 일부 반영됐을 수 있으니 오늘 날짜 시트를 확인한 뒤 다시 마감하세요.');
    return;
  }
  const itemIds=checkedRows.map(r=>r.itemId);
  const summary=Object.entries(exportData.exported).map(([iso,info])=>`${labelFor(iso)}: ${info.sheet}시트 ${info.rows}행`).join(' / ');
  try{
    const res=await fetch('/api/items/delete', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'삭제 실패')}
  }catch(err){
    toast('엑셀 저장은 됐지만('+summary+') workbench.db 정리는 실패: '+err.message);
    return;
  }
  if(typeof recordFinalizedCounts==='function') recordFinalizedCounts(checkedRows);
  ROWS = ROWS.filter(r=>!checkedRowKeys.has(r.rowKey));
  isos.forEach(iso=>EXPORT_DATES.delete(iso));
  renderAll();
  toast(checkedRows.length+'건 마감 완료 - '+summary+' / workbench.db 정리됨 (7일간 복구 가능)');
}

// -- Delete: removes checked rows from the board AND from workbench.db
// (order_items rows, cascading to their parent order if that was its last
// item). A checked date-export checkbox also counts as "delete this whole
// date" (it already checks every row underneath it - see renderAll's
// exportCb listener - this dateKeys union is a belt-and-suspenders fallback
// in case some row somehow doesn't have a live checkbox in the DOM). --
async function deleteChecked(){
  const rowKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  const checkedDates=[...document.querySelectorAll('.date-export-check:checked')]
    .map(cb=>cb.closest('tbody.date-group').dataset.date);
  const dateKeys=ROWS.filter(r=>checkedDates.includes(r.date)).map(r=>r.rowKey);
  const keys=[...new Set([...rowKeys, ...dateKeys])];
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    const res=await fetch('/api/items/delete', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'삭제 실패')}
  }catch(err){
    toast('삭제 실패 (DB 반영 안 됨): '+err.message);
    return;
  }
  ROWS = ROWS.filter(r=>!keys.includes(r.rowKey));
  checkedDates.forEach(iso=>EXPORT_DATES.delete(iso));
  renderAll();
  toast(keys.length+'건 삭제 (DB에 반영됨)');
}

// -- Copy: each checked row gets duplicated into a fresh standalone row
// (order_no blank, same safety rule as the manual memo row - see
// copy_items()'s docstring) inserted directly below itself. Position is
// computed server-side (re-densifies board_pos for the affected date
// section per copy), so the simplest correct way to pick it up is the same
// full reload submitManualRow()/mergeChecked() already use after a
// server-side row insert, rather than half-reconstructing board_row()'s
// highlights/mergeGroups shape here on the client. --
async function copyChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  let data;
  try{
    const res=await fetch('/api/items/copy', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds})
    });
    data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'복사 실패');
  }catch(err){toast('복사 실패: '+err.message);return}
  toast((data.items||[]).length+'건 복사됨');
  location.reload();
}

// -- Drag-and-drop between date sections (nearby dates) --
function onDragOver(e){e.preventDefault();e.currentTarget.classList.add('drag-over')}
document.addEventListener('dragleave',e=>{
  const grp=e.target.closest && e.target.closest('.date-group');
  if(grp && !grp.contains(e.relatedTarget)) grp.classList.remove('drag-over');
});
document.addEventListener('dragend',e=>{
  document.querySelectorAll('.date-group.drag-over').forEach(g=>g.classList.remove('drag-over'));
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
});
async function onDrop(e){
  // Fallback for dropping into empty space in a date section (not on a
  // specific row) - e.g. the section is empty, or the gap below its last
  // row. Row-level drops (onRowDrop) handle the precise-position case and
  // stopPropagation so this doesn't double-fire on top of them. Same
  // comma-joined multi-row payload onRowDrop reads (see grip's dragstart) -
  // moves every dragged row's whole group at once, keeping their relative
  // order, to the bottom of this date section.
  e.preventDefault();
  const tbody=e.currentTarget;
  tbody.classList.remove('drag-over');
  const sourceKeys=(e.dataTransfer.getData('text/plain')||'').split(',').filter(Boolean);
  if(!sourceKeys.length)return;
  const sourceSet=new Set(sourceKeys);
  const movedRows=ROWS.filter(r=>sourceSet.has(r.rowKey));
  if(!movedRows.length)return;
  const iso=tbody.dataset.date;
  if(movedRows.every(r=>r.date===iso))return;
  try{
    await persistItemDate(movedRows.map(r=>r.itemId), iso);
  }catch(err){toast('날짜 이동 저장 실패: '+err.message);return}
  movedRows.forEach(r=>{r.date=iso});
  renderAll();
  toast(movedRows.length+'건 다른 날짜로 이동 - DB에 반영됨');
}

// -- Same-date (and cross-date-to-a-specific-spot) reordering: drop directly
// on a row to insert the dragged group(s) immediately before/after the
// HOVERED ROW'S WHOLE GROUP - never mid-group, even when hovering the middle
// row of a 3-row order (2026-08-27 user report: the old row-by-row boundary
// could land inside a completely unrelated order's block, making the preview
// look like it was about to mix into it). Persists as board_pos for every
// item currently in the destination date section (see reorder_items()). --
function onRowDragOver(e){
  e.preventDefault();
  const tr=e.currentTarget;
  const tbody=tr.closest('tbody.date-group');
  const groupTrs=[...tbody.querySelectorAll('tr.item-row')].filter(el=>el.dataset.groupKey===tr.dataset.groupKey);
  const firstTr=groupTrs[0], lastTr=groupTrs[groupTrs.length-1];
  const groupTop=firstTr.getBoundingClientRect().top;
  const groupBottom=lastTr.getBoundingClientRect().bottom;
  const before=(e.clientY-groupTop) < (groupBottom-groupTop)/2;
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
  (before?firstTr:lastTr).classList.add(before?'drop-before':'drop-after');
}
async function onRowDrop(e){
  e.preventDefault();
  e.stopPropagation();
  const tr=e.currentTarget;
  const tbody=tr.closest('tbody.date-group');
  const groupTrs=[...tbody.querySelectorAll('tr.item-row')].filter(el=>el.dataset.groupKey===tr.dataset.groupKey);
  const firstTr=groupTrs[0], lastTr=groupTrs[groupTrs.length-1];
  const before=firstTr.classList.contains('drop-before');
  const after=lastTr.classList.contains('drop-after');
  document.querySelectorAll('.item-row.drop-before,.item-row.drop-after').forEach(el=>el.classList.remove('drop-before','drop-after'));
  if(!before && !after) return;

  const sourceKeys=(e.dataTransfer.getData('text/plain')||'').split(',').filter(Boolean);
  if(!sourceKeys.length) return;
  const sourceSet=new Set(sourceKeys);
  // Target is the whole hovered group, not just the one <tr> the browser
  // fired the drop on - anchor on its LAST row so a drop lands after every
  // row of that group when `after`, not spliced in after just its first row.
  const targetKey = before ? firstTr.dataset.rowKey : lastTr.dataset.rowKey;
  if(sourceSet.has(targetKey))return; // dropped back onto (part of) itself
  const movedRows=ROWS.filter(r=>sourceSet.has(r.rowKey)); // keeps original relative order
  if(!movedRows.length)return;
  ROWS=ROWS.filter(r=>!sourceSet.has(r.rowKey));

  const targetIso=tbody.dataset.date;
  const dateChanged=movedRows.some(r=>r.date!==targetIso);
  movedRows.forEach(r=>{r.date=targetIso});
  let targetIdx=ROWS.findIndex(r=>r.rowKey===targetKey);
  if(targetIdx===-1) targetIdx=ROWS.length;
  ROWS.splice(before?targetIdx:targetIdx+1, 0, ...movedRows);

  try{
    if(dateChanged) await persistItemDate(movedRows.map(r=>r.itemId), targetIso);
    await persistItemReorder(ROWS.filter(r=>r.date===targetIso).map(r=>r.itemId));
  }catch(err){toast('순서 저장 실패: '+err.message);return}
  renderAll();
  toast((dateChanged ? '날짜 이동 + 순서 변경' : '순서 변경')+' - '+movedRows.length+'건 - DB에 반영됨');
}

// -- Excel-panel date picker: check rows below, pick a date up top, they all
// jump there. Handles dates that don't have a section yet (e.g. board only
// shows 7/28 and 7/30 but the user wants 7/29) by creating a fresh date
// section in the right sorted position (renderAll() rebuilds every date in
// sorted order every time, so a brand-new iso just needs to exist on a
// row's `date` field to get its own section). --
async function moveCheckedToDate(iso){
  if(!iso)return;
  if(!/^\d{4}-\d{2}-\d{2}$/.test(iso) || isNaN(new Date(iso+'T00:00:00').getTime())){
    toast('날짜 형식이 올바르지 않습니다. 달력에서 다시 선택해주세요');
    return;
  }
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const toMove=ROWS.filter(r=>keys.includes(r.rowKey) && r.date!==iso);
  if(!toMove.length){toast('이미 해당 날짜에 있습니다');return}
  try{
    await persistItemDate(toMove.map(r=>r.itemId), iso);
  }catch(err){toast('날짜 이동 저장 실패: '+err.message);return}
  toMove.forEach(row=>row.date=iso);
  renderAll();
  toast(toMove.length+'건을 '+labelFor(iso)+'(으)로 이동 - DB에 반영됨');
}
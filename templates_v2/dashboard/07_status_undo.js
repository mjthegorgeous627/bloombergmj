// 2026-08-20: "초기화" - Done/Ready to process/Cancel을 누른 적 없는
// 원래 상태(흰색, board_status='')로 되돌림. 실사고 계기: ZRX 67079036 -
// 배송은 완료했지만 회수는 다음날로 미뤄졌는데 같은 오더 안에서 함께
// Done 처리돼 회수 행도 파랑으로 남아있던 걸 발견 (오더 단위가 아니라
// 행/아이템 단위 상태이므로, 되돌릴 대상도 체크된 행만 - 오더 전체가
// 아니라 회수 행 하나만 체크해서 초기화할 수 있음). set_items_board_status()는
// 이미 임의의 문자열(빈 문자열 포함, 로그 메시지도 '(초기화)'로 이미
// 대비돼 있었음)을 그대로 저장하므로 백엔드 변경 없이 프론트에서만 추가.
async function processReset(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    await persistItemStatus(itemIds, '');
  }catch(err){toast('초기화 저장 실패: '+err.message);return}
  keys.forEach(k=>{const row=ROWS.find(r=>r.rowKey===k); if(row) row.uiStatus=''});
  clearRowErrorFlags(itemIds);
  renderAll();
  toast(keys.length+'건 초기화(흰색) - DB에 반영됨');
}
// 2026-08-20: 실제 SAP 오더 수집 없이 텍스트만으로 board에 새 행을 만드는
// 기능 - 사용자 요청("메모용으로도 한줄씩 추가하고 싶을때가 있다"), 이후
// "구분/날짜도 고를 수 있게 해달라"는 후속 요청으로 단순 prompt() 대신
// 모달로 교체 - trashOverlay와 같은 오버레이 패턴 재사용. 새로 생긴 행은
// 서버가 만든 것이라 로컬 ROWS 배열에 없으므로(다른 버튼들처럼 로컬 상태만
// 패치할 수 없음) 저장 후 전체 새로고침.
// Type/check를 제외한 나머지 7개 board 열, 왼쪽부터의 순서 (add_manual_note()의
// cells 매핑과 반드시 같은 순서여야 함) - MANUAL_ROW_COLUMNS 선언은 파일
// 위쪽(ROWS 바로 아래)으로 옮겨짐: buildGroupRows()의 메모 행 렌더링이
// manualRowGroupStarts()를 통해 이 값을 쓰는데, 그 렌더링은 최초
// renderAll() 호출(이 지점보다 훨씬 앞)때부터 실행되므로 const가 원래
// 자리(여기)에 있으면 TDZ(temporal dead zone)로 "Cannot access
// 'MANUAL_ROW_COLUMNS' before initialization"가 나며 전체 스크립트가
// 죽는다 - 2026-08-27 실사고: 메모 행 추가 버튼을 누르자마자 보드 전체가
// (마지막 렌더된 날짜 한 줄만 남기고) 하얗게 사라짐, 콘솔에 정확히 이
// 에러가 찍힘.
// 2026-08-25 정정: "몇 칸을 쓸지" 물어서 그 수만큼 칸을 만든다는 게, 칸마다
// 원래 열 하나씩(좁게) 쓴다는 뜻이 아니라 - 예를 들어 1칸만 쓰면 나머지 6개
// 열 공간을 전부 병합해서 하나의 넓은 칸으로 쓴다는 뜻이었음(사용자 정정).
// 그래서 7개 열을 N개로 최대한 고르게 나눠 각 칸이 여러 열 너비를 차지하게
// 함 - manualRowGroupStarts(n)이 그 경계를 계산하고, 실제 저장은 각 그룹의
// "시작 열" 자리에만 값을 넣고 나머지는 빈 문자열로 둔다(add_manual_note()의
// cells와 동일한 7칸짜리 배열, 위치만 건너뛰며 채움). 렌더링(buildGroupRows)
// 쪽은 그 결과를 "값이 있는 열부터 시작해서 뒤따르는 빈 열들을 흡수해
// colspan으로 병합"하는 방식으로 다시 그리므로, 이 함수가 계산한 경계와
// 항상 자동으로 일치한다 - 즉 N을 별도로 저장할 필요가 없다(빈 칸이 아닌
// 열의 개수 자체가 곧 N이므로, 나중에 셀을 직접 편집해서 비우면 다음
// 렌더링 때 그 칸만큼 자연스럽게 옆 칸에 합쳐짐).
function manualRowGroupStarts(n){
  const total=MANUAL_ROW_COLUMNS.length;
  const base=Math.floor(total/n), rem=total%n;
  const starts=[]; let acc=0;
  for(let i=0;i<n;i++){ starts.push(acc); acc += base+(i<rem?1:0); }
  return starts;
}
function renderManualRowCells(){
  const countInput=document.getElementById('manualRowCellCount');
  const n=Math.max(1, Math.min(MANUAL_ROW_COLUMNS.length, parseInt(countInput.value)||1));
  countInput.value=n;
  const starts=manualRowGroupStarts(n);
  const wrap=document.getElementById('manualRowCells');
  const prevValues=[...wrap.querySelectorAll('.manualRowCellInput')].map(i=>i.value);
  wrap.innerHTML='';
  starts.forEach((start,i)=>{
    const end=(i+1<starts.length ? starts[i+1] : MANUAL_ROW_COLUMNS.length)-1;
    const rangeLabel=end>start ? MANUAL_ROW_COLUMNS[start]+'~'+MANUAL_ROW_COLUMNS[end] : MANUAL_ROW_COLUMNS[start];
    const label=document.createElement('label');
    label.style.cssText='display:block;margin-bottom:8px;font-size:13px;color:var(--muted,#555)';
    label.textContent=rangeLabel+' 칸'+(i===0?' (필수)':' (선택)');
    const input=document.createElement('input');
    input.type='text';
    input.className='manualRowCellInput';
    input.dataset.start=start;
    input.style.cssText='display:block;width:100%;margin-top:4px;padding:6px 8px;box-sizing:border-box';
    input.value=prevValues[i]||'';
    if(i===0) input.placeholder='제목 (필수)';
    label.appendChild(input);
    wrap.appendChild(label);
  });
  wrap.querySelector('.manualRowCellInput').focus();
}
function addManualRow(){
  document.getElementById('manualRowOrderLabel').value='';
  // 기본값 = 일반 오더 행과 같은 7칸 전부(칸마다 정확히 열 하나씩) - 2026-08-27
  // user request. 필요하면 이 값을 줄여서 병합된 넓은 칸 몇 개로 바꿔 쓸 수
  // 있음(기존 동작 그대로, manualRowGroupStarts() 참고).
  document.getElementById('manualRowCellCount').value=MANUAL_ROW_COLUMNS.length;
  renderManualRowCells();
  const sel=document.getElementById('manualRowType');
  sel.innerHTML='';
  allTypeOptions().forEach(v=>{
    const opt=document.createElement('option');
    opt.value=v; opt.textContent=v;
    sel.appendChild(opt);
  });
  sel.value='배송';
  document.getElementById('manualRowDate').value=new Date().toISOString().slice(0,10);
  document.getElementById('manualRowOverlay').classList.add('show');
}
function closeManualRowModal(){
  document.getElementById('manualRowOverlay').classList.remove('show');
}
async function submitManualRow(){
  const cells=new Array(MANUAL_ROW_COLUMNS.length).fill('');
  document.querySelectorAll('.manualRowCellInput').forEach(inp=>{
    cells[parseInt(inp.dataset.start)]=inp.value.trim();
  });
  const itemType=document.getElementById('manualRowType').value;
  const date=document.getElementById('manualRowDate').value;
  const orderLabel=document.getElementById('manualRowOrderLabel').value.trim();
  // 선택한 칸 수 그대로 서버에 저장(manual_cell_count) - 렌더링이 내용 유무가
  // 아니라 이 값으로 칸 경계를 그리도록(아래 buildGroupRows 참고, 2026-08-27
  // 버그 수정: 7칸으로 만들고 제목만 채워도 나머지 6칸이 빈 칸인 채로 남아있지
  // 않고 1칸으로 합쳐져 보이던 문제).
  const cellCount=parseInt(document.getElementById('manualRowCellCount').value)||cells.length;
  if(!cells[0]){toast('제목이 비어 있습니다');return}
  try{
    const res=await fetch('/api/orders/add_manual', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({cells, itemType, date, orderLabel, cellCount})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'추가 실패');
  }catch(err){toast('메모 행 추가 실패: '+err.message);return}
  closeManualRowModal();
  location.reload();
}

// -- Merge: manual "Excel merge cells", independently per field, for
// customer/phone/address/memo across different order numbers that are
// really the same customer (e.g. ZOR 7825897 / ZOR 7825832 / ZOR 7825838
// sharing one address even when the name on file differs per order - so
// each field is asked for and merged separately, not all 4 as one block).
// Which fields to merge comes from the persistent .merge-field-cb checkbox
// row (Customer/Phone/Address/Memo checked by default; Order# is present
// for layout symmetry but not wired to anything yet - filtered out
// server-side against MERGE_FIELDS regardless). Works on whichever rows are
// checked, same checkboxes as everything else - resolves them to their
// distinct orders server-side. Only renders merged once those orders are
// also adjacent on the board (drag rows next to each other first via the
// drag handle, same-date reorder). --
function selectedMergeFields(){
  return [...document.querySelectorAll('.merge-field-cb:checked')]
    .map(cb=>cb.value).filter(f=>MERGE_FIELDS.includes(f));
}
async function mergeChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const fields=selectedMergeFields();
  if(!fields.length){toast('병합할 칸을 하나 이상 체크하세요');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  let data;
  try{
    const res=await fetch('/api/items/merge', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, fields})
    });
    data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'병합 실패');
  }catch(err){toast('병합 실패: '+err.message);return}
  if(!data.merged){toast(data.message||'서로 다른 오더에 속한 행을 2건 이상 선택하세요');return}
  // mergeGroups lives on rows fetched fresh from the server (order_payload),
  // not something ROWS can patch itself - simplest correct way to pick it up.
  location.reload();
}
async function unmergeChecked(){
  const keys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!keys.length){toast('체크된 행이 없습니다');return}
  const fields=selectedMergeFields();
  if(!fields.length){toast('해제할 칸을 하나 이상 체크하세요');return}
  const itemIds=ROWS.filter(r=>keys.includes(r.rowKey)).map(r=>r.itemId);
  try{
    const res=await fetch('/api/items/unmerge', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({itemIds, fields})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'병합 해제 실패');
    toast(data.unmerged+'건 병합 해제 - DB에 반영됨');
  }catch(err){toast('병합 해제 실패: '+err.message);return}
  location.reload();
}

// -- Export: the one action in this app that actually writes a real file
// (the live shipping Excel workbook) instead of just changing what's on
// screen. Builds the current board state (post move/edit/merge/delete) for
// just the checked rows (`checkedRowKeys`) of a given date, and sends it to
// the server to write/append into that date's sheet. Checking a date's
// banner checkbox checks every row under it by default (see renderAll's
// exportCb listener) - unchecking a specific row before hitting 마감 excludes
// only that row, leaving it on the board for a later 마감 once it's actually
// resolved (2026-08-27 user request). --
function buildExportRows(iso, checkedRowKeys){
  const rows=ROWS.filter(r=>r.date===iso && checkedRowKeys.has(r.rowKey));
  const groups=groupRows(rows);
  const allSpans=computeAllFieldSpans(groups);
  const out=[];
  groups.forEach((group,gi)=>{
    const customer=resolvedField(group,'customer');
    const phone=resolvedField(group,'phone');
    const address=resolvedField(group,'address');
    const memo=resolvedField(group,'memo');
    group.rows.forEach((row, i)=>{
      out.push({
        // each field merges its own E-H column down its own independent run
        // (see _write_export_sheet) - a manually-merged field (mergeChecked())
        // can span a different set of orders than another merged field.
        mergeFirst: {
          customer: i===0 && allSpans.customer[gi].isFirst,
          phone: i===0 && allSpans.phone[gi].isFirst,
          address: i===0 && allSpans.address[gi].isFirst,
          memo: i===0 && allSpans.memo[gi].isFirst,
        },
        itemType: row.itemType,
        orderLabel: row.orderLabel,
        deliveryNo: row.deliveryNo,
        description: row.description,
        material: row.material,
        serial: row.serial,
        customer, phone, address, memo,
        // carries this row's manual cell highlight (see cell_highlights) into
        // the Excel export - only meaningful on the row that actually owns
        // each field's <td> (i===0 && isFirst, same rows mergeFirst already
        // marks), which is exactly where the highlight was ever recorded.
        highlights: row.highlights || {},
        // 2026-09-28: cancelled(회색) 행도 마감 대상에 포함되면서 추가 - 엑셀
        // 쪽에서 완료(파랑)와 구분되는 색으로 찍어야 vendor_dashboard.py의
        // 색상 기반 집계에서 잘못 카운트되지 않는다 (_write_export_sheet 참고).
        cancelled: row.uiStatus==='cancelled',
        billed: row.uiStatus==='billed',
      });
    });
  });
  return out;
}

// 비상용 엑셀 → workbench 불러오기. 평상시엔 SAP 수집이 직접 workbench.db에
// 반영되므로(import_sap_rows) 이 버튼은 자동 동기화를 대신하지 않는다 - 엑셀에
// 직접 손댄 내용을 다시 가져와야 하는 드문 경우에만, 시트+행 범위를 명시해서
// 딱 그만큼만 1회성으로 긁어온다 (통째로 다시 스캔하지 않음).
async function importExcelRange(){
  const sheet=document.getElementById('excelRangeSheet').value.trim();
  const rowStart=document.getElementById('excelRangeStart').value;
  const rowEnd=document.getElementById('excelRangeEnd').value;
  if(!sheet){toast('시트명을 입력하세요 (예: 8-6)');return}
  if(!rowStart||!rowEnd){toast('시작행/종료행을 입력하세요');return}
  if(!window.confirm(`"${sheet}" 시트의 ${rowStart}~${rowEnd}행을 workbench로 불러옵니다. 계속할까요?`))return;
  toast('엑셀에서 불러오는 중...');
  try{
    const res=await fetch('/api/import/range', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({sheet, rowStart:Number(rowStart), rowEnd:Number(rowEnd)})
    });
    const data=await res.json();
    if(!res.ok || data.error){throw new Error(data.error||'불러오기 실패')}
    toast(`불러오기 완료 - ${data.imported}개 오더 반영`);
    location.reload();
  }catch(err){
    toast('불러오기 실패: '+err.message);
  }
}

// -- 배송장 인쇄: Type/Order#/Item/M-N/S-N/Customer/Phone/Address/Memo만 뽑아
// A4 가로/흑백 전용 테이블을 새로 만들어 인쇄한다. 선택은 새 UI를 만들지 않고
// 이미 있는 두 체크박스를 그대로 재사용한다 - 날짜 배너 체크박스(EXPORT_DATES,
// 엑셀 내보내기와 같은 것)는 "그 날짜 전체", 행 체크박스(.row-check)는 체크된
// 행만(한쪽 leg만 체크했으면 그 leg만 인쇄되고 나머지 leg는 안 찍힘) - 다만
// 같은 오더의 체크된 행이 여럿이면 그 체크된 것들끼리는 Customer/Phone/
// Address/Memo를 보드 화면과 똑같이 rowSpan으로 한 번만 찍는다.
// computeAllFieldSpans는 인쇄 대상으로 뽑힌(체크된) 행들만 놓고 다시 돌려서 -
// 화면엔 붙어 있어도 체크 안 된 오더/행까지 병합 셀에 끼어들지 않게 한다. --
// 한 날짜의 모든 행(상태 무관)을 체크/해제한다 - 배송장 인쇄용 [전체선택] 버튼(PC 날짜 줄 / 모바일 날짜 제목)이 쓴다.
// 이미 전부 체크돼 있으면 전부 해제. 마감용 날짜 체크박스(EXPORT_DATES)와는 별개다.
function selectAllRowsOfDate(iso){
  const cbs=ROWS.filter(r=>r.date===iso).map(r=>document.querySelector('.row-check[data-row-key="'+r.rowKey+'"]')).filter(Boolean);
  if(!cbs.length) return false;
  const on=!cbs.every(cb=>cb.checked);
  cbs.forEach(cb=>{ if(cb.checked!==on){ cb.checked=on; cb.dispatchEvent(new Event('change')); } });
  return on;
}
function printChecked(){
  const checkedRowKeys=new Set([...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey));
  const sections=[];
  allDatesSorted().forEach(iso=>{
    const rows=ROWS.filter(r=>r.date===iso);
    const groups=groupRows(rows);
    const wholeDate=EXPORT_DATES.has(iso);
    // 2026-10-16: 날짜 줄 체크박스(마감용)를 눌러 '날짜 전체' 모드가 돼도, 그 체크박스가 자동으로
    // 체크해 주는 상태(완료/취소/청구)의 행을 사용자가 직접 해제했으면 그 행은 인쇄에서 뺀다.
    // (그 외 상태의 행은 날짜 체크박스가 원래 체크해 주지 않으므로 '체크 안 됨'이 '제외'를 뜻하지 않아 계속 포함.
    //  임의의 행을 빼고 싶으면 날짜 줄의 [전체선택] 버튼 -> 빼고 싶은 행만 해제 -> 인쇄.)
    const AUTO_CHECKED=new Set(['done','cancelled','billed']);
    const selectedGroups=(wholeDate
        ? groups.map(g=>({key:g.key, rows:g.rows.filter(r=>checkedRowKeys.has(r.rowKey) || !AUTO_CHECKED.has(r.uiStatus))}))
        : groups.map(g=>({key:g.key, rows:g.rows.filter(r=>checkedRowKeys.has(r.rowKey))})))
      .filter(g=>g.rows.length);
    if(!selectedGroups.length) return;
    sections.push({iso, groups:selectedGroups, allSpans:computeAllFieldSpans(selectedGroups)});
  });
  if(!sections.length){
    toast('인쇄할 날짜 또는 행이 선택되지 않았습니다 (날짜 줄 체크박스 또는 행 체크박스를 먼저 선택하세요)');
    return;
  }

  const printArea=document.getElementById('printArea');
  printArea.textContent='';
  const title=document.createElement('h2');
  title.textContent='배송장';
  printArea.appendChild(title);

  sections.forEach(({iso, groups, allSpans}, si)=>{
    const h3=document.createElement('h3');
    // 인쇄물에는 "(Today)"/"(Tomorrow)" 같은 화면 전용 문구를 빼고 순수 날짜만.
    h3.textContent=labelFor(iso).replace(/\s*\((Today|Tomorrow)\)\s*$/,'');
    // 2026-10-16: 날짜가 달라도 새 페이지로 넘기지 않고 이어서 한 장에 찍는다(종이 절약). 내용이 길면 알아서 다음 장으로 넘어감.
    h3.style.breakAfter='avoid';
    printArea.appendChild(h3);

    const table=document.createElement('table');
    table.className='print-table';
    const colgroup=document.createElement('colgroup');
    [5,10,15,8,9,8,9,21,15].forEach(w=>{
      const col=document.createElement('col'); col.style.width=w+'%'; colgroup.appendChild(col);
    });
    table.appendChild(colgroup);

    const thead=document.createElement('thead');
    const headTr=document.createElement('tr');
    ['Type','Order #','Item','M/N','S/N','Customer','Phone','Address','Memo'].forEach(h=>{
      const th=document.createElement('th'); th.textContent=h; headTr.appendChild(th);
    });
    thead.appendChild(headTr);
    table.appendChild(thead);

    const tbody=document.createElement('tbody');
    groups.forEach((group, gi)=>{
      const n=group.rows.length;
      group.rows.forEach((row, i)=>{
        const tr=document.createElement('tr');
        const typeTd=document.createElement('td'); typeTd.textContent=row.itemType||'';
        if(row.itemType==='배송') typeTd.className='print-type-delivery';
        tr.appendChild(typeTd);
        if(i===0){
          const ordTd=document.createElement('td'); ordTd.rowSpan=n; ordTd.textContent=row.orderLabel||'';
          tr.appendChild(ordTd);
        }
        const itemTd=document.createElement('td'); itemTd.textContent=row.description||''; tr.appendChild(itemTd);
        const mnTd=document.createElement('td'); mnTd.textContent=row.material||''; tr.appendChild(mnTd);
        const snTd=document.createElement('td'); snTd.textContent=row.serial||''; tr.appendChild(snTd);
        if(i===0){
          ['customer','phone','address','memo'].forEach(field=>{
            const fs=allSpans[field][gi];
            if(!fs.isFirst) return;
            const td=document.createElement('td');
            td.textContent=resolvedField(group, field);
            td.rowSpan=fs.span;
            tr.appendChild(td);
          });
        }
        tbody.appendChild(tr);
      });
    });
    table.appendChild(tbody);
    printArea.appendChild(table);
  });

  window.print();
}

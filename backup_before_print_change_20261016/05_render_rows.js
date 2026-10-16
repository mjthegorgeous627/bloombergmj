function buildGroupRows(group, allSpans, gi){
  const trs=[]; const n=group.rows.length;
  // "select this whole order" checkbox (added to ordTd below) - so merging
  // (or bulk Done/삭제/날짜이동) across orders with many items doesn't mean
  // hunting down and clicking every single item-row checkbox by hand; one
  // click selects/deselects every item this order-group owns. groupCheckboxes
  // collects each row's real .row-check as it's built; syncSelectAll keeps
  // the order-level box in check/indeterminate/unchecked sync if the user
  // still clicks individual item checkboxes directly.
  const groupCheckboxes=[];
  let selectAllCb=null;
  const syncSelectAll=()=>{
    if(!selectAllCb)return;
    const checkedCount=groupCheckboxes.filter(cb=>cb.checked).length;
    selectAllCb.checked=checkedCount>0 && checkedCount===groupCheckboxes.length;
    selectAllCb.indeterminate=checkedCount>0 && checkedCount<groupCheckboxes.length;
  };
  group.rows.forEach((row, i)=>{
    const tr=document.createElement('tr');
    tr.className='item-row'+(row.uiStatus==='done'?' row-done':row.uiStatus==='pending'?' row-pending':row.uiStatus==='cancelled'?' row-cancelled':row.uiStatus==='billed'?' row-billed':'')
      +((row.portalError||row.podError||row.snReloError)?' row-error':'');
    tr.dataset.rowKey=row.rowKey;
    // group.key (same value mergeKey(row) would give) stamped on every row of
    // this group - onRowDragOver/onRowDrop use it to snap the insert point to
    // the GROUP's own top/bottom edge instead of a boundary that can fall
    // inside a DIFFERENT, unrelated order's multi-row block (2026-08-27 user
    // report: the drag preview line could land mid-way through a completely
    // unrelated order - different order#, different company - making it look
    // like a drop there would splice into/mix with it).
    tr.dataset.groupKey=group.key;
    // same-date reordering: dragover/drop directly on a row inserts the
    // dragged row(s) immediately before/after its GROUP (see
    // onRowDragOver/onRowDrop). stopPropagation on drop only - dragover is
    // left to bubble up to the tbody handler too, so the whole date section
    // still gets its 'drag-over' background highlight while hovering any row
    // inside it.
    tr.addEventListener('dragover', onRowDragOver);
    tr.addEventListener('drop', onRowDrop);

    // drag handle - a dedicated grip so dragging doesn't fight with
    // clicking/selecting text inside the row's editable cells or dropdown
    const handleTd=document.createElement('td');
    handleTd.className='handle-cell';
    const grip=document.createElement('span');
    grip.className='drag-handle';
    grip.textContent='⠿';
    grip.draggable=true;
    // Dragging one row now drags its WHOLE order-group (every row sharing
    // its orderLabel - an order must never be split by a drag) plus, if the
    // grabbed row is itself checked, every OTHER checked row's whole group
    // too - so checking several rows (or whole orders via the "select this
    // order" box) and dragging any one of them moves the entire selection
    // together as one block, in their existing relative order (2026-08-27
    // user request: "선택하고 드래그 드롭을 하는게 뭉텅이로 쉽게 이동
    // 가능해야하는데"). Dragging an UNchecked row still only moves that one
    // row's own group, so a plain drag never sweeps up unrelated checked
    // rows left over from something else.
    grip.addEventListener('dragstart', e=>{
      const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
      const seedKeys = (checkedKeys.length && checkedKeys.includes(row.rowKey)) ? checkedKeys : [row.rowKey];
      const seedSet=new Set(seedKeys);
      // 2026-10-07 실사고(ZRX 67105999): 같은 오더의 배송(오늘)과 회수(10/13 예정)를 날짜 나눠
      // 놨는데, 배송 행을 드래그(같은 날짜 안 순서 변경이어도)하면 오더번호만 같다는 이유로
      // 10/13의 회수까지 같이 끌려와 오늘로 날짜가 바뀌어 버렸다. 묶음은 "같은 날짜 안의
      // 같은 오더"로 한정한다 - 화면에서 한 덩어리로 보이는 행들(date-group 내부)과 동일.
      const groupId=r=>r.date+'|'+mergeKey(r);
      const groupKeysInvolved=new Set(ROWS.filter(r=>seedSet.has(r.rowKey)).map(groupId));
      const fullKeys=ROWS.filter(r=>groupKeysInvolved.has(groupId(r))).map(r=>r.rowKey);
      e.dataTransfer.setData('text/plain', fullKeys.join(','));
      const fullKeySet=new Set(fullKeys);
      document.querySelectorAll('.item-row').forEach(el=>{
        if(fullKeySet.has(el.dataset.rowKey)) el.classList.add('dragging');
      });
    });
    grip.addEventListener('dragend', ()=>{
      document.querySelectorAll('.item-row.dragging').forEach(el=>el.classList.remove('dragging'));
    });
    handleTd.appendChild(grip);
    tr.appendChild(handleTd);

    // type - dropdown (배송/회수/Bloomberg/Delayed/회수 Delayed/... plus any
    // custom type added via "+ 추가"), per item (a merged block can still
    // mix a delivery line and a pickup line for the same order)
    const typeTd=document.createElement('td');
    typeTd.className='type-cell';
    const sel=document.createElement('select');
    const applySelClass=()=>{sel.className='type-select '+typeClassFor(sel.value)};
    allTypeOptions().forEach(v=>{
      const opt=document.createElement('option'); opt.value=v; opt.textContent=v;
      if(v===row.itemType) opt.selected=true;
      sel.appendChild(opt);
    });
    if(row.itemType && !allTypeOptions().includes(row.itemType)){
      const cur=document.createElement('option'); cur.value=row.itemType; cur.textContent=row.itemType;
      cur.selected=true; sel.insertBefore(cur, sel.firstChild);
    }
    const addOpt=document.createElement('option');
    addOpt.value='__add__'; addOpt.textContent='+ 추가';
    sel.appendChild(addOpt);
    applySelClass();
    sel.addEventListener('change', ()=>{
      if(sel.value==='__add__'){
        const newType=(window.prompt('새 Type 이름을 입력하세요 (모든 행의 드롭다운에 추가됩니다)')||'').trim();
        if(newType && !allTypeOptions().includes(newType)){
          CUSTOM_TYPES.push(newType);
          persistCustomType(newType)
            .catch(err=>toast('새 Type 목록 저장 실패 (이 세션에서만 보임): '+err.message));
        }
        if(newType){
          row.itemType=newType;
          persistItemField(row.itemId,'item_type',newType)
            .catch(err=>toast('Type 저장 실패 (DB 반영 안 됨): '+err.message));
        }
        renderAll();
        return;
      }
      // 2026-08-26 fix: 이전엔 여기서 row.itemType만 바꾸고 서버에 저장하는
      // 코드가 아예 없었음 - 화면에선 바뀐 것처럼 보이다가 다음 새로고침(또는
      // 워크벤치 재시작) 때 서버가 다시 내려주는 원래 값으로 조용히
      // 되돌아갔음. 배송/회수만 특별취급하던 resolve_item_type()의 별도
      // 버그와 겹쳐서 "Bloomberg로 바꿔도 자꾸 배송으로 돌아간다"는 증상으로
      // 나타남(사용자 신고, 8/26) - 그 버그도 같이 고침.
      row.itemType=sel.value;
      applySelClass();
      persistItemField(row.itemId,'item_type',sel.value)
        .catch(err=>toast('Type 저장 실패 (DB 반영 안 됨): '+err.message));
    });
    typeTd.appendChild(sel);
    applyHighlight(typeTd, row.itemId, 'type', row);
    tr.appendChild(typeTd);

    const checkTd=document.createElement('td');
    checkTd.className='check-cell';
    const cb=document.createElement('input');
    cb.type='checkbox'; cb.className='row-check'; cb.dataset.rowKey=row.rowKey;
    cb.addEventListener('change', syncSelectAll);
    groupCheckboxes.push(cb);
    checkTd.appendChild(cb);
    tr.appendChild(checkTd);

    if(i===0){
      const ordTd=document.createElement('td');
      ordTd.className='ord-cell';
      ordTd.rowSpan=n;

      selectAllCb=document.createElement('input');
      selectAllCb.type='checkbox';
      selectAllCb.className='order-select-all';
      selectAllCb.title='이 오더의 항목 전체 선택/해제 (병합·완료·삭제·날짜이동 등에 사용)';
      selectAllCb.addEventListener('change', ()=>{
        groupCheckboxes.forEach(cb=>{cb.checked=selectAllCb.checked});
      });
      ordTd.appendChild(selectAllCb);

      // main line: the SAP order code (ZOR/ZRE/ZRX/ZINP/ZINT/...) + order number - bold, black.
      // Persisted as of 2026-08-06 - but only the TYPE portion. order_no itself is never
      // sent to the server (see update_order_field()'s own comment for why: it's the real
      // matching key SAP resync/POD/Portal automation all key off, so silently renaming it
      // in place is too risky) - if the number portion is edited, it's reverted on-screen
      // and a toast explains why, while the type-portion edit (if any) still saves normally.
      const main=document.createElement('div');
      main.contentEditable='true'; main.spellcheck=false; main.textContent=row.orderLabel;
      const orderItemIds=group.rows.map(r=>r.itemId);
      main.addEventListener('blur', ()=>{
        const v=main.textContent.trim();
        const realNo=group.rows[0].orderNo||'';
        let newType, newNo;
        let orderNoChanged='';
        if(!realNo){
          // No real SAP order number tied to this row (memo rows, and any
          // other free-text label - e.g. "SDSK1333608359") - there is
          // nothing to protect, so the whole typed text IS the label as-is.
          // 2026-08-27 bug fix: this used to always split on the first space
          // and treat everything after it as "the order number", discarding
          // it whenever it didn't match a real order_no - for a blank-
          // order_no row that meant ANY space in the text silently ate
          // everything typed after it (e.g. "SDSK 1333608359" saved as just
          // "SDSK"). Splitting only makes sense when there IS a real order_no
          // to compare the typed "number" portion against.
          newType=v;
          newNo='';
        } else {
          const parts=v.split(/\s+/);
          newType=parts[0]||'';
          newNo=parts.slice(1).join(' ');
          if(newNo && newNo!==realNo){
            if(window.confirm('오더번호를 '+realNo+' → '+newNo+' 로 바꿉니다.\n\nSAP 수집·POD·Portal 자동화는 이 번호를 기준으로 동작하므로 실제 SAP 오더번호와 다르면 오류가 날 수 있습니다. 계속할까요?')){
              orderNoChanged=newNo;
            } else {
              toast('오더번호 수정을 취소했습니다 - 타입만 저장됩니다');
            }
          }
        }
        const savedLabel=(newType+' '+(orderNoChanged||realNo)).trim();
        group.rows.forEach(r=>{r.orderLabel=savedLabel; if(orderNoChanged) r.orderNo=orderNoChanged;});
        renderAll();
        persistOrderField(orderItemIds, 'order_type', newType)
          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
        if(orderNoChanged){
          persistOrderNo(orderItemIds, orderNoChanged)
            .catch(err=>toast('오더번호 저장 실패 (DB 반영 안 됨): '+err.message));
        }
      });
      main.addEventListener('keydown', e=>{ if(e.key==='Enter'){ e.preventDefault(); main.blur(); } });
      const ordLine=document.createElement('div');
      ordLine.className='ord-line';
      ordLine.appendChild(main);
      ordLine.appendChild(makeOrderIcons(group.rows[0].orderNo, row.orderLabel));
      ordTd.appendChild(ordLine);

      // 실패 배지 (todo b, 2026-08-18) - portalError/podError는 order 단위라
      // 그룹의 모든 행이 같은 값을 갖고 있음(row 대신 group.rows[0]에서 읽어도
      // 동일), 여기 ordTd는 rowSpan으로 그룹당 한 번만 그려지니 배지도 한 번만.
      // 클릭 시 실패 사유 팝업 - main(라벨 편집)의 contentEditable 영역과
      // 겹치지 않게 별도 span으로 붙이고 클릭 시 stopPropagation.
      const errMsg=row.portalError ? row.portalErrorMsg : (row.podError ? row.podErrorMsg : '');
      if(errMsg){
        const badge=document.createElement('span');
        badge.className='order-error-badge';
        badge.textContent='⚠';
        badge.title='클릭: 실패 사유 보기';
        badge.addEventListener('click', e=>{
          e.stopPropagation();
          window.alert((row.portalError ? 'Portal 등록(Serial/QR) 실패:\n\n' : 'POD(배송완료) 처리 실패:\n\n')+errMsg);
        });
        ordTd.appendChild(badge);
      }

      // secondary order-code line(s): Bloomberg-side codes (OBD/ORD/SDSK/...), one per
      // line, same small muted style, same free-text edit as everything else - not just
      // a bare number anymore, and more lines can be typed in directly if more than one
      // secondary code applies to this order. Fully persisted as of 2026-08-06 via
      // persistDeliveryCodes()/update_delivery_codes() - the server splits whichever
      // line starts with "OBD" into delivery_no (the real matching key) and joins
      // everything else (ORD/SDSK/etc) into extra_codes (plain text, no matching-key
      // role) - see update_delivery_codes()'s own comment.
      const codeLines=[...new Set(
        group.rows.flatMap(r=>String(r.deliveryNo||'').split('\n')).map(s=>s.trim()).filter(Boolean)
      )];
      const extra=document.createElement('div');
      extra.className='muted';
      extra.contentEditable='true'; extra.spellcheck=false;
      extra.textContent=codeLines.join('\n');
      extra.addEventListener('blur', ()=>{
        const v=extra.innerText.split('\n').map(s=>s.trim()).filter(Boolean).join('\n');
        group.rows.forEach(r=>r.deliveryNo=v);
        renderAll();
        persistDeliveryCodes(orderItemIds, v)
          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
      });
      const codeWrap=document.createElement('div');
      codeWrap.className='code-wrap';
      codeWrap.appendChild(extra);
      if(codeLines.length) codeWrap.appendChild(makeCodeIcons(codeLines));
      ordTd.appendChild(codeWrap);
      applyHighlight(ordTd, row.itemId, 'order', row);

      tr.appendChild(ordTd);
    }

    if(!row.orderNo && n===1){
      // 2026-08-25: 메모 행은 addManualRow() 모달에서 고른 개수만큼만
      // 제목/M-N/S-N/Customer/Phone/Address/Memo 7칸 중 일부에 값이 들어있고
      // 나머지는 빈 문자열로 저장돼 있다(manualRowGroupStarts() 참고, 정확히
      // 각 그룹의 "시작 열"에만 값이 있음). 2026-08-27 수정: 원래 이 조건이
      // orderLabel==='memo' 리터럴 문자열 매치였는데, Order#를 사용자가 직접
      // 입력할 수 있게 되면서(add_manual_note()의 order_label 참고) 라벨이
      // "memo"가 아닌 메모 행("SDSK1333608359" 등)에서 이 압축 렌더링이
      // 깨지는 문제가 생겨 !row.orderNo(실제 SAP 오더번호가 없는 행 전부)로
      // 바꿈 - 어차피 order_no는 메모 행에서 항상 빈 값이므로 라벨 텍스트와
      // 무관하게 정확히 같은 행 집합을 가리킨다. 다른 오더와 병합된 경우
      // (n>1, 드문 케이스)는 아래 일반 렌더링으로 그대로 둠.
      const fields=[
        ['description','item-cell', v=>row.description=v, v=>persistItemField(row.itemId,'description',v)],
        ['material','mn-cell', v=>row.material=v, v=>persistItemField(row.itemId,'material',v)],
        ['serial','sn-cell', v=>row.serial=v, v=>persistItemField(row.itemId,'serial',v)],
        ['customer','cust-cell customer-cell', v=>row.customer=v, v=>persistOrderField([row.itemId],'customer',v)],
        ['phone','cust-cell phone-cell', v=>row.phone=v, v=>persistOrderField([row.itemId],'phone',v)],
        ['address','cust-cell addr-cell', v=>row.address=v, v=>persistOrderField([row.itemId],'address',v)],
        ['memo','cust-cell memo-cell', v=>row.memo=v, v=>persistOrderField([row.itemId],'memo',v)],
      ];
      // 2026-08-27 버그 수정: 칸 경계를 "값이 있는 열부터 시작해서 뒤따르는
      // 빈 열들을 흡수" 방식으로만 정하면, 사용자가 일부러 7칸으로 만들고
      // 몇 칸은 선택사항이라 비워둔 경우까지 그 빈 칸들을 앞 칸에 합쳐버려
      // (신고 사례: 제목만 채운 7칸짜리 메모가 1칸으로 보임) 실제로 몇 칸을
      // 만들었는지와 화면이 안 맞았다. addManualRow()가 고른 칸 수를
      // manual_cell_count로 저장해두고(add_manual_note() 참고), 있으면 그
      // 값 그대로 manualRowGroupStarts()로 경계를 재현 - 내용 유무와 무관하게
      // 항상 정확히 그 칸 수만큼 그려짐. 그 값이 없는(이 저장 방식 이전에
      // 만들어진) 옛 메모 행만 예전처럼 내용 기준으로 추정.
      const storedCount=parseInt(row.manualCellCount, 10);
      const starts=(storedCount>=1 && storedCount<=fields.length)
        ? manualRowGroupStarts(storedCount)
        : null;
      // Cross-order merge (mergeChecked(), MERGE_FIELDS) reaches into a manual
      // memo row's own compact rendering too - a manual row (order_no blank)
      // can be picked as the merge anchor for customer/phone/address/memo
      // (e.g. one delivery address shared by both a memo row and its
      // matching real ZINP/ZRX order), or it can be the one that gets
      // absorbed into a DIFFERENT row's rowSpan. Before this, this branch
      // always drew row[field] standalone and never looked at allSpans/gi,
      // so a merged manual row rendered as if unmerged (real bug report,
      // 2026-09-17: "체크된 행 병합"이 안 먹힌다) - worse, an ABSORBED manual
      // row still drew its own (blank) cust/phone/addr/memo cells instead of
      // being skipped, breaking column count for every row after it in that
      // merged block. Only handled when this field's own compact-column slot
      // isn't ALSO combined with another field via colSpan (span===1) - a
      // manual row compacted down to fewer than 7 cells AND cross-merged on
      // one of those combined fields is rare enough to leave as a known gap
      // rather than reconcile colSpan+rowSpan math for it here.
      const mergedFieldTd=(field, cellClass)=>{
        const fs=allSpans[field][gi];
        if(!fs.isFirst) return null;  // absorbed into a preceding row's merged span - render nothing
        const td=makeEditableTd(
          resolvedField({rows: fs.mergedRows}, field), cellClass,
          v=>fs.mergedRows.forEach(r=>r[field]=v),
          v=>persistOrderField(fs.mergedRows.map(r=>r.itemId), field, v)
        );
        if(fs.span>1) td.rowSpan=fs.span;
        return td;
      };
      const isCrossMerged=field=>MERGE_FIELDS.includes(field) && row.mergeGroups && row.mergeGroups[field];
      if(starts){
        starts.forEach((start,ci)=>{
          const end=ci+1<starts.length ? starts[ci+1] : fields.length;
          const span=end-start;
          const [field, cellClass, setLocal, persist]=fields[start];
          let td;
          if(span===1 && isCrossMerged(field)){
            td=mergedFieldTd(field, cellClass);
          } else {
            td=makeEditableTd(row[field]||'', cellClass, setLocal, persist);
            if(span>1) td.colSpan=span;
          }
          if(!td) return;
          applyHighlight(td, row.itemId, field, row);
          tr.appendChild(td);
        });
      } else {
        let idx=0;
        while(idx<fields.length){
          const [field, cellClass, setLocal, persist]=fields[idx];
          let span=1;
          while(idx+span<fields.length && !(row[fields[idx+span][0]]||'').trim()) span++;
          let td;
          if(span===1 && isCrossMerged(field)){
            td=mergedFieldTd(field, cellClass);
          } else {
            td=makeEditableTd(row[field]||'', cellClass, setLocal, persist);
            if(span>1) td.colSpan=span;
          }
          if(td){
            applyHighlight(td, row.itemId, field, row);
            tr.appendChild(td);
          }
          idx+=span;
        }
      }
    } else {
    const descTd=makeEditableTd(row.description, 'item-cell', v=>row.description=v, v=>persistItemField(row.itemId,'description',v));
    applyHighlight(descTd, row.itemId, 'description', row);
    tr.appendChild(descTd);
    const matTd=makeEditableTd(row.material, 'mn-cell', v=>row.material=v, v=>persistItemField(row.itemId,'material',v));
    applyHighlight(matTd, row.itemId, 'material', row);
    tr.appendChild(matTd);
    // 2026-08-18: 토글 배지는 UI가 번거롭다는 사용자 피드백으로 제거함 -
    // 그냥 원래대로 편집만 되는 셀. "serial을 안 적으면 시리얼 없는
    // 물품"이라는 사용자의 실제 워크플로우는 저장 시점이 아니라 Portal
    // 등록 버튼(run_portal())이 클릭되는 시점에 그때의 serial 값을 보고
    // 판단함 - item["mode"]가 아니라 item["serial"]이 비었는지로 직접
    // 판정하도록 run_portal()을 고쳤음(위쪽 참고, ZINP 551059036
    // "COPYHOLDER FELLOWES BOOKLIFT" qty=1 계기). 그래서 여기 클라이언트
    // 쪽엔 mode를 건드릴 이유가 없음 - serial 저장만 하면 끝.
    const serTd=makeEditableTd(row.serial, 'sn-cell', v=>row.serial=v, v=>persistItemField(row.itemId,'serial',v));
    applyHighlight(serTd, row.itemId, 'serial', row);
    // S/N 오더 반영(paper relo 판별) 실패 배지 (2026-08-19) - portalError/
    // podError의 order-error-badge와 같은 스타일이지만 이건 항목 단위라
    // ordTd(그룹당 1회)가 아니라 이 항목의 sn-cell에 직접 붙인다.
    if(row.snReloError){
      const snBadge=document.createElement('span');
      snBadge.className='order-error-badge';
      snBadge.textContent='⚠';
      snBadge.title='클릭: S/N 오더 반영 실패 사유 보기';
      snBadge.addEventListener('click', e=>{
        e.stopPropagation();
        window.alert('S/N 오더 반영 실패 (paper relo 필요할 수 있음):\n\n'+row.snReloErrorMsg);
      });
      serTd.appendChild(snBadge);
    }
    tr.appendChild(serTd);

    if(i===0){
      // first non-blank value across the group, not just group.rows[0] - a
      // merged pickup leg with no customer/phone/address of its own (e.g.
      // ZRX 67066585's 회수 record) should still show/inherit the delivery
      // leg's values rather than rendering blank. Each field's own span
      // (allSpans[field][gi]) can extend across OTHER orders independently
      // when that field's been manually merged (mergeChecked()) - e.g.
      // address merged across 3 orders while customer stays per-order -
      // editing a merged cell then updates every row in THAT field's merged
      // block, not just this order-group's own rows.
      [['customer','customer-cell'],['phone','phone-cell'],['address','addr-cell'],['memo','memo-cell']]
        .forEach(([field, cellClass])=>{
          const fs=allSpans[field][gi];
          if(!fs.isFirst) return;
          // resolvedField({rows: fs.mergedRows}, ...), not resolvedField(group, ...) -
          // this order-group can be the merge anchor (isFirst) while a DIFFERENT
          // order later in the same merged block is the one actually holding the
          // real text (this group's own rows may all be blank for this field) -
          // only mergedRows covers every row the span actually renders over.
          const td=makeEditableTd(
            resolvedField({rows: fs.mergedRows}, field), 'cust-cell '+cellClass, v=>fs.mergedRows.forEach(r=>r[field]=v),
            v=>persistOrderField(fs.mergedRows.map(r=>r.itemId), field, v)
          );
          td.rowSpan=fs.span;
          applyHighlight(td, row.itemId, field, row);
          if(field==='memo') decorateMemoTd(td, fs.mergedRows);
          tr.appendChild(td);
        });
    }
    }

    trs.push(tr);
  });
  return trs;
}

function renderAll(){
  const table=document.getElementById('boardTable');
  table.querySelectorAll('tbody.date-group').forEach(tb=>tb.remove());
  allDatesSorted().forEach(iso=>{
    const tbody=document.createElement('tbody');
    tbody.className='date-group';
    tbody.dataset.date=iso;
    tbody.addEventListener('drop', onDrop);
    tbody.addEventListener('dragover', onDragOver);

    const collapsed=COLLAPSED_DATES.has(iso);
    if(collapsed) tbody.classList.add('collapsed');

    const bannerTd=document.createElement('td');
    bannerTd.colSpan=11;
    const bannerInner=document.createElement('div');
    bannerInner.className='date-banner-inner';
    const labelSpan=document.createElement('span');
    labelSpan.textContent=labelFor(iso);

    const rightGroup=document.createElement('div');
    rightGroup.className='date-banner-right';
    const exportCb=document.createElement('input');
    exportCb.type='checkbox';
    exportCb.className='date-export-check';
    exportCb.title='엑셀로 내보낼 날짜로 선택 (이 날짜의 행 체크박스가 전부 함께 체크/해제됨 - 마감 전에 아직 해결 안 된 행만 개별로 체크 해제하면 이번 마감에서 빠짐)';
    exportCb.checked=EXPORT_DATES.has(iso);
    exportCb.addEventListener('change', ()=>{
      if(exportCb.checked) EXPORT_DATES.add(iso); else EXPORT_DATES.delete(iso);
      // 날짜 체크박스 = "이 날짜 전체를 기본 선택" 편의 기능 (2026-08-27 user
      // request) - 개별 행(.row-check)을 전부 이 상태로 맞춘 뒤, 특정 행만
      // 다시 체크 해제해서 이번 마감에서 뺄 수 있게 한다. 반대 방향(행을
      // 개별로 건드려도 날짜 체크박스 자체는 되돌리지 않음)은 order-select-all
      // 패턴과 달리 의도적으로 안 함 - 이건 "부분 상태"를 표시하는 용도가
      // 아니라 순수 일괄-선택 버튼이라 단순하게 유지.
      //
      // 2026-09-02 실사고로 수정: 예전엔 상태 무관하게 그 날짜 행을 전부
      // 체크했는데, 오더 67083939처럼 같은 오더의 항목이 날짜별로 갈려있고
      // (배송 9/1 완료=파랑, 회수는 9/8로 옮겨졌지만 아직 미확정) 사용자가
      // 그 날짜 체크박스를 개별 언체크하는 걸 깜빡하면, 아직 완료 안 된
      // 행까지 마감 파일에 그대로 들어가버렸다("파란색 마감된 것만 들어가야
      // 함" 사용자 지적). 이제 체크(on)할 때는 board_status가 'done'(파랑)인
      // 행만 자동으로 체크하고, 나머지(대기/취소/미지정)는 그대로 둔다 -
      // 끌 때(off)는 상태 무관하게 전부 해제(선택 취소는 항상 안전하므로).
      //
      // 2026-09-28: 'cancelled'(회색)도 'done'과 함께 자동 체크 대상에 추가.
      // 취소 건(예: ZRX 67099736 bunit 캔슬)은 위의 "아직 미확정" 케이스가
      // 아니라 이미 처리가 끝난 상태라 done과 똑같이 마감 대상이어야 하는데,
      // 09-02 수정이 done만 남기면서 cancelled까지 같이 걸러져 마감에서
      // 누락되고 있었다(사용자 지적으로 발견).
      tbody.querySelectorAll('.row-check').forEach(cb=>{
        if(!exportCb.checked){ cb.checked=false; return; }
        const row=ROWS.find(r=>r.rowKey===cb.dataset.rowKey);
        cb.checked = !!(row && (row.uiStatus==='done' || row.uiStatus==='cancelled' || row.uiStatus==='billed'));
      });
    });
    const toggleBtn=document.createElement('button');
    toggleBtn.type='button';
    toggleBtn.className='date-toggle';
    toggleBtn.textContent=collapsed ? '▸' : '▾';
    toggleBtn.title=collapsed ? '펼치기' : '접기';
    toggleBtn.addEventListener('click', ()=>{
      if(COLLAPSED_DATES.has(iso)) COLLAPSED_DATES.delete(iso); else COLLAPSED_DATES.add(iso);
      renderAll();
    });
    rightGroup.appendChild(exportCb);
    rightGroup.appendChild(toggleBtn);

    bannerInner.appendChild(labelSpan);
    bannerInner.appendChild(rightGroup);
    bannerTd.appendChild(bannerInner);
    const bannerTr=document.createElement('tr');
    bannerTr.className='date-banner-row';
    bannerTr.appendChild(bannerTd);
    tbody.appendChild(bannerTr);

    const rows=ROWS.filter(r=>r.date===iso && flagMatch(r));
    if(BOARD_FILTER!=='all' && !rows.length) return;  // 모아보기 중에는 해당 없는 날짜는 숨김
    if(!rows.length){
      tbody.appendChild(emptyRowTr());
    } else {
      const groups=groupRows(rows);
      const allSpans=computeAllFieldSpans(groups);
      groups.forEach((g,gi)=>buildGroupRows(g, allSpans, gi).forEach(tr=>tbody.appendChild(tr)));
    }
    table.appendChild(tbody);
  });
  if(typeof updateSideViews==='function') updateSideViews();
}
renderAll();

// Shift-click range select on row checkboxes (Excel-style): click one row's
// checkbox, then shift-click another, and every row in between is set to
// match the shift-clicked box's new state. Plain clicks need no extra
// handling for "select several" - checkboxes are already additive, each
// click just toggles that one row on top of whatever else is checked.
// Delegated on the table (which itself is never recreated - only its
// tbody.date-group sections are, on every renderAll()) so this keeps working
// across re-renders instead of needing to be re-attached per row.
{
  let lastCheckedCb=null;
  document.getElementById('boardTable').addEventListener('click', e=>{
    const cb=e.target.closest('.row-check');
    if(!cb) return;
    if(e.shiftKey && lastCheckedCb && lastCheckedCb.isConnected){
      const all=[...document.querySelectorAll('.row-check')];
      const from=all.indexOf(lastCheckedCb), to=all.indexOf(cb);
      if(from!==-1 && to!==-1){
        const [lo,hi]=from<to?[from,to]:[to,from];
        const state=cb.checked;
        for(let i=lo;i<=hi;i++){
          if(all[i].checked!==state){ all[i].checked=state; all[i].dispatchEvent(new Event('change')); }
        }
      }
    }
    lastCheckedCb=cb;
  });
}

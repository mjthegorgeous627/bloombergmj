async function pollPidRunning(pid, timeoutMs){
  const start=Date.now();
  while(Date.now()-start < timeoutMs){
    await new Promise(r=>setTimeout(r, 3000));
    const res=await fetch('/api/process/'+pid+'/running');
    const data=await res.json();
    if(!data.running) return true;
  }
  return false;
}

// -- POD (Proof of Delivery / "mark as delivered") - ported from
// launcher.py's "3. 수동 배송 완료 처리" panel. Grouped by the board's
// rendered orderLabel ("ZRX 12345"), NOT raw orderId - real bug hit live
// (2026-08-03): a 배송/회수 pair that LOOKS like one order on screen (one
// shared label, one select-all checkbox) can actually be two separate
// `orders` DB rows (Excel sometimes imports a ZRX's two legs as distinct
// order records - groupRows() only unifies them visually by matching label
// text, same mechanism the merge feature and an earlier stale-duplicate-
// order issue both already hit). Grouping by orderId split one checked
// block into two incomplete POD runs; grouping by label and sending every
// order id in the group lets the server resolve which one actually carries
// the real Portal delivery# (_resolve_pod_order()). For each label group:
// preview (no browser touch) -> a confirm dialog showing exactly what will
// be saved -> if confirmed, the real "click Update POD" run (async, polled
// via PID liveness rather than orders.status - see run_pod_update()'s own
// comment server-side for why). Declining or failing one group's confirm
// just moves on to the next, unlike runPortalBatch() which aborts the whole
// batch - POD is a per-order human sign-off, not a pipeline where one
// failure implies something systemic is wrong. --
// POD 미리보기 조회 - 성공하면 preview 객체, 실패하면 null(+토스트).
// orderIds/itemIds/pickupDone은 podForLabel()/runSpareOneshot() 둘 다 같은
// 방식(체크된 행 그룹)으로 계산해서 넘긴다.
async function podFetchPreview(label, orderIds, itemIds, pickupDone){
  try{
    const res=await fetch('/api/pod/preview', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({orderIds, itemIds, pickupDone})
    });
    const preview=await res.json();
    if(!res.ok || preview.error) throw new Error(preview.error||'미리보기 실패');
    return preview;
  }catch(err){
    toast(label+' POD 미리보기 실패: '+err.message+' - 건너뜁니다');
    return null;
  }
}
function podConfirmText(preview, actionLabel){
  return `오더: ${preview.order}\n`+
    `Delivery#: ${preview.delivery}\n`+
    `배송시간: ${preview.delivery_datetime}\n`+
    `수신인: ${preview.signed_by}\n`+
    `배송 Serial: ${(preview.delivery_serials||[]).join(', ')||'-'}\n`+
    `회수 Serial: ${(preview.pickup_serials||[]).join(', ')||'-'}\n`+
    `Remarks: ${preview.remarks}\n\n`+
    actionLabel;
}
// 실제 POD 저장(Update POD 클릭) 실행 + 완료까지 대기 + 결과 토스트. 확인창은
// 이미 호출한 쪽(podForLabel 또는 runSpareOneshot)에서 끝난 뒤라 여기서는
// 다시 묻지 않는다.
async function podRunUpdate(label, orderIds, itemIds, pickupDone, preview){
  let data;
  try{
    const res=await fetch('/api/pod/update', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        orderIds, itemIds, pickupDone,
        signedBy: preview.signed_by, datetime: preview.delivery_datetime, remarks: preview.remarks,
      })
    });
    data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
  }catch(err){
    toast(label+' POD 저장 실패: '+err.message);
    return false;
  }
  if(!data.started){
    toast(label+': '+(data.message||'건너뜀'));
    return false;
  }
  toast(label+' POD 저장 진행 중 (PID '+data.pid+') - 완료될 때까지 대기...');
  const finished=await pollPidRunning(data.pid, 120000);
  if(!finished){
    toast(label+' POD 저장 시간 초과 (2분) - 상태를 직접 확인하세요');
    return false;
  }
  // small grace period - the watcher thread's own add_event() write can
  // land a beat after the OS process itself is gone (which is what
  // pollPidRunning just observed), so fetching events immediately can
  // occasionally race and show a stale prior event instead.
  await new Promise(r=>setTimeout(r, 500));
  try{
    const res=await fetch('/api/orders/'+data.podOrderId);
    const od=await res.json();
    const msg=(od.events && od.events[0] && od.events[0].message || '').split('\n')[0];
    toast(label+': '+(msg||'POD 처리 완료'));
  }catch(err){
    toast(label+': POD 처리 완료 (상태 확인 실패)');
  }
  return true;
}
// runPodBatch용: 미리보기 -> 확인창 -> 저장, 기존 그대로.
async function podForLabel(label, groupRows_){
  const orderIds=[...new Set(groupRows_.map(r=>r.orderId))];
  // pickupDone comes straight from which of THIS group's rows are
  // actually checked, not a separate question - 배송/회수 each have their
  // own row+checkbox now, so checking only the 배송 row means "delivery
  // only", checking both means "delivery and pickup both done", exactly
  // like checking rows already drives every other bulk action on the
  // board (Done/삭제/날짜이동/병합 etc).
  const pickupDone=groupRows_.some(r=>r.itemType==='회수' || r.itemType==='회수 Delayed');
  // Only the CHECKED rows' item ids go to the server - the remarks text
  // (Delivered/collected item counts) is built from exactly these, not
  // Excel and not every item the order happens to have, so checking just
  // the 배송 row correctly leaves the 회수 item out of "Delivered ..." too.
  const itemIds=groupRows_.map(r=>r.itemId);

  const preview=await podFetchPreview(label, orderIds, itemIds, pickupDone);
  if(!preview) return false;

  if(!window.confirm(podConfirmText(preview, '이 내용으로 POD(배송완료) 저장할까요?'))){
    toast(label+': POD 저장 건너뜀');
    return false;
  }
  return podRunUpdate(label, orderIds, itemIds, pickupDone, preview);
}
async function runPodBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const labels=[...new Set(checkedRows.map(r=>r.orderLabel))];

  for(const label of labels){
    const groupRows_=checkedRows.filter(r=>r.orderLabel===label);
    await podForLabel(label, groupRows_);
  }
  toast('POD 처리 완료');
}
// -- 스페어 배송처리 (todo a, 2026-08-18): Serial 등록 + POD(배송완료)를 한
// 번의 클릭으로 처리. 확인창은 버튼을 누른 그 순간, Serial 등록을 시작하기
// *전에* POD 미리보기 내용으로 딱 한 번만 뜬다 - 실사용해보니(2026-08-18)
// Serial 등록이 끝난 뒤에야 POD 확인창이 뜨는 구조라 결국 두 번 확인해야
// 완료됐는데, 그 확인창을 앞으로 당겨서 한 번 확인하면 Serial 등록부터
// POD 저장까지 중간에 아무 팝업 없이 끝까지 진행되도록 바꿈 (사용자 확정).
// 그래서 순서가 podForLabel()과 반대: 여기서는 미리보기→확인을 먼저 하고,
// 그 확인에서 얻은 preview 값(signed_by/datetime/remarks)을 그대로 들고
// Serial 등록 → podRunUpdate()로 이어간다 - 확인 이후에는 preview를 다시
// 조회하지 않는다(값이 그 사이 바뀔 이유가 없고, 다시 조회하면 그 결과를
// 또 확인해야 하니 원샷의 의미가 없어짐). QR 인쇄는 항상 생략(--skip-qr) -
// 스페어는 라벨 재인쇄가 보통 불필요해서 버튼 자체에서 뺐다 (사용자 확정).
async function runSpareOneshot(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const labels=[...new Set(checkedRows.map(r=>r.orderLabel))];

  for(const label of labels){
    const groupRows_=checkedRows.filter(r=>r.orderLabel===label);
    const orderIds=[...new Set(groupRows_.map(r=>r.orderId))];
    const pickupDone=groupRows_.some(r=>r.itemType==='회수' || r.itemType==='회수 Delayed');
    const itemIds=groupRows_.map(r=>r.itemId);

    const preview=await podFetchPreview(label, orderIds, itemIds, pickupDone);
    if(!preview) continue;
    if(!window.confirm(podConfirmText(preview, '이 내용으로 Serial 등록 후 POD(배송완료)까지 한 번에 진행할까요?'))){
      toast(label+': 스페어 배송처리 건너뜀');
      continue;
    }

    let portalOk=true;
    for(const orderId of orderIds){
      let data;
      try{
        const res=await fetch('/api/orders/'+orderId+'/run_portal', {
          method:'POST', headers:{'Content-Type':'application/json'},
          body:JSON.stringify({skipQr: true})
        });
        data=await res.json();
        if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
      }catch(err){
        toast(label+' Serial 등록 실행 실패: '+err.message+' - 이 오더는 건너뜁니다');
        portalOk=false; break;
      }
      if(!data.started){
        // 이미 'printed'(완료)인 경우는 실패가 아니라 그대로 POD 단계로
        // 진행 - Serial 등록만 먼저 성공해둔 뒤 원샷을 재실행하는 경우 등.
        // '처리 중'(다른 프로세스가 살아있음)인 경우만 진짜로 건너뜀.
        if((data.message||'').includes('이미 완료')){
          toast(label+': Serial 등록 이미 완료 - POD로 진행');
        } else {
          toast(label+': '+(data.message||'Serial 등록 건너뜀')+' - 이 오더는 건너뜁니다');
          portalOk=false; break;
        }
      } else {
        toast(label+' Serial 등록 진행 중 (PID '+data.pid+') - 완료 대기...');
        const result=await pollOrderDone(orderId, 360000);
        if(result.status==='printed'){
          toast(label+' Serial 등록 완료');
        } else if(result.status==='timeout'){
          toast(label+' Serial 등록 시간 초과 (6분) - 이 오더는 건너뜁니다');
          portalOk=false; break;
        } else {
          toast(label+' Serial 등록 실패: '+(result.detail||('상태: '+result.status))+' - 이 오더는 건너뜁니다');
          portalOk=false; break;
        }
      }
    }
    if(!portalOk) continue;

    await podRunUpdate(label, orderIds, itemIds, pickupDone, preview);
  }
  toast('스페어 배송처리 완료');
}

// -- S/N 오더 반영 (2026-08-19): 회수 오더 수집 시 SAP상 S/N이 없어서('X')
// 워크벤치에 X로 찍힌 항목을 위한 단계. 사용자가 S/N 칸에 실제로 회수한
// S/N을 먼저 입력해두고 체크한 뒤 이 버튼을 누르면, 그 항목들을 순서대로
// VA02 오더에 입력·저장 시도한다 - ZREC 준비(runZrecPrepare)와 같은 "체크
// → 버튼 한 번 → 항목마다 순서대로 처리" 패턴. 저장되면(firm/cust 일치)
// 그걸로 끝 - 이 항목은 이제 일반 회수 항목과 똑같은 상태이니, 이어서
// 5번 ZREC 처리 버튼을 사람이 직접 눌러 진행한다(자동 연결 안 함, 사용자
// 확정). 저장 실패(firm/cust 불일치 추정)는 paper relo가 필요하다는 뜻 -
// S/N 칸 옆에 ⚠ 배지 + 행 주황 배경으로 표시된다.
async function runSnReloBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const targets=checkedRows.filter(r=>r.serial && r.serial.trim().toUpperCase()!=='X');
  if(!targets.length){
    window.alert('체크된 항목 중 S/N 칸이 채워진(X가 아닌) 항목이 없습니다.\n실제 회수한 S/N을 먼저 S/N 칸에 입력한 뒤 다시 시도하세요.');
    return;
  }
  if(!window.confirm(
    `체크된 항목 중 S/N이 입력된 ${targets.length}개를 대상으로, 순서대로 VA02 오더를 열어 `+
    `그 S/N을 Technical Objects에 입력·저장 시도합니다.\n`+
    `저장되면 firm/cust 일치(ZREC 진행 가능), 저장 실패면 paper relo가 필요하다는 뜻입니다.\n\n진행할까요?`
  )) return;

  let okCount=0, errCount=0;
  for(const row of targets){
    let data;
    try{
      const res=await fetch('/api/sn_relo/fill_one', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({itemId: row.itemId})
      });
      data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    }catch(err){
      errCount++;
      row.snReloError=true; row.snReloErrorMsg=err.message;
      if(!window.confirm(row.orderLabel+' S/N 오더 반영 실패: '+err.message+'\n\n다음 항목으로 계속할까요?')) break;
      continue;
    }
    if(data.status==='saved'){
      okCount++;
      row.snReloError=false; row.snReloErrorMsg='';
    }else{
      errCount++;
      row.snReloError=true; row.snReloErrorMsg=data.message||'알 수 없는 실패';
    }
  }
  renderAll();
  toast(`S/N 오더 반영 완료 - 성공 ${okCount}건, 실패/확인필요 ${errCount}건`+(errCount?' (⚠ 배지 클릭으로 사유 확인)':''));
}

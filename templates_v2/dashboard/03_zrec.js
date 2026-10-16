// -- ZREC 준비 (todo c, 2026-08-18/19, 개편 2026-08-21): 사용자의 실제
// 하루 업무 방식("오늘 회수된 모든 시리얼을 한번에 넣고 확인한 뒤, ZREC는
// 하나씩") + 2026-08-19 확정 사항("확인 누르면 Receive Equipment는
// 자동으로 클릭, 다 끝나면 ZIH08로 6507/0052 확인")을 반영한 흐름.
//   1단계: 체크된 모든 항목의 시리얼을 모아 ZIH08에 딱 한 번 조회
//          (/api/zrec/lookup_batch) - ZREC 화면은 전혀 안 건드리는
//          read-only 조회라 확인창 없이 바로 진행.
//   2단계: 조회 결과를 오더번호/M-N/S-N/수량/아이템명 한 표로 모아 확인창
//          "한 번만" 띄운다 - 원래는 항목마다 ZREC 채우기 직후 개별
//          확인창이 떴는데, 여러 건을 한 번에 처리할 때 매번 팝업이 뜨는
//          게 번거롭다는 지적(2026-08-21, 사용자) 반영. 이 확인 하나로
//          전체 배치를 승인하는 것 - "확인"을 누르면 항목마다 순서대로
//          ZREC 채우기 + Receive Equipment 자동 클릭까지 중간 팝업 없이
//          진행한다(개별 실패는 그 항목만 건너뛰고 계속). 사람이 SAP
//          ZREC 화면에 실제로 채워지는 6507/오더번호/시리얼 값을 눈으로
//          보고 승인하는 지점은 없다 - 승인은 이 확인창에서 ZIH08 조회
//          데이터를 보고 한 번에 이뤄진다(2026-08-27 확인).
//   3단계: 실제로 접수(commit)된 항목들의 시리얼을 ZIH08로 재조회해서
//          Plant/Location이 6507/0052로 들어왔는지 확인(/api/zrec/verify_batch)
//          - 확인된 항목만 board_status를 'done'(파랑)으로 자동 전환한다.
//          준비 상태(Ready to process, 빨강)는 ZIH08로 접수가 실제 확인될
//          때까지 그대로 유지 (사용자 확정, 2026-08-19: "zrec할 준비가
//          된것이 빨간색이고, zrec가 완료되면 파란색으로 처리"). 이게
//          진짜 성공/실패 판정이고, 아래 commit_one 결과 문구는 참고용일
//          뿐이니 "실패로 보인다" 같은 불확실한 문구를 화면에 띄우지
//          않는다(2026-08-19: SAP 팝업 문구 판정이 실제 성공을 실패로
//          오판하는 게 실사용에서 확인됨 - 사용자 지적, 2026-08-21: "성공
//          했는데 팝업에 실패한 것 같다는 문구가 뜬다").
async function runZrecPrepare(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));

  // 1단계: 일괄 ZIH08 조회 (read-only - 확인 없이 바로 진행)
  let plans;
  try{
    const res=await fetch('/api/zrec/lookup_batch', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({itemIds: checkedRows.map(r=>r.itemId)})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
    plans=data.plans||[];
  }catch(err){
    toast('ZIH08 일괄 조회 실패: '+err.message);
    return;
  }
  if(!plans.length){toast('ZREC 대상 항목이 없습니다');return}

  // 2단계: 전체를 한 표(줄바꿈 나열)로 모아 확인창을 딱 한 번만 띄움.
  // 항목마다 빈 줄로 섹션을 나눠 한눈에 읽히게 하고, ZIH08 경고가 있는
  // 항목은 그 아래 ⚠로 붙여서 개별 확인창 없이도 문제를 미리 볼 수 있게 함.
  const summaryLines=plans.map(p=>{
    const sn=p.nonSerial ? '(S/N 없음)' : (p.serial||'-');
    const warn=(p.warnings||[]).map(w=>'   ⚠ '+w).join('\n');
    return `${p.orderLabel}  |  M/N ${p.material||'-'}  |  S/N ${sn}  |  수량 ${p.qty}  |  ${p.description||'-'}`
      +(warn ? '\n'+warn : '');
  });
  const confirmMsg=
    `오늘 ZREC 접수할 ${plans.length}건 (오더번호 | M/N | S/N | 수량 | 아이템명):\n\n`+
    summaryLines.join('\n\n')+
    `\n\n이 내용으로 확인을 누르면 항목마다 순서대로 자동으로 Receive Equipment까지 `+
    `클릭합니다(중간 확인창 없음) - 실제로 SAP에 접수되니 내용을 꼭 확인하고 눌러주세요.\n`+
    `모두 끝나면 ZIH08로 Plant/Location(6507/0052) 접수 여부를 다시 확인합니다.`;
  if(!window.confirm(confirmMsg)) return;

  // 3단계: 항목마다 순서대로 ZREC 채우기 → 자동 Receive Equipment.
  // 위 2단계에서 이미 전체 승인을 받았으므로 항목별 확인창은 없음 -
  // 개별 실패는 토스트로만 알리고 다음 항목으로 계속 진행.
  //
  // 2026-08-27: 예전엔 여기서 fill_one(dry-run, 화면만 채움) →
  // commit_one(세션 리셋 후 똑같은 값 재입력 → 클릭) 두 번을 호출했다.
  // 그런데 fill_one이 채운 화면을 사람이 보고 판단할 시점이 이미 없다
  // (승인은 위 2단계 확인창에서 ZIH08 데이터만으로 끝남) - 재입력은
  // 안전 효과 없이 SAP 왕복만 두 배로 냈다(사용자 지적). commit_one 한
  // 번으로 채우기+검증+클릭까지 끝낸다.
  const committedItemIds=[];
  for(const plan of plans){
    try{
      const res=await fetch('/api/zrec/commit_one', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({itemId: plan.itemId, orderNo: plan.filledOrder})
      });
      const cdata=await res.json();
      if(!res.ok || cdata.error) throw new Error(cdata.error||'Receive Equipment 클릭 실패');
      // Receive Equipment 버튼은 실제로 눌렸으므로(SAP 트랜잭션 시도 완료)
      // cdata.success 값(SAP 팝업 문구 판정)과 무관하게 일단 3단계 ZIH08
      // verify_batch로 검증 대상에 넣는다 - 문구 판정은 실사용에서 실제
      // 성공을 실패로 오판하는 게 확인됐으므로(2026-08-19) 여기서는 그
      // 판정 결과를 화면에 노출하지 않고 중립적으로만 알린다. ZIH08
      // 재조회(6507/0052 확인)만이 진짜 성공/실패를 가리는 최종 판정.
      committedItemIds.push(plan.itemId);
      toast(plan.orderLabel+' Receive Equipment 클릭 완료 - ZIH08로 곧 확인합니다');
    }catch(err){
      toast(plan.orderLabel+' Receive Equipment 클릭 실패: '+err.message+' - 이 항목은 건너뜁니다');
      continue;
    }
  }

  if(!committedItemIds.length){
    toast('ZREC 처리 완료 (실제로 접수된 항목 없음)');
    return;
  }

  // 3단계: ZIH08 재조회로 Plant/Location 6507/0052 확인
  toast('ZIH08로 접수 완료 여부 확인 중...');
  try{
    const res=await fetch('/api/zrec/verify_batch', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({itemIds: committedItemIds})
    });
    const vdata=await res.json();
    if(!res.ok || vdata.error) throw new Error(vdata.error||'확인 실패');
    const results=vdata.results||[];
    const ok=results.filter(r=>r.receivedOk);
    const bad=results.filter(r=>!r.receivedOk);

    // ZREC 완료 확인된 항목은 파랑(done)으로 자동 전환 - 준비된 상태(빨강,
    // Ready to process)는 그대로 두고, 실제 접수가 ZIH08로 확인된 것만
    // 파랑으로 바뀐다 (사용자 확정, 2026-08-19: "zrec할 준비가 된것이
    // 빨간색이고, zrec가 완료되면 파란색으로 처리"). 확인 안 된 항목은
    // 빨강 그대로 남아서 다시 확인해야 함을 표시. "해당 오더를 모두"
    // 파랑으로 바꾸라는 요청이라 확인된 시리얼 자신뿐 아니라 같은
    // orderId를 공유하는 다른 행(예: 같은 오더의 다른 품목)까지 전부
    // 포함해서 done 처리한다.
    if(ok.length){
      const okItemIds=new Set(ok.map(r=>r.itemId));
      const okOrderIds=new Set(
        ok.map(r=>{ const row=ROWS.find(x=>x.itemId===r.itemId); return row ? row.orderId : null; }).filter(Boolean)
      );
      // 2026-10-07: "같은 오더 전부 파랑"이 날짜가 안 온(예정) 형제 행까지 파랑으로 만들어,
      // 그 행이 완료로 오인돼 마감까지 나가던 경로(ZRX 67103429 회수 10/8). 확인된 행
      // 자신은 날짜 무관하게, 같은 오더의 다른 행은 오늘 이전/당일 날짜일 때만 포함한다.
      const todayIsoLocal=localTodayIso();
      const doneRows=ROWS.filter(r=>okItemIds.has(r.itemId) || (okOrderIds.has(r.orderId) && r.date<=todayIsoLocal));
      try{
        await persistItemStatus(doneRows.map(r=>r.itemId), 'done');
        doneRows.forEach(r=>{ r.uiStatus='done'; });
        renderAll();
      }catch(err){
        toast('완료 항목 파랑 처리 저장 실패: '+err.message);
      }
    }

    let summary=`ZREC 완료 확인: ${ok.length}/${results.length}건 Plant 6507/Location 0052 확인됨(해당 오더 전체 파랑 처리).`;
    if(bad.length){
      summary+=`\n\n⚠ 아직 확인 안 된 항목(빨강 유지):\n`+bad.map(r=>`${r.orderLabel} (${r.serial}) - Plant='${r.plant}' Location='${r.location}'`).join('\n');
      window.alert(summary);
    } else {
      toast(summary);
    }
  }catch(err){
    toast('ZIH08 완료 확인 조회 실패: '+err.message);
  }
}

// ── Board data model ─────────────────────────────────────────────────────
// ROWS is the single in-memory source of truth for the whole board (one
// entry per item line). Every action (move, edit, delete, type change,
// Done/Ready) mutates ROWS then calls renderAll() to rebuild the table
// from scratch - simplest way to keep same-order merging (rowspan) correct
// after arbitrary edits, rather than patching DOM rowspans incrementally.
let ROWS = [];
let BOARD_FILTER = 'all';  // 메모 라벨 모아보기: all | delayed | check
let SIDE_FIN = {};           // 달력: { 날짜: 마감 완료 오더 수 } (첫 renderAll 이전에 선언돼야 함)
let SIDE_FIN_LOADED = {};    // 달력: 이미 불러온 달
const DATE_LABELS = {};
const COLLAPSED_DATES = new Set();
const EXPORT_DATES = new Set();
// Type/check를 제외한 나머지 7개 board 열, 왼쪽부터의 순서 - addManualRow()
// 모달과 add_manual_note()의 cells 매핑과 반드시 같은 순서여야 함. 여기,
// ROWS 선언 바로 아래에 있어야 한다: buildGroupRows()의 메모 행 렌더링이
// manualRowGroupStarts()를 통해 이 값을 참조하고, 그 렌더링은 최초
// renderAll() 호출(파일 훨씬 아래쪽) 시점에 이미 실행되므로, 이 const가 그
// 호출 지점보다 뒤에 있으면 TDZ(temporal dead zone)로 스크립트 전체가 죽는다
// - 2026-08-27 실사고: 메모 행 추가 버튼을 누르자마자 보드 전체가 하얗게
// 사라짐, 콘솔에 "Cannot access 'MANUAL_ROW_COLUMNS' before initialization".
const MANUAL_ROW_COLUMNS=['제목','M/N','S/N','Customer','Phone','Address','Memo'];

// Type dropdown options. BASE_TYPES is fixed; CUSTOM_TYPES accumulates
// whatever the user adds via the dropdown's "+ 추가" entry - shared across
// every row's dropdown, not just the one it was added from. 2026-08-26 fix:
// this used to be session-only (reset to [] on every load) even though each
// ROW's own chosen type was already being saved - the *list of known custom
// types offered in the dropdown* was the part that wasn't, so a custom type
// had to be retyped via "+ 추가" every time it was needed on a different row
// after a reload. Now seeded from the server (custom_types table via
// __CUSTOM_TYPES__, see render_dashboard_page()) and every newly-added one
// is persisted immediately - see the '__add__' branch below.
const BASE_TYPES = ['배송','회수','배송_퀵','회수_퀵','Bloomberg','Delayed','회수 Delayed'];
// 회수 계열(회수/회수_퀵/회수 Delayed)과 퀵 계열(배송_퀵/회수_퀵) 판별 - 2026-10-16
function isPickupType(t){ return String(t||'').startsWith('회수'); }
function isQuickType(t){ return /_퀵$/.test(String(t||'')); }
let CUSTOM_TYPES = __CUSTOM_TYPES__;
function allTypeOptions(){ return [...BASE_TYPES, ...CUSTOM_TYPES]; }

// 2026-09-13 사용자 요청: 날짜 배너를 "자동화가 안 도는 날"(주말/한국
// 공휴일) + "직접 지정한 휴무일"이면 빨간색으로. 한국 공휴일 이름은 서버
// (holiday_check.kr_holiday_labels, 음력 계산 포함)가 내려주고, 주말은
// iso 문자열만으로 클라이언트가 바로 계산할 수 있어 서버 왕복 없이 처리.
const NONWORKING = __NONWORKING_DAYS__;
const KR_HOLIDAYS = NONWORKING.holidays || {};
const MANUAL_HOLIDAYS = new Set(NONWORKING.manual || []);
function nonworkingInfo(iso){
  if(MANUAL_HOLIDAYS.has(iso)) return {non:true, label:'수동 지정 휴무일'};
  if(KR_HOLIDAYS[iso]) return {non:true, label:KR_HOLIDAYS[iso]};
  const dow = new Date(iso+'T00:00:00').getDay();
  if(dow===0 || dow===6) return {non:true, label:'주말'};
  return {non:false, label:''};
}
function typeClassFor(v){
  if(isPickupType(v)) return 'type-pickup';
  if(v==='배송' || v==='배송_퀵') return 'type-delivery';
  return 'type-other';
}

(function loadInitialData(){
  const data = __BOARD_DATA__;
  data.forEach(d=>{
    DATE_LABELS[d.iso] = d.label;
    d.rows.forEach(r=>ROWS.push({...r, date: d.iso, uiStatus: r.boardStatus || ''}));
  });
})();

// 새 오더 폴링 - SAP 수집이 workbench.db에 오더를 넣어도 이미 열려있는 탭은
// 새로고침 전까진 몰랐다(사용자 지적, 2026-08-14). 20초마다 서버 board_data()를
// 다시 가져와서 로컬 ROWS에 없는 itemId만 "새 오더"로 판단해 추가한다 - 기존
// 행은 절대 건드리지 않음(진행 중인 드래그/편집/선택 상태를 깨뜨리지 않기
// 위해 add-only로 제한). 예외는 사용자가 텍스트로 직접 편집하는 필드가 아니라
// 버튼/드롭다운/우클릭 메뉴로 원자적으로 바뀌는 상태값들 - 다른 탭/세션에서
// 트리거된 변화도 새로고침 없이 20초 안에 반영되도록 이 필드들만 기존 행에도
// 동기화한다: portalError/podError/snReloError(todo b, 2026-08-18 / S/N
// 반영, 2026-08-19), boardStatus(Done/Ready/Cancel/초기화 행 색칠)와
// highlights(우클릭 셀 강조), itemType(타입 드롭다운) - 2026-09-01, 사용자
// 지적("업데이트가 새로고침 없이 안 보임") 계기로 확대. memo/customer/
// address 등 자유 텍스트 필드는 여전히 건드리지 않음(다른 탭에서 타이핑
// 중인 내용을 덮어쓸 위험).
let _pollingBoard = false;
async function pollForNewOrders(){
  if(_pollingBoard) return;  // 이전 폴링이 아직 안 끝났으면 겹쳐서 돌리지 않음
  _pollingBoard = true;
  try{
    const res = await fetch('/api/board_data');
    if(!res.ok) return;
    const data = await res.json();
    const byItemId = new Map(ROWS.map(r=>[r.itemId, r]));
    let added = 0;
    let changed = false;
    data.forEach(d=>{
      d.rows.forEach(r=>{
        const existing = byItemId.get(r.itemId);
        if(existing){
          if(existing.portalError!==r.portalError || existing.portalErrorMsg!==r.portalErrorMsg
             || existing.podError!==r.podError || existing.podErrorMsg!==r.podErrorMsg
             || existing.snReloError!==r.snReloError || existing.snReloErrorMsg!==r.snReloErrorMsg){
            existing.portalError=r.portalError; existing.portalErrorMsg=r.portalErrorMsg;
            existing.podError=r.podError; existing.podErrorMsg=r.podErrorMsg;
            existing.snReloError=r.snReloError; existing.snReloErrorMsg=r.snReloErrorMsg;
            changed = true;
          }
          if(existing.boardStatus!==r.boardStatus){
            existing.boardStatus=r.boardStatus;
            existing.uiStatus=r.boardStatus || '';
            changed = true;
          }
          if(existing.itemType!==r.itemType){
            existing.itemType=r.itemType;
            changed = true;
          }
          if(JSON.stringify(existing.highlights||{})!==JSON.stringify(r.highlights||{})){
            existing.highlights=r.highlights;
            changed = true;
          }
          return;
        }
        if(!DATE_LABELS[d.iso]) DATE_LABELS[d.iso] = d.label;
        const newRow={...r, date: d.iso, uiStatus: r.boardStatus || ''};
        ROWS.push(newRow);
        byItemId.set(r.itemId, newRow);
        added++;
      });
    });
    if(added > 0){
      renderAll();
      toast('새 오더 '+added+'건 반영됨');
    } else if(changed){
      renderAll();
    }
  }catch(err){
    // 조용히 무시 - 네트워크 순간 끊김 등으로 매 폴링마다 토스트 띄우면 방해됨
  }finally{
    _pollingBoard = false;
  }
}
// 2026-09-01: 20초 폴링은 "바뀌면 바로 반영돼야지"라기엔 너무 느리다는 지적.
// 그렇다고 board_data() 전체(오더+아이템 조인 여러 개)를 1~2초마다 통째로
// 다시 불러오는 건 낭비 - 대신 /api/board_version(그냥 SQLite data_version
// pragma 하나, 사실상 공짜)을 1.5초마다 찔러보고 값이 바뀐 걸 감지했을 때만
// 진짜 pollForNewOrders()를 부른다. data_version은 이 서버의 HTTP 핸들러를
// 거치지 않고 SAP 자동루프/ZREC/Portal 등 다른 프로세스가 workbench.db에
// 직접 쓴 변경까지도 잡아내므로(각 write 경로마다 "바뀜"을 따로 통지하게
// 만드는 것보다 훨씬 안전함), 체감상 거의 즉시(최대 1.5초) 반영된다. 혹시
// 이 감지를 놓치는 경우에 대비한 belt-and-suspenders로 30초 전체 폴백도
// 유지.
let _lastBoardVersion = null;
let _checkingVersion = false;
async function checkBoardVersion(){
  if(_checkingVersion) return;
  _checkingVersion = true;
  try{
    const res = await fetch('/api/board_version');
    if(!res.ok) return;
    const {version} = await res.json();
    if(_lastBoardVersion===null){ _lastBoardVersion = version; return; }
    if(version !== _lastBoardVersion){
      _lastBoardVersion = version;
      pollForNewOrders();
    }
  }catch(err){
    // 조용히 무시
  }finally{
    _checkingVersion = false;
  }
}
setInterval(checkBoardVersion, 1500);
setInterval(pollForNewOrders, 30000);
const DELETED_ITEM_RETENTION_DAYS = __DELETED_ITEM_RETENTION_DAYS__;

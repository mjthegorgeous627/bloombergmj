/* /m 모바일 화면 동작 - PC(v2) 화면과 같은 전역(ROWS, renderAll, processDone, runPortalBatch, persistItemField …)을 그대로 쓴다.
   카드는 PC 표의 숨겨진 .row-check 체크박스를 대신 눌러 주는 리모컨 역할이라, 버튼이 하는 일은 PC와 완전히 같다. */
(function(){
  const $=id=>document.getElementById(id);
  const mSel=new Set();        // 선택된 rowKey (PC의 .row-check:checked 와 같은 것)
  const mOpen=new Set();       // 펼친 카드
  let mErrOnly=false;
  let mPage='board';

  function el(tag,cls,text){const e=document.createElement(tag); if(cls) e.className=cls; if(text!==undefined) e.textContent=text; return e}

  // ---- 카톡 전송 문구와 같은 제품명 변환 (kakao_handler.py PRODUCT_MAP) ----
  const PRODUCT_MAP=[[['KEYBOARD','5'],'키보드5'],[['KEYBOARD'],'일반키보드'],[['MONITOR'],'모니터'],[['PC'],'PC'],[['ROUTER'],'라우터'],[['SERVER'],'서버']];
  function productName(desc,material){
    const d=String(desc||'').toUpperCase(); const m=String(material||'').trim();
    if(m==='10045246'||(d.includes('KEYBOARD')&&/\bUS\b/.test(d))) return 'US일반키보드';
    for(const [kw,name] of PRODUCT_MAP) if(kw.every(k=>d.includes(k))) return name;
    return String(desc||'').trim().split(/\s+/)[0]||'';
  }
  function actionOf(t){
    t=String(t||'');
    if(t.includes('회수')) return '회수';
    if(t==='배송'||t.startsWith('Delayed')) return '배송';
    return t||'배송';
  }
  function summaryOf(rows){
    const first=rows[0];
    if(!first.orderNo){ // 메모 행
      const t=rows.map(r=>r.description||'').filter(Boolean).join(' / ');
      return t||first.memo||'';
    }
    const cnt=new Map();
    rows.forEach(r=>{
      const key=productName(r.description,r.material)+'\u0000'+actionOf(r.itemType);
      cnt.set(key,(cnt.get(key)||0)+(parseInt(r.qty,10)||1));
    });
    return '['+[...cnt].map(([k,n])=>{const [p,a]=k.split('\u0000'); return `${p} ${n}개 ${a}`}).join(' / ')+']';
  }
  function groupClass(rows){
    const st=rows.map(r=>r.uiStatus||'');
    let c='';
    if(st.some(s=>s==='pending')) c='ready';
    else if(st.every(s=>s==='cancelled')) c='cancel';
    else if(st.every(s=>s==='done'||s==='billed')) c=st.some(s=>s==='billed')?'billed':'done';
    if(rows.some(r=>r.portalError||r.podError||r.snReloError)) c+=' err';
    return c;
  }

  // ---- 선택(체크) 동기화 ----
  function checkboxFor(k){ return document.querySelector('.row-check[data-row-key="'+(window.CSS&&CSS.escape?CSS.escape(k):k)+'"]'); }
  function syncFromDom(){
    mSel.clear();
    document.querySelectorAll('.row-check:checked').forEach(cb=>mSel.add(cb.dataset.rowKey));
  }
  function setGroup(rows,on){
    rows.forEach(r=>{
      const cb=checkboxFor(r.rowKey); if(!cb) return;
      cb.checked=on; cb.dispatchEvent(new Event('change'));
    });
    syncFromDom(); refreshBar();
  }
  function refreshBar(){
    const n=mSel.size;
    $('mBarN').textContent=n;
    $('mBar').classList.toggle('show', n>0 && mPage==='board');
    document.querySelectorAll('.m-card').forEach(c=>{
      const keys=(c.dataset.keys||'').split(',');
      c.classList.toggle('sel', keys.some(k=>mSel.has(k)));
    });
  }

  // ---- 카드 ----
  function matches(rows,q){
    if(!q) return true;
    const hay=rows.map(r=>[r.orderLabel,r.customer,r.address,r.memo,r.description,r.material,r.serial,r.deliveryNo].join(' ')).join(' ').toLowerCase();
    return hay.includes(q);
  }
  function itemField(row,label,field,ph){
    const wrap=el('div'); wrap.appendChild(el('div','m-lbl',label));
    const inp=el('input','m-in'); inp.value=row[field]||''; inp.placeholder=ph||'';
    if(field==='serial' && !inp.value) inp.classList.add('empty');
    inp.addEventListener('change',()=>{
      const v=inp.value.trim(); row[field]=v; inp.classList.toggle('empty', field==='serial' && !v);
      persistItemField(row.itemId,field,v).then(()=>{ if(typeof toast==='function') toast('저장됨'); })
        .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
    });
    wrap.appendChild(inp); return wrap;
  }
  function buildCard(iso,g){
    const rows=g.rows, first=rows[0], gk=iso+'|'+g.key;
    const card=el('div','m-card '+groupClass(rows));
    card.dataset.keys=rows.map(r=>r.rowKey).join(',');
    const chk=el('div','m-chk'); const cb=el('input'); cb.type='checkbox';
    cb.checked=rows.some(r=>mSel.has(r.rowKey));
    cb.addEventListener('change',()=>setGroup(rows,cb.checked)); chk.appendChild(cb); card.appendChild(chk);
    const body=el('div','m-body');
    const l1=el('div','m-l1'); l1.appendChild(el('span','m-on',first.orderLabel||'(번호없음)'));
    const tags=el('div','m-tags');
    [...new Set(rows.map(r=>r.itemType).filter(Boolean))].forEach(t=>tags.appendChild(el('span','m-tag',t)));
    if(rows.some(r=>r.portalError)) tags.appendChild(el('span','m-tag r','Portal 오류'));
    if(rows.some(r=>r.podError)) tags.appendChild(el('span','m-tag r','POD 오류'));
    if(rows.some(r=>r.snReloError)) tags.appendChild(el('span','m-tag r','S/N 오류'));
    l1.appendChild(tags); body.appendChild(l1);
    body.appendChild(el('div','m-sum',summaryOf(rows)));
    const sub=[first.customer,first.address].filter(Boolean).join(' · ');
    if(sub) body.appendChild(el('div','m-sub',sub));
    if(first.memo) body.appendChild(el('span','m-memo',first.memo));

    const open=mOpen.has(gk);
    body.addEventListener('click',e=>{
      if(e.target.closest('input,textarea,a,button')) return;
      if(mOpen.has(gk)) mOpen.delete(gk); else mOpen.add(gk);
      renderMobile();
    });
    if(open){
      const x=el('div','m-exp');
      const kv=(k,v,link)=>{ if(!v) return; const d=el('div'); d.appendChild(el('span','k',k));
        if(link){const a=el('a',null,v); a.href='tel:'+String(v).replace(/[^0-9+]/g,''); d.appendChild(a);} else d.appendChild(document.createTextNode(v));
        x.appendChild(d); };
      kv('고객',first.customer); kv('전화',first.phone,true); kv('주소',first.address); kv('코드',first.deliveryNo);
      rows.forEach(r=>{
        const it=el('div','m-item');
        const head=el('div','m-it'); head.appendChild(el('span','m-st '+(r.uiStatus||'')));
        head.appendChild(document.createTextNode(`${r.itemType||''} · ${r.description||'(품목)'} · M/N ${r.material||'-'}`));
        it.appendChild(head);
        if(r.orderNo){ it.appendChild(itemField(r,'S/N (Serial 등록에 사용)','serial','S/N 입력…')); }
        else { it.appendChild(itemField(r,'내용','description','')); it.appendChild(itemField(r,'S/N','serial','')); }
        x.appendChild(it);
      });
      const errs=[]; const seen=new Set();
      rows.forEach(r=>{ [['Portal',r.portalErrorMsg],['POD',r.podErrorMsg],['S/N',r.snReloErrorMsg]].forEach(([n,m])=>{ if(m&&!seen.has(n+m)){seen.add(n+m); errs.push(`⚠ ${n}: ${m}`)} }) });
      errs.forEach(t=>x.appendChild(el('div','m-err',t)));
      const ml=el('div'); ml.appendChild(el('div','m-lbl','메모 (수정 가능)'));
      const ta=el('textarea','m-in m-ta'); ta.value=first.memo||''; ta.rows=2;
      ta.addEventListener('change',()=>{
        const v=ta.value.trim(); rows.forEach(r=>r.memo=v);
        persistOrderField(rows.map(r=>r.itemId),'memo',v).then(()=>{ toast('메모 저장됨'); renderMobile(); })
          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
      });
      ml.appendChild(ta); x.appendChild(ml);
      body.appendChild(x);
    }
    card.appendChild(body); return card;
  }

  function renderMobile(){
    if(typeof ROWS==='undefined') return;
    const list=$('mList'), chips=$('mChips');
    const q=($('mSearch').value||'').trim().toLowerCase();
    const todayIso=new Date().toLocaleDateString('sv-SE');
    const keepScroll=window.scrollY;
    list.textContent=''; chips.textContent='';
    // 긴급 집계(오더 단위)
    const errOrders={p:new Set(),d:new Set(),s:new Set()};
    ROWS.forEach(r=>{ if(r.portalError) errOrders.p.add(r.orderId); if(r.podError) errOrders.d.add(r.orderId); if(r.snReloError) errOrders.s.add(r.orderId); });
    const tot=new Set([...errOrders.p,...errOrders.d,...errOrders.s]).size;
    const al=$('mAlert');
    if(tot){ al.hidden=false;
      $('mAlertTxt').textContent='⚠ 긴급 '+tot+'건'+(errOrders.p.size?' · Portal 오류 '+errOrders.p.size:'')+(errOrders.d.size?' · POD 오류 '+errOrders.d.size:'')+(errOrders.s.size?' · S/N 오류 '+errOrders.s.size:'')+(mErrOnly?' (전체 보기›)':'');
    } else { al.hidden=true; mErrOnly=false; }
    let shown=0;
    allDatesSorted().forEach(iso=>{
      let groups=groupRows(ROWS.filter(r=>r.date===iso));
      const total=groups.length;
      groups=groups.filter(g=>matches(g.rows,q) && (!mErrOnly || g.rows.some(r=>r.portalError||r.podError||r.snReloError)));
      const chip=el('a','m-chip'+(iso===todayIso?' today':''),iso.slice(5).replace('-','/')+(iso===todayIso?' 오늘 ':' ')); chip.appendChild(el('i',null,String(total)));
      chip.href='#m-d-'+iso; chips.appendChild(chip);
      if(!groups.length) return;
      const h=el('div','m-dh',labelFor(iso)+' · '+groups.length+'건'); h.id='m-d-'+iso; h.style.scrollMarginTop='108px';
      list.appendChild(h);
      groups.forEach(g=>{ list.appendChild(buildCard(iso,g)); shown++; });
    });
    if(!shown) list.appendChild(el('div','m-sub',q||mErrOnly?'검색 결과가 없습니다.':'표시할 오더가 없습니다.'));
    refreshBar();
    window.scrollTo(0,keepScroll);
  }

  // PC 화면이 새로 그릴 때마다(상태 변경, 새 오더 감지 등) 모바일도 같이 다시 그린다
  const origRenderAll=renderAll;
  renderAll=function(){ const r=origRenderAll.apply(this,arguments); try{ syncFromDom(); renderMobile(); }catch(e){ console.error('[m] render',e); } return r; };

  // ---- 탭 ----
  const TITLES={board:['▤ 보드','t-board'],work:['⚡ 작업','t-work'],ticket:['🎫 티켓','t-ticket']};
  function showPage(p){
    mPage=p;
    ['board','work','ticket'].forEach(n=>{ $('pg'+n[0].toUpperCase()+n.slice(1)).hidden=(n!==p); });
    $('mTitle').textContent=TITLES[p][0];
    $('mTop').className='m-top '+TITLES[p][1];
    document.querySelectorAll('.m-tab').forEach(t=>t.classList.toggle('on',t.dataset.p===p));
    refreshBar(); window.scrollTo(0,0);
  }
  document.querySelectorAll('.m-tab').forEach(t=>t.addEventListener('click',()=>showPage(t.dataset.p)));
  $('mAlert').addEventListener('click',()=>{ mErrOnly=!mErrOnly; renderMobile(); });
  $('mSearch').addEventListener('input',renderMobile);

  // ---- 상태 표시줄 / 시계 ----
  function tick(){ $('mClock').textContent=new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'}); }
  setInterval(tick,5000); tick();
  async function loopStatus(){
    try{ const d=await (await fetch('/api/sap/status')).json();
      $('mLoopDot').className='m-dot '+(d.loopRunning?'on':'off'); $('mLoopTxt').textContent=d.loopRunning?'루프 실행 중':'루프 중지됨';
    }catch(e){ $('mLoopDot').className='m-dot'; $('mLoopTxt').textContent='상태 확인 실패'; }
  }
  loopStatus(); setInterval(loopStatus,8000);

  // ---- 티켓: PC 사이드바의 숨은 입력/버튼과 연결 ----
  window.mReadTicketNum=function(){
    const v=$('mTkNum').value; const inp=document.getElementById('ticketNumInput');
    if(!inp){ toast('티켓 입력칸을 찾지 못했습니다'); return; }
    inp.value=v; startTicketIntake();
  };
  const TK_DEFAULT=$('mTkCheck').innerHTML;
  setInterval(()=>{
    const src=document.getElementById('ticketCheckBtn'), b=$('mTkCheck');
    if(src){ const t=(src.textContent||'').trim(); if(t && t!=='티켓 확인'){ b.textContent=t; } else if(b.innerHTML!==TK_DEFAULT){ b.innerHTML=TK_DEFAULT; } b.disabled=!!src.disabled; }
    const msg=document.getElementById('ticketIntakeMsg'); $('mTkMsg').textContent=(msg&&msg.style.display!=='none')?(msg.textContent||''):'';
    const nb=document.getElementById('ticketNumBtn'); if(nb) $('mTkNumBtn').disabled=!!nb.disabled;
  },1000);

  syncFromDom(); renderMobile(); showPage('board');
})();

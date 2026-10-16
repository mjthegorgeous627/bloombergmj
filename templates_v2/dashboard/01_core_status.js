function tick(){document.getElementById('clock').textContent=new Date().toLocaleString('ko-KR',{weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',second:'2-digit'})}
setInterval(tick,1000);tick();

// 콘솔 창이 없어진(2026-08-25) 뒤로 SAP 자동루프가 실제로 돌고 있는지 화면에서
// 확인할 방법이 없어져서 추가 - 시계 옆에 점+문구로 표시. 백엔드 /api/sap/status는
// 이미 있었는데(2026-08-25) 프론트에서 아무도 안 부르고 있었음.
async function updateLoopStatus(){
  const el=document.getElementById('sapLoopStatus');
  try{
    const res=await fetch('/api/sap/status');
    const data=await res.json();
    el.classList.remove('running','stopped');
    if(data.loopRunning){
      el.classList.add('running');
      el.querySelector('.txt').textContent='SAP 루프 실행 중 (PID '+data.pids.join(', ')+')';
    }else{
      el.classList.add('stopped');
      el.querySelector('.txt').textContent='SAP 루프 중지됨';
    }
  }catch(err){
    el.classList.remove('running','stopped');
    el.querySelector('.txt').textContent='SAP 루프 상태 확인 실패';
  }
}
updateLoopStatus();
setInterval(updateLoopStatus,5000);

function toggleCtrlCol(id){
  const col=document.getElementById(id);
  const collapsed=col.classList.toggle('collapsed');
  col.querySelector('.ctrl-col-toggle').textContent=collapsed?'▸':'▾';
  syncStickyTopHeight();
}

// Keeps the board's sticky column-header row (already position:sticky) glued
// directly beneath the sticky-top control bar instead of overlapping it.
// Re-measured whenever the control bar's own height can change (column
// collapse/expand, window resize) since that height isn't fixed.
function syncStickyTopHeight(){
  const el=document.querySelector('.sticky-top');
  if(el) document.documentElement.style.setProperty('--sticky-top-h', el.offsetHeight+'px');
}
window.addEventListener('resize', syncStickyTopHeight);
syncStickyTopHeight();

function toast(msg){
  const t=document.createElement('div');
  t.className='toast';
  t.textContent=msg;
  document.body.appendChild(t);
  requestAnimationFrame(()=>t.classList.add('show'));
  setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},1900);
}

// -- Control panel: structure-only placeholders, wired to real actions next --
async function runMorningRoutine(){
  try{
    const res=await fetch('/api/sap/morning_routine', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('아침 루틴 강제 실행됨 (PID '+data.pid+') - 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('아침 루틴 실행 실패: '+err.message);
  }
}
// 2026-10-15: 티켓 확인 - 서버가 숨김 프로세스로 Claude를 띄워 수신함을 읽고 "보드 반영 제안"을 만든다(보드는 수정 안 함).
let _ticketPoll=null;
function _ticketBtn(text, busy){
  const b=document.getElementById('ticketCheckBtn');
  if(!b) return;
  b.textContent=text; b.disabled=!!busy;
}
function _closeTicketReport(){ const o=document.getElementById('ticketReportOverlay'); if(o) o.remove(); }
function _ticketActionText(p){
  if(p.action==='append_memo') return (p.row_label||('행 '+p.row_id))+' Memo에 추가: '+(p.memo_append||'')+'\n+ Order# 아래에 SDSK '+String(p.ticket||'')+' 번호 추가';
  if(p.action==='set_date') return (p.row_label||('행 '+p.row_id))+' 날짜 → '+(p.date||'');
  if(p.action==='new_row'){ const n=p.new_row||{}; return '새 행 만들기 (Order# SDSK '+String(p.ticket||'')+')'+String.fromCharCode(10)+[n.type,n.customer,n.item,n.date].filter(Boolean).join(' / '); }
  return '';
}
async function _applyTicketProposal(p, btn, skipBtn){
  if(!window.confirm('배송장에 반영할까요?\n\n'+_ticketActionText(p)+'\n\n(Ctrl+Z로 되돌릴 수 있습니다)')) return false;
  btn.disabled=true;
  try{
    const res=await fetch('/api/tickets/apply',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({proposal:p})});
    const data=await res.json();
    if(!res.ok||data.error) throw new Error(data.error||'실패');
    toast(data.message);
    if(data.applied===false){ btn.disabled=false; return false; }
    return true;
  }catch(err){ btn.disabled=false; toast('적용 실패: '+err.message); return false; }
}
function _fmtTicketWhen(iso){
  const m=String(iso||'').match(/(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)/); if(!m) return '';
  const h=+m[4]; return (+m[2])+'월 '+(+m[3])+'일 '+(h<12?'오전 ':'오후 ')+(h%12||12)+'시 '+m[5]+'분';
}
function _tkDateInput(init, required){
  // 할 일 날짜: 배송장에 내보낼 때(새 행 날짜/날짜 변경)와 보관함 정렬(날짜 오름차순)에 쓴다
  const di=document.createElement('input'); di.type='date'; di.title='할 일 날짜 (배송장 날짜 / 보관함 정렬 기준)';
  if(/^\d{4}-\d{2}-\d{2}$/.test(init||'')) di.value=init;
  const paint=()=>{ const today=new Date(); const t=today.getFullYear()+'-'+String(today.getMonth()+1).padStart(2,'0')+'-'+String(today.getDate()).padStart(2,'0');
    di.style.border=(!di.value&&required)?'1px solid #d9534f':(di.value&&di.value<t)?'1px solid #d9822b':''; };
  di.style.padding='2px 4px'; paint(); di.addEventListener('input',paint);
  return di;
}
function _tkWithDate(b,di){
  const v=di?di.value:'';
  if(b.action==='new_row') return Object.assign({},b,{new_row:Object.assign({},b.new_row,{date:v})});
  if(b.action==='set_date' && v) return Object.assign({},b,{date:v});
  return b;
}
function _tkReopenBtn(it, card, opts){
  const x=document.createElement('button'); x.className='btn'; x.textContent='처리할 것으로'; x.title='잘못 옮겼을 때 처리할 것 탭으로 되돌립니다 (배송장에 이미 반영한 내용은 되돌려지지 않습니다)';
  x.onclick=async()=>{ try{ const r=await fetch('/api/tickets/reopen',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key})}); const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패'); card.style.transition='opacity .25s'; card.style.opacity='0'; setTimeout(()=>{ card.remove(); if(opts.onChange) opts.onChange(); if(window._tkRefreshPending) window._tkRefreshPending(); },260); toast('처리할 것으로 되돌렸습니다'); }catch(err){ toast('되돌리기 실패: '+err.message); } };
  return x;
}
function _tkProposalBox(it, b, card, opts){
  // 배송장에 만들 행 미리보기(읽기 전용) + 고칠 점을 말로 적는 한 줄. [다시 정리]는 저장된 티켓 글로 제안만 다시 만든다(터미널 안 씀)
  const wrap=document.createElement('div'); wrap.style.cssText='margin:8px 0;padding:8px 10px;background:#f6f8fa;border-radius:6px;font-size:13px';
  const n=b&&b.action==='new_row'?(b.new_row=b.new_row||{}):null;
  if(n){
    const head=document.createElement('div'); head.style.cssText='color:#555;margin-bottom:4px'; head.textContent='새 행 미리보기  (Order# SDSK '+it.key+')'; wrap.appendChild(head);
    const line=(label,node)=>{ const r=document.createElement('div'); r.style.cssText='display:flex;gap:10px;margin:2px 0'; const l=document.createElement('span'); l.style.cssText='flex:none;width:56px;color:#777'; l.textContent=label; r.appendChild(l); node.style.flex='1'; r.appendChild(node); wrap.appendChild(r); };
    const sel=document.createElement('select'); sel.style.cssText='flex:none;width:auto;padding:1px 4px'; const cur=n.type||'배송';
    ['배송','회수','Bloomberg','업무'].concat(['배송','회수','Bloomberg','업무'].includes(cur)?[]:[cur]).forEach(v=>{ const o=document.createElement('option'); o.value=o.textContent=v; if(v===cur) o.selected=true; sel.appendChild(o); });
    n.type=sel.value; sel.onchange=()=>{ n.type=sel.value; };
    const sw=document.createElement('span'); sw.appendChild(sel); line('Type',sw);
    [['품목','item'],['MN','mn'],['S/N','sn'],['받는 곳','customer'],['주소','address'],['메모','memo']].forEach(([lab,k])=>{
      let v=String(n[k]||''); if(k==='sn'||k==='mn') v=v.replace(/\s*,\s*/g,String.fromCharCode(10)); if(!v) return;
      const sp=document.createElement('span'); sp.style.whiteSpace='pre-wrap'; sp.textContent=v; line(lab,sp); if(k==='sn'||k==='mn') n[k]=v; });
  }
  const row=document.createElement('div'); row.style.cssText='display:flex;gap:6px;margin-top:6px';
  const inp=document.createElement('input'); inp.placeholder='(선택) 고칠 점이 있으면 말로 적으세요  예) 서울오피스 리턴함 픽업, 회수 행, S/N 한 줄씩, 날짜 10/20'; inp.style.cssText='flex:1;min-width:0;padding:4px 6px;font:inherit';
  const bt=document.createElement('button'); bt.className='btn'; bt.textContent='다시 정리'; bt.title='저장된 티켓 내용과 이 지시로 TO DO·배송장 제안을 다시 만듭니다 (약 1분, 터미널 사용 안 함)';
  const go=async()=>{
    bt.disabled=true; inp.disabled=true; bt.textContent='정리 중…';
    try{
      const r=await fetch('/api/tickets/regroup',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key,instruction:inp.value})});
      const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패');
      const tmp=document.createElement('div'); _ticketCards(tmp,[d.card],opts); card.replaceWith(tmp.firstChild); toast('다시 정리했습니다');
    }catch(err){ toast('다시 정리 실패: '+err.message); bt.disabled=false; inp.disabled=false; bt.textContent='다시 정리'; }
  };
  bt.onclick=go; inp.onkeydown=(ev)=>{ if(ev.key==='Enter'){ ev.preventDefault(); go(); } };
  row.appendChild(inp); row.appendChild(bt); wrap.appendChild(row);
  return wrap;
}
function _tkEl(tag, css, txt){ const e=document.createElement(tag); if(css) e.style.cssText=css; if(txt!=null) e.textContent=txt; return e; }
function _ticketCards(box, items, opts){
  // 티켓 한 줄(저장소 기록) = 카드 1개. 본 글 = Summary, 안 본 글 = NEW. 글자 14px 한 가지, 굵게는 제목/섹션 이름만.
  opts=opts||{};
  const accent={'New':'#d9534f','Updated':'#4a6fa5','Closed':'#9aa0a6','Mail':'#4f9a6b','Done':'#6b7b8c','List':'#7a5ea8'};
  const INK='#222', SUB='#7a808a';
  (items||[]).forEach(it=>{
    const tag=opts.archived ? (it.closed?'Closed':'Done') : (it.tag||'Updated');
    const card=_tkEl('div','border:1px solid #e4e6ea;border-left:4px solid '+(accent[tag]||SUB)+';border-radius:8px;padding:12px 16px;margin-bottom:14px;background:#fff;color:'+INK+';font-size:14px;line-height:1.55;font-weight:400');
    card.setAttribute('data-tk','1');
    const head=_tkEl('div','font-weight:700;margin-bottom:2px');
    head.appendChild(_tkEl('span','display:inline-block;background:'+(accent[tag]||SUB)+';color:#fff;border-radius:4px;padding:1px 9px;margin-right:10px;font-size:12.5px;font-weight:700;letter-spacing:.2px;vertical-align:1px',tag));
    if(it.kind==='SDSK'){
      const a=_tkEl('a','color:#1a56b0;text-decoration:underline;cursor:pointer;margin-right:8px','SDSK '+it.key);
      a.title='클릭하면 터미널에서 이 티켓을 엽니다';
      a.onclick=async(e)=>{ e.preventDefault(); try{ const r=await fetch('/api/tickets/open',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({number:it.key})}); const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패'); toast(d.message); }catch(err){ toast('티켓 열기 실패: '+err.message); } };
      head.appendChild(a);
    }
    head.appendChild(document.createTextNode(it.title||''));
    card.appendChild(head);
    if(it.srep) card.appendChild(_tkEl('div','color:'+SUB+';margin-bottom:8px',(it.kind==='Mail'?'보낸 사람  ':'담당  ')+it.srep));
    const posts=it.posts||[];
    const seen=posts.filter(p=>p.seen||opts.archived), fresh=posts.filter(p=>!p.seen&&!opts.archived);
    const rowEl=(p,muted)=>{
      const me=/minjung|^mj$/i.test(p.who||'');
      const row=_tkEl('div','display:grid;grid-template-columns:34px 82px 1fr 78px;gap:8px;align-items:baseline;padding:3px 0;color:'+(muted?SUB:INK));
      row.appendChild(_tkEl('span','color:'+SUB,p.n?('#'+p.n):''));
      row.appendChild(_tkEl('span','',(p.who||'')+(me?' (나)':'')));
      row.appendChild(_tkEl('span','white-space:pre-wrap;word-break:break-word',p.text||''));
      row.appendChild(_tkEl('span','color:'+SUB+';text-align:right;white-space:nowrap',p.when||''));
      return row;
    };
    if(seen.length){
      const d=document.createElement('details'); d.open=seen.length<=3;
      const sm=_tkEl('summary','cursor:pointer;font-weight:700;color:'+SUB+';margin-top:10px','Summary');
      d.appendChild(sm); const sw=_tkEl('div','padding:0 10px'); seen.forEach(p=>sw.appendChild(rowEl(p,true))); d.appendChild(sw); card.appendChild(d);
    }
    if(fresh.length){
      const nb=_tkEl('div','margin-top:10px;background:#fff8e1;border-radius:6px;padding:6px 10px');
      nb.appendChild(_tkEl('div','font-weight:700;color:#8a6d1a','NEW'));
      fresh.forEach(p=>nb.appendChild(rowEl(p,false))); card.appendChild(nb);
    }
    const reflected=((it.history||[]).some(x=>x.result==='Memo 반영'))||/배송장에\s*이미\s*반영/.test(it.todo||'');
    const todoTxt=String(it.todo||'').replace(/\s*[\(（]?\s*배송장에\s*이미\s*반영됨\s*[\)）]?/g,'').trim();
    if(todoTxt){
      const none=/^\s*조치 없음/.test(todoTxt);
      const tb=_tkEl('div','margin-top:10px;border-radius:6px;padding:6px 10px;background:'+(none?'#f3f4f6':'#e9f5ee'));
      tb.appendChild(_tkEl('div','font-weight:700;color:'+(none?SUB:'#2e7d4f'),'TO DO'));
      tb.appendChild(_tkEl('div','white-space:pre-wrap',todoTxt)); card.appendChild(tb);
    }
    const bar=_tkEl('div','margin-top:12px;display:flex;gap:8px;align-items:center;flex-wrap:wrap');
    if(opts.watching){
      const ago=_tkAgo(it.last_checked);
      const stamp=_tkEl('span','flex:1 1 200px;color:'+(!ago?'#c0392b':ago.days>=7?'#d9822b':SUB),'마지막 확인  '+(ago?(ago.txt+(ago.days>0?'  ('+ago.days+'일 전)':'  (오늘)')):'확인 안 함')+(reflected?'  ·  배송장에 반영함':''));
      bar.appendChild(stamp);
      const wdate=_tkDateInput(it.due||it.due_auto||((it.board&&it.board.new_row&&it.board.new_row.date)||((it.board&&it.board.action==='new_row')?new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,10):'')), false);
      wdate.onchange=async()=>{ try{ const r=await fetch('/api/tickets/due',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key,due:wdate.value})}); const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패'); toast(wdate.value?('할 일 날짜 '+wdate.value+' 저장 (다음에 열면 날짜순으로 정렬됩니다)'):'할 일 날짜를 지웠습니다'); }catch(err){ toast('날짜 저장 실패: '+err.message); } };
      const wwrap=_tkEl('span','display:inline-flex;gap:6px;align-items:center'); wwrap.appendChild(_tkEl('span','color:'+SUB+';font-size:13px','할 일 날짜')); wwrap.appendChild(wdate);
      if(!it.due&&it.due_auto&&wdate.value===it.due_auto){ const h2=_tkEl('span','color:'+SUB+';font-size:12px','(To do 기준)'); wwrap.appendChild(h2); wdate.addEventListener('input',()=>h2.remove()); }
      bar.appendChild(wwrap);
      const ok=document.createElement('button'); ok.className='btn'; ok.textContent='확인함';
      ok.title='업데이트가 없어도 다시 확인했다는 날짜를 찍습니다';
      ok.onclick=async()=>{ try{ const r=await fetch('/api/tickets/recheck',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key})}); const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패'); const a2=_tkAgo(d.at); stamp.style.color=SUB; stamp.textContent='마지막 확인  '+a2.txt+'  (오늘)'+(reflected?'  ·  배송장에 반영함':''); if(opts.onChange) opts.onChange('stamp'); }catch(err){ toast('저장 실패: '+err.message); } };
      const rs=document.createElement('button'); rs.className='btn'; rs.textContent='Done';
      rs.onclick=async()=>{ if(!window.confirm('Done(완료함)으로 옮길까요?\n해결(Resolved)·종료(Closed)되었거나 내가 더 확인하지 않아도 되는 티켓입니다.')) return;
        try{ const r=await fetch('/api/tickets/resolve',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key})}); const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'실패'); card.style.transition='opacity .25s'; card.style.opacity='0'; setTimeout(()=>{ card.remove(); if(opts.onChange) opts.onChange(); },260); }catch(err){ toast('저장 실패: '+err.message); } };
      const wb=it.board&&it.board.action&&it.board.action!=='none' ? it.board : null;
      card.appendChild(_tkProposalBox(it,wb,card,opts));
      if(wb){
        const lab={append_memo:'Memo에 반영',set_date:'날짜 반영',new_row:'새 행 반영'}[wb.action]||'배송장에 반영';
        const mb=document.createElement('button'); mb.className='btn'; mb.textContent=lab;
        mb.title='반영 안 됐을 수 있어 다시 반영할 수 있습니다. 같은 줄이 이미 Memo에 있으면 중복으로 넣지 않습니다. '+(wb.row_label||'');
        mb.onclick=async()=>{ if(wb.action==='new_row'&&!wdate.value){ toast('날짜를 먼저 입력하세요'); wdate.focus(); return; } if(await _applyTicketProposal(Object.assign({ticket:it.key},_tkWithDate(wb,wdate)),mb,mb)){ try{ const r=await fetch('/api/tickets/recheck',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key,result:'Memo 반영'})}); const d=await r.json(); const a3=_tkAgo(d.at); stamp.style.color=SUB; stamp.textContent='마지막 확인  '+a3.txt+'  (오늘)  ·  배송장에 반영함'; mb.disabled=false; }catch(e){} } };
        bar.appendChild(mb);
      }
      bar.appendChild(ok); bar.appendChild(rs); bar.appendChild(_tkReopenBtn(it,card,opts));
    } else if(opts.archived){
      const h=(it.history||[]).map(x=>String(x.at||'').replace('T',' ').slice(5,16)+' '+x.result).join('  ·  ');
      bar.appendChild(_tkEl('span','flex:1 1 200px;color:'+SUB,'처리 이력: '+(h||'-'))); bar.appendChild(_tkReopenBtn(it,card,opts));
    } else {
      const b=it.board&&it.board.action&&it.board.action!=='none' ? it.board : null;
      card.appendChild(_tkProposalBox(it,b,card,opts));
      bar.appendChild(_tkEl('span','flex:1 1 200px',(!b&&reflected)?'배송장에 반영함':b?('배송장에 반영: '+String(b.question||'').replace(/\s*(하겠습니까|할까요)\?/,'')+(b.row_label?('  ('+b.row_label+')'):'')):''));
      const initD=it.due||it.due_auto||(b&&((b.new_row&&b.new_row.date)||(b.action==='set_date'&&b.date)))||((b&&b.action==='new_row')?new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,10):'');
      const dateIn=_tkDateInput(initD, !!(b&&b.action==='new_row'));
      const lab0=_tkEl('span','color:'+SUB+';font-size:13px','할 일 날짜'); const dwrap=_tkEl('span','display:inline-flex;gap:6px;align-items:center'); dwrap.appendChild(lab0); dwrap.appendChild(dateIn);
      if(!it.due&&it.due_auto&&dateIn.value===it.due_auto){ const hint=_tkEl('span','color:'+SUB+';font-size:12px','(To do 기준)'); dwrap.appendChild(hint); dateIn.addEventListener('input',()=>hint.remove()); }
      bar.appendChild(dwrap);
      const done=async(result,to)=>{
        try{
          const r=await fetch('/api/tickets/handle',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:it.key,result,to,due:(dateIn&&dateIn.value)||''})});
          const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'이동 실패');
          card.style.transition='opacity .25s'; card.style.opacity='0';
          setTimeout(()=>{ card.remove(); if(opts.onChange) opts.onChange(); },260);
        }catch(err){ toast('이동 실패: '+err.message); }
      };
      const btn=(label,title)=>{ const x=document.createElement('button'); x.className='btn'; x.textContent=label; if(title) x.title=title; return x; };
      if(b){
        const lab={append_memo:'Memo에 반영',set_date:'날짜 반영',new_row:'새 행 반영'}[b.action]||'배송장에 반영';
        const ok=btn(lab,'배송장에 반영한 뒤 보관함으로 옮깁니다');
        ok.onclick=async()=>{ if(b.action==='new_row'&&!dateIn.value){ toast('날짜를 먼저 입력하세요'); dateIn.focus(); return; } if(await _applyTicketProposal(Object.assign({ticket:it.key},_tkWithDate(b,dateIn)),ok,ok)) done('Memo 반영','watching'); };
        bar.appendChild(ok);
      }
      const bo=btn('보관함','배송장 반영 없이 보관함으로 옮깁니다. 해결될 때까지 중간중간 다시 확인할 티켓');
      bo.onclick=()=>done('보관함','watching'); bar.appendChild(bo);
      const dn=btn('Done','완료함으로 옮깁니다. 해결(Resolved)·종료(Closed)됐거나 나와 상관없는 티켓');
      dn.onclick=()=>done('Done','done'); bar.appendChild(dn);
    }
    card.appendChild(bar);
    box.appendChild(card);
  });
}
function _tkOpenTicket(num){
  return fetch('/api/tickets/open',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({number:num})})
    .then(r=>r.json().then(d=>{ if(!r.ok||d.error) throw new Error(d.error||'실패'); toast(d.message); }))
    .catch(err=>toast('티켓 열기 실패: '+err.message));
}
function _tkAgo(iso){
  if(!iso) return null;
  const d=new Date(iso); if(isNaN(d)) return null;
  const days=Math.floor((Date.now()-d.getTime())/86400000);
  return {txt:(d.getMonth()+1)+'/'+d.getDate()+' '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0'), days};
}
function _openTicketReport(when){
  _closeTicketReport();
  const o=document.createElement('div');
  o.id='ticketReportOverlay';
  o.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;display:flex;align-items:center;justify-content:center';
  o.onclick=(e)=>{ if(e.target===o) _closeTicketReport(); };
  const box=document.createElement('div');
  box.style.cssText='font-size:14px;background:#fff;color:#222;width:min(860px,94vw);height:min(82vh,820px);display:flex;flex-direction:column;overflow:hidden;border-radius:10px;padding:16px 20px;box-shadow:0 8px 30px rgba(0,0,0,.3)';
  const head=document.createElement('div');
  head.style.cssText='flex:none;display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px;flex-wrap:wrap';
  const tabs=document.createElement('div'); tabs.style.cssText='display:flex;gap:6px';
  const mk=(id,label)=>{ const t=document.createElement('button'); t.className='btn'; t.textContent=label; t.dataset.tab=id; tabs.appendChild(t); return t; };
  const tPend=mk('pending','처리할 것'), tWatch=mk('watch','보관함'), tArch=mk('archive','완료함');
  head.appendChild(tabs);
  const info=document.createElement('span'); info.style.cssText='color:#7a808a;margin-left:auto'; head.appendChild(info);
  const x=document.createElement('button'); x.textContent='닫기'; x.className='btn'; x.onclick=_closeTicketReport; head.appendChild(x);
  box.appendChild(head);
  const search=document.createElement('input'); search.placeholder='검색 (SDSK 번호, 이름, 내용)';
  search.style.cssText='flex:none;display:none;width:100%;box-sizing:border-box;margin-bottom:10px;padding:6px 8px'; box.appendChild(search);
  const list=document.createElement('div'); list.style.cssText='flex:1 1 auto;min-height:0;overflow-y:auto;padding-right:4px'; box.appendChild(list);
  let mode='pending';
  const setActive=()=>{ [tPend,tWatch,tArch].forEach(t=>{ const on=t.dataset.tab===mode; t.style.fontWeight=on?'800':'400'; t.style.background=on?'#e9eef7':''; t.style.borderColor=on?'#4a6fa5':''; }); };
  const whenTxt=()=>(when?('마지막 조회 '+_fmtTicketWhen(when)):'');
  async function showPending(){
    mode='pending'; setActive(); search.style.display='none'; list.innerHTML='';
    const d=await (await fetch('/api/tickets/pending')).json();
    const upd=()=>{
      const n=list.querySelectorAll('[data-tk]').length; info.textContent='처리할 것 '+n+'건  ·  '+whenTxt();
      tPend.textContent='처리할 것 '+n;
      if(!n && !list.querySelector('.tk-empty')){ const e=_tkEl('div','color:#7a808a;padding:20px 0;text-align:center','처리할 것이 없습니다.'); e.className='tk-empty'; list.appendChild(e); }
    };
    _ticketCards(list, d.cards||[], {onChange:upd}); upd();
  }
  async function showWatch(){
    mode='watch'; setActive(); search.style.display='none'; list.innerHTML='';
    const d=await (await fetch('/api/tickets/watching')).json();
    const upd=()=>{
      const n=list.querySelectorAll('[data-tk]').length; tWatch.textContent='보관함 '+n;
      info.textContent='보관함 '+n+'건  ·  할 일 날짜가 빠른 순서 (날짜 직접 입력 > To do의 날짜, 없으면 맨 아래)';
      if(!n && !list.querySelector('.tk-empty')){ const e=_tkEl('div','color:#7a808a;padding:20px 0;text-align:center','보관함이 비어 있습니다. 처리할 것에서 [보관함]을 누르면 여기로 옵니다.'); e.className='tk-empty'; list.appendChild(e); }
    };
    _ticketCards(list, d.cards||[], {watching:true, onChange:(w)=>{ if(w!=='stamp') upd(); }}); upd();
  }
  async function showArchive(){
    mode='archive'; setActive(); search.style.display='block'; list.innerHTML='';
    const d=await (await fetch('/api/tickets/archive?q='+encodeURIComponent(search.value||''))).json();
    info.textContent='완료함 (티켓 '+d.total+'개)  ·  해결·종료되었거나 나와 상관없는 티켓';
    _ticketCards(list, d.items||[], {archived:true});
    if(!(d.items||[]).length) list.appendChild(_tkEl('div','color:#7a808a;padding:20px 0;text-align:center','완료함이 비어 있습니다.'));
  }
  tPend.onclick=()=>showPending(); tWatch.onclick=()=>showWatch(); tArch.onclick=()=>showArchive();
  let tmr=null; search.oninput=()=>{ clearTimeout(tmr); tmr=setTimeout(()=>{ showArchive(); },250); };
  o.appendChild(box); document.body.appendChild(o);
  window._tkRefreshPending=()=>{ if(document.getElementById('ticketReportOverlay')) showPending().catch(()=>{}); };
  showPending().catch(err=>toast('불러오기 실패: '+err.message));
  // 탭 이름에 건수 미리 표시
  fetch('/api/tickets/watching').then(r=>r.json()).then(d=>{ tWatch.textContent='보관함 '+(d.cards||[]).length; }).catch(()=>{});
}
let _tkIntakeTimer=null;
async function startTicketIntake(){
  // 사이드바 '번호로 읽기': SDSK 숫자 / 숫자만, 여러 개는 쉼표·공백·줄바꿈
  const inp=document.getElementById('ticketNumInput'), btn=document.getElementById('ticketNumBtn'), msg=document.getElementById('ticketIntakeMsg');
  const nums=(inp.value.match(/\d{8,12}/g)||[]).filter((v,i,a)=>a.indexOf(v)===i);
  if(!nums.length){ toast('티켓 번호를 찾지 못했습니다. 예: SDSK 1339513698 또는 1339513698'); inp.focus(); return; }
  if(!window.confirm('다음 티켓을 터미널에서 읽습니다 (티켓당 1~2분):\n\n'+nums.map(n=>'SDSK '+n).join('\n')+'\n\n읽는 동안 마우스·키보드를 건드리지 마세요. 시작할까요?')) return;
  try{
    const r=await fetch('/api/tickets/check',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({numbers:nums.join(',')})});
    const d=await r.json(); if(!r.ok||d.error) throw new Error(d.error||'시작 실패');
    if(!d.started){ toast(d.message); return; }
  }catch(err){ toast('시작 실패: '+err.message); return; }
  btn.disabled=true; inp.disabled=true; msg.style.display='block'; msg.textContent='읽는 중… (SDSK '+nums.join(', ')+')';
  const t0=Date.now(); clearInterval(_tkIntakeTimer);
  _tkIntakeTimer=setInterval(async()=>{
    try{
      const st=await (await fetch('/api/tickets/status')).json();
      const sec=Math.round((Date.now()-t0)/1000);
      if(st.state==='waiting'||st.state==='running'){ msg.textContent='읽는 중… '+sec+'초 (SDSK '+nums.join(', ')+') 마우스·키보드를 건드리지 마세요'; return; }
      clearInterval(_tkIntakeTimer); btn.disabled=false; inp.disabled=false;
      if(st.state==='done'){ msg.textContent='완료 · '+(st.message||'')+' — [티켓 결과]에서 확인'; inp.value=''; toast('티켓 읽기 완료 - 티켓 결과 > 처리할 것에 올렸습니다'); if(window._tkRefreshPending) window._tkRefreshPending(); }
      else { msg.textContent='실패 · '+(st.message||''); toast('티켓 읽기 실패: '+(st.message||'')); }
    }catch(e){ /* 서버 재시작 중 등 - 다음 주기에 재시도 */ }
  },3000);
}
async function showTicketReport(){
  try{
    const st=await (await fetch('/api/tickets/status')).json();
    _openTicketReport(st.finished);
    if(false) toast('아직 보여줄 티켓 확인 결과가 없습니다 ('+(st.state||'idle')+')');
  }catch(err){ toast('결과를 불러오지 못했습니다: '+err.message); }
}
async function _pollTicketCheck(){
  try{
    const st=await (await fetch('/api/tickets/status')).json();
    if(st.state==='waiting'){ _ticketBtn('티켓 확인 (SAP 사이클 대기중…)', true); return; }
    if(st.state==='running'){ _ticketBtn('티켓 확인 중… (PC 만지지 마세요)', true); return; }
    clearInterval(_ticketPoll); _ticketPoll=null; _ticketBtn('티켓 확인', false);
    if(st.state==='done'){ toast('티켓 확인 완료 - 결과를 엽니다'); _openTicketReport(st.finished); }
    else if(st.state==='error'){ toast('티켓 확인 실패: '+(st.message||'')); }
  }catch(err){ /* 서버 재시작 중 등 - 다음 폴링에서 재시도 */ }
}
async function runTicketCheck(){
  if(!window.confirm('블룸버그 터미널에서 새 티켓을 읽습니다 (보통 2~5분).\n\n- 읽는 동안 마우스·키보드를 건드리지 마세요.\n- SAP 자동루프는 이 동안 잠깐 대기합니다.\n- 보드는 바뀌지 않고, 반영 제안만 만들어집니다.\n\n시작할까요?')) return;
  try{
    const res=await fetch('/api/tickets/check', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'시작 실패');
    if(!data.started){ toast(data.message); return; }
    _ticketBtn('티켓 확인 (시작 중…)', true);
    if(_ticketPoll) clearInterval(_ticketPoll);
    _ticketPoll=setInterval(_pollTicketCheck, 3000);
  }catch(err){ toast('티켓 확인 시작 실패: '+err.message); }
}
async function runSap(kind){
  try{
    if(kind==='start'){
      const res=await fetch('/api/sap/start_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'시작 실패');
      toast(data.started ? 'SAP 자동루프 시작됨 (PID '+data.pid+')' : data.message);
      updateLoopStatus();
    } else if(kind==='stop'){
      if(!window.confirm('SAP 자동루프를 지금 중지할까요? (워치독도 다시 켤 때까지 재시작하지 않습니다)')) return;
      const res=await fetch('/api/sap/stop_loop', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'중지 실패');
      toast(data.message);
      updateLoopStatus();
    } else if(kind==='all'){
      const res=await fetch('/api/sap/run_all', {method:'POST'});
      const data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'조회 실패');
      toast(data.message);
    }
  }catch(err){
    toast('SAP '+kind+' 실패: '+err.message);
  }
}
// 2026-08-28: 트레이 아이콘 우클릭 메뉴와 같은 조작을 시계 옆 상태 점에서도
// 할 수 있게 - 클릭하면 열리는 드롭다운. 문서 아무 곳이나 클릭하면 닫힘.
function toggleSapLoopMenu(ev){
  ev.stopPropagation();
  document.getElementById('sapLoopMenu').classList.toggle('open');
}
function closeSapLoopMenu(){
  document.getElementById('sapLoopMenu').classList.remove('open');
}
document.addEventListener('click', closeSapLoopMenu);

// 2026-09-13 사용자 요청: 시계 옆 📅 버튼 - 주말/한국 공휴일(자동화 비운영일)
// 을 달력 숫자로 빨갛게 보여주고, 클릭으로 수동 지정/해제도 가능하게. 색
// 판단은 이미 있는 nonworkingInfo()/MANUAL_HOLIDAYS를 그대로 재사용 -
// 서버의 /api/nonworking_days/set에 저장되는 것도 동일.
let CAL_VIEW = (() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; })();

function _calIso(y, m, day){
  return `${y}-${String(m + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function toggleCalendarPanel(ev){
  ev.stopPropagation();
  const panel = document.getElementById('calPanel');
  const opening = !panel.classList.contains('open');
  panel.classList.toggle('open');
  if(opening){
    const today = new Date();
    CAL_VIEW = { y: today.getFullYear(), m: today.getMonth() };
    renderCalendarPanel();
  }
}
function closeCalendarPanel(){
  document.getElementById('calPanel').classList.remove('open');
}
document.addEventListener('click', closeCalendarPanel);

function calShiftMonth(delta, ev){
  ev.stopPropagation();
  CAL_VIEW.m += delta;
  if(CAL_VIEW.m < 0){ CAL_VIEW.m = 11; CAL_VIEW.y -= 1; }
  if(CAL_VIEW.m > 11){ CAL_VIEW.m = 0; CAL_VIEW.y += 1; }
  renderCalendarPanel();
}

function calToggleDay(iso, ev){
  ev.stopPropagation();
  const willAdd = !MANUAL_HOLIDAYS.has(iso);
  if(willAdd) MANUAL_HOLIDAYS.add(iso); else MANUAL_HOLIDAYS.delete(iso);
  fetch('/api/nonworking_days/set', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({date: iso, on: willAdd}),
  }).catch(() => {});
  renderCalendarPanel();
}

function renderCalendarPanel(){
  const {y, m} = CAL_VIEW;
  const panel = document.getElementById('calPanel');
  const startOffset = new Date(y, m, 1).getDay();  // 0=일요일
  const daysInMonth = new Date(y, m + 1, 0).getDate();
  const now = new Date();
  const todayIso = _calIso(now.getFullYear(), now.getMonth(), now.getDate());

  let html = `<div class="cal-head">`
    + `<button type="button" onclick="calShiftMonth(-1,event)">◂</button>`
    + `<span class="cal-title">${y}. ${String(m + 1).padStart(2, '0')}</span>`
    + `<button type="button" onclick="calShiftMonth(1,event)">▸</button>`
    + `</div><div class="cal-grid">`;
  ['일', '월', '화', '수', '목', '금', '토'].forEach(w => { html += `<div class="cal-wd">${w}</div>`; });
  for(let i = 0; i < startOffset; i++) html += `<div></div>`;
  for(let day = 1; day <= daysInMonth; day++){
    const iso = _calIso(y, m, day);
    const info = nonworkingInfo(iso);
    const cls = ['cal-day'];
    if(info.non) cls.push('nonworking');
    if(iso === todayIso) cls.push('today');
    const title = info.non ? info.label : '클릭: 이 날짜를 수동으로 휴무일(빨간색) 지정';
    html += `<div class="${cls.join(' ')}" title="${title}" onclick="calToggleDay('${iso}',event)">${day}</div>`;
  }
  html += `</div><div class="cal-legend">빨간 숫자 = 그날 SAP 자동화 비운영일(주말/공휴일/수동 지정). 숫자를 클릭하면 수동으로 지정·해제할 수 있습니다.</div>`;
  panel.innerHTML = html;
}
async function restartWorkbench(){
  if(!window.confirm('Workbench를 강제로 재시작할까요?\n1~2초간 서버가 끊겼다가 새로 뜨고, 이 페이지는 자동으로 새로고침됩니다.')) return;
  try{
    const res=await fetch('/api/workbench/restart', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'재시작 트리거 실패');
    toast(data.message);
    setTimeout(()=>location.reload(), 3000);
  }catch(err){
    toast('Workbench 재시작 실패: '+err.message);
  }
}
async function runSapSession(idx){
  try{
    const res=await fetch('/api/sap/run_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({session:idx})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast(data.message);
  }catch(err){
    toast('세션 '+idx+' 조회 실패: '+err.message);
  }
}
async function runOrderNoAction(){
  const input=document.getElementById('orderNoInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/run_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    toast('오더번호 반영 시작됨: '+v+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더번호 반영 실패: '+err.message);
  }
}
async function runOrderJumpAction(){
  const input=document.getElementById('orderJumpInput');
  const v=input.value.trim();
  if(!v){toast('오더번호를 입력하세요');return}
  try{
    const res=await fetch('/api/sap/open_order', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({orderNo:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    const tcode = v.startsWith('6') ? 'VA02' : 'VA03';
    toast('오더 '+v+' ('+tcode+') 여는 중 (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('오더 바로가기 실패: '+err.message);
  }
}
async function runSerialLookupAction(){
  const input=document.getElementById('serialLookupInput');
  const v=input.value.trim();
  if(!v){toast('시리얼을 입력하세요');return}
  try{
    const res=await fetch('/api/sap/open_serial', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({serial:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    if(!data.started){toast(data.message);return}
    toast('시리얼 '+v+' ZIH08 조회 중 (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('시리얼 조회 실패: '+err.message);
  }
}
async function runNewSessionAction(){
  const input=document.getElementById('tcodeInput');
  const v=input.value.trim();
  try{
    const res=await fetch('/api/sap/open_session', {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({tcode:v})
    });
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('새 SAP 세션 여는 중'+(v?' ('+v+')':'')+' (PID '+data.pid+')');
    input.value='';
  }catch(err){
    toast('새 세션 열기 실패: '+err.message);
  }
}
async function openPortalLoginManual(){
  try{
    const res=await fetch('/api/portal/login_manual', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('Chrome 열림 (PID '+data.pid+') - 직접 로그인하세요. 이 창을 켜둔 채로 계속 재사용하면 됨');
  }catch(err){
    toast('Portal 로그인 창 열기 실패: '+err.message);
  }
}
// -- Bloomberg Portal: real automation, not a stub - runs
// portal_ship_and_print.py per distinct order among the checked rows, ONE
// AT A TIME (waits for each to fully finish via pollOrderDone() before
// starting the next), and stops on the first failure/timeout rather than
// plowing through the rest - if one order's run failed, something is likely
// systemically wrong (not logged into Portal, Chrome not running, a bad
// scrape) and the same failure would probably just repeat. Confirms the
// exact order list up front, same safety habit launcher.py's own
// messagebox.askyesno already has for this same action. --
async function pollOrderDone(orderId, timeoutMs){
  const start=Date.now();
  while(Date.now()-start < timeoutMs){
    await new Promise(r=>setTimeout(r, 3000));
    const res=await fetch('/api/orders/'+orderId);
    const data=await res.json();
    if(data.status && data.status!=='working'){
      // events[0] is the most recent event (newest-first) - on failure this
      // is run_portal()'s own add_event() call, whose first line is always
      // a clean one-sentence summary of what actually went wrong (see
      // run_portal's watcher()), not just a bare exit code.
      const detail=(data.events && data.events[0] && data.events[0].message || '').split('\n')[0];
      return {status: data.status, detail};
    }
  }
  return {status: 'timeout', detail: ''};
}
async function runPortalBatch(){
  const checkedKeys=[...document.querySelectorAll('.row-check:checked')].map(cb=>cb.dataset.rowKey);
  if(!checkedKeys.length){toast('체크된 행이 없습니다');return}
  const checkedRows=ROWS.filter(r=>checkedKeys.includes(r.rowKey));
  const orderIds=[...new Set(checkedRows.map(r=>r.orderId))];
  const labelFor_=id=>{ const r=checkedRows.find(r=>r.orderId===id); return r ? r.orderLabel : ('오더#'+id); };
  const summary=orderIds.map(labelFor_).join(', ');
  if(!window.confirm(
    `다음 ${orderIds.length}개 오더에 대해 "Serial 등록 & QR 인쇄"를 순서대로 실행합니다:\n\n${summary}\n\n진행할까요?`
  )) return;

  for(const orderId of orderIds){
    const label=labelFor_(orderId);
    let data;
    try{
      const res=await fetch('/api/orders/'+orderId+'/run_portal', {method:'POST'});
      data=await res.json();
      if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    }catch(err){
      toast(label+' 실행 실패: '+err.message+' - 나머지 오더는 중단합니다');
      return;
    }
    if(!data.started){
      toast(label+': '+(data.message||'건너뜀'));
      continue;
    }
    toast(label+' 진행 중 (PID '+data.pid+') - 완료될 때까지 대기...');
    const result=await pollOrderDone(orderId, 360000);
    if(result.status==='printed'){
      toast(label+' 완료');
      autoPrintPod(orderId);
    } else if(result.status==='timeout'){
      toast(label+' 시간 초과 (6분) - 상태를 직접 확인하세요. 나머지 오더는 중단합니다');
      return;
    } else {
      toast(label+' 실패: '+(result.detail||('상태: '+result.status))+' - 나머지 오더는 중단합니다');
      return;
    }
  }
  toast('선택한 오더 처리 완료');
}
async function runLatestZpl(){
  try{
    const res=await fetch('/api/portal/print_latest_zpl', {method:'POST'});
    const data=await res.json();
    if(!res.ok || data.error) throw new Error(data.error||'실행 실패');
    toast('ZPL 인쇄 시작됨 (PID '+data.pid+') - Downloads 폴더 최신 라벨을 프린터로 전송, 완료되면 윈도우 알림으로 결과 확인');
  }catch(err){
    toast('ZPL 인쇄 실패: '+err.message);
  }
}

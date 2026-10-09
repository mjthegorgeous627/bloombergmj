const WEEKDAY_KR = ['월','화','수','목','금','토','일'];
function formatDateLabel(iso){
  const d=new Date(iso+'T00:00:00');
  const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  const todayIso=new Date().toLocaleDateString('sv-SE');
  const t=new Date(); t.setDate(t.getDate()+1);
  const tomorrowIso=t.toLocaleDateString('sv-SE');
  let suffix='';
  if(iso===todayIso) suffix=' (Today)';
  else if(iso===tomorrowIso) suffix=' (Tomorrow)';
  return `${y}. ${m}. ${day} ${WEEKDAY_KR[(d.getDay()+6)%7]}요일${suffix}`;
}
function labelFor(iso){
  if(!DATE_LABELS[iso]) DATE_LABELS[iso]=formatDateLabel(iso);
  return DATE_LABELS[iso];
}
function allDatesSorted(){
  // Only dates that currently have rows get a section - no forced "today"
  // anchor, so finishing today's date (e.g. via "마감") makes its banner
  // disappear entirely instead of sticking around empty. DATE_LABELS is just
  // a label-text cache and intentionally excluded here.
  const set=new Set();
  ROWS.forEach(r=>set.add(r.date));
  return [...set].sort();
}

// -- Same-order merge: rows sharing the same order label (SAP order type +
// number) are grouped into one visual block (rowspan on the shared
// columns) - this is what lets two separate order records for the same
// real-world order (e.g. a delivery row and a pickup row imported from
// Excel under the same order number, like ZRX 67066585's delivery/pickup
// pair) collapse into a single block, even when one of those records is
// missing customer/phone/address/memo (Excel only carries those on the
// delivery leg; the pickup leg's row inherits them via the first-non-blank
// fallback in sharedTd() below rather than requiring an exact match). --
function mergeKey(row){
  return (row.orderLabel||'').trim();
}
function groupRows(rows){
  const groups=[]; const indexByKey=new Map();
  rows.forEach(row=>{
    const key=mergeKey(row);
    if(indexByKey.has(key)) groups[indexByKey.get(key)].rows.push(row);
    else { indexByKey.set(key, groups.length); groups.push({key, rows:[row]}); }
  });
  return groups;
}
// first non-blank value for a shared field across a merged group - same
// backfill rule the display uses (see sharedTd in buildGroupRows), reused
// by the Excel export so a blank pickup-leg row exports with the same
// customer/phone/address/memo the board shows for it, not a blank cell.
function resolvedField(group, field){
  return group.rows.map(r=>r[field]).find(v=>v && String(v).trim()) || '';
}

// -- Cross-order merge ("Excel merge cells" for customer/phone/address/memo
// across DIFFERENT order numbers that are really the same customer) - manual,
// via the "고객정보 병합" button (mergeChecked()/unmergeChecked() further
// down). Each of the 4 fields merges INDEPENDENTLY (its own group id, its
// own span) rather than one all-or-nothing block, because a real case can
// share e.g. just the address while the customer name on file differs.
// Persisted as orders.merge_group_<field>; a field only actually renders as
// one rowspan block when the merged orders' groups also end up adjacent on
// the board (same constraint real Excel has - you can only merge a
// contiguous range), which is what the drag-to-reorder-within-a-date
// feature is for. --
const MERGE_FIELDS=['customer','phone','address','memo'];
function fieldMergeKeyOf(group, field){
  // Checks EVERY row in the group, not just rows[0] - a "group" here can
  // already be a concatenation of multiple different underlying orders that
  // happen to render under the same order-number label (groupRows() collapses
  // same-labeled orders together regardless of which one is actually merged,
  // same fallback logic resolvedField() already uses for customer/phone/etc).
  // If only rows[0] were checked, a group whose FIRST underlying order isn't
  // merged would report "not merged" even when a later order folded into the
  // same block genuinely is - exactly what happened with a stale duplicate
  // "ZOR 7825897" order sitting before the real merged one.
  const row=group.rows.find(r=>r.mergeGroups && r.mergeGroups[field]);
  return row ? String(row.mergeGroups[field]).trim() : '';
}
function computeFieldSpans(groups, field){
  const spans=new Array(groups.length);
  let i=0;
  while(i<groups.length){
    const key=fieldMergeKeyOf(groups[i], field);
    let j=i+1;
    let total=groups[i].rows.length;
    const mergedRows=[...groups[i].rows];
    if(key){
      while(j<groups.length && fieldMergeKeyOf(groups[j], field)===key){
        total+=groups[j].rows.length;
        mergedRows.push(...groups[j].rows);
        j++;
      }
    }
    spans[i]={isFirst:true, span:total, mergedRows};
    for(let k=i+1;k<j;k++) spans[k]={isFirst:false, span:0, mergedRows:[]};
    i=j;
  }
  return spans;
}
function computeAllFieldSpans(groups){
  const out={};
  MERGE_FIELDS.forEach(field=>{ out[field]=computeFieldSpans(groups, field); });
  return out;
}

function emptyRowTr(){
  const td=document.createElement('td');
  td.colSpan=11;
  td.textContent='여기로 행을 드래그하거나, 체크 후 상단 Excel 칸의 날짜로 이동할 수 있습니다 (화면상으로만 - 저장 연결은 다음 단계)';
  const tr=document.createElement('tr');
  tr.className='empty-row';
  tr.appendChild(td);
  return tr;
}

function makeEditableTd(value, className, onCommit, onPersist){
  const td=document.createElement('td');
  if(className) td.className=className;
  td.contentEditable='true';
  td.spellcheck=false;
  td.textContent=value||'';
  td.addEventListener('blur', ()=>{
    const v=td.textContent.trim();
    onCommit(v);
    renderAll();
    // onPersist is optional (still-screen-only fields like orderLabel/
    // deliveryNo don't pass one yet) - fires after the optimistic on-screen
    // update so typing stays instant; a failure just warns instead of
    // silently leaving the DB out of sync with what's shown.
    if(onPersist) onPersist(v).catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));
  });
  td.addEventListener('keydown', e=>{
    if(e.key!=='Enter') return;
    e.preventDefault();
    if(e.altKey){
      // Alt+Enter = line break within the cell, same shortcut Excel uses -
      // insertText keeps it a plain '\n' character (matches how
      // deliveryNo/address already store multi-line text), not a nested
      // <div>/<br>, so td.textContent stays a simple string.
      document.execCommand('insertText', false, '\n');
      return;
    }
    td.blur();
  });
  return td;
}

// -- Manual cell highlight (whole-cell only, not per-character - see
// project memory round 15): right-click any highlightable <td> for a small
// menu (배경 노랑/핑크, 글자 빨강, Bold). Persists to workbench.db
// (row.highlights, keyed by {itemId, field}) and also carries into the
// Excel export (buildExportRows/_write_export_sheet) - deliberately does
// NOT show up in the 배송장 인쇄 output (print stays plain black-on-white
// per its own spec). --
const HL_BG_COLOR={yellow:'#fde047', pink:'#f9a8d4'};
const HL_TEXT_COLOR={red:'#dc2626'};
function applyHighlight(td, itemId, field, row){
  td.dataset.hlItem=itemId;
  td.dataset.hlField=field;
  const style=(row.highlights && row.highlights[field]) || null;
  td.style.background = style && HL_BG_COLOR[style.bg] ? HL_BG_COLOR[style.bg] : '';
  td.style.color = style && HL_TEXT_COLOR[style.color] ? HL_TEXT_COLOR[style.color] : '';
  td.style.fontWeight = style && style.bold ? '700' : '';
}
async function setHighlight(itemId, field, style){
  const res=await fetch('/api/items/highlight', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body:JSON.stringify({itemId, field, style})
  });
  const data=await res.json();
  if(!res.ok || data.error){ toast('강조 표시 저장 실패: '+(data.error||'')); return; }
  const row=ROWS.find(r=>r.itemId===itemId);
  if(row){
    row.highlights=row.highlights||{};
    if(Object.keys(style).length) row.highlights[field]=style; else delete row.highlights[field];
  }
  renderAll();
}
let hlMenuEl=null;
function closeHlMenu(){
  if(hlMenuEl){ hlMenuEl.remove(); hlMenuEl=null; }
  document.removeEventListener('click', closeHlMenu);
}
function openHlMenu(x, y, itemId, field, current){
  closeHlMenu();
  const menu=document.createElement('div');
  menu.className='hl-menu';
  menu.style.left=x+'px'; menu.style.top=y+'px';
  const mk=(label, cls, onClick)=>{
    const b=document.createElement('button');
    b.type='button'; b.className='hl-menu-btn'+(cls?' '+cls:''); b.textContent=label;
    b.addEventListener('click', e=>{ e.stopPropagation(); onClick(); closeHlMenu(); });
    menu.appendChild(b);
  };
  const apply=patch=>{
    const next={...current, ...patch};
    Object.keys(next).forEach(k=>{ if(!next[k]) delete next[k]; });
    setHighlight(itemId, field, next);
  };
  mk('배경 없음', '', ()=>apply({bg:undefined}));
  mk('배경 노랑', 'hl-swatch-yellow', ()=>apply({bg:'yellow'}));
  mk('배경 핑크', 'hl-swatch-pink', ()=>apply({bg:'pink'}));
  mk('글자 기본', '', ()=>apply({color:undefined}));
  mk('글자 빨강', 'hl-swatch-red', ()=>apply({color:'red'}));
  mk(current.bold ? 'Bold 끄기' : 'Bold 켜기', '', ()=>apply({bold:!current.bold}));
  mk('전체 지우기', '', ()=>setHighlight(itemId, field, {}));
  document.body.appendChild(menu);
  hlMenuEl=menu;
  setTimeout(()=>document.addEventListener('click', closeHlMenu), 0);
}
document.getElementById('boardTable').addEventListener('contextmenu', e=>{
  const td=e.target.closest('[data-hl-field]');
  if(!td) return;
  e.preventDefault();
  const itemId=parseInt(td.dataset.hlItem, 10);
  const field=td.dataset.hlField;
  const row=ROWS.find(r=>r.itemId===itemId);
  const current=(row && row.highlights && row.highlights[field]) || {};
  openHlMenu(e.clientX, e.clientY, itemId, field, current);
});

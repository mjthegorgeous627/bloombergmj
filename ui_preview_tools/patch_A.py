"""실제 workbench 템플릿을 복사해 와서(원본은 건드리지 않음) 추가 기능을 붙인 작업용 사본 A_mod/를 만든다."""
import os, re, shutil
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = r'C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\templates'
DST = os.path.join(HERE, 'A_mod', 'templates')
if os.path.exists(os.path.join(HERE, 'A_mod')):
    shutil.rmtree(os.path.join(HERE, 'A_mod'))
shutil.copytree(SRC, DST)
shutil.copy(r'C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\dashboard_assets.py', os.path.join(HERE, 'A_mod'))

def rd(p): return open(os.path.join(DST, p), encoding='utf-8').read()
def wr(p, s): open(os.path.join(DST, p), 'w', encoding='utf-8', newline='').write(s)
def sub1(s, old, new, label):
    assert s.count(old) == 1, (label, s.count(old))
    return s.replace(old, new, 1)
def rsub1(s, pat, new, label):
    assert len(re.findall(pat, s, re.S)) == 1, (label, len(re.findall(pat, s, re.S)))
    return re.sub(pat, lambda m: new, s, count=1, flags=re.S)

# ---- 03_zrec.js : 필터 상태 (첫 renderAll 이전에 있어야 함) ----
s = rd('dashboard/03_zrec.js')
s = sub1(s, 'let ROWS = [];', 'let ROWS = [];\nlet BOARD_FILTER = \'all\';  // 메모 라벨 모아보기: all | delayed | check', 'ROWS')
wr('dashboard/03_zrec.js', s)

# ---- 04_board_model.js : 편집 칸 안의 비편집 요소(라벨/아이콘)는 저장 값에서 제외 ----
s = rd('dashboard/04_board_model.js')
s = sub1(s, "    const v=td.textContent.trim();\n    onCommit(v);", "    const v=editableText(td);\n    onCommit(v);", 'blur')
wr('dashboard/04_board_model.js', s)

# ---- 05_render_rows.js ----
s = rd('dashboard/05_render_rows.js')
# (a) 오더번호 수정 허용(확인창) + 아이콘
s = sub1(s, "      let newType, newNo;", "      let newType, newNo;\n        let orderNoChanged='';", 'newType')
s = rsub1(s, r"          if\(newNo && newNo!==realNo\)\{\s*toast\('오더번호는 여기서 수정할 수 없습니다 - 타입만 저장됩니다'\);\s*\}",
"""          if(newNo && newNo!==realNo){
            if(window.confirm('오더번호를 '+realNo+' → '+newNo+' 로 바꿉니다.\\n\\nSAP 수집·POD·Portal 자동화는 이 번호를 기준으로 동작하므로 실제 SAP 오더번호와 다르면 오류가 날 수 있습니다. 계속할까요?')){
              orderNoChanged=newNo;
            } else {
              toast('오더번호 수정을 취소했습니다 - 타입만 저장됩니다');
            }
          }""", 'orderno-toast')
s = sub1(s, "        const savedLabel=(newType+' '+realNo).trim();\n        group.rows.forEach(r=>r.orderLabel=savedLabel);",
         "        const savedLabel=(newType+' '+(orderNoChanged||realNo)).trim();\n        group.rows.forEach(r=>{r.orderLabel=savedLabel; if(orderNoChanged) r.orderNo=orderNoChanged;});", 'savedLabel')
s = sub1(s, "        persistOrderField(orderItemIds, 'order_type', newType)\n          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));",
         "        persistOrderField(orderItemIds, 'order_type', newType)\n          .catch(err=>toast('저장 실패 (DB 반영 안 됨): '+err.message));\n        if(orderNoChanged){\n          persistOrderNo(orderItemIds, orderNoChanged)\n            .catch(err=>toast('오더번호 저장 실패 (DB 반영 안 됨): '+err.message));\n        }", 'persist-type')
s = sub1(s, "      ordTd.appendChild(main);",
         "      const ordLine=document.createElement('div');\n      ordLine.className='ord-line';\n      ordLine.appendChild(main);\n      ordLine.appendChild(makeOrderIcons(group.rows[0].orderNo, row.orderLabel));\n      ordTd.appendChild(ordLine);", 'ordLine')
s = sub1(s, "      ordTd.appendChild(extra);",
         "      const codeWrap=document.createElement('div');\n      codeWrap.className='code-wrap';\n      codeWrap.appendChild(extra);\n      if(codeLines.length) codeWrap.appendChild(makeCodeIcons(codeLines));\n      ordTd.appendChild(codeWrap);", 'codeWrap')
# (b) 메모 칸 라벨
s = sub1(s, "          td.rowSpan=fs.span;\n          applyHighlight(td, row.itemId, field, row);\n          tr.appendChild(td);",
         "          td.rowSpan=fs.span;\n          applyHighlight(td, row.itemId, field, row);\n          if(field==='memo') decorateMemoTd(td, fs.mergedRows);\n          tr.appendChild(td);", 'memo-decorate')
# (c) renderAll: 필터 + 갱신 훅
s = sub1(s, "    const rows=ROWS.filter(r=>r.date===iso);\n    if(!rows.length){",
         "    const rows=ROWS.filter(r=>r.date===iso && flagMatch(r));\n    if(BOARD_FILTER!=='all' && !rows.length) return;  // 모아보기 중에는 해당 없는 날짜는 숨김\n    if(!rows.length){", 'rows-filter')
s = rsub1(s, r"    table\.appendChild\(tbody\);\n  \}\);\n\}\nrenderAll\(\);",
          "    table.appendChild(tbody);\n  });\n  if(typeof updateSideViews==='function') updateSideViews();\n}\nrenderAll();", 'renderAll-end')
wr('dashboard/05_render_rows.js', s)

# ---- dashboard.html ----
s = rd('dashboard.html')
exec(open(os.path.join(HERE, 'patch_layout.py'), encoding='utf-8').read())
side = '''<aside id="sideCol" class="side-col">
  <div class="side-head"><b>달력 · 메모</b><button type="button" id="sideToggle" class="side-toggle" onclick="toggleSideCol()" title="오른쪽 열 접기/펼치기">▸</button></div>
  <div class="side-body">
    <div class="side-sec"><div id="sideCal"></div></div>
    <div class="side-sec">
      <h4>메모 <i id="sideMemoCount"></i></h4>
      <textarea id="sideMemoNew" class="sm-new" placeholder="메모 (Enter 저장, Ctrl+M)"></textarea>
      <input id="sideMemoQ" class="sm-q" placeholder="검색">
      <div id="sideMemoList"></div>
    </div>
  </div>
  <div class="side-rail">
    <button type="button" onclick="toggleSideCol()" title="달력 열기">달력</button>
    <button type="button" onclick="openSideMemo()" title="빠른 메모 (Ctrl+M)">메모</button>
  </div>
</aside>

<div id="printArea"></div>'''
s = sub1(s, '<div id="printArea"></div>', SIDE_HTML + '\n\n<div id="printArea"></div>', 'side')
s = sub1(s, '//#include dashboard/08_finalize_init.js', '//#include dashboard/08_finalize_init.js\n//#include dashboard/09_side_panel.js', 'include09')
# 2026-10-14: 상단 📅(휴일 달력) 버튼 제거 - 휴일은 오른쪽 새 달력에 표시. calToggleDay가 #calPanel을 쓰므로 숨김 상태로 남김
s, n = re.subn(r'<div id="calWrap"><button type="button" class="cal-toggle-btn".*?</button><div id="calPanel"></div></div>', '<div id="calWrap" style="display:none"><div id="calPanel"></div></div>', s, flags=re.S)
assert n == 1, 'calWrap'
# ver.2 표시: 상단에 "v2 테스트" 배지 + v1(원래 화면)로 돌아가는 링크 (긴급 복귀용)
s, n = re.subn(r'(<div class="topbar-right">)', r'<a href="/" title="원래 화면(v1)으로 돌아갑니다 - 데이터는 같습니다" style="color:#fde68a;text-decoration:none;font-size:11.5px;border:1px solid #b45309;background:#78350f;border-radius:5px;padding:3px 8px;font-weight:700">v2 테스트 · v1로 돌아가기</a>', s, count=1)
assert n == 1, 'v2badge'
wr('dashboard.html', s)

# ---- style.css / 09 ----
s = rd('dashboard/style.css')
s += open(os.path.join(HERE, 'A_extra.css'), encoding='utf-8').read() + open(os.path.join(HERE, 'quick.css'), encoding='utf-8').read()
wr('dashboard/style.css', s)
wr('dashboard/09_side_panel.js', open(os.path.join(HERE, 'A_09_side_panel.js'), encoding='utf-8').read() + open(os.path.join(HERE, 'A_quick.js'), encoding='utf-8').read())
# 실제 서버용 사본: sap_automation\templates_v2 (PREVIEW_ONLY 환경에서는 건너뜀)
V2 = os.environ.get('V2_DIR')
if V2:
    if os.path.exists(V2): shutil.rmtree(V2)
    os.makedirs(V2)
    shutil.copy(os.path.join(DST, 'dashboard.html'), V2)
    shutil.copytree(os.path.join(DST, 'dashboard'), os.path.join(V2, 'dashboard'))
    print('v2 exported', V2)
print('patched ok')

# patch_A.py에서 exec 됨: s = dashboard.html 텍스트. 기존 상단 4칸(.controls)을 오른쪽 열/모아보기 줄로 재배치한다.
# 버튼은 원본 HTML에서 onclick 기준으로 그대로 가져온다 (이름/툴팁/동작 변경 없음).
ORIG = s

def B(token, label=None):
    m = re.findall(r'<button class="btn[^"]*" onclick="' + re.escape(token) + r'[^>]*>.*?</button>', ORIG, re.S)
    assert len(m) >= 1, token
    h = m[0]
    if label:
        h = re.sub(r'>[^<]*</button>$', '>' + label + '</button>', h)
    return h

def INP(id_, ph, fn):
    return ('<input id="%s" placeholder="%s" onkeydown="if(event.key===\'Enter\'){event.preventDefault();%s();}">' % (id_, ph, fn))

# ---- 1) 상단 .controls 를 숨김 보관함으로 교체 + 모아보기 줄(sticky 안) ----
st = s.index('<div class="controls">')
tail = '  </div>\n</div>\n</div>\n'
en = s.index(tail, st)
assert s.count(tail) == 1
old_block = s[st:en + len('  </div>\n</div>\n')]

HIDDEN = '''<div class="controls" id="hiddenCtl" style="display:none">
  <div id="mergeFieldHome"><div class="merge-field-row">
    <label><input type="checkbox" class="merge-field-cb" value="order">Order#</label>
    <label><input type="checkbox" class="merge-field-cb" value="customer" checked>Customer</label>
    <label><input type="checkbox" class="merge-field-cb" value="phone" checked>Phone</label>
    <label><input type="checkbox" class="merge-field-cb" value="address" checked>Address</label>
    <label><input type="checkbox" class="merge-field-cb" value="memo" checked>Memo</label>
  </div></div>
  <input type="date" id="moveDateInput" onchange="moveCheckedToDate(this.value); this.value='';">
  <div class="hidden-tools">
    ''' + B('runMorningRoutine()') + B('runSapSession(0)') + B('runSapSession(1)') + B('runSapSession(2)') + B('runSapSession(3)') + '''
    <input type="text" id="excelRangeSheet" placeholder="시트명 (예: 8-6)" size="10">
    <input type="number" id="excelRangeStart" placeholder="시작행" min="1" style="width:5em">
    <input type="number" id="excelRangeEnd" placeholder="종료행" min="1" style="width:5em">
    ''' + B('importExcelRange()') + '''
  </div>
</div>
<div class="flag-filter-bar" id="flagFilterBar">
  <span id="ffChips" class="ff-chips"></span>
  <span class="ff-sep"></span>
  <span class="ff-act">''' + B('processDone()') + B('processPickupReady()', 'Ready') + B('processCancel()') + B('processReset()') + B('processBilled()') + '''</span>
  <span class="ff-hist"><input id="histQ" type="search" placeholder="이력 검색" autocomplete="off" spellcheck="false" title="오더번호 / 시리얼 / 고객 등으로 예전 배송 이력 조회"></span>
  <span class="ff-spacer"></span>
  <button type="button" class="ff-ico" onclick="undoLastAction()" title="실행취소 (Ctrl+Z) - 삭제/날짜이동/완료처리를 최근 20건까지 되돌림">↶</button>
  <button type="button" class="ff-ico" onclick="openTrash()" title="삭제된 항목 보기 (휴지통) - 7일 안이면 복구 가능">🗑</button>
</div>
'''
s = s.replace(old_block, HIDDEN, 1)

# ---- 2) 오른쪽 열: 달력 · 메모 · SAP · ZREC · (가림) · Portal · Excel ----
SIDE_HTML = '''<aside id="sideCol" class="side-col">
  <div class="side-head"><b>달력 · 메모 · 도구</b><button type="button" id="sideToggle" class="side-toggle" onclick="toggleSideCol()" title="오른쪽 열 접기/펼치기">▸</button></div>
  <div class="side-body">
    <div class="side-sec"><div id="sideCal"></div></div>
    <div class="side-sec">
      <h4><span>메모 <i id="sideMemoCount"></i></span><button type="button" class="pin-open" onclick="togglePinPanel()" title="고정 메모 (탭별 기록)">📌</button></h4>
      <textarea id="sideMemoNew" class="sm-new" placeholder="메모 (Enter 저장, Ctrl+M)"></textarea>
      <input id="sideMemoQ" class="sm-q" placeholder="검색">
      <div id="sideMemoList"></div>
    </div>
    <div class="side-sec tools">
      <h4>SAP</h4>
      <div class="srow">''' + B('runSap(\'all\')', '전체조회') + B('runSap(\'stop\')') + B('runSap(\'start\')', 'Start&amp;Loop') + '''</div>
      <div class="srow">''' + B('runZrecPrepare()', 'ZREC 처리') + B('runSnReloBatch()', 'S/N 오더 반영 시도') + '''</div>
      <details class="side-det"><summary>SAP 도구 (오더 수집 · 오더 조회 · SN 조회 · SAP Tab+)</summary>
        <div class="ctrl-inline">''' + B('runOrderNoAction()', '오더 수집') + INP('orderNoInput', '오더번호', 'runOrderNoAction') + '''</div>
        <div class="ctrl-inline">''' + B('runOrderJumpAction()', '오더 조회') + INP('orderJumpInput', '오더번호', 'runOrderJumpAction') + '''</div>
        <div class="ctrl-inline">''' + B('runSerialLookupAction()', 'SN 조회') + INP('serialLookupInput', 'Serial Number', 'runSerialLookupAction') + '''</div>
        <div class="ctrl-inline">''' + B('runNewSessionAction()', 'SAP Tab+') + INP('tcodeInput', 'T-code (선택)', 'runNewSessionAction') + '''</div>
      </details>
    </div>
    <div class="side-sec tools">
      <h4>Bloomberg Portal</h4>
      <div class="srow">''' + B('openPortalLoginManual()') + B('runLatestZpl()', '최근 ZPL 수동 인쇄') + '''</div>
      <div class="srow">''' + B('runPortalBatch()', 'Serial 등록 &amp; QR 인쇄') + B('runPodBatch()', 'POD (배송완료) 처리') + B('runSpareOneshot()', '스페어 배송처리 (serial+POD처리)') + '''</div>
    </div>
    <div class="side-sec tools">
      <h4>Excel</h4>
      <div class="srow">''' + B('finalizeChecked()') + B('printChecked()') + '''</div>
      <details class="side-det"><summary>월초</summary>
        <div class="srow">''' + B('runQuickSettlement()') + '''</div>
      </details>
    </div>
  </div>
  <div class="side-rail">
    <button type="button" onclick="toggleSideCol()" title="달력 열기">달력</button>
    <button type="button" onclick="openSideMemo()" title="빠른 메모 (Ctrl+M)">메모</button>
  </div>
</aside>'''

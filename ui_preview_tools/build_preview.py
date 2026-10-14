"""patch_A.py가 만든 A_mod를 실제 DB 사본의 데이터로 렌더링한 단일 HTML(preview_A.html)을 만든다.
실제 DB/실제 서버는 건드리지 않는다 (읽기 전용 백업 사본 + fetch 가짜 응답)."""
import os, sys, json, sqlite3, pathlib, re
HERE = os.path.dirname(os.path.abspath(__file__))
SAP = r'C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation'
copy = os.path.join(HERE, 'wb_copy.db')
if os.path.exists(copy):
    os.remove(copy)
src = sqlite3.connect('file:' + SAP.replace('\\', '/') + '/workbench.db?mode=ro', uri=True)
dst = sqlite3.connect(copy)
src.backup(dst)
dst.close(); src.close()

sys.path.insert(0, SAP)
import db
db.DB_PATH = pathlib.Path(copy)
import items
board = json.dumps(items.board_data(), ensure_ascii=False)
custom = json.dumps(items.list_custom_types(), ensure_ascii=False)
nonwork = json.dumps(items.nonworking_days_payload(), ensure_ascii=False)

sys.path.insert(0, os.path.join(HERE, 'A_mod'))
import importlib.util
spec = importlib.util.spec_from_file_location('dashboard_assets_mod', os.path.join(HERE, 'A_mod', 'dashboard_assets.py'))
da = importlib.util.module_from_spec(spec); spec.loader.exec_module(da)
page = da.assemble('dashboard.html', base=pathlib.Path(HERE) / 'A_mod' / 'templates')
page = (page.replace('__BOARD_DATA__', board).replace('__CUSTOM_TYPES__', custom)
        .replace('__NONWORKING_DAYS__', nonwork).replace('__DELETED_ITEM_RETENTION_DAYS__', '7'))

import history_search
_hc = history_search._connect()
_hr = _hc.execute("SELECT r.id,r.sheet_id,r.group_no,r.excel_row,r.kind,r.order_no,r.cells_json,s.sheet_date,s.sheet_name,s.source_file FROM rows r JOIN sheets s ON s.id=r.sheet_id ORDER BY s.sheet_date DESC, r.sheet_id DESC, r.excel_row LIMIT 4000").fetchall()
HIST_DATA = json.dumps([{'rowId':r['id'],'sheetId':r['sheet_id'],'groupNo':r['group_no'],'excelRow':r['excel_row'],'kind':r['kind'] or '','orderNo':r['order_no'] or '','cells':json.loads(r['cells_json']),'sheetDate':r['sheet_date'] or '','sheetName':r['sheet_name'],'sourceFile':os.path.basename(r['source_file'])} for r in _hr], ensure_ascii=False)
_hc.close()
PAGES_DATA = json.dumps(items.list_pages(), ensure_ascii=False)

STUB = r'''<script>
/* 미리보기 전용: 서버 호출은 전부 가짜 응답 (실제 DB/SAP/터미널에 아무것도 보내지 않음) */
(function(){
  var memos=[],nid=1;
  var HIST=__HIST__, PAGES=__PAGES__, pid=1000;
  function ok(o){return Promise.resolve(new Response(JSON.stringify(o||{}),{status:200,headers:{'Content-Type':'application/json'}}));}
  window.fetch=function(url,opts){
    var body={};try{body=JSON.parse((opts&&opts.body)||'{}')}catch(e){}
    if(url==='/api/memos') return ok({memos:memos});
    if(url==='/api/memos/add'){var m={id:nid++,body:body.body,pinned:0,done:0,created_at:new Date().toISOString().slice(0,10)+' '+new Date().toTimeString().slice(0,5)};memos.push(m);return ok({id:m.id});}
    if(url==='/api/memos/update'){memos.forEach(function(m){if(m.id===body.id){['body','pinned','done'].forEach(function(k){if(k in body)m[k]=body[k]})}});return ok({});}
    if(url==='/api/memos/delete'){memos=memos.filter(function(m){return m.id!==body.id});return ok({});}
    if(url.indexOf('/api/history/search')===0){var q=decodeURIComponent(url.split('q=')[1]||'').toLowerCase().split(/\s+/).filter(Boolean);var hit=HIST.filter(function(h){var t=(h.kind+' '+h.orderNo+' '+h.cells.join(' ')).toLowerCase();return q.every(function(w){return t.indexOf(w)>=0})});return ok({results:hit.slice(0,200),total:hit.length});}
    if(url.indexOf('/api/history/group')===0){var u=new URLSearchParams(url.split('?')[1]);var sid=+u.get('sheet_id'),gn=+u.get('group_no');var rs=HIST.filter(function(h){return h.sheetId===sid&&h.groupNo===gn});return ok(rs.length?{sheetDate:rs[0].sheetDate,sheetName:rs[0].sheetName,sourceFile:rs[0].sourceFile,rows:rs}:{error:'not found'});}
    if(url.indexOf('/api/quick/suggest')===0){var cq=decodeURIComponent(url.split('company=')[1]||'').toLowerCase();var seenq={},outq=[];HIST.forEach(function(h){var a=(h.cells[6]||'');if(outq.length>=3||a.toLowerCase().indexOf(cq)<0)return;var ls=a.split('\n').filter(function(x){return /[가-힣]/.test(x)&&/\d/.test(x)&&x.indexOf('cust#')<0&&x.indexOf('연락')<0});if(!ls.length)return;var t=ls.join(' ');if(seenq[t])return;seenq[t]=1;outq.push({text:t,date:h.sheetDate})});return ok({suggestions:outq});}
    if(url==='/api/pages') return ok({pages:PAGES});
    if(url==='/api/pages/add'){var np={id:pid++,title:body.title||'새 메모장',body:''};PAGES.push(np);return ok({added:true,id:np.id,title:np.title});}
    if(url==='/api/pages/edit'){PAGES.forEach(function(p){if(p.id===body.id)p[body.field]=body.value});return ok({updated:true});}
    if(url==='/api/pages/delete'){PAGES=PAGES.filter(function(p){return p.id!==body.id});return ok({deleted:true});}
    if(url==='/api/forms/print') return ok({started:false,message:'미리보기 화면에서는 실제로 인쇄되지 않습니다 - 실제 workbench 주소(http://127.0.0.1:8765/v2)에서 사용하세요'});
    if(url==='/api/sap/open_order'||url==='/api/sap/open_obd'||url==='/api/terminal/open_ticket') return ok({started:true,message:'미리보기: 실제로는 열리지 않습니다'});
    return ok({});
  };
  window.__PREVIEW__=true;
  window.__QUICK_LOGO__='__LOGO__';
})();
</script>'''

# artifact wants no <html>/<head>/<body> wrapper: keep <style> blocks + body content
style = ''.join(re.findall(r'<style>.*?</style>', page, re.S))
body = re.search(r'<body[^>]*>(.*)</body>', page, re.S).group(1)
import base64
LOGO = 'data:image/png;base64,' + base64.b64encode(open(r'C:\Users\bloomberg\Documents\MJSuh\mjbg\sap_automation\assets\bloomberg_logo.png','rb').read()).decode()
STUB = STUB.replace('__LOGO__', LOGO).replace('__HIST__', HIST_DATA).replace('__PAGES__', PAGES_DATA)
out = '<title>Workbench 추가기능 미리보기</title>\n' + style + '\n<style>body{background:#f1f4f9}</style>\n' + STUB + '\n' + body
open(os.path.join(HERE, 'preview_A.html'), 'w', encoding='utf-8').write(out)
js = re.findall(r'<script>(.*?)</script>', out, re.S)
open(os.path.join(HERE, 'chk_A.js'), 'w', encoding='utf-8').write('\n'.join(js))
print('ok', len(out), 'rows', sum(len(d['rows']) for d in json.loads(board)))

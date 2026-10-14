import os,sys,json
from playwright.sync_api import sync_playwright
f=sys.argv[1]; patched=sys.argv[2]=='A'
def L(*a): print(*a,flush=True)
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1700,'height':1000})
    errs=[]; pg.on('pageerror',lambda e:errs.append(str(e))); pg.on('dialog',lambda d:d.accept())
    calls=[]
    pg.goto('file:///'+os.path.abspath(f).replace(chr(92),'/')); pg.wait_for_timeout(1500)
    pg.evaluate("(()=>{window.__c=[];const _f=window.fetch;window.fetch=function(u,o){window.__c.push(u+' '+((o&&o.body)||'').slice(0,80));return _f.apply(this,arguments)};})()")
    N=lambda: pg.evaluate("document.querySelectorAll('tr.item-row').length")
    RS=lambda: pg.evaluate("document.querySelectorAll('td[rowspan]').length")
    def uncheck(): pg.evaluate("document.querySelectorAll('.row-check:checked,.merge-field-cb').forEach(c=>{if(c.classList.contains('row-check'))c.checked=false})")
    def chk(i): pg.locator('.row-check').nth(i).check()
    def act(fn,idx=None,label=None):
        if patched and label:
            pg.locator('.handle-cell').nth(idx).click(button='right'); pg.wait_for_timeout(150)
            pg.locator('.row-menu').get_by_text(label,exact=False).first.click()
        else: pg.evaluate(fn)
        pg.wait_for_timeout(600)
    L('rows',N(),'rowspans',RS())
    uncheck(); chk(1); n=N(); act('copyChecked()',1,'복사'); L('copy',n,'->',N())
    uncheck(); chk(2); n=N(); act('deleteChecked()',2,'삭제'); L('delete',n,'->',N())
    pg.evaluate("undoLastAction()"); pg.wait_for_timeout(500); L('undo ->',N())
    uncheck(); chk(3); chk(4); r=RS(); act('mergeChecked()',3,'병합'); L('merge',r,'->',RS())
    uncheck(); chk(3); r=RS(); act('unmergeChecked()',3,'병합 해제'); L('unmerge',r,'->',RS())
    uncheck(); chk(5)
    cl=lambda: pg.evaluate("document.querySelectorAll('tr.item-row')[5].className")
    for fn in ['processDone','processPickupReady','processCancel','processReset','processBilled']:
        pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)"); chk(5)
        pg.evaluate(fn+"()"); pg.wait_for_timeout(500); L(fn,'->',cl())
    # status via real button in patched
    if patched:
        uncheck(); chk(5); pg.locator('.ff-act button',has_text='Done').first.click(); pg.wait_for_timeout(400); L('Done button ->',cl())
    L('errs',errs[:4])
    b.close()

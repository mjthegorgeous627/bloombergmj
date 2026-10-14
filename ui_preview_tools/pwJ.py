import os
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1700,'height':1000})
    errs=[]; pg.on('pageerror',lambda e:errs.append(str(e))); pg.on('dialog',lambda d:d.accept())
    pg.goto('file:///'+os.path.abspath('preview_A2.html').replace(chr(92),'/')); pg.wait_for_timeout(1500)
    pg.evaluate("""(()=>{window.__s=[];['copyChecked','deleteChecked','mergeChecked','unmergeChecked','addManualRow','processDone','processPickupReady','processCancel','processReset','processBilled','undoLastAction','openTrash','finalizeChecked','printChecked','runSap','runZrecPrepare','runSnReloBatch','runPortalBatch','runPodBatch','runSpareOneshot','runLatestZpl','openPortalLoginManual','runQuickSettlement','runOrderNoAction','runOrderJumpAction','runSerialLookupAction','runNewSessionAction'].forEach(n=>{try{window[n]=function(){window.__s.push(n+':'+[].slice.call(arguments).join(','))}}catch(e){}});})()""")
    # inline onclick resolves globals at call time -> spy works for function declarations replaced on window
    pg.locator('.row-check').nth(1).check()
    for lab in ['복사','삭제','병합','병합 해제']:
        pg.locator('.handle-cell').nth(1).click(button='right'); pg.wait_for_timeout(150)
        pg.locator('.row-menu').get_by_text(lab,exact=True).first.click(); pg.wait_for_timeout(200)
    pg.locator('.add-row-btn').first.click(force=True)
    for sel in ['.ff-act .btn','.ff-ico','.side-col .btn']:
        for i in range(pg.locator(sel).count()):
            e=pg.locator(sel).nth(i)
            if not e.is_visible():
                pg.evaluate("document.querySelectorAll('.side-det').forEach(d=>d.open=true)")
            try: e.click(timeout=1500)
            except Exception as x: print('noclick',sel,i)
    print(pg.evaluate("window.__s.join(String.fromCharCode(10))"))
    print('errs',errs[:3])
    b.close()

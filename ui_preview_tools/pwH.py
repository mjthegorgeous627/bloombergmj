import os
from playwright.sync_api import sync_playwright
def log(*a): print(*a, flush=True)
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1700,'height':1000})
    errs=[]; pg.on('pageerror',lambda e:errs.append('PAGEERR '+str(e)))
    pg.on('dialog',lambda d:d.accept())
    pg.goto('file:///'+os.path.abspath('preview_A.html').replace(chr(92),'/')); pg.wait_for_timeout(1500)
    N=lambda: pg.evaluate("document.querySelectorAll('tr.item-row').length")
    log('rows',N())
    # pick first date group rows
    def chk(idx):
        pg.locator('.row-check').nth(idx).check()
    def menu(idx,label):
        pg.locator('.handle-cell').nth(idx).click(button='right'); pg.wait_for_timeout(150)
        pg.locator('.row-menu').get_by_text(label,exact=False).first.click(); pg.wait_for_timeout(500)
    n0=N(); chk(1); menu(1,'복사'); log('copy',n0,'->',N())
    n0=N(); pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)")
    # delete the copy (row index 2 probably)
    chk(2); menu(2,'삭제'); log('delete',n0,'->',N())
    pg.wait_for_timeout(300)
    log('trash open'); pg.locator('.ff-ico').nth(1).click(); pg.wait_for_timeout(400)
    log(' overlay', pg.evaluate("[...document.querySelectorAll('[id*=rash]')].map(e=>e.id+':'+getComputedStyle(e).display).join(',')")); pg.evaluate("closeTrash()")
    pg.locator('.ff-ico').nth(0).click(); pg.wait_for_timeout(500); log('undo ->',N())
    # merge
    pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)")
    chk(3); chk(4)
    rs0=pg.evaluate("[...document.querySelectorAll('td[rowspan]')].length")
    menu(3,'병합'); log('merge rowspans',rs0,'->',pg.evaluate("[...document.querySelectorAll('td[rowspan]')].length"))
    pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)"); chk(3)
    menu(3,'병합 해제'); log('unmerge rowspans ->',pg.evaluate("[...document.querySelectorAll('td[rowspan]')].length"))
    # status buttons
    pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)")
    chk(5)
    def cls(): return pg.evaluate("document.querySelectorAll('tr.item-row')[5].className")
    log('cls0',cls())
    for lab in ['Done','Ready','Cancel','초기화']:
        pg.locator('.flag-filter-bar .ff-act button',has_text=lab).first.click(); pg.wait_for_timeout(400)
        log(lab,'->',cls()); 
        chk(5) if not pg.locator('.row-check').nth(5).is_checked() else None
    # + button
    pg.locator('.add-row-btn').first.click(force=True); pg.wait_for_timeout(300)
    log('manual modal', pg.evaluate("document.getElementById('manualRowOverlay').classList.contains('show')"))
    pg.fill('#manualRowTitle','테스트') if pg.locator('#manualRowTitle').count() else None
    pg.evaluate("submitManualRow()"); pg.wait_for_timeout(500); log('after add rows',N())
    # highlight right click on cell
    pg.locator('[data-hl-field]').first.click(button='right'); pg.wait_for_timeout(200)
    log('hl menu', pg.evaluate("!!document.querySelector('.hl-menu, #hlMenu, .highlight-menu')"))
    pg.mouse.click(5,500)
    # calendar click moves checked rows
    pg.evaluate("document.querySelectorAll('.row-check:checked').forEach(c=>c.checked=false)"); chk(0)
    d0=pg.evaluate("document.querySelectorAll('tr.item-row')[0].closest('tbody').dataset.date")
    pg.locator('.sc-day:not(.non)').nth(20).click(); pg.wait_for_timeout(500)
    log('calendar move: first row date',d0,'->',pg.evaluate("[...document.querySelectorAll('tbody.date-group')].map(t=>t.dataset.date).join(',')")[:200])
    # filter chips
    log('chips', pg.evaluate("[...document.querySelectorAll('#ffChips button')].map(b=>b.textContent.trim()).join('|')"))
    # collapse date
    pg.locator('.date-banner-row').first.click(); pg.wait_for_timeout(200)
    log('collapse rows',N()); pg.locator('.date-banner-row').first.click()
    # ctrl+z
    pg.keyboard.press('Control+z'); pg.wait_for_timeout(300)
    # type dropdown
    log('type options', pg.evaluate("[...document.querySelectorAll('select')].filter(s=>[...s.options].some(o=>o.text==='회수')).slice(0,1).map(s=>[...s.options].map(o=>o.text).join('/'))"))
    log('errs',errs[:6])
    b.close()

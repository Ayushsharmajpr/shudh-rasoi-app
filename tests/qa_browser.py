import json
from playwright.sync_api import sync_playwright
import os
B=os.environ.get('QA_URL','http://localhost:5077')   # start the app first: TIFFIN_DB=$(copy of tiffin.db) TIFFIN_PORT=5077 python3 app.py
errs=[]; fails=[]
def ok(n,c,x=''):
    print(('PASS' if c else 'FAIL'),n,x); 
    if not c: fails.append(n)
with sync_playwright() as p:
    br=p.chromium.launch(args=['--no-sandbox'])
    ctx=br.new_context(viewport={'width':1360,'height':860}); pg=ctx.new_page()
    pg.on('console',lambda m: errs.append((pg.url,m.text)) if m.type in('error','warning') else None)
    pg.on('pageerror',lambda e: errs.append((pg.url,'PAGEERROR '+str(e))))
    pg.on('requestfailed',lambda r: errs.append((pg.url,'REQFAIL '+r.url)))
    pg.on('response',lambda r: errs.append((pg.url,f'HTTP {r.status} {r.url}')) if r.status>=400 and r.status!=401 else None)
    pg.goto(B+'/'); ok('redirect to login', '/login' in pg.url)
    pg.screenshot(path='shots/login.png')
    pg.fill('#username','admin'); pg.fill('#password','wrong'); pg.click('button[type=submit]'); pg.wait_for_selector('.auth-error')
    ok('bad login message', 'Incorrect' in pg.inner_text('.auth-error'))
    pg.fill('#password','admin'); pg.click('button[type=submit]'); pg.wait_for_url(B+'/')
    pg.wait_for_selector('#tourDlg[open]'); ok('tour auto-opens first time', True); pg.screenshot(path='shots/tour.png')
    for _ in range(6): pg.click("#tourNext")
    pg.wait_for_selector('#tourDlg:not([open])',state='attached'); ok('tour finishes', not pg.evaluate("document.getElementById('tourDlg').open"))
    pg.reload(); pg.wait_for_timeout(400); ok('tour not shown again', not pg.evaluate("document.getElementById('tourDlg').open"))
    ok('default-credential banner visible', pg.locator('.notice.warn').count()==1)
    pg.screenshot(path='shots/dashboard.png',full_page=True)
    # search palette
    pg.keyboard.press('Control+k'); pg.wait_for_selector('#palette[open]'); pg.fill('#paletteInput','ayush'); pg.wait_for_selector('.pal-label:has-text("Ayush")',timeout=5000); ok('search finds customer', True)
    pg.screenshot(path='shots/search.png'); pg.keyboard.press('Escape')
    # help drawer
    pg.click('body'); pg.keyboard.type('?'); pg.wait_for_timeout(300); ok('help drawer opens with ?', pg.evaluate("document.getElementById('helpDrawer').classList.contains('open')")); pg.screenshot(path='shots/help_drawer.png'); pg.keyboard.press('Escape')
    # invoices list
    pg.goto(B+'/billing'); pg.screenshot(path='shots/billing.png',full_page=True)
    ok('list has 15 rows', pg.locator('tbody tr').count()==15)
    pg.fill('#q','rahul'); pg.press('#q','Enter'); pg.wait_for_load_state(); ok('search filters list', pg.locator('tbody tr').count()==2, pg.locator('tbody tr').count())
    pg.goto(B+'/billing?month=2026-10'); pg.locator('.pay-btn').first.click(); pg.wait_for_selector('#payDlg[open]'); pg.screenshot(path='shots/pay_dialog.png'); pg.click('#payDlg [data-close-dialog]')
    # invoice detail + edit
    pg.goto(B+'/invoices/18'); pg.screenshot(path='shots/invoice_detail.png',full_page=True)
    pg.click('text=Edit invoice >> nth=0'); pg.wait_for_selector('#editForm')
    before=pg.inner_text('#liveTotal')
    pg.fill('#subscription_amount','2500'); pg.click('#addLine'); pg.fill('[name=line_desc]','Delivery'); pg.fill('[name=line_qty]','2'); pg.fill('[name=line_price]','40')
    ok('live total updates', pg.inner_text('#liveTotal').replace(',','')=='₹2580', pg.inner_text('#liveTotal'))
    pg.screenshot(path='shots/invoice_edit.png',full_page=True)
    pg.click('.live-total button[type=submit]'); pg.wait_for_url('**/invoices/18'); ok('saved & redirected', '2,580' in pg.inner_text('body'))
    pg.reload(); ok('persists after reload', '2,580' in pg.inner_text('body'))
    # delete dialog
    pg.click('button:has-text("Delete invoice") >> nth=0'); pg.wait_for_selector('#deleteDlg[open]'); pg.screenshot(path='shots/delete_dialog.png')
    pg.click('#deleteDlg [data-close-dialog]'); ok('cancel keeps invoice', pg.goto(B+'/invoices/18').status==200)
    # locked invoice
    pg.goto(B+'/invoices/2'); ok('locked banner + unlock button', pg.locator('text=Unlock Sep 2026').count()>=1); pg.screenshot(path='shots/invoice_locked.png')
    # new invoice autofill
    pg.goto(B+'/invoices/new'); pg.select_option('#customer_id',label='abhishek'); pg.click('.mp-btn'); pg.click('.mp-m[data-k="2026-11"]'); pg.wait_for_selector('#previewBox strong')
    ok('new invoice auto-fills amount', pg.input_value('#subscription_amount')!=''); pg.screenshot(path='shots/invoice_new.png',full_page=True)
    # confirm modal replaces confirm()
    pg.goto(B+'/billing'); pg.click('text=Generate month'); pg.wait_for_selector('#genDlg[open]'); pg.click('#genDlg [data-close-dialog]')
    pg.goto(B+'/pgs'); 
    # daily toggle with csrf
    pg.goto(B+'/daily'); btn=pg.locator('.daily-meal-btn').first
    if btn.count():
        t0=btn.inner_text(); btn.click(); pg.wait_for_timeout(600); t1=btn.inner_text(); ok('daily toggle works (CSRF header)', t0!=t1, f'{t0}->{t1}'); 
        for _ in range(2): btn.click(); pg.wait_for_timeout(500)
    pg.screenshot(path='shots/daily.png')
    # account change
    pg.goto(B+'/account'); pg.fill('#current_password','admin'); pg.fill('#new_password','Sunrise-2026'); pg.fill('#confirm_password','Sunrise-2026'); pg.click('text=Update sign-in details'); pg.wait_for_selector('.toast.success'); ok('credential change toast', True)
    pg.screenshot(path='shots/account.png'); ok('banner gone', pg.locator('.notice.warn').count()==0)
    pg.goto(B+'/help'); pg.screenshot(path='shots/help.png',full_page=True)
    # walk every nav page
    pages=['/','/customers','/pgs','/daily','/extra-orders','/billing','/payments','/dues','/expenses','/savings','/personal-income','/accounts','/reports','/monthly','/history','/settings','/help','/account','/customers/3','/customers/3/edit','/history/2026-09']
    for u in pages:
        r=pg.goto(B+u); pg.wait_for_load_state('networkidle'); 
        if r.status!=200: fails.append('page '+u)
    # logout
    pg.click('summary.user-chip'); pg.click('text=Sign out'); pg.wait_for_url('**/login'); ok('logout', True)
    # mobile
    m=br.new_context(viewport={'width':390,'height':800},device_scale_factor=2,has_touch=True,is_mobile=True); mp=m.new_page()
    mp.on('pageerror',lambda e: errs.append(('mobile','PAGEERROR '+str(e))))
    mp.on('console',lambda x: errs.append(('mobile',x.text)) if x.type=='error' else None)
    mp.goto(B+'/login'); mp.screenshot(path='shots/m_login.png')
    mp.fill('#username','admin'); mp.fill('#password','Sunrise-2026'); mp.click('button[type=submit]'); mp.wait_for_url(B+'/'); 
    mp.evaluate("fetch('/onboarding/complete',{method:'POST',headers:{'X-CSRF-Token':CSRF}})")
    for n,u in [('dash','/'),('billing','/billing'),('inv','/invoices/15'),('edit','/invoices/15/edit'),('new','/invoices/new'),('customers','/customers'),('daily','/daily')]:
        mp.goto(B+u); mp.wait_for_load_state('networkidle')
        sw=mp.evaluate("document.documentElement.scrollWidth"); ok(f'mobile no horizontal overflow {u}', sw<=392, sw)
        mp.screenshot(path=f'shots/m_{n}.png')
    mp.goto(B+'/'); mp.click('[data-nav-open]'); mp.wait_for_timeout(350); mp.screenshot(path='shots/m_nav.png')
    br.close()
print('\nCONSOLE/NETWORK ISSUES:',len(errs))
for e in errs[:25]: print(e)
print('FAILS:',fails)

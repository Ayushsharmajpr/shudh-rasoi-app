"""Self-contained QA script. Run from the project folder:
    QA_SOURCE_DB=tiffin.db python3 tests/qa_crud.py
It works on a throw-away COPY of the database (your real file is never touched).
Written against the shipped tiffin.db (it uses that data's customer/invoice ids)."""
import os, shutil, tempfile
_SRC = os.environ.get("QA_SOURCE_DB", "tiffin.db")
_TMP = os.path.join(tempfile.mkdtemp(), "qa.db")
shutil.copy(_SRC, _TMP)
import os,sys,re,sqlite3,io,zipfile
os.environ['TIFFIN_DB']=_TMP
sys.path.insert(0,os.getcwd())
import app as A
c=A.app.test_client()
q=lambda s,*a:[r for r in sqlite3.connect(_TMP).execute(s,a).fetchall()]
res=[]
def chk(n,x,e=''): res.append(bool(x)); print(('PASS' if x else 'FAIL'),n,e)
def csrf(): return re.search(r'name="csrf-token" content="([^"]+)"',c.get('/help').data.decode()).group(1)
def post(u,d=None,**k):
    d=dict(d or {}); d['csrf_token']=csrf(); return c.post(u,data=d,**k)
t=re.search(r'name="csrf_token" value="([^"]+)"',c.get('/login').data.decode()).group(1); c.post('/login',data={'username':'admin','password':'admin','csrf_token':t})
# PG
post('/pgs',{'name':'Test PG','address':'X','contact_person':'Y','contact_number':'9876543210'}); pid=q("select id from pgs where name='Test PG'")[0][0]; chk('pg add',pid)
post(f'/pgs/{pid}/edit',{'name':'Test PG 2','address':'X','contact_person':'Y','contact_number':'9876543210'}); chk('pg edit',q('select name from pgs where id=?',pid)[0][0]=='Test PG 2')
# customer
r=post('/customers',{'name':'QA Customer','category':'Regular','contact':'9999999999','area':'A','plan_type':'flat','meal_lunch':'1','monthly_rate':'3000','join_date':'2026-10-01','status':'Active','pg_id':str(pid)})
cr=q("select id,monthly_rate,pg_id from customers where name='QA Customer'"); chk('customer add',cr,cr)
cid=cr[0][0] if cr else None
if cid:
    r=post(f'/customers/{cid}/edit',{'name':'QA Customer','category':'Regular','contact':'9999999999','area':'B','join_date':'2026-10-01','status':'Active','meal_lunch':'1','monthly_rate':'3000','effective_month':'2026-10','pg_id':str(pid)}); chk('customer edit',q('select area from customers where id=?',cid)[0][0]=='B',r.status_code)
    r=post('/customers',{'name':'Bad','category':'Regular','contact':'abc','meal_lunch':'1','monthly_rate':'100','join_date':'2026-10-01','status':'Active'}); chk('customer bad phone rejected',not q("select 1 from customers where name='Bad'"))
    r=post(f'/customers/{cid}/status',{'status':'Paused','pause_from':'2026-10-10','pause_to':'2026-10-12'}); chk('customer pause',q('select status from customers where id=?',cid)[0][0]=='Paused',q('select status from customers where id=?',cid))
    # extra order
    post('/extra-orders',{'date':'2026-10-09','customer_id':str(cid),'description':'box','amount':'120','payment_status':'Unpaid','payment_mode':'Cash'}); eo=q("select id,amount from extra_orders where description='box'"); chk('extra add',eo)
    post(f'/extra-orders/{eo[0][0]}/mark-paid',{'payment_mode':'UPI'}); chk('extra mark paid',q('select payment_status from extra_orders where id=?',eo[0][0])[0][0]=='Paid')
    post(f'/extra-orders/{eo[0][0]}/delete'); chk('extra delete',not q('select 1 from extra_orders where id=?',eo[0][0]))
    # payment advance
    post('/payments',{'date':'2026-10-09','customer_id':str(cid),'amount':'500','payment_mode':'Cash','notes':'adv'}); pm=q("select id from revenue where notes='adv'"); chk('payment add (advance)',pm)
    post(f'/payments/{pm[0][0]}/delete'); chk('payment delete',not q("select 1 from revenue where notes='adv'"))
    post('/customers/%d/delete'%cid); chk('customer delete (no history)',not q('select 1 from customers where id=?',cid))
post(f'/pgs/{pid}/delete'); chk('pg delete',not q('select 1 from pgs where id=?',pid))
# expenses
post('/expenses',{'date':'2026-10-09','category':'Gas Cylinder','description':'qa','amount':'900','payment_mode':'Cash','expense_type':'Business'}); ex=q("select id from expenses where description='qa'"); chk('expense add',ex)
post(f'/expenses/{ex[0][0]}/edit',{'date':'2026-10-09','category':'Gas Cylinder','description':'qa2','amount':'950','payment_mode':'Cash','expense_type':'Business'}); chk('expense edit',q('select amount from expenses where id=?',ex[0][0])[0][0]==950)
r=post('/expenses',{'date':'2026-10-09','category':'Gas Cylinder','description':'neg','amount':'-5','payment_mode':'Cash'}); chk('expense negative rejected',not q("select 1 from expenses where description='neg'"))
post(f'/expenses/{ex[0][0]}/delete'); chk('expense delete',not q('select 1 from expenses where id=?',ex[0][0]))
post('/expenses/recurring/add',{'category':'Rent','description':'rent','amount':'5000','day_of_month':'1','payment_mode':'Cash'}); rc=q('select id from recurring_expenses'); chk('recurring add',rc)
if rc: post(f'/expenses/recurring/{rc[0][0]}/delete'); chk('recurring delete',not q('select 1 from recurring_expenses'))
# savings / personal income
post('/savings',{'date':'2026-10-09','amount':'1000','entry_type':'Deposit','description':'qa','payment_mode':'UPI'}); sv=q("select id from savings where description='qa'"); chk('savings add',sv)
post(f'/savings/{sv[0][0]}/delete'); chk('savings delete',not q('select 1 from savings where id=?',sv[0][0]))
post('/personal-income',{'date':'2026-10-09','source':'Other','amount':'700','description':'qa','payment_mode':'UPI'}); pi=q("select id from personal_income"); chk('personal income add',pi)
post(f'/personal-income/{pi[0][0]}/delete'); chk('personal income delete',not q('select 1 from personal_income'))
# closed-month writes protected
r=post('/expenses',{'date':'2026-09-10','category':'Gas Cylinder','description':'late','amount':'10','payment_mode':'Cash'}); chk('closed-month expense blocked',not q("select 1 from expenses where description='late'"))
# settings / backup / exports
post('/settings',{'business_name':"Shuddh Rasoi by Sharma's",'invoice_prefix':'INV','invoice_due_days':'7'}); chk('settings save',q("select value from settings where key='business_name'")[0][0].startswith('Shuddh'))
r=c.get('/settings/backup'); chk('backup download is valid sqlite',r.status_code==200 and r.data[:15]==b'SQLite format 3')
r=c.get('/export/excel'); chk('excel export',r.status_code==200 and zipfile.is_zipfile(io.BytesIO(r.data)))
r=c.get('/history/2026-09/export.json'); chk('month json export',r.status_code==200)
r=c.get('/nope'); chk('404 page',r.status_code==404 and b'404' in r.data)
r=c.get('/invoices/9999'); chk('missing invoice 404',r.status_code==404)
# data integrity vs original
o=sqlite3.connect(_SRC); n=sqlite3.connect(_TMP)
def cols(c,t): return ','.join(r[1] for r in o.execute(f'pragma table_info({t})'))
same=all(o.execute(f'select {cols(o,t)} from {t} order by 1').fetchall()==n.execute(f'select {cols(o,t)} from {t} order by 1').fetchall() for t in ['customers','invoices','revenue','expenses','extra_orders','daily_log','pgs','savings','settings','customer_rates','month_snapshots','invoice_snapshots','brand_assets'])
chk('original records unchanged after CRUD round-trip',same)
print('RESULT',sum(res),'/',len(res))

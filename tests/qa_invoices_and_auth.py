"""Self-contained QA script. Run from the project folder:
    QA_SOURCE_DB=tiffin.db python3 tests/qa_invoices_and_auth.py
It works on a throw-away COPY of the database (your real file is never touched).
Written against the shipped tiffin.db (it uses that data's customer/invoice ids)."""
import os, shutil, tempfile
_SRC = os.environ.get("QA_SOURCE_DB", "tiffin.db")
_TMP = os.path.join(tempfile.mkdtemp(), "qa.db")
shutil.copy(_SRC, _TMP)
import os,sys,re,sqlite3,io
os.environ['TIFFIN_DB']=_TMP
sys.path.insert(0,os.getcwd())
import app as A, auth
c=A.app.test_client()
db=lambda: sqlite3.connect(_TMP)
def q(sql,*a): 
    d=db(); r=d.execute(sql,a).fetchall(); d.close(); return r
res=[]
def chk(name,cond,extra=''):
    res.append(cond); print(('PASS' if cond else 'FAIL'),name,extra)
def csrf(): 
    return re.search(r'name="csrf-token" content="([^"]+)"',c.get('/help').data.decode()).group(1)
def post(url,data=None,**kw):
    d=dict(data or {}); d['csrf_token']=csrf(); return c.post(url,data=d,**kw)
r=c.get('/login'); t=re.search(r'name="csrf_token" value="([^"]+)"',r.data.decode()).group(1)
c.post('/login',data={'username':'admin','password':'admin','csrf_token':t})

before={t:q(f'select count(*) from {t}')[0][0] for t in ['customers','invoices','revenue','expenses','extra_orders','daily_log','pgs','savings','month_snapshots','customer_rates']}
chk('csrf: POST without token rejected', c.post('/invoices/20/delete').status_code==302 and q('select count(*) from invoices where id=20')[0][0]==1)

# --- EDIT open-month invoice (id 17 Piyush Oct 4000, no payments)
r=post('/invoices/17/edit',{'generated_date':'2026-10-06','due_date':'2026-10-20','notes':'Thanks!','bill_to_name':'Piyush K','bill_to_area':'JTM mall, Jaipur','bill_to_contact':'8905783973','subscription_amount':'4500',
   'line_id':['',''],'line_desc':['Delivery fee','Festival discount'],'line_qty':['2','1'],'line_price':['50','-100']})
row=q('select subscription_amount,extra_amount,total_amount,generated_date,due_date,notes,bill_to_name from invoices where id=17')[0]
chk('edit: totals recalculated (4500 + 2*50 - 100 = 4500)', row[0]==4500 and row[1]==0 and row[2]==4500, row)
chk('edit: dates/notes/billto saved', row[3:]==('2026-10-06','2026-10-20','Thanks!','Piyush K'), row[3:])
chk('edit: snapshot final total synced', q("select final_total,final_subscription_amount from invoice_snapshots where invoice_id=17 and status='active'")[0]==(4500,4500))
chk('edit: lines stored', q('select description,quantity,unit_price,amount from invoice_lines where invoice_id=17 order by id')==[('Delivery fee',2,50,100),('Festival discount',1,-100,-100)])
html=c.get('/invoices/17').data.decode(); chk('view shows edited bill-to, lines & notes','Piyush K' in html and 'Festival discount' in html and 'Thanks!' in html)
pdf=c.get('/invoices/17/pdf'); chk('pdf renders after edit',pdf.status_code==200 and pdf.data[:4]==b'%PDF')
# amount only change + reload persistence
lid=[r[0] for r in q('select id from invoice_lines where invoice_id=17 order by id')]
r=post('/invoices/17/edit',{'generated_date':'2026-10-06','due_date':'','bill_to_name':'Piyush K','bill_to_area':'x','bill_to_contact':'','subscription_amount':'4000','line_id':[str(lid[0])],'line_desc':['Delivery fee'],'line_qty':['3'],'line_price':['50']})
row=q('select subscription_amount,extra_amount,total_amount,due_date from invoices where invoice_id is null or id=17')[0] if False else q('select subscription_amount,extra_amount,total_amount,due_date from invoices where id=17')[0]
chk('edit: remove line + change qty + clear due date', row==(4000,150,4150,None) and len(q('select 1 from invoice_lines where invoice_id=17'))==1, row)
# validation
r=post('/invoices/17/edit',{'generated_date':'2026-10-06','due_date':'2026-10-01','bill_to_name':'P','subscription_amount':'-5'})
chk('edit: bad input rejected with errors, data unchanged', r.status_code==400 and b'earlier than the invoice date' in r.data and b'cannot be negative' in r.data and q('select total_amount from invoices where id=17')[0][0]==4150)
r=post('/invoices/17/edit',{'generated_date':'2026-10-06','bill_to_name':'P','subscription_amount':'10','line_desc':['disc'],'line_qty':['1'],'line_price':['-500']})
chk('edit: negative total rejected', r.status_code==400 and b'negative' in r.data)
# --- EDIT amount on invoice WITH payment (id 15, paid 24900)
r=post('/invoices/15/edit',{'generated_date':'2026-10-04','due_date':'2026-11-07','bill_to_name':'Ayush arora','bill_to_area':'Malviya Nagar NXT home','bill_to_contact':'9587194127','subscription_amount':'80000'})
chk('edit: amount on paid invoice updates total, keeps payments', q('select total_amount from invoices where id=15')[0][0]==80000 and q('select sum(amount) from revenue where invoice_id=15')[0][0]==24900)
# --- extras attach/detach on invoice 18 (prakriti) 
d=db(); d.execute("insert into extra_orders(date,customer_id,description,amount,payment_status) values('2026-10-05',6,'guest box',150,'Unpaid')"); d.commit(); eid=d.execute('select max(id) from extra_orders').fetchone()[0]; d.close()
r=post('/invoices/18/edit',{'generated_date':'2026-10-04','due_date':'2026-11-07','bill_to_name':'prakriti','subscription_amount':'2100','attach_extra':[str(eid)]})
chk('edit: attach unbilled extra', q('select extra_amount,total_amount from invoices where id=18')[0]==(150,2250) and q('select invoice_id from extra_orders where id=?',eid)[0][0]==18)
r=post('/invoices/18/edit',{'generated_date':'2026-10-04','due_date':'2026-11-07','bill_to_name':'prakriti','subscription_amount':'2100','extra_id':[str(eid)],'extra_date':['2026-10-05'],'extra_desc':['guest box'],'extra_amount':['150'],'extra_remove':[str(eid)]})
chk('edit: detach extra returns it to unbilled', q('select extra_amount,total_amount from invoices where id=18')[0]==(0,2100) and q('select invoice_id from extra_orders where id=?',eid)[0][0] is None)

# --- LOCKED month (Sept, id 1)
r=post('/invoices/1/edit',{'generated_date':'2026-09-07','bill_to_name':'x','subscription_amount':'1'})
chk('locked: edit refused, nothing changed', q('select total_amount from invoices where id=1')[0][0]==49600 and r.status_code==400)
r=post('/invoices/1/delete',{'payments':'keep'}); chk('locked: delete refused', q('select count(*) from invoices where id=1')[0][0]==1)
r=post('/invoices/1/unlock'); chk('unlock month', q("select status from month_snapshots where ym='2026-09' order by id desc limit 1")[0][0]=='superseded')
r=post('/invoices/1/edit',{'generated_date':'2026-09-07','due_date':'2026-10-07','bill_to_name':'Ayush arora','subscription_amount':'50000'})
chk('after unlock: edit works', q('select total_amount from invoices where id=1')[0][0]==50000)

# --- DELETE
r=post('/invoices/20/delete'); chk('delete unpaid invoice (id 20)', q('select count(*) from invoices where id=20')[0][0]==0 and q("select status from invoice_snapshots where invoice_id=20")[0][0]=='voided')
r=post('/invoices/16/delete'); chk('delete paid invoice without choice refused', q('select count(*) from invoices where id=16')[0][0]==1)
r=post('/invoices/16/delete',{'payments':'keep'}); chk('delete paid invoice, keep payments as advance', q('select count(*) from invoices where id=16')[0][0]==0 and q('select invoice_id,amount from revenue where id=41')[0]==(None,3900))
r=post('/invoices/19/delete',{'payments':'delete'}); chk('delete invoice + payments', q('select count(*) from invoices where id=19')[0][0]==0 and q('select count(*) from revenue where id=32')[0][0]==0)
r=c.get('/invoices/20'); chk('deleted invoice 404', r.status_code==404)
r=c.get('/billing'); chk('billing list renders after deletes', r.status_code==200 and b'INV-2026-0013' not in r.data)

# --- CREATE
pv=c.get('/invoices/preview?customer_id=6&month=2026-10').get_json(); chk('preview returns calc', pv['ok'] and pv['calculated_amount']>0 and pv['existing'] is not None, pv['calculated_amount'])
r=post('/invoices/new',{'customer_id':'6','month':'2026-10','generated_date':'2026-10-09','subscription_amount':'2100'}); chk('create: duplicate refused', b'already has invoice' in r.data)
r=post('/invoices/new',{'customer_id':'8','month':'2026-10','generated_date':'2026-10-09','due_date':'2026-11-07','subscription_amount':'1700','line_desc':['GST'],'line_qty':['1'],'line_price':['85']},follow_redirects=False)
new=q('select id,invoice_number,total_amount,subscription_amount,extra_amount from invoices where customer_id=8 and period_start="2026-10-01"')
chk('create: invoice for deleted slot (cust 8)', len(new)==1 and new[0][2]==1785, new)
chk('create: gets snapshot & unique number', q("select count(*) from invoice_snapshots where invoice_id=? and status='active'",new[0][0])[0][0]==1 and new[0][1] not in ('INV-2026-0013',), new[0][1])
r=post('/invoices/new',{'customer_id':'8','month':'2026-08','generated_date':'2026-10-09','subscription_amount':'1'}); 
# quick pay
bal=c.get('/billing?status=Unpaid').status_code; chk('filter by status',bal==200)
r=post(f'/invoices/{new[0][0]}/pay',{'amount':'785','date':'2026-10-09','payment_mode':'UPI'}); chk('pay: partial payment', q('select sum(amount) from revenue where invoice_id=?',new[0][0])[0][0]==785)
r=post(f'/invoices/{new[0][0]}/pay',{'amount':'99999','date':'2026-10-09','payment_mode':'UPI'}); chk('pay: overpayment refused', q('select sum(amount) from revenue where invoice_id=?',new[0][0])[0][0]==785)
r=post('/invoices/22/edit',{'generated_date':'2026-10-05','due_date':'2026-11-07','bill_to_name':'abhishek','bill_to_area':'govindam residency','bill_to_contact':'8949803079','subscription_amount':'4800','update_customer':'1'}); chk('update_customer propagates', q('select name from customers where id=10')[0][0]=='abhishek')

# --- credentials
r=post('/account/credentials',{'current_password':'wrong','new_password':'Better-pass-1','confirm_password':'Better-pass-1'}); chk('creds: wrong current pw refused', q('select must_change from users')[0][0]==1)
r=post('/account/credentials',{'current_password':'admin','new_password':'admin','confirm_password':'admin'}); chk('creds: weak pw refused', q('select must_change from users')[0][0]==1)
r=post('/account/credentials',{'current_password':'admin','new_password':'abc','confirm_password':'abc'}); chk('creds: short pw refused', q('select must_change from users')[0][0]==1)
r=post('/account/credentials',{'current_password':'admin','new_password':'Better-pass-1','confirm_password':'nope'}); chk('creds: mismatch refused', q('select must_change from users')[0][0]==1)
r=post('/account/credentials',{'current_password':'admin','new_username':'owner','new_password':'Better-pass-1','confirm_password':'Better-pass-1'})
u=q('select username,password_hash,must_change from users')[0]
chk('creds: changed, hashed, flag cleared', u[0]=='owner' and u[2]==0 and 'Better-pass-1' not in u[1] and u[1].split('$')[0] in ('scrypt','pbkdf2:sha256') or u[1].startswith(('scrypt','pbkdf2')), u[1][:12])
chk('creds: still signed in after change', c.get('/customers').status_code==200)
c2=A.app.test_client(); t2=re.search(r'name="csrf_token" value="([^"]+)"',c2.get('/login').data.decode()).group(1)
chk('creds: old login fails', b'Incorrect username or password' in c2.post('/login',data={'username':'admin','password':'admin','csrf_token':t2}).data)
chk('creds: new login works', c2.post('/login',data={'username':'OWNER','password':'Better-pass-1','csrf_token':t2}).status_code==302)
# lockout
auth.reset_throttle(); c3=A.app.test_client(); t3=re.search(r'name="csrf_token" value="([^"]+)"',c3.get('/login').data.decode()).group(1)
for i in range(5): c3.post('/login',data={'username':'owner','password':'x'+str(i),'csrf_token':t3})
r=c3.post('/login',data={'username':'owner','password':'Better-pass-1','csrf_token':t3}); chk('lockout after 5 failures (even right pw blocked)', b'Too many failed attempts' in r.data)
auth.reset_throttle()
# open redirect
r=c2.get('/login?next=//evil.com'); chk('login while signed in ignores external next', 'evil' not in (r.headers.get('Location') or ''))
c4=A.app.test_client(); t4=re.search(r'name="csrf_token" value="([^"]+)"',c4.get('/login?next=//evil.com').data.decode()).group(1)
r=c4.post('/login',data={'username':'owner','password':'Better-pass-1','csrf_token':t4,'next':'https://evil.com'}); chk('no open redirect', r.headers['Location'].endswith('/'))
# logout
r=post('/logout'); chk('logout then protected', c.get('/customers').status_code==302)
c.post('/login',data={'username':'owner','password':'Better-pass-1','csrf_token':re.search(r'name="csrf_token" value="([^"]+)"',c.get('/login').data.decode()).group(1)})
# restore keeps account
bk=open(_SRC,'rb').read()
r=post('/settings/restore',{'confirm':'1','backup_file':(io.BytesIO(bk),'old.db')},content_type='multipart/form-data')
chk('restore old backup keeps current sign-in', q('select username from users')[0][0]=='owner' and q('select count(*) from invoices')[0][0]==15, q('select username from users'))
chk('restored data present', q('select count(*) from customers')[0][0]==8)
print('\nRESULT',sum(res),'/',len(res))

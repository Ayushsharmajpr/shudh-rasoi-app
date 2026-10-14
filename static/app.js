/* Shudh Rasoi — shared front-end behaviour (no dependencies). */
(function(){
'use strict';
var $=function(s,r){return (r||document).querySelector(s)}, $$=function(s,r){return [].slice.call((r||document).querySelectorAll(s))};
var CSRF=(($('meta[name=csrf-token]')||{}).content)||'';
window.CSRF=CSRF;

/* ---------- toasts ---------- */
function mkToast(msg,type){
  var st=$('#toastStack'); if(!st) return alert(msg);
  var t=document.createElement('div'); t.className='toast '+(type||'info'); t.setAttribute('role',type==='error'?'alert':'status');
  var sp=document.createElement('span'); sp.textContent=msg;
  var b=document.createElement('button'); b.className='close-toast'; b.type='button'; b.setAttribute('aria-label','Dismiss'); b.textContent='×';
  b.onclick=function(){t.remove()}; t.appendChild(sp); t.appendChild(b); st.appendChild(t); life(t,type==='error'?9000:4800);
}
function life(t,ms){setTimeout(function(){t.classList.add('out');setTimeout(function(){t.remove()},220)},ms)}
window.toast=mkToast;
$$('.toast').forEach(function(t){life(t,t.classList.contains('error')?9000:4800)});

/* ---------- confirm dialog (replaces window.confirm, keeps the same call sites) ---------- */
var cdlg=$('#confirmDlg'), cb=null;
function openConfirm(msg,ok,opts){
  opts=opts||{}; if(!cdlg){ if(confirm(msg)) ok(); return; }
  $('#confirmMsg').textContent=msg; $('#confirmTitle').textContent=opts.title||'Please confirm';
  var yes=$('#confirmYes'); yes.textContent=opts.yes||(/delete|remove|erase|replace/i.test(msg)?'Yes, continue':'Continue');
  yes.className='btn '+((opts.danger!==undefined?opts.danger:/delete|remove|erase|replace/i.test(msg))?'danger':'');
  cb=ok; cdlg.showModal(); yes.focus();
}
window.openConfirm=openConfirm;
if(cdlg){
  $('#confirmYes').addEventListener('click',function(){var f=cb;cb=null;cdlg.close();if(f)f()});
  $('#confirmNo').addEventListener('click',function(){cb=null;cdlg.close()});
  cdlg.addEventListener('click',function(e){if(e.target===cdlg){cb=null;cdlg.close()}});
}
window.confirmAction=function(msg){
  var ev=window.event, el=ev&&(ev.currentTarget&&ev.currentTarget.tagName?ev.currentTarget:ev.target);
  var form=el&&(el.tagName==='FORM'?el:(el.form||(el.closest&&el.closest('form'))));
  if(!form) return confirm(msg);
  if(form.dataset.confirmed==='1'){form.dataset.confirmed='';return true}
  var submitter=(el.tagName==='FORM')?(ev&&ev.submitter):el;
  openConfirm(msg,function(){form.dataset.confirmed='1'; if(form.requestSubmit){submitter&&submitter.form===form?form.requestSubmit(submitter):form.requestSubmit()}else{form.submit()}});
  return false;
};
/* generic dialog helpers: <button data-open-dialog="id"> and <button data-close-dialog> */
document.addEventListener('click',function(e){
  var o=e.target.closest('[data-open-dialog]'); if(o){var d=document.getElementById(o.dataset.openDialog); if(d){e.preventDefault(); d.showModal();}}
  var c=e.target.closest('[data-close-dialog]'); if(c){var dd=c.closest('dialog'); if(dd) dd.close();}
  if(e.target.tagName==='DIALOG'&&e.target.hasAttribute('data-light-dismiss')) e.target.close();
});

/* ---------- navigation drawer ---------- */
var shell=$('#appShell');
$$('[data-nav-open]').forEach(function(b){b.addEventListener('click',function(){shell.classList.add('nav-open')})});
$$('[data-nav-close]').forEach(function(b){b.addEventListener('click',function(){shell.classList.remove('nav-open')})});

/* ---------- dropdown menus (<details class="menu">) ---------- */
document.addEventListener('click',function(e){$$('details.menu[open]').forEach(function(d){if(!d.contains(e.target))d.removeAttribute('open')})});
document.addEventListener('keydown',function(e){if(e.key==='Escape'){$$('details.menu[open]').forEach(function(d){d.removeAttribute('open')});shell&&shell.classList.remove('nav-open');closeDrawer()}});

/* ---------- help drawer ---------- */
var drawer=$('#helpDrawer');
function openDrawer(){if(!drawer)return;drawer.classList.add('open');drawer.setAttribute('aria-hidden','false');$('#helpClose').focus()}
function closeDrawer(){if(!drawer)return;drawer.classList.remove('open');drawer.setAttribute('aria-hidden','true')}
$$('[data-help-open]').forEach(function(b){b.addEventListener('click',function(e){e.preventDefault();drawer.classList.contains('open')?closeDrawer():openDrawer()})});
var hc=$('#helpClose'); if(hc) hc.addEventListener('click',closeDrawer);

/* ---------- tooltips ---------- */
var pop=document.createElement('div'); pop.className='tip-pop'; pop.setAttribute('role','tooltip'); pop.hidden=true; document.body.appendChild(pop);
function showTip(el){pop.textContent=el.dataset.tip;pop.hidden=false;var r=el.getBoundingClientRect(),w=pop.offsetWidth,h=pop.offsetHeight;
  var x=Math.min(Math.max(8,r.left+r.width/2-w/2),window.innerWidth-w-8), y=r.top-h-8; if(y<8)y=r.bottom+8; pop.style.left=x+'px';pop.style.top=y+'px'}
function hideTip(){pop.hidden=true}
document.addEventListener('mouseover',function(e){var t=e.target.closest&&e.target.closest('.help-tip');if(t)showTip(t)});
document.addEventListener('mouseout',function(e){if(e.target.closest&&e.target.closest('.help-tip'))hideTip()});
document.addEventListener('focusin',function(e){var t=e.target.closest&&e.target.closest('.help-tip');if(t)showTip(t)});
document.addEventListener('focusout',hideTip);
document.addEventListener('click',function(e){var t=e.target.closest&&e.target.closest('.help-tip');if(t){e.preventDefault();pop.hidden&&true?showTip(t):hideTip()}else hideTip()});
window.addEventListener('scroll',hideTip,true);

/* ---------- command palette / global search ---------- */
var pal=$('#palette'), pin=$('#paletteInput'), plist=$('#paletteList'), ptimer=null, pitems=[], psel=0, pseq=0;
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function renderPal(data){
  var html='',n=0; pitems=[];
  [['customers','Customers'],['invoices','Invoices'],['pages','Go to']].forEach(function(g){
    var arr=data[g[0]]||[]; if(!arr.length)return; html+='<li class="pal-group">'+g[1]+'</li>';
    arr.forEach(function(it){pitems.push(it);html+='<li role="option" id="pal-'+n+'" data-i="'+n+'" class="pal-item"><span class="pal-label">'+esc(it.label)+'</span><span class="pal-hint">'+esc(it.hint||'')+'</span></li>';n++});
  });
  plist.innerHTML=html||'<li class="pal-empty">No matches. Try a customer name or invoice number.</li>'; psel=0; markPal();
}
function markPal(){$$('.pal-item',plist).forEach(function(li,i){li.classList.toggle('sel',i===psel);li.setAttribute('aria-selected',i===psel)}); var s=$('.pal-item.sel',plist); if(s){s.scrollIntoView({block:'nearest'});pin.setAttribute('aria-activedescendant',s.id)}}
function doSearch(){
  var q=pin.value.trim(), my=++pseq; plist.classList.add('loading');
  fetch('/api/search?q='+encodeURIComponent(q),{headers:{'X-Requested-With':'fetch'},credentials:'same-origin'})
   .then(function(r){if(r.status===401){location.href='/login';throw 0}return r.json()})
   .then(function(d){if(my===pseq){plist.classList.remove('loading');renderPal(d)}})
   .catch(function(){plist.classList.remove('loading')});
}
function openPal(){if(!pal)return;pal.showModal();pin.value='';doSearch();pin.focus()}
$$('[data-search-open]').forEach(function(b){b.addEventListener('click',openPal)});
if(pal){
  pin.addEventListener('input',function(){clearTimeout(ptimer);ptimer=setTimeout(doSearch,150)});
  pin.addEventListener('keydown',function(e){
    if(e.key==='ArrowDown'){e.preventDefault();psel=Math.min(psel+1,pitems.length-1);markPal()}
    else if(e.key==='ArrowUp'){e.preventDefault();psel=Math.max(psel-1,0);markPal()}
    else if(e.key==='Enter'&&pitems[psel]){e.preventDefault();location.href=pitems[psel].url}
  });
  plist.addEventListener('click',function(e){var li=e.target.closest('.pal-item');if(li)location.href=pitems[+li.dataset.i].url});
  pal.addEventListener('click',function(e){if(e.target===pal)pal.close()});
}
document.addEventListener('keydown',function(e){
  var typing=/^(INPUT|TEXTAREA|SELECT)$/.test((document.activeElement||{}).tagName)||(document.activeElement||{}).isContentEditable;
  if((e.key==='k'||e.key==='K')&&(e.ctrlKey||e.metaKey)){e.preventDefault();openPal()}
  else if(e.key==='/'&&!typing&&!e.ctrlKey&&!e.metaKey){e.preventDefault();openPal()}
  else if(e.key==='?'&&!typing){e.preventDefault();drawer&&(drawer.classList.contains('open')?closeDrawer():openDrawer())}
});

/* ---------- first-run tour ---------- */
var tour=$('#tourDlg');
if(tour){
  var steps=JSON.parse($('#tourData').textContent), ti=0;
  var tt=$('#tourTitle'), tb=$('#tourBody'), tp=$('#tourProg'), tn=$('#tourNext'), tprev=$('#tourPrev');
  function drawTour(){var s=steps[ti];tt.textContent=s.title;tb.textContent=s.body;
    tp.innerHTML=steps.map(function(_,i){return '<span class="dot'+(i===ti?' on':'')+'"></span>'}).join('');
    tprev.hidden=ti===0; tn.textContent=ti===steps.length-1?'Start using the app':'Next'}
  function finishTour(){fetch('/onboarding/complete',{method:'POST',headers:{'X-CSRF-Token':CSRF,'X-Requested-With':'fetch'},credentials:'same-origin'});tour.close()}
  tn.addEventListener('click',function(){if(ti<steps.length-1){ti++;drawTour()}else finishTour()});
  tprev.addEventListener('click',function(){if(ti>0){ti--;drawTour()}});
  $('#tourSkip').addEventListener('click',finishTour);
  tour.addEventListener('cancel',function(){finishTour()});
  drawTour(); if(tour.dataset.auto==='1') tour.showModal();
  $$('[data-tour-open]').forEach(function(b){b.addEventListener('click',function(){ti=0;drawTour();closeDrawer();tour.showModal()})});
}

/* ---------- never submit twice; show progress ---------- */
document.addEventListener('submit',function(e){
  var f=e.target; if(f.method&&f.method.toLowerCase()!=='post'||f.dataset.noLock) return;
  setTimeout(function(){ if(e.defaultPrevented) return; $$('button[type=submit],button:not([type])',f).forEach(function(b){b.classList.add('is-loading');b.disabled=true}) },0);
});
window.addEventListener('pageshow',function(e){if(e.persisted)$$('button.is-loading').forEach(function(b){b.classList.remove('is-loading');b.disabled=false})});

/* ---------- live money formatting helper for pages ---------- */
window.fmtMoney=function(n,sym){var neg=n<0;n=Math.abs(n);var s=n.toLocaleString('en-IN',{minimumFractionDigits:0,maximumFractionDigits:2});return (neg?'−':'')+(sym||'₹')+s};

/* Calendar-style month picker: progressively enhances every month <select> */
(function(){
var MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], open=null;
function closeAll(){if(open){open.classList.remove('open');open=null}}
document.addEventListener('click',function(e){if(open&&!open.contains(e.target))closeAll()});
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeAll()});
document.querySelectorAll('select[name="month"],select[name="effective_month"],select[data-month-picker]').forEach(function(sel){
  var opts=[].slice.call(sel.options),avail={},allLbl=null;
  opts.forEach(function(o){if(/^\d{4}-\d{2}$/.test(o.value))avail[o.value]=o.text.trim();else if(o.value==='')allLbl=o.text.trim()});
  var keys=Object.keys(avail).sort(); if(!keys.length)return;
  var minY=+keys[0].slice(0,4),maxY=+keys[keys.length-1].slice(0,4);
  var cur=sel.value&&avail[sel.value]?sel.value:(sel.value===''&&allLbl?'':keys[keys.length-1]);
  var year=+(cur||keys[keys.length-1]).slice(0,4);
  var w=document.createElement('div');w.className='mp';
  var b=document.createElement('button');b.type='button';b.className='mp-btn';b.setAttribute('aria-haspopup','dialog');
  var pop=document.createElement('div');pop.className='mp-pop';pop.setAttribute('role','dialog');pop.setAttribute('aria-label','Choose month');
  sel.parentNode.insertBefore(w,sel);w.appendChild(b);w.appendChild(pop);w.appendChild(sel);sel.style.display='none';
  var cal='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>';
  function label(){return cur===''?(allLbl||'All time'):avail[cur]}
  function render(){
    b.innerHTML=cal+'<span>'+label()+'</span><span class="mp-caret">▾</span>';
    var h='<div class="mp-head"><button type="button" class="mp-nav" data-d="-1" aria-label="Previous year"'+(year<=minY?' disabled':'')+'>‹</button><strong>'+year+'</strong><button type="button" class="mp-nav" data-d="1" aria-label="Next year"'+(year>=maxY?' disabled':'')+'>›</button></div><div class="mp-grid">';
    for(var i=0;i<12;i++){var k=year+'-'+('0'+(i+1)).slice(-2),ok=!!avail[k];
      h+='<button type="button" class="mp-m'+(k===cur?' sel':'')+'" data-k="'+k+'"'+(ok?'':' disabled')+'>'+MN[i]+'</button>'}
    h+='</div><div class="mp-foot">'+(allLbl!==null?'<button type="button" class="mp-link" data-k="">'+allLbl+'</button>':'<span></span>');
    var now=new Date(),tk=now.getFullYear()+'-'+('0'+(now.getMonth()+1)).slice(-2);
    h+=(avail[tk]?'<button type="button" class="mp-link" data-k="'+tk+'">This month</button>':'')+'</div>';
    pop.innerHTML=h;
  }
  b.addEventListener('click',function(e){e.stopPropagation();var was=w.classList.contains('open');closeAll();if(!was){year=+(cur||keys[keys.length-1]).slice(0,4);render();w.classList.add('open');open=w}});
  pop.addEventListener('click',function(e){
    var t=e.target.closest('button');if(!t)return;e.stopPropagation();
    if(t.classList.contains('mp-nav')){year+=+t.dataset.d;render();return}
    if(t.dataset.k===undefined||t.disabled)return;
    cur=t.dataset.k;sel.value=cur;render();closeAll();
    sel.dispatchEvent(new Event('change',{bubbles:true}));
  });
  render();
});
})();

})();

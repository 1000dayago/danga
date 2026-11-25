/* 단가 관리 — shared data model, Google Drive sync and shared views.
   No price data lives in this file or anywhere in the repository:
   data comes from importing 단가.xlsx and is kept in the user's Google Drive. */
(function(){
"use strict";
const A=window.Danga={};
const APP='danga-manager',SCHEMA=2;
const FOLDER='application/vnd.google-apps.folder';
const SCOPE_RW='https://www.googleapis.com/auth/drive',SCOPE_RO='https://www.googleapis.com/auth/drive.readonly';
const tkey=()=>'danga.token'+(window.Danga&&window.Danga.readOnly?'.sales':'');
const REASONS=['원자재 가격 변동','연간 단가 협상','환율 변동','신규 거래','거래 중단','물량 조건 변경','입력 오류 수정','기타'];
const KNOWN={'하나테크':'hana','덕우':'dukwoo','삼성 광주':'sam_gj','태성전기':'taesung','유남전기':'yunam','DCE':'dce','그린테크':'green','네오 글로벌':'neo','오비오':'obio','유니크 대성':'unique','세고스':'segos','경인본부':'gyeongin','삼성 태국':'sam_th','삼성 폴란드':'sam_pl','삼성 멕시코':'sam_mx','삼성 베트남':'sam_vn','히타치':'hitachi','왈튼':'walton','베스텔':'vestel','엔터캅':'entercop','텐사이':'tensai','R F L':'rfl'};
const XLSX_SRC='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

/* ---------- utils ---------- */
const $=(s,el)=>(el||document).querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const LS={get(k){try{const v=localStorage.getItem(k);return v?JSON.parse(v):null}catch(e){return null}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}},del(k){try{localStorage.removeItem(k)}catch(e){}}};
const pad=n=>String(n).padStart(2,'0');
const isoD=d=>d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
const today=()=>isoD(new Date());
const addDays=(s,n)=>{const d=new Date(s+'T00:00:00');d.setDate(d.getDate()+n);return isoD(d)};
const nowIso=()=>{const d=new Date();return isoD(d)+'T'+pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds())};
const nf=new Intl.NumberFormat('ko-KR',{maximumFractionDigits:2});
const fmtN=v=>v==null||isNaN(v)?'':nf.format(v);
const uid=p=>p+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const shortDT=s=>s?String(s).replace('T',' ').slice(0,16):'';
const localDT=s=>{if(!s)return '';const d=new Date(s);if(isNaN(d))return shortDT(s);return isoD(d)+' '+pad(d.getHours())+':'+pad(d.getMinutes())};
const clone=o=>JSON.parse(JSON.stringify(o));
const norm=s=>String(s).replace(/\s+/g,' ').trim();
Object.assign(A,{$,esc,LS,today,addDays,nowIso,fmtN,shortDT,localDT,REASONS,uid,norm,pad,isoD,openModal});

let toastT;
A.toast=function(msg,opt){let t=$('#toast');if(!t){t=document.createElement('div');t.id='toast';t.className='toast';t.setAttribute('role','status');document.body.appendChild(t)}t.textContent=msg;if(opt&&opt.action){const b=document.createElement('button');b.type='button';b.className='toast-act';b.textContent=opt.action.label;b.addEventListener('click',()=>{t.hidden=true;opt.action.fn()});t.appendChild(b)}t.hidden=false;clearTimeout(toastT);toastT=setTimeout(()=>t.hidden=true,opt&&opt.action?8000:3800)};

/* ---------- state ---------- */
const emptyDb=()=>({columns:[],items:[],entries:[],fx:[],deleted:[],removedItems:[],removedColumns:[],log:[],fxCurs:[],groups:[],sheetLog:[]});
const validDb=d=>d&&Array.isArray(d.columns)&&Array.isArray(d.items)&&Array.isArray(d.entries);
const normDb=d=>{for(const k of ['fx','deleted','removedItems','removedColumns','log','fxCurs','groups','sheetLog'])if(!Array.isArray(d[k]))d[k]=[];return d};
A.settings=Object.assign({clientId:'',path:'회사 폴더/단가관리',prefix:'단가DB',autoSync:true,folderId:null,folderFor:null,connected:false,lastBackup:null,userName:'',company:'',sheetNote:'',warnPct:30,lowMargin:10,basis:'hana',viewKrw:false,favs:[],salesMode:false,sharePath:'회사 폴더/단가관리/영업용',autoShare:false,shareFolderId:null,shareFolderFor:null,lastShare:null},LS.get('danga.settings')||{});
A.readOnly=!!A.settings.salesMode;
A.saveSettings=()=>LS.set('danga.settings',A.settings);
const DKEY=A.readOnly?'danga.draft.sales':'danga.draft';
const draft=LS.get(DKEY);
A.DB=draft&&validDb(draft.db)?normDb(draft.db):emptyDb();
A.dirty=draft?draft.dirty||0:0;
A.sync=Object.assign({fileId:null,remoteModified:null,at:null,title:null,overwrite:false},draft&&draft.sync||{});
A.persist=()=>LS.set(DKEY,{db:A.DB,dirty:A.dirty,sync:A.sync});
const listeners=[];
A.onChange=fn=>listeners.push(fn);
A.emit=()=>listeners.forEach(f=>{try{f()}catch(e){console.error(e)}});
/* undo: one snapshot per user action; kept until the next save or load */
A.undoStack=[];let prevStr=null,batching=0,batchPushed=false,batchLabel='';
const snap=()=>JSON.stringify(A.DB);
A.resetUndo=()=>{A.undoStack=[];prevStr=snap()};
A.touch=(n,label)=>{
  const push=lb=>{A.undoStack.push({str:prevStr,dirty:A.dirty,label:lb||'변경'});if(A.undoStack.length>30)A.undoStack.shift()};
  if(batching){if(!batchPushed){push(batchLabel);batchPushed=true}}else push(label);
  A.dirty+=n||1;A.persist();prevStr=snap();A.emit();
};
let batchTag=null;
A.batch=(label,fn)=>{batching++;batchLabel=label;batchPushed=false;const mine=!batchTag;if(mine)batchTag={id:uid('bt'),label};try{return fn()}finally{batching--;if(mine)batchTag=null}};
A.undo=()=>{const e=A.undoStack.pop();if(!e)return null;A.DB=normDb(JSON.parse(e.str));A.reindex();A.dirty=e.dirty;A.persist();prevStr=e.str;A.emit();document.dispatchEvent(new Event('danga:undo'));return e.label};
A.resetUndo();
A.isEmpty=()=>!A.DB.items.length;

/* ---------- index & lookups ---------- */
let IDX=new Map();A.COL=new Map();A.ITEM=new Map();
const cmpE=(a,b)=>a.from<b.from?-1:a.from>b.from?1:(a.at<b.at?-1:a.at>b.at?1:0);
A.cmpE=cmpE;
A.reindex=function(){
  IDX=new Map();A.COL=new Map(A.DB.columns.map(c=>[c.key,c]));A.ITEM=new Map(A.DB.items.map(i=>[i.id,i]));
  for(const e of A.DB.entries){const k=e.item+'|'+e.col;let a=IDX.get(k);if(!a)IDX.set(k,a=[]);a.push(e)}
  for(const a of IDX.values())a.sort(cmpE);
};
A.hist=(i,c)=>IDX.get(i+'|'+c)||[];
A.entryAt=(i,c,d)=>{let r=null;for(const e of A.hist(i,c)){if(e.from<=d)r=e;else break}return r};
A.prevOf=e=>{const h=A.hist(e.item,e.col);const k=h.indexOf(e);return k>0?h[k-1]:null};
A.hasFuture=(i,c,d)=>{const h=A.hist(i,c);return h.length>0&&h[h.length-1].from>d};
const isNum=e=>!!e&&typeof e.v==='number';A.isNum=isNum;
A.hasVal=e=>!!e&&(isNum(e)||!!e.t);
A.eHtml=(e,d)=>{if(!e)return '';if(isNum(e)){const c=e.cur||'KRW';if(c==='KRW')return esc(fmtN(e.v));if(A.view.krw&&d){const r=A.fxAt(d,c);if(r)return '<span class="conv" title="'+esc(fmtN(e.v)+' '+c+' × '+fmtN(r)+'원')+'">'+esc(fmtN(Math.round(e.v*r)))+'</span>';return '<span class="cur">'+esc(c)+'</span>'+esc(fmtN(e.v))+'<span class="nofx" title="환율이 없어 원래 통화로 표시합니다">!</span>'}return '<span class="cur">'+esc(c)+'</span>'+esc(fmtN(e.v))}return esc(e.t||'')};
A.eText=e=>{if(!e)return '없음';if(isNum(e))return fmtN(e.v)+(e.cur&&e.cur!=='KRW'?' '+e.cur:'원');return e.t?e.t:'없음'};
/* ---------- currencies ---------- */
const CUR_CODES='KRW USD EUR RMB JPY GBP AUD CAD CHF HKD SGD TWD THB VND IDR MYR PHP INR PKR BDT LKR NPR MMK KHR LAK MNT KZT UZS RUB UAH BYN TRY AED SAR QAR KWD BHD OMR JOD ILS EGP MAD TND DZD NGN GHS KES TZS UGX ZAR ETB XOF XAF NZD FJD PGK MXN BRL ARS CLP COP PEN UYU BOB PYG CRC GTQ DOP JMD TTD PAB SEK NOK DKK ISK PLN CZK HUF RON BGN RSD MKD ALL BAM GEL AMD AZN MDL IRR IQD LBP SYP YER AFN BND MOP MVR BTN KGS TJS TMT AOA MZN ZMW BWP NAD MUR SCR GMD'.split(' ');
const FAV_CUR=['KRW','USD','EUR','RMB','JPY'];
let _dn=null;try{_dn=new Intl.DisplayNames(['ko'],{type:'currency'})}catch(e){}
const curKo=c=>{if(c==='KRW')return '대한민국 원';if(c==='RMB')return '중국 위안 (CNY)';try{const n=_dn&&_dn.of(c);return n&&n!==c?n:''}catch(e){return ''}};
A.curList=()=>CUR_CODES.map(c=>({code:c,ko:curKo(c)}));
A.isCur=c=>CUR_CODES.includes(c);
A.normCur=c=>{c=String(c||'').toUpperCase();if(c==='CNY')c='RMB';return A.isCur(c)?c:null};
A.curLabel=c=>!c||c==='KRW'?'원':c;
A.usedCurs=()=>{const s=new Set();for(const e of A.DB.entries)if(e.cur&&e.cur!=='KRW')s.add(e.cur);for(const f of A.DB.fx)for(const k of Object.keys(f))if(/^[A-Z]{3}$/.test(k)&&f[k]!=null)s.add(k);for(const c of A.DB.fxCurs||[])s.add(c);return s};
A.curOptions=v=>{const list=[...new Set([...FAV_CUR,...A.usedCurs()])];if(v&&!list.includes(v))list.push(v);return list.map(c=>'<option value="'+c+'"'+(c===v?' selected':'')+'>'+esc(A.curLabel(c))+'</option>').join('')+'<option value="__more__">다른 통화 검색…</option>'};
A.setCur=(sel,c)=>{c=c||'KRW';if(![...sel.options].some(o=>o.value===c)){const o=document.createElement('option');o.value=c;o.textContent=A.curLabel(c);sel.insertBefore(o,sel.querySelector('option[value="__more__"]'))}sel.value=c;sel.dataset.prev=c};
A.pickCurrency=function(cb){
  const old=document.getElementById('curpick');if(old)old.remove();
  const ov=document.createElement('div');ov.id='curpick';ov.className='modal';ov.style.zIndex=70;
  ov.innerHTML='<div class="modal-card"><h2>통화 선택</h2><input type="search" id="curq" placeholder="코드나 이름 (예: JPY, 엔, 달러)" autocomplete="off"><div class="curlist"></div><div class="row-end"><button class="btn" type="button" data-x>취소</button></div></div>';
  document.body.appendChild(ov);
  const list=A.curList(),q=ov.querySelector('#curq'),box=ov.querySelector('.curlist');
  const close=c=>{ov.remove();cb(c)};
  const draw=()=>{const s=q.value.trim().toLowerCase(),used=A.usedCurs();
    const rows=list.filter(x=>!s||x.code.toLowerCase().includes(s)||x.ko.toLowerCase().includes(s)).sort((a,b)=>((FAV_CUR.includes(b.code)||used.has(b.code))?1:0)-((FAV_CUR.includes(a.code)||used.has(a.code))?1:0));
    box.innerHTML=rows.length?rows.slice(0,100).map(x=>'<button type="button" class="currow" data-c="'+x.code+'"><b>'+esc(x.code)+'</b><span>'+esc(x.ko)+'</span></button>').join(''):'<p class="hint">찾는 통화가 없습니다.</p>'};
  q.addEventListener('input',draw);
  ov.addEventListener('click',ev=>{if(ev.target===ov||ev.target.closest('[data-x]')){close(null);return}const b=ev.target.closest('[data-c]');if(b)close(b.dataset.c)});
  ov.addEventListener('keydown',ev=>{if(ev.key==='Escape'){ev.stopPropagation();close(null)}});
  draw();q.focus();
};
document.addEventListener('change',ev=>{const s=ev.target;if(!s.classList||!s.classList.contains('cursel'))return;
  if(s.value==='__more__'){const prev=s.dataset.prev||'KRW';A.pickCurrency(code=>{if(!code){s.value=prev;return}A.setCur(s,code);s.dispatchEvent(new Event('change',{bubbles:true}))})}else s.dataset.prev=s.value});
A.groupKeys=id=>{const g=A.DB.groups.find(x=>x.id===id);return g?g.keys.filter(k=>A.COL.get(k)):[]};
A.colMatch=(val,key)=>!val||val===key||(String(val).startsWith('g:')&&A.groupKeys(val.slice(2)).includes(key));
A.colLabel=val=>{if(!val)return '전체';if(String(val).startsWith('g:')){const g=A.DB.groups.find(x=>x.id===val.slice(2));return g?g.name:'그룹'}const c=A.COL.get(val);return c?c.name:val};
A.saveGroup=function(g){
  const now=new Date().toISOString(),ex=g.id&&A.DB.groups.find(x=>x.id===g.id);
  if(ex){const before=ex.name;Object.assign(ex,{name:g.name,keys:g.keys.slice(),upd:now});logChange('그룹 수정: '+(before!==g.name?before+' → ':'')+g.name+' ('+g.keys.length+'곳)')}
  else{A.DB.groups.push({id:uid('g'),name:g.name,keys:g.keys.slice(),upd:now});logChange('그룹 만들기: '+g.name+' ('+g.keys.length+'곳)')}
  commit();
};
A.removeGroup=id=>{const g=A.DB.groups.find(x=>x.id===id);if(!g)return;A.DB.groups=A.DB.groups.filter(x=>x.id!==id);A.DB.deleted.push(id);logChange('그룹 삭제: '+g.name);commit()};
A.suggestGroups=()=>{const m=new Map();for(const c of A.DB.columns){if(c.group==='buy')continue;const w=c.name.split(/\s+/)[0];if(w.length<2||w===c.name)continue;if(!m.has(w))m.set(w,[]);m.get(w).push(c.key)}
  return [...m].filter(([w,ks])=>ks.length>=2&&!A.DB.groups.some(g=>g.name===w||g.name===w+' 그룹')).map(([w,ks])=>({name:w+' 그룹',keys:ks}))};
/* what-if: effect of changes on margin alerts, evaluated at date d */
A.impact=function(list,d){
  const a0=A.alerts(d),tmp=list.map((e,i)=>Object.assign({id:'tmp'+i,at:nowIso()},e));
  A.DB.entries.push(...tmp);A.reindex();let a1;
  try{a1=A.alerts(d)}finally{A.DB.entries=A.DB.entries.filter(x=>!String(x.id).startsWith('tmp'));A.reindex()}
  const key=x=>x.it.id+'|'+x.c.key,was=new Set(a0.neg.map(key));
  return{neg0:a0.neg.length,neg1:a1.neg.length,low0:a0.low.length,low1:a1.low.length,newNeg:a1.neg.filter(x=>!was.has(key(x)))};
};
A.logSheet=rec=>{if(A.readOnly)return;A.DB.sheetLog.push(Object.assign({id:uid('sl'),at:nowIso(),by:A.settings.userName||''},rec));if(A.DB.sheetLog.length>300)A.DB.sheetLog=A.DB.sheetLog.slice(-300);A.touch(1,'단가표 발행 기록')};
A.fxColumns=()=>{const s=new Set(['EUR','RMB','USD']);for(const c of A.usedCurs())s.add(c);return [...s]};

/* display currency (원래 통화 / 원화 환산) */
A.view={krw:!!A.settings.viewKrw};
A.setKrw=b=>{A.view.krw=!!b;A.settings.viewKrw=!!b;A.saveSettings();A.emit()};
A.krwValue=(e,d)=>{if(!A.view.krw||!isNum(e))return null;const c=e.cur||'KRW';if(c==='KRW')return null;const r=A.fxAt(d,c);return r?Math.round(e.v*r):null};
A.krwNote=(e,d)=>{if(!isNum(e)||!e.cur||e.cur==='KRW')return '';const r=A.fxAt(d,e.cur);return r?'≈ '+fmtN(Math.round(e.v*r))+'원 (1 '+e.cur+' = '+fmtN(r)+'원)':'환율 미입력'};

/* thresholds, alerts, favorites */
A.lowMargin=()=>{const v=Number(A.settings.lowMargin);return A.settings.lowMargin===''||A.settings.lowMargin==null||isNaN(v)?10:v};
A.basis=()=>A.settings.basis==='dukwoo'?'dukwoo':'hana';
A.setBasis=b=>{A.settings.basis=b;A.saveSettings()};
A.isBig=(p,n)=>{const th=Number(A.settings.warnPct);if(!(th>0)||!isNum(p)||!isNum(n))return false;if((p.cur||'KRW')!==(n.cur||'KRW'))return true;if(!p.v)return false;return Math.abs((n.v-p.v)/p.v*100)>=th};
A.marginList=(d,basis)=>{const out=[];for(const it of A.DB.items)for(const c of A.DB.columns){if(c.group==='buy')continue;const r=A.marginOf(it,c,d,basis||A.basis());if(r)out.push({it,c,r})}return out};
A.alerts=(d,basis)=>{d=d||today();const all=A.marginList(d,basis),lm=A.lowMargin(),ok=all.filter(x=>!x.r.unknown);
  let fxCells=0;const fxCur=new Set();for(const it of A.DB.items)for(const c of A.DB.columns){const e=A.entryAt(it.id,c.key,d);if(isNum(e)&&e.cur&&e.cur!=='KRW'&&!A.fxAt(d,e.cur)){fxCells++;fxCur.add(e.cur)}}
  return{neg:ok.filter(x=>x.r.m<0),low:ok.filter(x=>x.r.m>=0&&x.r.p<lm),unknown:all.filter(x=>x.r.unknown),fxCells,fxCur:[...fxCur]}};
A.itemMarginNotice=(id,d)=>{const it=A.ITEM.get(id);if(!it)return '';const lm=A.lowMargin(),bad=[];for(const c of A.DB.columns){if(c.group==='buy')continue;const r=A.marginOf(it,c,d,A.basis());if(r&&!r.unknown&&(r.m<0||r.p<lm))bad.push(c.name+' '+(r.m<0?'역마진 ':'저마진 ')+fmtN(Math.round(r.m*100)/100)+'원('+r.p.toFixed(1)+'%)')}return bad.length?'⚠ 이 품목의 마진을 확인하세요: '+bad.join(', '):''};
A.isFav=id=>(A.settings.favs||[]).includes(id);
A.toggleFav=id=>{const s=new Set(A.settings.favs||[]);if(s.has(id))s.delete(id);else s.add(id);A.settings.favs=[...s];A.saveSettings();A.emit()};

A.delta=(p,e)=>{if(!isNum(p)||!isNum(e)||(p.cur||'KRW')!==(e.cur||'KRW'))return null;const d=e.v-p.v;return{d,p:p.v?d/p.v*100:null}};
A.deltaHtml=x=>{if(!x)return '';const cls=x.d>0?'up':x.d<0?'down':'';const s=x.d>0?'+':'';return '<span class="'+cls+' num">'+s+esc(fmtN(x.d))+(x.p!=null?' ('+s+x.p.toFixed(1)+'%)':'')+'</span>'};
A.sameVal=(a,b)=>JSON.stringify(a?[a.v,a.t,a.cur]:null)===JSON.stringify(b?[b.v,b.t,b.cur]:null);
A.isChanged=(e,since)=>!!(e&&since&&!e.base&&e.from>since&&!A.sameVal(A.prevOf(e),e));
A.itemLabel=it=>[it.cat1,it.cat2].filter(Boolean).join(' › ');
A.itemName=it=>it.model||it.cat2||it.cat1;
A.cats=()=>[...new Set(A.DB.items.map(i=>i.cat1))];
A.fxAt=(d,cur)=>{if(!cur||cur==='KRW')return 1;let r=null;for(const f of [...A.DB.fx].sort((a,b)=>a.date<b.date?-1:1)){if(f.date<=d&&f[cur]>0)r=f[cur]}return r};
A.matchItem=(it,q,cat)=>{if(cat&&it.cat1!==cat)return false;if(!q)return true;q=q.toLowerCase().replace(/\s+/g,'');return (it.model+' '+it.cat1+' '+it.cat2).toLowerCase().replace(/\s+/g,'').includes(q)};
A.costAt=(it,d,basis)=>{
  const order=basis==='dukwoo'?['dukwoo','hana']:['hana','dukwoo'];
  const buys=A.DB.columns.filter(c=>c.group==='buy').map(c=>c.key);
  for(const k of order.concat(buys.filter(k=>!order.includes(k)))){if(!A.COL.get(k))continue;const e=A.entryAt(it.id,k,d);if(isNum(e)){const r=A.fxAt(d,e.cur||'KRW');return{e,krw:r==null?null:e.v*r,needFx:r==null}}}
  return null;
};
A.marginOf=(it,c,d,basis)=>{
  const s=A.entryAt(it.id,c.key,d);if(!isNum(s))return null;
  const cost=A.costAt(it,d,basis);const sr=A.fxAt(d,s.cur||'KRW');
  if(!cost||cost.needFx||sr==null)return{s,cost,unknown:true};
  const sk=s.v*sr,m=sk-cost.krw;return{s,cost,m,p:sk?m/sk*100:0};
};

/* ---------- mutations ---------- */
A.addEntry=e=>{e.id=uid('e');e.at=nowIso();if(A.settings.userName)e.by=A.settings.userName;if(batchTag){e.batch=batchTag.id;e.bl=batchTag.label}A.DB.entries.push(e);A.reindex();const it=A.ITEM.get(e.item),c=A.COL.get(e.col);A.touch(1,'단가 기록: '+(it?A.itemName(it):'')+(c?' · '+c.name:''))};
A.updateEntry=function(id,patch,opts){
  const e=A.DB.entries.find(x=>x.id===id);if(!e||e.base)return false;
  const lab=x=>A.eText({v:x.v,t:x.t,cur:x.cur}),diffs=[];
  const nv={v:patch.v!==undefined?patch.v:e.v,t:patch.t!==undefined?patch.t:e.t,cur:patch.cur||e.cur||'KRW'};
  const nFrom=patch.from||e.from,nReason=patch.reason!=null?patch.reason:(e.reason||''),nNote=patch.note!=null?patch.note:(e.note||'');
  if(nFrom!==e.from)diffs.push('적용일 '+e.from+' → '+nFrom);
  if(JSON.stringify([nv.v,nv.t,nv.cur])!==JSON.stringify([e.v,e.t,e.cur||'KRW']))diffs.push('단가 '+lab(e)+' → '+lab(nv));
  if(nReason!==(e.reason||''))diffs.push('사유 '+(e.reason||'없음')+' → '+(nReason||'없음'));
  if(nNote!==(e.note||''))diffs.push('비고 수정');
  if(!diffs.length)return false;
  Object.assign(e,{from:nFrom,v:nv.v,t:nv.t,cur:nv.cur,reason:nReason,note:nNote});
  e.upd=new Date().toISOString();(e.edits=e.edits||[]).push({at:nowIso(),by:A.settings.userName||'',text:diffs.join(', ')});
  A.reindex();A.touch(1,(opts&&opts.label)||'기록 수정');return true;
};
A.deleteEntry=(id,quiet)=>{const e=A.DB.entries.find(x=>x.id===id);A.DB.entries=A.DB.entries.filter(x=>x.id!==id);A.DB.deleted.push(id);if(e&&!quiet){const it=A.ITEM.get(e.item),c=A.COL.get(e.col);logChange('기록 삭제: '+(it?A.itemName(it):'')+' · '+(c?c.name:'')+' '+e.from+' '+A.eText(e))}A.reindex();A.touch(1,'기록 삭제')};
A.addItem=it=>{
  it.id=uid('i');let idx=-1;
  A.DB.items.forEach((x,k)=>{if(x.cat1===it.cat1&&x.cat2===it.cat2)idx=k});
  if(idx<0)A.DB.items.forEach((x,k)=>{if(x.cat1===it.cat1)idx=k});
  if(idx<0)A.DB.items.push(it);else A.DB.items.splice(idx+1,0,it);
  A.reindex();A.touch();return it;
};
A.addColumn=c=>{c.key=uid('c');if(c.group==='buy'){let k=0;A.DB.columns.forEach((x,i)=>{if(x.group==='buy')k=i+1});A.DB.columns.splice(k,0,c)}else A.DB.columns.push(c);A.reindex();A.touch();return c};

/* merge two versions: entries are immutable and carry unique ids, so a union minus tombstones is safe */
/* keep items of one 대구분/소구분 together, in first-seen order */
function regroup(items){
  const o1=new Map(),o2=new Map();items.forEach((it,i)=>{if(!o1.has(it.cat1))o1.set(it.cat1,i);const k=it.cat1+'\u0000'+it.cat2;if(!o2.has(k))o2.set(k,i)});
  return items.map((it,i)=>({it,i})).sort((a,b)=>(o1.get(a.it.cat1)-o1.get(b.it.cat1))||(o2.get(a.it.cat1+'\u0000'+a.it.cat2)-o2.get(b.it.cat1+'\u0000'+b.it.cat2))||(a.i-b.i)).map(x=>x.it);
}
const sortCols=cols=>cols.map((c,i)=>({c,i})).sort((a,b)=>((a.c.group==='buy'?0:1)-(b.c.group==='buy'?0:1))||(a.i-b.i)).map(x=>x.c);
const insertInGroup=(items,it)=>{let idx=-1;items.forEach((x,k)=>{if(x.cat1===it.cat1&&x.cat2===it.cat2)idx=k});if(idx<0)items.forEach((x,k)=>{if(x.cat1===it.cat1)idx=k});if(idx<0)items.push(it);else items.splice(idx+1,0,it)};
/* merge two versions: entries are immutable with unique ids (union minus tombstones);
   items/columns carry 'upd' so the most recent rename/delete/restore wins */
A.mergeDb=function(local,remote){
  local=normDb(clone(local));remote=normDb(clone(remote));
  const tomb=new Set([...local.deleted,...remote.deleted]);
  const byId=(mine,theirs)=>{const m=new Map();for(const o of theirs)m.set(o.id,o);for(const o of mine)m.set(o.id,o);return [...m.values()].filter(o=>!tomb.has(o.id))};
  const pickE=(mine,theirs)=>{const m=new Map();for(const o of theirs)m.set(o.id,o);for(const o of mine){const x=m.get(o.id);m.set(o.id,!x||(o.upd||'')>=(x.upd||'')?o:x)}return [...m.values()].filter(o=>!tomb.has(o.id))};
  const newer=(a,b)=>!a?b:!b?a:((b.upd||'')>(a.upd||'')?b:a);
  const pool=(act,rem,key)=>{const m=new Map();for(const o of act){delete o.removed;m.set(o[key],o)}for(const o of rem){o.removed=true;m.set(o[key],o)}return m};
  const mergeSet=(la,lr,ra,rr,key,place)=>{
    const L=pool(la,lr,key),R=pool(ra,rr,key);const ids=new Set([...L.keys(),...R.keys()]);const chosen=new Map();
    for(const id of ids){if(tomb.has(id))continue;chosen.set(id,newer(L.get(id),R.get(id)))}
    const act=[],done=new Set();
    for(const o of la){const c=chosen.get(o[key]);if(c&&!c.removed){act.push(c);done.add(o[key])}}
    for(const [id,c] of chosen){if(!c.removed&&!done.has(id))place(act,c)}
    return{act,rem:[...chosen.values()].filter(c=>c.removed)};
  };
  const I=mergeSet(local.items,local.removedItems,remote.items,remote.removedItems,'id',insertInGroup);
  const C=mergeSet(local.columns,local.removedColumns,remote.columns,remote.removedColumns,'key',(a,c)=>a.push(c));
  return{columns:sortCols(C.act),items:regroup(I.act),entries:pickE(local.entries,remote.entries),fx:byId(local.fx,remote.fx),deleted:[...tomb],removedItems:I.rem,removedColumns:C.rem,log:byId(local.log,remote.log),fxCurs:[...new Set([...local.fxCurs,...remote.fxCurs])],groups:pickE(local.groups,remote.groups),sheetLog:byId(local.sheetLog,remote.sheetLog)};
};

/* ---------- structure edits: rename, move, delete (kept for restore), restore ---------- */
const updStamp=()=>new Date().toISOString();
let lastLog='';const logChange=text=>{lastLog=text;A.DB.log.push({id:uid('l'),at:nowIso(),by:A.settings.userName||'',text})};
A.logChange=logChange;
const itemPath=it=>[it.cat1,it.cat2,it.model].filter(Boolean).join(' › ');
const commit=()=>{A.reindex();A.touch(1,lastLog)};
A.findDupItem=(cat1,cat2,model,exceptId)=>A.DB.items.find(x=>x.id!==exceptId&&x.cat1===cat1&&x.cat2===cat2&&x.model===model);
A.updateItem=function(id,patch){
  const it=A.ITEM.get(id);if(!it)return;const before=itemPath(it);const moved=patch.cat1!==it.cat1||patch.cat2!==it.cat2;
  Object.assign(it,patch,{upd:updStamp()});
  if(moved){A.DB.items.splice(A.DB.items.indexOf(it),1);insertInGroup(A.DB.items,it)}
  A.DB.items=regroup(A.DB.items);logChange('품목 수정: '+before+' → '+itemPath(it));commit();
};
A.removeItem=function(id){const k=A.DB.items.findIndex(x=>x.id===id);if(k<0)return;const it=A.DB.items.splice(k,1)[0];it.removed=true;it.upd=updStamp();A.DB.removedItems.push(it);logChange('품목 삭제: '+itemPath(it));commit()};
A.restoreItem=function(id){const k=A.DB.removedItems.findIndex(x=>x.id===id);if(k<0)return;const it=A.DB.removedItems.splice(k,1)[0];
  if(A.findDupItem(it.cat1,it.cat2,it.model))throw new Error('같은 이름의 품목이 이미 있어 복원할 수 없습니다. 먼저 그 품목의 이름을 바꾸세요.');
  delete it.removed;it.upd=updStamp();insertInGroup(A.DB.items,it);logChange('품목 복원: '+itemPath(it));commit()};
A.updateColumn=function(key,patch){const c=A.COL.get(key);if(!c)return;const before=c.name+(c.group==='buy'?'(매입)':'(판매)');Object.assign(c,patch,{upd:updStamp()});A.DB.columns=sortCols(A.DB.columns);logChange('거래처 수정: '+before+' → '+c.name+(c.group==='buy'?'(매입)':'(판매)'));commit()};
A.removeColumn=function(key){const k=A.DB.columns.findIndex(x=>x.key===key);if(k<0)return;const c=A.DB.columns.splice(k,1)[0];c.removed=true;c.upd=updStamp();A.DB.removedColumns.push(c);logChange('거래처 삭제: '+c.name);commit()};
A.restoreColumn=function(key){const k=A.DB.removedColumns.findIndex(x=>x.key===key);if(k<0)return;const c=A.DB.removedColumns[k];
  if(A.DB.columns.some(x=>x.name===c.name))throw new Error('같은 이름의 거래처가 이미 있어 복원할 수 없습니다. 먼저 그 거래처의 이름을 바꾸세요.');
  A.DB.removedColumns.splice(k,1);delete c.removed;c.upd=updStamp();A.DB.columns.push(c);A.DB.columns=sortCols(A.DB.columns);logChange('거래처 복원: '+c.name);commit()};
A.renameCat=function(level,cat1,oldName,newName){
  newName=norm(newName);if(!newName&&level===1)throw new Error('대구분 이름을 입력하세요.');
  const all=A.DB.items.concat(A.DB.removedItems);const hit=all.filter(it=>level===1?it.cat1===oldName:(it.cat1===cat1&&it.cat2===oldName));
  for(const it of hit){const nc1=level===1?newName:it.cat1,nc2=level===2?newName:it.cat2;if(!it.removed&&A.findDupItem(nc1,nc2,it.model,it.id)&&!hit.includes(A.findDupItem(nc1,nc2,it.model,it.id)))throw new Error('합치면 같은 모델이 겹칩니다: '+(it.model||'(모델 없음)'))}
  for(const it of hit){if(level===1)it.cat1=newName;else it.cat2=newName;it.upd=updStamp()}
  A.DB.items=regroup(A.DB.items);logChange((level===1?'대구분':'소구분')+' 이름 변경: '+(level===2?cat1+' › ':'')+(oldName||'(소구분 없음)')+' → '+(newName||'(소구분 없음)')+' (품목 '+hit.length+'개)');commit();return hit.length;
};

A.replaceDb=(db,opts)=>{A.DB=normDb(db);A.reindex();opts=opts||{};A.dirty=opts.dirty||0;if(opts.overwrite)A.sync.overwrite=true;A.persist();A.resetUndo();A.emit()};

/* ---------- excel ---------- */
let xlsxP=null;
A.loadXlsx=()=>{if(window.XLSX)return Promise.resolve(window.XLSX);if(!xlsxP)xlsxP=new Promise((res,rej)=>{const s=document.createElement('script');s.src=XLSX_SRC;s.onload=()=>res(window.XLSX);s.onerror=()=>{xlsxP=null;rej(new Error('엑셀 도구를 불러오지 못했습니다. 인터넷 연결을 확인하세요.'))};document.head.appendChild(s)});return xlsxP};
A.parseWorkbook=function(X,wb){
  const sheetName=wb.SheetNames[0],ws=wb.Sheets[sheetName];
  if(!ws||!ws['!ref'])throw new Error('첫 번째 시트가 비어 있습니다.');
  const ref=X.utils.decode_range(ws['!ref']);
  const cell=(r,c)=>ws[X.utils.encode_cell({r,c})];
  const val=(r,c)=>{const x=cell(r,c);if(!x||x.v==null)return null;if(typeof x.v==='string'&&!x.v.trim())return null;return x.v};
  const cols=[];
  for(let c=3;c<=ref.e.c;c++){const h=val(0,c);if(h==null)continue;const raw=String(h);const name=norm(raw).replace(/ ?매입가$/,'');cols.push({key:KNOWN[name]||('c_'+name.replace(/\s+/g,'')),name,group:/매입가/.test(raw)?'buy':'sell',c})}
  if(!cols.length)throw new Error('1행에서 거래처 이름(D열부터)을 찾지 못했습니다. 단가표 양식이 맞는지 확인하세요.');
  const m=/(\d{4})[.\-\/](\d{1,2})/.exec(sheetName);
  const base=m?m[1]+'-'+pad(m[2])+'-01':today();
  const items=[],entries=[];let cat1='',cat2='',n=0;
  for(let r=1;r<=ref.e.r;r++){
    const a=val(r,0),b=val(r,1),md=val(r,2);
    const vals=cols.map(c=>val(r,c.c));
    if(a==null&&b==null&&md==null&&vals.every(v=>v==null))continue;
    if(a!=null){cat1=norm(a);cat2=''}
    if(b!=null)cat2=norm(b);
    const iid='i'+r;
    items.push({id:iid,cat1,cat2,model:md==null?'':norm(md)});
    cols.forEach((c,k)=>{
      const v=vals[k];if(v==null)return;n++;
      const cm=cell(r,c.c).c;const ct=cm?cm.map(x=>x.t||'').join(' '):'';
      const cmm=/\b(EUR|RMB|CNY|USD|JPY)\b/i.exec(ct);let cur=cmm?cmm[1].toUpperCase():'KRW';if(cur==='CNY')cur='RMB';
      const e={id:'b'+n,item:iid,col:c.key,from:base,at:base+'T00:00:00',reason:'기준 단가 ('+sheetName+' 단가표)',base:true,cur};
      if(typeof v==='number'){e.v=v;e.t=null}else{e.v=null;e.t=norm(v)}
      entries.push(e);
    });
  }
  return{db:{columns:cols.map(({key,name,group})=>({key,name,group})),items,entries,fx:[],deleted:[]},sheetName,base,count:entries.length};
};
A.importXlsxBuffer=async function(buf){const X=await A.loadXlsx();const wb=X.read(buf,{type:'array'});return A.parseWorkbook(X,wb)};
A.exportXlsx=async function(d){
  const X=await A.loadXlsx();const cols=A.DB.columns,its=A.DB.items;
  const aoa=[['대구분','소구분','MODEL'].concat(cols.map(c=>c.name+(c.group==='buy'?'\n매입가':'')))];
  const notes=[];
  its.forEach((it,r)=>{const row=[it.cat1,it.cat2,it.model];cols.forEach((c,k)=>{const e=A.entryAt(it.id,c.key,d);const cv=A.krwValue(e,d);row.push(e?(isNum(e)?(cv!=null?cv:e.v):(e.t||null)):null);if(isNum(e)&&e.cur&&e.cur!=='KRW'&&cv==null)notes.push([r+1,k+3,e.cur])});aoa.push(row)});
  const merges=[];
  for(let r=0;r<its.length;){let n=1;while(r+n<its.length&&its[r+n].cat1===its[r].cat1)n++;if(n>1)merges.push({s:{r:r+1,c:0},e:{r:r+n,c:0}});r+=n}
  for(let r=0;r<its.length;){let n=1;while(r+n<its.length&&its[r+n].cat1===its[r].cat1&&its[r+n].cat2===its[r].cat2)n++;if(n>1)merges.push({s:{r:r+1,c:1},e:{r:r+n,c:1}});r+=n}
  const ws=X.utils.aoa_to_sheet(aoa);ws['!merges']=merges;
  notes.forEach(([r,c,t])=>{const a=X.utils.encode_cell({r,c});if(ws[a])ws[a].c=[{a:'단가 관리',t}]});
  for(let r=1;r<=its.length;r++)for(let c=3;c<3+cols.length;c++){const a=X.utils.encode_cell({r,c});if(ws[a]&&ws[a].t==='n')ws[a].z='#,##0.##'}
  ws['!cols']=[{wch:12},{wch:12},{wch:20}].concat(cols.map(()=>({wch:10})));
  const wb=X.utils.book_new();X.utils.book_append_sheet(wb,ws,d.slice(0,7).replace('-','.'));
  A.saveBlob('단가표_'+d+'.xlsx',new Blob([X.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
};
/* ---------- reports: Excel + print-to-PDF ---------- */
const curName=c=>!c||c==='KRW'?'원':c;
const numOrText=e=>!e?null:(isNum(e)?e.v:(e.t||null));
const stampStr=()=>{const d=new Date();return isoD(d)+' '+pad(d.getHours())+':'+pad(d.getMinutes())};
/* generic workbook: sheets=[{name,title,subtitle,head,rows,widths,pct:[colIdx],num:[colIdx]}] */
A.exportBook=async function(filename,sheets){
  const X=await A.loadXlsx();const wb=X.utils.book_new();
  for(const s of sheets){
    const aoa=[[s.title],[s.subtitle+' · 출력 '+stampStr()],[],s.head].concat(s.rows);
    const ws=X.utils.aoa_to_sheet(aoa);const top=4;
    for(let r=0;r<s.rows.length;r++)for(let c=0;c<s.head.length;c++){const a=X.utils.encode_cell({r:r+top,c});const cell=ws[a];if(!cell||cell.t!=='n')continue;if((s.pct||[]).includes(c))cell.z='0.0"%"';else if((s.num||[]).includes(c))cell.z='#,##0.##'}
    if(s.merges)ws['!merges']=s.merges.map(m=>({s:{r:m.s.r+top,c:m.s.c},e:{r:m.e.r+top,c:m.e.c}}));
    ws['!cols']=(s.widths||s.head.map(()=>12)).map(w=>({wch:w}));
    X.utils.book_append_sheet(wb,ws,s.name.slice(0,31));
  }
  A.saveBlob(filename,new Blob([X.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
};
/* print-to-PDF: renders a clean report into #print and opens the print dialog */
A.printReport=function(title,subtitle,bodyHtml){
  let el=document.getElementById('print');if(!el){el=document.createElement('div');el.id='print';document.body.appendChild(el)}
  el.innerHTML='<header><h1>'+esc(title)+'</h1><div>'+esc(subtitle)+'</div><div class="pr-meta">출력 '+esc(stampStr())+' · 단가 관리</div></header>'+bodyHtml;
  document.body.classList.add('printing');
  const done=()=>{document.body.classList.remove('printing');window.removeEventListener('afterprint',done)};
  window.addEventListener('afterprint',done);
  setTimeout(()=>{try{window.print()}catch(e){done();A.toast('이 브라우저에서는 인쇄 창을 열 수 없습니다. 브라우저 메뉴의 인쇄를 이용하세요.')}},60);
};
const prTable=(head,rows,opts)=>{opts=opts||{};return '<table class="pr"><thead><tr>'+head.map((h,i)=>'<th'+((opts.right||[]).includes(i)?' class="r"':'')+'>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+(rows.length?rows.map(r=>'<tr>'+r.map((v,i)=>'<td'+((opts.right||[]).includes(i)?' class="r"':'')+'>'+(v&&v.html!=null?v.html:esc(v==null?'':v))+'</td>').join('')+'</tr>').join(''):'<tr><td colspan="'+head.length+'" class="empty-row">내용이 없습니다.</td></tr>')+'</tbody></table>'};
const pctTxt=p=>p==null?'':(p>0?'+':'')+p.toFixed(1)+'%';
const dTxt=d=>d==null?'':(d>0?'+':'')+fmtN(d);

/* 단가표 (grid, merged categories) */
function gridPrint(d,cellsFn,cols,opts){
  opts=opts||{};const its=opts.items||A.DB.items;
  let h='<table class="pr prgrid"><thead><tr><th>대구분</th><th>소구분</th><th>MODEL</th>'+cols.map(c=>'<th class="'+(c.group==='buy'?'buy':'')+'">'+esc(c.name)+'</th>').join('')+'</tr></thead><tbody>';
  for(let r=0;r<its.length;r++){const it=its[r];h+='<tr'+((r===0||its[r-1].cat1!==it.cat1)&&r>0?' class="grp"':'')+'>';
    const new1=r===0||its[r-1].cat1!==it.cat1,new2=new1||its[r-1].cat2!==it.cat2;
    h+='<td class="cat'+(new1?'':' cont')+'">'+(new1?esc(it.cat1):'')+'</td><td class="cat'+(new2?'':' cont')+'">'+(new2?esc(it.cat2):'')+'</td>';
    h+='<td class="md">'+esc(it.model)+'</td>'+cols.map(c=>cellsFn(it,c)).join('')+'</tr>'}
  return h+'</tbody></table>';
}
A.printPriceTable=function(d){
  const cols=A.DB.columns.filter(c=>A.DB.items.some(it=>A.hasVal(A.entryAt(it.id,c.key,d))));
  const since=addDays(d,-30);
  const html=gridPrint(d,(it,c)=>{const e=A.entryAt(it.id,c.key,d);const txt=e&&!isNum(e)&&e.t&&e.t.length>3;return '<td class="'+(txt?'txt':'r')+(c.group==='buy'?' buy':'')+(A.isChanged(e,since)?' chg':'')+'">'+A.eHtml(e,d)+'</td>'},cols);
  A.printReport('단가표',d+' 기준 · 거래가 있는 거래처 '+cols.length+'곳 · 음영: 최근 30일 변경 · 표시 없는 금액은 원',html+'<p class="pr-note">'+(A.view.krw?'외화 단가는 '+d+' 기준 환율로 원화 환산한 값입니다.':'EUR·RMB 등 통화가 표시된 금액은 외화 단가입니다.')+'</p>');
};

/* 마진표 */
function marginRows(d,basis){
  const out=[];
  for(const it of A.DB.items)for(const c of A.DB.columns){if(c.group==='buy')continue;const r=A.marginOf(it,c,d,basis);if(!r)continue;out.push({it,c,r})}
  return out;
}
A.exportMarginXlsx=function(d,basis){
  const rows=marginRows(d,basis).map(({it,c,r})=>[it.cat1,it.cat2,it.model,c.name,r.s.v,curName(r.s.cur),r.cost?A.COL.get(r.cost.e.col).name:'',r.cost?r.cost.e.v:null,r.cost?curName(r.cost.e.cur):'',r.cost&&!r.cost.needFx?Math.round(r.cost.krw*100)/100:null,r.unknown?null:Math.round(r.m*100)/100,r.unknown?null:Math.round(r.p*10)/10,r.unknown?(r.cost?'환율 필요':'매입가 없음'):'']);
  return A.exportBook('마진표_'+d+'.xlsx',[{name:'마진표',title:'마진표',subtitle:d+' 기준 · 마진 = 판매가 − 원가(원 환산) · 마진율은 판매가 대비',head:['대구분','소구분','모델','거래처','판매가','통화','원가 출처','매입가','매입 통화','원가(원)','마진(원)','마진율(%)','비고'],rows,widths:[12,12,18,12,10,6,10,10,8,11,11,10,12],num:[4,7,9,10],pct:[11]}]);
};
A.printMargin=function(d,basis){
  const rowsAll=marginRows(d,basis);const sells=A.DB.columns.filter(c=>c.group!=='buy'&&rowsAll.some(x=>x.c===c));
  const items=A.DB.items.filter(it=>rowsAll.some(x=>x.it===it));
  const cols=[{key:'__cost',name:'원가(원)',group:'buy'}].concat(sells);
  const html=gridPrint(d,(it,c)=>{
    if(c.key==='__cost'){const cost=A.costAt(it,d,basis);if(!cost)return '<td class="r buy">-</td>';if(cost.needFx)return '<td class="r buy">'+esc(cost.e.cur)+' 환율 필요</td>';return '<td class="r buy">'+esc(fmtN(Math.round(cost.krw*100)/100))+'</td>'}
    const r=A.marginOf(it,c,d,basis);if(!r)return '<td></td>';if(r.unknown)return '<td class="r">?</td>';
    return '<td class="r'+(r.m<0?' neg':r.p<A.lowMargin()?' low':'')+'">'+esc(fmtN(Math.round(r.m*100)/100))+'<br><small>'+r.p.toFixed(1)+'%</small></td>'},cols,{items});
  A.printReport('마진표',d+' 기준 · 칸 위: 마진(원), 아래: 마진율(판매가 대비) · 원가 기준 '+(basis==='dukwoo'?'덕우 우선':'하나테크 우선'),html+'<p class="pr-note">?: 원가 또는 환율이 없어 계산하지 못한 칸. 빨강: 역마진.</p>');
};

/* 변경 이력 */
const histHead=['적용일','대구분','소구분','모델','거래처','구분','이전 단가','변경 단가','통화','증감','증감률(%)','사유','비고','입력 일시','입력자'];
const histRow=({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return [e.from,it.cat1,it.cat2,it.model,c.name,c.group==='buy'?'매입':'판매',numOrText(p),numOrText(e),curName(e.cur),x?x.d:null,x&&x.p!=null?Math.round(x.p*10)/10:null,e.base?'기준 단가':(e.reason||''),e.note||'',shortDT(e.at),e.by||'']};
A.exportHistoryXlsx=function(rows,desc){return A.exportBook('단가이력_'+today()+'.xlsx',[{name:'변경 이력',title:'단가 변경 이력',subtitle:desc,head:histHead,rows:rows.map(histRow),widths:[11,12,12,18,12,6,10,10,6,9,9,16,20,16,10],num:[6,7,9],pct:[10]}])};
A.printHistory=function(rows,desc){
  A.printReport('단가 변경 이력',desc,prTable(['적용일','구분','모델','거래처','이전','변경','증감','사유','비고'],rows.map(({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return [e.from,A.itemLabel(it),it.model,c.name,p?A.eText(p):'',{html:'<b>'+esc(A.eText(e))+'</b>'},{html:x?'<span class="'+(x.d>0?'up':x.d<0?'down':'')+'">'+esc(dTxt(x.d)+(x.p!=null?' ('+pctTxt(x.p)+')':''))+'</span>':''},e.base?'기준 단가':(e.reason||''),e.note||'']}),{right:[4,5,6]}));
};

/* 두 날짜 비교 */
A.exportDiffXlsx=function(diffs,Ad,Bd,desc){
  const rows=diffs.map(d=>[d.it.cat1,d.it.cat2,d.it.model,d.c.name,d.c.group==='buy'?'매입':'판매',numOrText(d.a),numOrText(d.b),curName((d.b||d.a||{}).cur),d.x?d.x.d:null,d.x&&d.x.p!=null?Math.round(d.x.p*10)/10:null,!A.hasVal(d.a)?'신규':!A.hasVal(d.b)?'거래 없음으로':'',d.b?d.b.from:'',d.b&&!d.b.base?(d.b.reason||''):'']);
  return A.exportBook('단가비교_'+Ad+'_'+Bd+'.xlsx',[{name:'두 날짜 비교',title:'단가 비교 '+Ad+' → '+Bd,subtitle:desc,head:['대구분','소구분','모델','거래처','구분',Ad,Bd,'통화','증감','증감률(%)','변화','최근 적용일','사유'],rows,widths:[12,12,18,12,6,11,11,6,9,9,12,11,16],num:[5,6,8],pct:[9]}]);
};
A.printDiff=function(diffs,Ad,Bd,desc){
  A.printReport('단가 비교 '+Ad+' → '+Bd,desc,prTable(['구분','모델','거래처',Ad,Bd,'증감','최근 적용일','사유'],diffs.map(d=>[A.itemLabel(d.it),d.it.model,d.c.name,A.eText(d.a),{html:'<b>'+esc(A.eText(d.b))+'</b>'},{html:d.x?'<span class="'+(d.x.d>0?'up':d.x.d<0?'down':'')+'">'+esc(dTxt(d.x.d)+(d.x.p!=null?' ('+pctTxt(d.x.p)+')':''))+'</span>':(!A.hasVal(d.a)?'신규':!A.hasVal(d.b)?'거래 없음으로':'')},d.b?d.b.from:'',d.b&&!d.b.base?(d.b.reason||''):'']),{right:[3,4,5]}));
};
const expRow='<div class="row exp" data-exp><button class="btn small" type="button" data-x="xlsx">엑셀로 내보내기</button><button class="btn small" type="button" data-x="pdf">PDF로 저장</button></div>';
const expErr=e=>A.toast(e&&e.message?e.message:'내보내기에 실패했습니다.');

A.saveBlob=(name,blob)=>{const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.rel='noopener';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),30000)};
A.downloadText=(name,text)=>A.saveBlob(name,new Blob([text],{type:'application/json'}));
A.pack=()=>JSON.stringify({app:APP,schema:SCHEMA,savedAt:nowIso(),db:A.DB});
A.unpack=text=>{let o;try{o=JSON.parse(text)}catch(e){return null}if(o&&o.app===APP&&validDb(o.db)){if(o.kind==='share'&&!A.readOnly)return null;return normDb(o.db)}return null};

/* ---------- Google Drive ---------- */
const D={token:null,exp:0};
const gisReady=()=>!!(window.google&&google.accounts&&google.accounts.oauth2);
let gisP=null;
A.loadGis=()=>{if(gisReady())return Promise.resolve();if(!gisP)gisP=new Promise((res,rej)=>{const s=document.createElement('script');s.src='https://accounts.google.com/gsi/client';s.async=true;s.onload=()=>res();s.onerror=()=>{gisP=null;rej(new Error('구글 로그인 도구를 불러오지 못했습니다. 인터넷 연결을 확인하세요.'))};document.head.appendChild(s)});return gisP};
function cachedToken(){if(D.token&&Date.now()<D.exp-60000)return D.token;const s=LS.get(tkey());if(s&&Date.now()<s.exp-60000){D.token=s.t;D.exp=s.exp;return D.token}return null}
A.hasToken=()=>!!cachedToken();
/* must be called directly from a click so the Google popup is allowed */
function getToken(interactive){
  const t=cachedToken();if(t)return Promise.resolve(t);
  if(!interactive)return Promise.reject(new Error('구글 로그인이 필요합니다.'));
  if(!A.settings.clientId)return Promise.reject(new Error('설정에서 구글 클라이언트 ID를 먼저 입력하세요. (README의 설정 방법 참고)'));
  if(!gisReady())return A.loadGis().then(()=>Promise.reject(new Error('구글 로그인 준비가 끝났습니다. 버튼을 한 번 더 눌러 주세요.')));
  return new Promise((res,rej)=>{
    const c=google.accounts.oauth2.initTokenClient({client_id:A.settings.clientId.trim(),scope:A.readOnly?SCOPE_RO:SCOPE_RW,
      callback:r=>{if(r.error){rej(new Error('구글 로그인 실패: '+(r.error_description||r.error)));return}D.token=r.access_token;D.exp=Date.now()+(Number(r.expires_in)||3600)*1000;LS.set(tkey(),{t:D.token,exp:D.exp});A.fetchUser();res(D.token)},
      error_callback:e=>rej(new Error(e&&e.type==='popup_closed'?'구글 로그인 창이 닫혔습니다.':e&&e.type==='popup_failed_to_open'?'로그인 창이 차단됐습니다. 브라우저의 팝업 차단을 해제하세요.':'구글 로그인에 실패했습니다.'))});
    c.requestAccessToken({prompt:''});
  });
}
A.fetchUser=()=>{if(A.settings.userName)return;api('GET','https://www.googleapis.com/drive/v3/about?fields=user(displayName)').then(r=>{const n=r&&r.user&&r.user.displayName;if(n&&!A.settings.userName){A.settings.userName=n;A.saveSettings();A.emit()}}).catch(()=>{})};
A.signOut=()=>{const t=D.token;D.token=null;D.exp=0;LS.del(tkey());if(t&&gisReady())google.accounts.oauth2.revoke(t,()=>{});A.emit()};
async function api(method,url,body,headers,type){
  const t=cachedToken();if(!t)throw new Error('구글 로그인이 만료됐습니다. 다시 로그인하세요.');
  const r=await fetch(url,{method,headers:Object.assign({Authorization:'Bearer '+t},headers||{}),body});
  if(r.status===401){D.token=null;LS.del(tkey());throw new Error('구글 로그인이 만료됐습니다. 다시 로그인하세요.')}
  if(!r.ok){let m='';try{const j=await r.json();m=j.error&&j.error.message||''}catch(e){}
    if(r.status===403&&/has not been used|disabled/i.test(m))throw new Error('Google Cloud에서 Google Drive API가 꺼져 있습니다. README의 2단계를 확인하세요.');
    if(r.status===404)throw Object.assign(new Error('드라이브에서 파일이나 폴더를 찾지 못했습니다.'),{status:404});
    throw new Error('드라이브 오류 ('+r.status+') '+m)}
  if(type==='text')return r.text();if(type==='buf')return r.arrayBuffer();return r.json();
}
const qs=s=>String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
const DRV='https://www.googleapis.com/drive/v3/files',UP='https://www.googleapis.com/upload/drive/v3/files';
async function find(q,order){const p=new URLSearchParams({q,fields:'files(id,name,mimeType,modifiedTime,webViewLink,parents)',orderBy:order||'modifiedTime desc',pageSize:'50',supportsAllDrives:'true',includeItemsFromAllDrives:'true'});return (await api('GET',DRV+'?'+p)).files||[]}
async function mkFolder(name,parent){return api('POST',DRV+'?fields=id&supportsAllDrives=true',JSON.stringify({name,mimeType:FOLDER,parents:[parent]}),{'Content-Type':'application/json'})}
async function upload(meta,content,fileId){
  const b='danga'+Math.random().toString(36).slice(2);
  const body='--'+b+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+JSON.stringify(meta)+'\r\n--'+b+'\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n'+content+'\r\n--'+b+'--';
  return api(fileId?'PATCH':'POST',UP+(fileId?'/'+fileId:'')+'?uploadType=multipart&supportsAllDrives=true&fields=id,name,modifiedTime',body,{'Content-Type':'multipart/related; boundary='+b});
}
async function subFolder(parent,name,create){const f=(await find("'"+qs(parent)+"' in parents and name = '"+qs(name)+"' and mimeType = '"+FOLDER+"' and trashed = false"))[0];if(f)return f.id;if(!create)return null;return (await mkFolder(name,parent)).id}
async function resolveFolder(create,pathStr){
  const raw=(pathStr!=null?pathStr:(A.settings.path||'')).trim();
  const m=raw.match(/folders\/([A-Za-z0-9_-]{10,})/);if(m)return{id:m[1],label:'폴더 링크'};
  if(/^[A-Za-z0-9_-]{25,}$/.test(raw))return{id:raw,label:'폴더 ID'};
  const segs=raw.split('/').map(s=>s.trim()).filter(Boolean);if(segs[0]==='내 드라이브'||segs[0]==='My Drive')segs.shift();
  let cur='root';
  for(let k=0;k<segs.length;k++){const id=await subFolder(cur,segs[k],create);if(!id)return{id:null,missing:segs.slice(k),found:segs.slice(0,k)};cur=id}
  return{id:cur,label:'내 드라이브'+(segs.length?' / '+segs.join(' / '):'')};
}
A.activePath=()=>A.readOnly?(A.settings.sharePath||''):(A.settings.path||'');
async function ensureFolderFor(ro,create){
  const path=ro?(A.settings.sharePath||''):(A.settings.path||''),idK=ro?'shareFolderId':'folderId',forK=ro?'shareFolderFor':'folderFor';
  if(A.settings[idK]&&A.settings[forK]===path)return A.settings[idK];
  const r=await resolveFolder(create,path);if(!r.id)return null;
  A.settings[idK]=r.id;A.settings[forK]=path;A.saveSettings();return r.id;
}
const ensureFolder=create=>ensureFolderFor(A.readOnly,create);
const mainName=()=>(A.settings.prefix||'단가DB').trim()+(A.readOnly?'_영업용':'')+'.json';
const shareName=()=>(A.settings.prefix||'단가DB').trim()+'_영업용.json';
async function findMain(fid){
  const f=(await find("'"+qs(fid)+"' in parents and name = '"+qs(mainName())+"' and trashed = false"))[0];if(f)return{file:f,main:true};
  if(A.readOnly)return null;
  const pre=(A.settings.prefix||'단가DB').trim();
  const l=(await find("'"+qs(fid)+"' in parents and name contains '"+qs(pre)+"' and trashed = false")).filter(x=>/\.json$/i.test(x.name)&&!/_영업용/.test(x.name)&&x.mimeType!==FOLDER);
  return l[0]?{file:l[0],main:false}:null;
}
async function readDb(id){const t=await api('GET',DRV+'/'+id+'?alt=media&supportsAllDrives=true',null,null,'text');const db=A.unpack(t);if(!db)throw new Error('드라이브 파일을 단가 데이터로 읽지 못했습니다.');return db}

let busy=false;
A.isBusy=()=>busy;
async function run(fn){if(busy)throw new Error('이전 작업이 진행 중입니다.');busy=true;A.status('작업 중…');A.emit();try{return await fn()}finally{busy=false;A.emit()}}
A.statusText={msg:'',cls:''};
A.status=(msg,cls)=>{A.statusText={msg,cls:cls||''};document.querySelectorAll('[data-status]').forEach(el=>{el.textContent=msg;el.className='status-line'+(cls?' '+cls:'')})};

/* save: merge with whatever another device saved, then write the main file; one dated backup per day */
A.buildShare=function(){
  const sells=A.DB.columns.filter(c=>c.group!=='buy'),keys=new Set(sells.map(c=>c.key));
  const entries=A.DB.entries.filter(e=>keys.has(e.col)&&A.ITEM.get(e.item)).map(e=>({id:e.id,item:e.item,col:e.col,v:e.v,t:e.t,cur:e.cur,from:e.from,at:e.at,base:!!e.base}));
  const used=new Set(entries.map(e=>e.item));
  const db={columns:sells.map(c=>({key:c.key,name:c.name,group:c.group})),items:A.DB.items.filter(i=>used.has(i.id)).map(i=>({id:i.id,cat1:i.cat1,cat2:i.cat2,model:i.model})),entries,fx:A.DB.fx.map(f=>clone(f)),groups:A.DB.groups.map(g=>({id:g.id,name:g.name,keys:g.keys.filter(k=>keys.has(k))})).filter(g=>g.keys.length),deleted:[],removedItems:[],removedColumns:[],log:[],fxCurs:[...A.DB.fxCurs],sheetLog:[]};
  return JSON.stringify({app:APP,schema:SCHEMA,kind:'share',savedAt:nowIso(),db});
};
A.saveShare=function(quiet){
  if(A.readOnly)return Promise.resolve();
  if(!A.settings.sharePath){A.toast('설정에서 영업용 파일 저장 폴더를 먼저 입력하세요.');return Promise.resolve()}
  return getToken(!quiet).then(()=>run(async()=>{
    A.status('영업용 파일을 저장하는 중…');
    const fid=await ensureFolderFor(true,true),name=shareName();
    const ex=(await find("'"+qs(fid)+"' in parents and name = '"+qs(name)+"' and trashed = false"))[0],content=A.buildShare();
    const r=ex?await upload({},content,ex.id):await upload({name,parents:[fid],mimeType:'application/json'},content);
    A.settings.lastShare=nowIso();A.saveSettings();
    A.status('영업용 파일을 저장했습니다: '+name+' ('+localDT(r.modifiedTime)+')','ok');A.toast('영업용 파일을 갱신했습니다.');A.emit();
  })).catch(e=>{if(e&&e.message){A.status(e.message,'err');if(!quiet)A.toast(e.message)}});
};
A.save=function(){
  if(A.readOnly){A.toast('영업용 읽기 전용 모드에서는 저장하지 않습니다. 최신본 불러오기를 이용하세요.');return Promise.resolve()}
  return getToken(true).then(()=>run(async()=>{
    A.status('저장 폴더를 확인하는 중…');
    const fid=await ensureFolder(true);
    const found=await findMain(fid);
    let merged=false;
    if(found&&found.main&&!A.sync.overwrite&&found.file.modifiedTime!==A.sync.remoteModified){
      A.status('다른 기기에서 저장한 내용을 합치는 중…');
      const remote=await readDb(found.file.id);
      const before=A.DB.entries.length;A.DB=A.mergeDb(A.DB,remote);A.reindex();merged=A.DB.entries.length!==before;
    }
    A.status('저장하는 중…');
    const content=A.pack();
    const r=found&&found.main?await upload({},content,found.file.id):await upload({name:mainName(),parents:[fid],mimeType:'application/json'},content);
    A.sync={fileId:r.id,remoteModified:r.modifiedTime,at:nowIso(),title:r.name,overwrite:false};A.dirty=0;
    A.settings.connected=true;A.saveSettings();A.persist();A.resetUndo();
    const day=today();
    if(A.settings.lastBackup!==day){
      try{const bf=await subFolder(fid,'백업',true);const bn=(A.settings.prefix||'단가DB').trim()+'_'+day.replace(/-/g,'')+'.json';const ex=(await find("'"+qs(bf)+"' in parents and name = '"+qs(bn)+"' and trashed = false"))[0];await upload(ex?{}:{name:bn,parents:[bf],mimeType:'application/json'},content,ex&&ex.id);A.settings.lastBackup=day;A.saveSettings()}catch(e){console.warn('backup',e)}
    }
    A.status('저장했습니다: '+r.name+' ('+localDT(r.modifiedTime)+')','ok');
    A.toast(merged?'다른 기기의 변경과 합쳐서 저장했습니다.':'드라이브에 저장했습니다.');
    A.emit();
    if(A.settings.autoShare&&A.settings.sharePath)setTimeout(()=>A.saveShare(true),300);
  })).catch(e=>{A.status(e.message,'err');A.toast(e.message);if(e.status===404){A.settings.folderId=null;A.saveSettings()}});
};
/* pull: bring the newest Drive copy in, keeping unsaved local changes */
A.pull=function(interactive){
  const p=interactive?getToken(true):(A.hasToken()?Promise.resolve():Promise.reject(new Error('')));
  return p.then(()=>run(async()=>{
    A.status('드라이브 최신본을 확인하는 중…');
    const fid=await ensureFolder(false);
    if(!fid){A.status('저장 폴더가 아직 없습니다. 저장하면 폴더를 만듭니다.','err');return false}
    const found=await findMain(fid);
    if(!found){A.status('폴더에 저장된 단가 파일이 없습니다. 엑셀을 가져온 뒤 저장하세요.','err');return false}
    if(found.file.modifiedTime===A.sync.remoteModified&&!A.isEmpty()){A.status('이미 최신입니다. ('+localDT(found.file.modifiedTime)+')','ok');return false}
    const remote=await readDb(found.file.id);
    if(A.dirty>0&&!A.isEmpty()){A.DB=A.mergeDb(A.DB,remote)}else{A.DB=normDb(remote);A.dirty=0}
    A.reindex();
    A.sync=Object.assign(A.sync,{fileId:found.main?found.file.id:null,remoteModified:found.main?found.file.modifiedTime:null,at:nowIso(),title:found.file.name});
    A.settings.connected=true;A.saveSettings();A.persist();A.resetUndo();
    A.status('불러왔습니다: '+found.file.name+' ('+localDT(found.file.modifiedTime)+')'+(A.dirty?' · 저장 안 된 변경 '+A.dirty+'건은 유지했습니다':''),'ok');
    A.emit();return true;
  })).catch(e=>{if(e.message){A.status(e.message,'err');if(interactive)A.toast(e.message)}return false});
};
A.listSaves=function(){
  return getToken(true).then(()=>run(async()=>{
    const fid=await ensureFolder(false);if(!fid)return [];
    const pre=(A.settings.prefix||'단가DB').trim();
    const top=(await find("'"+qs(fid)+"' in parents and name contains '"+qs(pre)+"' and trashed = false")).filter(f=>/\.json$/i.test(f.name));
    const bf=await subFolder(fid,'백업',false);
    const bk=bf?(await find("'"+qs(bf)+"' in parents and trashed = false")).filter(f=>/\.json$/i.test(f.name)).map(f=>Object.assign(f,{backup:true})):[];
    A.status('저장 기록 '+(top.length+bk.length)+'개를 찾았습니다.','ok');
    return top.concat(bk).sort((a,b)=>a.modifiedTime<b.modifiedTime?1:-1);
  }));
};
A.loadFile=function(f){
  return getToken(true).then(()=>run(async()=>{
    const db=await readDb(f.id);
    const isMain=f.name===mainName()&&!f.backup;
    A.DB=db;A.reindex();A.dirty=isMain?0:1;
    A.sync=Object.assign(A.sync,{fileId:isMain?f.id:A.sync.fileId,remoteModified:isMain?f.modifiedTime:A.sync.remoteModified,at:nowIso(),title:f.name,overwrite:!isMain});
    A.persist();A.resetUndo();A.status(isMain?'불러왔습니다: '+f.name:f.name+'을(를) 불러왔습니다. 저장하면 이 내용으로 현재 파일을 덮어씁니다.','ok');A.emit();
  }));
};
A.checkPath=function(){
  return getToken(true).then(()=>run(async()=>{
    A.settings.folderId=null;const r=await resolveFolder(false);
    if(r.id){A.settings.folderId=r.id;A.settings.folderFor=A.settings.path;A.settings.connected=true;A.saveSettings();A.status('폴더를 찾았습니다: '+r.label,'ok')}
    else A.status('"'+r.missing.join(' / ')+'" 폴더가 없습니다'+(r.found.length?' ('+r.found.join(' / ')+'까지는 있음)':'')+'. 저장하면 새로 만듭니다.','err');
  })).catch(e=>A.status(e.message,'err'));
};
A.importFromDrive=function(link){
  const m=String(link).match(/\/d\/([A-Za-z0-9_-]{10,})/)||String(link).match(/[?&]id=([A-Za-z0-9_-]{10,})/)||String(link).trim().match(/^([A-Za-z0-9_-]{25,})$/);
  if(!m)return Promise.reject(new Error('드라이브 파일 링크를 확인하세요.'));
  return getToken(true).then(()=>run(async()=>{A.status('엑셀 파일을 내려받는 중…');const buf=await api('GET',DRV+'/'+m[1]+'?alt=media&supportsAllDrives=true',null,null,'buf');return A.importXlsxBuffer(buf)}));
};

/* auto sync on open, only when a still-valid login exists (no popup without a click) */
A.autoSync=function(){if(A.settings.clientId)A.loadGis().catch(()=>{});if(A.settings.autoSync&&A.settings.connected&&A.hasToken())A.pull(false).then(ch=>{if(ch)A.toast('드라이브 최신본을 불러왔습니다.')})};

/* ---------- pill ---------- */
A.pillState=()=>{
  if(busy)return{t:'드라이브 작업 중…',c:''};
  if(A.readOnly)return A.sync.at?{t:'영업용 · 불러옴 '+shortDT(A.sync.at).slice(5),c:'ok'}:{t:'영업용 · 불러오기 전',c:''};
  if(A.dirty>0)return{t:'저장 안 된 변경 '+A.dirty+'건',c:'dirty'};
  if(A.sync.at)return{t:'저장됨 · '+shortDT(A.sync.at).slice(5),c:'ok'};
  return{t:'드라이브 저장 전',c:''};
};
A.renderPill=el=>{const s=A.pillState();el.className='pill'+(s.c?' '+s.c:'');el.textContent=s.t};

/* ---------- shared views ---------- */
function stepChart(es){
  const W=400,H=130,L=56,R=10,T=12,B=24,t=today();
  const xs=es.map(e=>Date.parse(e.from));const x0=Math.min(...xs),x1=Math.max(Date.parse(t),...xs)+864e5*5;
  const vs=es.map(e=>e.v);let lo=Math.min(...vs),hi=Math.max(...vs);if(lo===hi){lo=lo*0.9;hi=hi*1.1||1}const pd=(hi-lo)*0.15;lo-=pd;hi+=pd;
  const X=v=>L+(v-x0)/(x1-x0)*(W-L-R),Y=v=>T+(1-(v-lo)/(hi-lo))*(H-T-B);
  let p='';es.forEach((e,k)=>{const x=X(Date.parse(e.from)),y=Y(e.v);p+=(k?'H'+x.toFixed(1)+'V'+y.toFixed(1):'M'+x.toFixed(1)+' '+y.toFixed(1))});p+='H'+(W-R);
  const area=p+'V'+(H-B)+'H'+X(Date.parse(es[0].from)).toFixed(1)+'Z';const last=es[es.length-1];
  const tx=(x,y,s,anchor)=>'<text x="'+x+'" y="'+y+'" text-anchor="'+(anchor||'start')+'" font-size="10" fill="var(--muted)" font-family="IBM Plex Mono,monospace">'+esc(s)+'</text>';
  let g='';[lo+pd,hi-pd].forEach(v=>{const y=Y(v);g+='<line x1="'+L+'" x2="'+(W-R)+'" y1="'+y+'" y2="'+y+'" stroke="var(--line)" stroke-dasharray="3 3"/>'+tx(L-6,y+4,fmtN(Math.round(v*100)/100),'end')});
  const dots=es.map(e=>'<circle cx="'+X(Date.parse(e.from)).toFixed(1)+'" cy="'+Y(e.v).toFixed(1)+'" r="'+(e===last?4:2.5)+'" fill="var(--accent)"/>').join('');
  return '<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="단가 추이"><path d="'+area+'" fill="var(--accent-weak)" opacity=".7"/>'+g+'<path d="'+p+'" fill="none" stroke="var(--accent)" stroke-width="2"/>'+dots+tx(L,H-6,es[0].from)+tx(W-R,H-6,t,'end')+'</svg>';
}

/* cell history + change form; opts.date = date currently viewed */
A.renderCell=function(el,i,c,opts){
  opts=opts||{};const it=A.ITEM.get(i),col=A.COL.get(c);if(!it||!col){el.innerHTML='';return}
  const h=A.hist(i,c),cur=A.entryAt(i,c,today()),t=today();
  const lastCur=(h.length?h[h.length-1].cur:null)||'KRW';
  let b=(opts.notice?'<div class="status-line err">'+esc(opts.notice)+'</div>':'')+'<div><div class="hint">오늘 적용 중인 단가</div><div class="now"><span class="big">'+esc(cur?A.eText(cur):'없음')+'</span>'+(cur&&A.krwNote(cur,t)?'<span class="hint">'+esc(A.krwNote(cur,t))+'</span>':'')+(cur?'<span class="tag">'+esc(cur.from)+'부터</span>':'')+'</div></div>';
  const nums=h.filter(isNum).filter(e=>(e.cur||'KRW')===lastCur);
  if(nums.length>=2)b+='<div class="chart">'+stepChart(nums)+'</div>';
  b+='<div><h2 style="margin-bottom:6px">변경 이력 <span class="hint">('+h.length+'건)</span></h2>';
  if(!h.length)b+='<p class="hint">기록이 없습니다. 아래에서 첫 단가를 입력하세요.</p>';
  else{b+='<ol class="timeline">';for(const e of [...h].reverse()){const p=A.prevOf(e);
    b+='<li class="'+(e.from>t?'future':'')+'"><span class="d">'+esc(e.from)+(e.from>t?'<br>적용 예정':'')+'</span><span><span class="val">'+esc(A.eText(e))+'</span> '+(p?'<span class="hint">← '+esc(A.eText(p))+'</span> ':'')+A.deltaHtml(A.delta(p,e))+'</span>'+
    '<span class="meta">'+(e.base?'<span class="tag base">기준 단가</span>':'')+(e.reason&&!e.base?'<span class="tag">'+esc(e.reason)+'</span>':'')+(e.note?'<span>'+esc(e.note)+'</span>':'')+'<span>입력 '+esc(shortDT(e.at))+(e.by?' · '+esc(e.by):'')+'</span>'+(e.edits&&e.edits.length?'<span class="tag" title="'+esc(e.edits.map(x=>shortDT(x.at)+(x.by?' '+x.by:'')+': '+x.text).join('\n'))+'">수정됨</span><span class="hint">'+esc(e.edits[e.edits.length-1].text)+'</span>':'')+(e.base?'':'<button class="btn small" type="button" data-edit="'+esc(e.id)+'">수정</button><button class="btn small danger" type="button" data-del="'+esc(e.id)+'">삭제</button>')+'</span></li>'}
    b+='</ol>'}
  b+='</div><form class="section" autocomplete="off" data-cellform><h2>단가 변경 기록</h2><div class="form">'+
    '<label class="field"><span>새 단가</span><input type="text" inputmode="decimal" name="v" placeholder="숫자"></label>'+
    '<label class="field"><span>통화</span><select name="cur" class="cursel">'+A.curOptions('KRW')+'</select></label>'+
    '<label class="field"><span>적용일</span><input type="date" name="from" required></label>'+
    '<label class="field"><span>사유</span><select name="reason">'+REASONS.map(r=>'<option>'+esc(r)+'</option>').join('')+'</select></label>'+
    '<label class="field full"><span>비고</span><input type="text" name="note" placeholder="예: 5월 출고분부터, 견적서 번호"></label>'+
    '<label class="check full"><input type="checkbox" name="none"> 거래 없음으로 변경 (칸 비우기)</label></div>'+
    '<p class="hint" data-msg>적용일 이전 날짜로 보면 이전 단가가, 이후 날짜로 보면 새 단가가 나옵니다.</p><div class="row-end"><button class="btn primary" type="submit">기록하기</button></div></form>';
  el.innerHTML=b;
  if(opts.readonly){const fm=el.querySelector('[data-cellform]');if(fm)fm.remove();el.querySelectorAll('[data-edit],[data-del]').forEach(x=>x.remove());return}
  const f=el.querySelector('[data-cellform]');A.setCur(f.cur,lastCur);const d=opts.date||t;f.addEventListener('input',()=>{if(f.dataset.ok){delete f.dataset.ok;f.querySelector('button[type=submit]').textContent='기록하기'}});f.from.value=d>t?d:t;
  f.addEventListener('submit',ev=>{ev.preventDefault();const msg=f.querySelector('[data-msg]');
    const none=f.none.checked,raw=f.v.value.trim().replace(/,/g,''),from=f.from.value;
    if(!from){msg.textContent='적용일을 입력하세요.';msg.className='hint up';return}
    let v=null,tx=null;
    if(!none){if(raw==='-')tx='-';else{v=Number(raw);if(raw===''||isNaN(v)){msg.textContent='새 단가를 숫자로 입력하세요. 거래를 끊는 경우 "거래 없음"을 체크하세요.';msg.className='hint up';return}}}
    const e={item:i,col:c,v,t:tx,cur:f.cur.value,from,reason:f.reason.value,note:f.note.value.trim()};
    const prevE=A.entryAt(i,c,from),okKey=JSON.stringify([v,tx,e.cur,from]);
    if(f.dataset.ok!==okKey){
      const parts=[];
      if(A.isBig(prevE,e)){const same=(prevE.cur||'KRW')===(e.cur||'KRW'),pe=same?(e.v-prevE.v)/prevE.v*100:null;
        parts.push(same?'⚠ '+A.eText(prevE)+' → '+A.eText(e)+' ('+(pe>0?'+':'')+pe.toFixed(0)+'%)로 크게 바뀝니다.':'⚠ 통화가 '+A.curLabel(prevE.cur)+' → '+A.curLabel(e.cur)+'(으)로 바뀝니다.')}
      if(isNum(e)){const im=A.impact([e],from);if(im.newNeg.length)parts.push('⚠ 이 변경으로 역마진 칸이 '+im.newNeg.length+'개 늘어납니다 ('+im.newNeg.slice(0,3).map(x=>A.itemName(x.it)+'·'+x.c.name).join(', ')+(im.newNeg.length>3?' 외':'')+').')}
      if(parts.length){msg.textContent=parts.join(' ')+' 맞으면 한 번 더 눌러 주세요.';msg.className='hint up';f.dataset.ok=okKey;f.querySelector('button[type=submit]').textContent='맞습니다, 기록하기';return}}
    A.addEntry(e);const notice=A.itemMarginNotice(i,from);
    A.toast(from+'부터 '+A.eText(e)+'(으)로 기록했습니다.',{action:{label:'되돌리기',fn:()=>{A.undo();A.renderCell(el,i,c,opts)}}});
    A.renderCell(el,i,c,Object.assign({},opts,{notice}));
  });
  el.querySelectorAll('[data-del]').forEach(btn=>btn.addEventListener('click',()=>{
    if(btn.dataset.armed){A.deleteEntry(btn.dataset.del);A.toast('기록을 삭제했습니다.');A.renderCell(el,i,c,opts)}
    else{btn.dataset.armed='1';btn.textContent='한 번 더 누르면 삭제';setTimeout(()=>{if(btn.isConnected){delete btn.dataset.armed;btn.textContent='삭제'}},4000)}
  }));
  el.querySelectorAll('[data-edit]').forEach(btn=>btn.addEventListener('click',()=>{
    const e=A.DB.entries.find(x=>x.id===btn.dataset.edit);if(!e)return;
    const li=btn.closest('li');li.classList.add('editing');
    li.innerHTML='<form class="editentry" autocomplete="off" style="grid-column:1/-1"><b>기록 수정 · '+esc(e.from)+'</b><div class="form">'+
      '<label class="field"><span>적용일</span><input type="date" name="from" value="'+esc(e.from)+'" required></label>'+
      '<label class="field"><span>단가</span><input type="text" inputmode="decimal" name="v" value="'+esc(isNum(e)?String(e.v):(e.t||''))+'"></label>'+
      '<label class="field"><span>통화</span><select name="cur" class="cursel">'+A.curOptions(e.cur||'KRW')+'</select></label>'+
      '<label class="field"><span>사유</span><select name="reason">'+(e.reason&&!REASONS.includes(e.reason)?'<option>'+esc(e.reason)+'</option>':'')+(e.reason?'':'<option value="">(없음)</option>')+REASONS.map(r=>'<option'+(r===e.reason?' selected':'')+'>'+esc(r)+'</option>').join('')+'</select></label>'+
      '<label class="field full"><span>비고</span><input type="text" name="note" value="'+esc(e.note||'')+'"></label>'+
      '<label class="check full"><input type="checkbox" name="none"'+((!isNum(e)&&!e.t)?' checked':'')+'> 거래 없음 (칸 비우기)</label></div>'+
      '<p class="hint up" data-emsg></p><div class="row-end"><button class="btn" type="button" data-ecancel>취소</button><button class="btn primary" type="submit">수정 저장</button></div></form>';
    const f=li.querySelector('form'),msg=m=>{f.querySelector('[data-emsg]').textContent=m};
    A.setCur(f.cur,e.cur||'KRW');f.from.focus();
    f.querySelector('[data-ecancel]').addEventListener('click',()=>A.renderCell(el,i,c,opts));
    f.addEventListener('submit',ev=>{ev.preventDefault();
      const none=f.none.checked,raw=f.v.value.trim().replace(/,/g,''),from=f.from.value;
      if(!from){msg('적용일을 입력하세요.');return}
      let v=null,tx=null;
      if(!none){if(raw==='-')tx='-';else{v=Number(raw);if(raw===''||isNaN(v)){msg('단가를 숫자로 입력하세요. 거래를 끊는 기록이면 "거래 없음"을 체크하세요.');return}}}
      if(!A.updateEntry(e.id,{from,v,t:tx,cur:f.cur.value,reason:f.reason.value,note:f.note.value.trim()})){msg('바뀐 내용이 없습니다.');return}
      A.toast('기록을 수정했습니다.',{action:{label:'되돌리기',fn:()=>{A.undo();A.renderCell(el,i,c,opts)}}});
      A.renderCell(el,i,c,Object.assign({},opts,{notice:A.itemMarginNotice(i,from)}));
    });
  }));
};

function catOptions(){return '<option value="">전체</option>'+A.cats().map(c=>'<option>'+esc(c)+'</option>').join('')}
function colOptions(){const gs=A.DB.groups.filter(g=>A.groupKeys(g.id).length);return '<option value="">전체</option>'+(gs.length?'<optgroup label="그룹">'+gs.map(g=>'<option value="g:'+esc(g.id)+'">'+esc(g.name)+' ('+A.groupKeys(g.id).length+'곳)</option>').join('')+'</optgroup>':'')+'<optgroup label="거래처">'+A.DB.columns.map(c=>'<option value="'+esc(c.key)+'">'+esc(c.name)+(c.group==='buy'?' (매입)':'')+'</option>').join('')+'</optgroup>'}
A.colOptions=colOptions;
function keepVal(sel,html){const v=sel.value;sel.innerHTML=html;sel.value=v}

/* 이력 검색 */
A.mountHistory=function(el,opts){
  opts=opts||{};
  el.innerHTML='<div class="filters"><label class="field"><span>모델 · 구분</span><input type="search" name="q" placeholder="모델명 일부"></label><label class="field"><span>거래처</span><select name="col"></select></label><label class="field"><span>대구분</span><select name="cat"></select></label><label class="field"><span>적용일 시작</span><input type="date" name="from"></label><label class="field"><span>적용일 끝</span><input type="date" name="to"></label><label class="field"><span>사유 · 비고</span><input type="search" name="reason" placeholder="예: 원자재"></label><label class="check"><input type="checkbox" name="base"> 기준 단가 포함</label></div><div class="row" style="justify-content:space-between"><div class="stats" data-stats></div>'+expRow+'</div><div data-out></div>';
  let lastRows=[],lastDesc='';
  const g=n=>el.querySelector('[name="'+n+'"]');
  function refresh(){
    keepVal(g('col'),colOptions());keepVal(g('cat'),catOptions());
    const q=g('q').value.trim(),col=g('col').value,cat=g('cat').value,f=g('from').value,t=g('to').value,rq=g('reason').value.trim().toLowerCase(),base=g('base').checked;
    const match=(e,o)=>{if(!o.base&&e.base)return false;const it=A.ITEM.get(e.item);if(!it||!A.COL.get(e.col))return false;if(!A.colMatch(o.col,e.col))return false;if(!A.matchItem(it,o.q,o.cat))return false;if(o.f&&e.from<o.f)return false;if(o.t&&e.from>o.t)return false;if(o.rq&&!((e.reason||'')+' '+(e.note||'')).toLowerCase().includes(o.rq))return false;return true};
    const cond={q,col,cat,f,t,rq,base};
    const list=A.DB.entries.filter(e=>match(e,cond)).sort((a,b)=>-cmpE(a,b));
    let up=0,down=0;const rows=list.map(e=>{const p=A.prevOf(e),x=A.delta(p,e);if(x&&x.d>0)up++;if(x&&x.d<0)down++;return{e,p,x}});
    const lastCh=A.DB.entries.filter(e=>!e.base).sort(cmpE).pop();
    el.querySelector('[data-stats]').innerHTML='<span class="stat">결과<b>'+list.length+'</b></span><span class="stat">인상<b class="up">'+up+'</b></span><span class="stat">인하<b class="down">'+down+'</b></span><span class="stat">마지막 변경<b>'+esc(lastCh?lastCh.from:'없음')+'</b></span>';
    const out=el.querySelector('[data-out]');
    if(!list.length){lastRows=[];
      if(A.isEmpty()){out.innerHTML='<div class="empty"><b>아직 단가 데이터가 없습니다.</b>설정에서 단가.xlsx를 가져오거나 드라이브에서 불러오세요.</div>';return}
      const any=A.DB.entries.some(e=>!e.base);
      const active=[];const cn=col?{name:A.colLabel(col)}:null;
      if(q)active.push('모델 "'+q+'"');if(cn)active.push('거래처 '+cn.name);if(cat)active.push('대구분 '+cat);if(f||t)active.push('적용일 '+(f||'처음')+' ~ '+(t||'오늘 이후'));if(rq)active.push('사유 "'+rq+'"');
      const nBase=base?0:A.DB.entries.filter(e=>match(e,Object.assign({},cond,{base:true}))).length;
      const nOnlyQ=active.length>1||(active.length===1&&!q)?A.DB.entries.filter(e=>match(e,{q,base:true})).length:0;
      let h='<div class="empty"><b>'+(any||base?'조건에 맞는 기록이 없습니다.':'아직 기준 단가 이후 변경 기록이 없습니다.')+'</b>';
      if(active.length)h+='<span>지금 조건: '+esc(active.join(' · '))+' (모두 만족해야 나옵니다)</span>';
      if(!any&&!base)h+='<span>단가 칸을 눌러 변경을 기록하면 여기에 쌓입니다. 처음 가져온 단가는 "기준 단가 포함"을 켜야 보입니다.</span>';
      h+='<div class="row" style="justify-content:center">';
      if(nBase)h+='<button class="btn primary" type="button" data-hx="base">기준 단가 포함해서 보기 ('+nBase+'건)</button>';
      if(nOnlyQ&&q)h+='<button class="btn'+(nBase?'':' primary')+'" type="button" data-hx="onlyq">모델 "'+esc(q)+'"만으로 찾기 ('+nOnlyQ+'건)</button>';
      if(active.length&&!(nOnlyQ&&q))h+='<button class="btn" type="button" data-hx="reset">조건 초기화</button>';
      out.innerHTML=h+'</div></div>';return}
    lastRows=rows;const act=[];const cnm=A.COL.get(col);if(q)act.push('모델 "'+q+'"');if(col)act.push('거래처 '+A.colLabel(col));if(cat)act.push('대구분 '+cat);if(f||t)act.push('적용일 '+(f||'처음')+' ~ '+(t||'끝'));if(rq)act.push('사유 "'+rq+'"');act.push(base?'기준 단가 포함':'기준 단가 제외');
    lastDesc='조건: '+act.join(' · ')+' · '+rows.length+'건 (인상 '+up+', 인하 '+down+')';
    const shown=rows.slice(0,600);
    if(opts.cards){
      out.innerHTML='<div class="cards">'+shown.map(({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return '<div class="card click" data-i="'+esc(e.item)+'" data-c="'+esc(e.col)+'"><div class="top"><span class="model">'+esc(A.itemName(it))+'</span><span class="num hint">'+esc(e.from)+(e.from>today()?' <span class="tag fut">예정</span>':'')+'</span></div><div class="sub">'+esc(c.name)+' · '+esc(A.itemLabel(it))+'</div><div class="vals">'+(p?'<span class="hint">'+esc(A.eText(p))+' →</span>':'')+'<b>'+esc(A.eText(e))+'</b>'+A.deltaHtml(x)+'</div><div class="sub">'+(e.base?'<span class="tag base">기준 단가</span>':'<span class="tag">'+esc(e.reason||'')+'</span>'+(e.edits&&e.edits.length?' <span class="tag" title="'+esc(e.edits.map(x=>x.text).join(' / '))+'">수정됨</span>':''))+(e.note?' '+esc(e.note):'')+(e.by?' · '+esc(e.by):'')+'</div></div>'}).join('')+'</div>';
    }else{
      out.innerHTML='<div class="tbl-wrap section" style="padding:6px"><table class="list"><thead><tr><th>적용일</th><th>구분</th><th>모델</th><th>거래처</th><th class="r">이전</th><th class="r">변경</th><th class="r">증감</th><th>사유</th><th>비고</th><th>입력</th></tr></thead><tbody>'+
      shown.map(({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return '<tr class="click" data-i="'+esc(e.item)+'" data-c="'+esc(e.col)+'"><td class="num">'+esc(e.from)+(e.from>today()?' <span class="tag fut">예정</span>':'')+'</td><td>'+esc(A.itemLabel(it))+'</td><td class="num">'+esc(it.model)+'</td><td>'+esc(c.name)+'</td><td class="r num">'+(p?esc(A.eText(p)):'<span class="hint">-</span>')+'</td><td class="r num"><b>'+esc(A.eText(e))+'</b></td><td class="r">'+A.deltaHtml(x)+'</td><td>'+(e.base?'<span class="tag base">기준 단가</span>':'<span class="tag">'+esc(e.reason||'')+'</span>'+(e.edits&&e.edits.length?' <span class="tag" title="'+esc(e.edits.map(x=>x.text).join(' / '))+'">수정됨</span>':''))+'</td><td>'+esc(e.note||'')+'</td><td class="num hint">'+esc(shortDT(e.at))+(e.by?' · '+esc(e.by):'')+'</td></tr>'}).join('')+'</tbody></table></div>';
    }
    if(rows.length>600)out.insertAdjacentHTML('beforeend','<p class="hint">처음 600건만 표시합니다. 조건을 좁혀 주세요.</p>');
  }
  el.addEventListener('input',refresh);
  el.addEventListener('click',ev=>{const x=ev.target.closest('[data-x]');if(x){if(!lastRows.length){A.toast('내보낼 기록이 없습니다. 검색 조건을 확인하세요.');return}if(x.dataset.x==='xlsx')A.exportHistoryXlsx(lastRows,lastDesc).then(()=>A.toast('엑셀 파일을 내려받았습니다.')).catch(expErr);else A.printHistory(lastRows,lastDesc);return}
    const b=ev.target.closest('[data-hx]');if(!b)return;const k=b.dataset.hx;
    if(k==='onlyq'||k==='reset'){['col','cat','from','to','reason'].forEach(n=>g(n).value='');if(k==='reset')g('q').value=''}
    g('base').checked=true;if(k==='reset'&&!A.DB.entries.some(e=>!e.base))g('base').checked=true;refresh()});
  refresh();return{refresh};
};

/* 변경 분석 */
function barChart(labels,vals){
  const W=520,H=180,L=28,R=8,T=14,B=26;const mx=Math.max(4,...vals);const step=Math.ceil(mx/4);const top=step*4;const bw=(W-L-R)/labels.length;let s='';
  const tx=(x,y,str,anchor,fill)=>'<text x="'+x+'" y="'+y+'" text-anchor="'+(anchor||'middle')+'" font-size="10" fill="'+(fill||'var(--muted)')+'" font-family="IBM Plex Mono,monospace">'+esc(str)+'</text>';
  for(let k=0;k<=4;k++){const v=step*k,y=T+(1-v/top)*(H-T-B);s+='<line x1="'+L+'" x2="'+(W-R)+'" y1="'+y+'" y2="'+y+'" stroke="var(--line)"/>'+tx(L-5,y+3.5,v,'end')}
  labels.forEach((m,k)=>{const v=vals[k],h=v/top*(H-T-B),x=L+k*bw+bw*.18,y=H-B-h;
    if(v)s+='<rect x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+(bw*.64).toFixed(1)+'" height="'+h.toFixed(1)+'" rx="2" fill="var(--accent)"><title>'+m+': '+v+'건</title></rect>'+tx((x+bw*.32).toFixed(1),(y-3).toFixed(1),v,'middle','var(--fg)');
    s+=tx((L+k*bw+bw/2).toFixed(1),H-8,(k===0||m.endsWith('-01'))?m.slice(2).replace('-','.'):m.slice(5))});
  return '<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="월별 변경 건수">'+s+'</svg>';
}
A.mountAnalysis=function(el,opts){
  opts=opts||{};
  el.innerHTML='<div class="section"><div class="filters"><h2 style="align-self:center">두 날짜 비교</h2><label class="field"><span>이전 날짜</span><input type="date" name="a"></label><label class="field"><span>이후 날짜</span><input type="date" name="b"></label><label class="field"><span>거래처</span><select name="col"></select></label></div><div class="row" style="justify-content:space-between"><div class="stats" data-stats></div>'+expRow+'</div><div data-diff></div></div>'+
    '<div class="section"><h2>월별 변경 건수</h2><p class="hint">적용일 기준 최근 12개월, 기준 단가 제외.</p><div class="chart" data-monthly></div></div>'+
    '<div class="section"><h2>거래처별 변경 요약</h2><div class="tbl-wrap" data-bycol></div></div>'+
    '<div class="section"><h2>변경이 잦은 품목</h2><div class="tbl-wrap" data-byitem></div></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  function refresh(){
    keepVal(g('col'),colOptions());
    if(!g('a').value){const bases=A.DB.entries.filter(e=>e.base).map(e=>e.from).sort();g('a').value=bases[0]||addDays(today(),-30)}
    if(!g('b').value)g('b').value=today();
    const Ad=g('a').value,Bd=g('b').value,colF=g('col').value;const diffs=[];last={diffs,Ad,Bd,colF};
    for(const it of A.DB.items)for(const c of A.DB.columns){if(!A.colMatch(colF,c.key))continue;const a=A.entryAt(it.id,c.key,Ad),b=A.entryAt(it.id,c.key,Bd);if(!A.sameVal(a&&A.hasVal(a)?a:null,b&&A.hasVal(b)?b:null))diffs.push({it,c,a,b,x:A.delta(a,b)})}
    const up=diffs.filter(d=>d.x&&d.x.d>0).length,down=diffs.filter(d=>d.x&&d.x.d<0).length,nw=diffs.filter(d=>!A.hasVal(d.a)).length,gone=diffs.filter(d=>!A.hasVal(d.b)).length;
    el.querySelector('[data-stats]').innerHTML='<span class="stat">바뀐 칸<b>'+diffs.length+'</b></span><span class="stat">인상<b class="up">'+up+'</b></span><span class="stat">인하<b class="down">'+down+'</b></span><span class="stat">신규<b>'+nw+'</b></span><span class="stat">거래 없음으로<b>'+gone+'</b></span>';
    const dEl=el.querySelector('[data-diff]');
    if(!diffs.length)dEl.innerHTML='<div class="empty"><b>두 날짜 사이에 바뀐 단가가 없습니다.</b>날짜를 바꾸거나, 단가 칸을 눌러 변경을 기록해 보세요.</div>';
    else if(opts.cards)dEl.innerHTML='<div class="cards">'+diffs.map(d=>'<div class="card click" data-i="'+esc(d.it.id)+'" data-c="'+esc(d.c.key)+'"><div class="top"><span class="model">'+esc(A.itemName(d.it))+'</span><span class="hint">'+esc(d.c.name)+'</span></div><div class="vals"><span class="hint">'+esc(A.eText(d.a))+' →</span><b>'+esc(A.eText(d.b))+'</b>'+A.deltaHtml(d.x)+'</div></div>').join('')+'</div>';
    else dEl.innerHTML='<div class="tbl-wrap"><table class="list"><thead><tr><th>구분</th><th>모델</th><th>거래처</th><th class="r">'+esc(Ad)+'</th><th class="r">'+esc(Bd)+'</th><th class="r">증감</th><th>최근 변경일</th></tr></thead><tbody>'+diffs.map(d=>'<tr class="click" data-i="'+esc(d.it.id)+'" data-c="'+esc(d.c.key)+'"><td>'+esc(A.itemLabel(d.it))+'</td><td class="num">'+esc(d.it.model)+'</td><td>'+esc(d.c.name)+'</td><td class="r num">'+esc(A.eText(d.a))+'</td><td class="r num"><b>'+esc(A.eText(d.b))+'</b></td><td class="r">'+A.deltaHtml(d.x)+'</td><td class="num">'+esc(d.b?d.b.from:'')+'</td></tr>').join('')+'</tbody></table></div>';
    const ch=A.DB.entries.filter(e=>!e.base);
    const months=[];const d0=new Date();for(let k=11;k>=0;k--){const d=new Date(d0.getFullYear(),d0.getMonth()-k,1);months.push(d.getFullYear()+'-'+pad(d.getMonth()+1))}
    el.querySelector('[data-monthly]').innerHTML=barChart(months,months.map(m=>ch.filter(e=>e.from.slice(0,7)===m).length));
    const byCol=A.DB.columns.map(c=>{const es=ch.filter(e=>e.col===c.key);const ps=es.map(e=>A.delta(A.prevOf(e),e)).filter(x=>x&&x.p!=null).map(x=>x.p);return{c,n:es.length,avg:ps.length?ps.reduce((a,b)=>a+b,0)/ps.length:null,last:es.length?es.map(e=>e.from).sort().pop():''}}).filter(r=>r.n);
    el.querySelector('[data-bycol]').innerHTML=byCol.length?'<table class="list"><thead><tr><th>거래처</th><th class="r">건수</th><th class="r">평균 증감률</th><th>최근 변경</th></tr></thead><tbody>'+byCol.sort((a,b)=>b.n-a.n).map(r=>'<tr><td>'+esc(r.c.name)+'</td><td class="r num">'+r.n+'</td><td class="r num '+(r.avg>0?'up':r.avg<0?'down':'')+'">'+(r.avg==null?'-':(r.avg>0?'+':'')+r.avg.toFixed(1)+'%')+'</td><td class="num">'+esc(r.last)+'</td></tr>').join('')+'</tbody></table>':'<p class="hint">변경 기록이 쌓이면 거래처별 건수와 평균 증감률이 나옵니다.</p>';
    const m=new Map();for(const e of ch)m.set(e.item,(m.get(e.item)||0)+1);
    const top=[...m.entries()].sort((a,b)=>b[1]-a[1]).slice(0,15).filter(([i])=>A.ITEM.get(i));
    el.querySelector('[data-byitem]').innerHTML=top.length?'<table class="list"><thead><tr><th>모델</th><th>구분</th><th class="r">건수</th><th>최근 변경</th></tr></thead><tbody>'+top.map(([i,n])=>{const it=A.ITEM.get(i);const last=ch.filter(e=>e.item===i).sort(cmpE).pop();return '<tr><td class="num">'+esc(A.itemName(it))+'</td><td>'+esc(A.itemLabel(it))+'</td><td class="r num">'+n+'</td><td class="num">'+esc(last.from)+' · '+esc(A.COL.get(last.col)?A.COL.get(last.col).name:'')+'</td></tr>'}).join('')+'</tbody></table>':'<p class="hint">아직 변경 기록이 없습니다.</p>';
  }
  let last=null;
  el.addEventListener('input',refresh);
  el.addEventListener('click',ev=>{const x=ev.target.closest('[data-x]');if(!x||!last)return;if(!last.diffs.length){A.toast('두 날짜 사이에 바뀐 단가가 없습니다.');return}
    const L=last,cn=L.colF?{name:A.colLabel(L.colF)}:null;const up=L.diffs.filter(d=>d.x&&d.x.d>0).length,down=L.diffs.filter(d=>d.x&&d.x.d<0).length;
    const desc=(cn?'거래처 '+cn.name+' · ':'')+'바뀐 칸 '+L.diffs.length+'개 (인상 '+up+', 인하 '+down+')';
    if(x.dataset.x==='xlsx')A.exportDiffXlsx(L.diffs,L.Ad,L.Bd,desc).then(()=>A.toast('엑셀 파일을 내려받았습니다.')).catch(expErr);else A.printDiff(L.diffs,L.Ad,L.Bd,desc)});
  refresh();return{refresh};
};

/* 환율 */
A.mountFx=function(el){
  function refresh(){
    const cols=A.fxColumns(),fx=[...A.DB.fx].sort((a,b)=>a.date<b.date?1:-1);
    const inp=(f,k)=>'<input type="number" step="any" data-k="'+k+'" value="'+esc(f[k]==null?'':f[k])+'" style="width:96px;text-align:right">';
    el.innerHTML='<div class="row" style="justify-content:space-between"><h2>환율 (외화 → 원)</h2><div class="row"><button class="btn small" type="button" data-fxcur>통화 추가</button><button class="btn small" type="button" data-fxadd>환율 추가</button></div></div><p class="hint">적용일부터 다음 환율 전까지 쓰입니다. 모든 통화는 <b>1단위당 원</b>으로 입력하세요(예: 1 JPY = 9.2원). 100엔 기준으로 고시된 값은 100으로 나눠 넣습니다.</p>'+
      (fx.length?'<div class="tbl-wrap"><table class="list"><thead><tr><th>적용일</th>'+cols.map(c=>'<th class="r">1 '+esc(c)+'</th>').join('')+'<th></th></tr></thead><tbody>'+fx.map(f=>'<tr data-fx="'+esc(f.id)+'"><td><input type="date" data-k="date" value="'+esc(f.date)+'" style="width:140px"></td>'+cols.map(c=>'<td class="r">'+inp(f,c)+'</td>').join('')+'<td class="r"><button class="btn small danger" type="button" data-fxdel="'+esc(f.id)+'">삭제</button></td></tr>').join('')+'</tbody></table></div>'
      :'<div class="empty"><b>환율이 아직 없습니다.</b>외화 단가의 마진과 원화 환산을 보려면 환율을 추가하세요.</div>');
  }
  el.addEventListener('click',ev=>{
    if(ev.target.closest('[data-fxadd]')){A.DB.fx.push({id:uid('fx'),date:today()});A.touch(1,'환율 추가');refresh();return}
    if(ev.target.closest('[data-fxcur]')){A.pickCurrency(code=>{if(!code||code==='KRW')return;if(!A.fxColumns().includes(code)){A.DB.fxCurs.push(code);A.touch(1,'환율 통화 추가')}refresh();A.toast(code+' 칸을 추가했습니다.')});return}
    const b=ev.target.closest('[data-fxdel]');if(!b)return;
    if(b.dataset.armed){A.DB.fx=A.DB.fx.filter(x=>x.id!==b.dataset.fxdel);A.DB.deleted.push(b.dataset.fxdel);A.touch(1,'환율 삭제');refresh()}else{b.dataset.armed='1';b.textContent='한 번 더'}
  });
  el.addEventListener('change',ev=>{const i=ev.target.closest('input[data-k]');if(!i)return;const f=A.DB.fx.find(x=>x.id===i.closest('tr').dataset.fx);if(!f)return;const k=i.dataset.k;const n0=A.alerts(today()).neg.length;f[k]=k==='date'?i.value:(i.value===''?null:Number(i.value));A.touch(1,'환율 수정');const n1=A.alerts(today()).neg.length;if(n1>n0)A.toast('환율 변경으로 역마진 칸이 '+(n1-n0)+'개 늘었습니다.',{action:{label:'되돌리기',fn:()=>A.undo()}})});
  refresh();return{refresh};
};

/* 설정 · 저장 · 가져오기 */
function mountSalesSettings(el){
  el.innerHTML='<div class="section"><h2>영업용 보기 설정</h2><p class="hint">이 기기는 <b>영업용 읽기 전용</b>입니다. 판매 단가만 볼 수 있고, 매입가와 마진은 보이지 않으며 수정이나 저장은 할 수 없습니다.</p>'+
    '<label class="field"><span>구글 클라이언트 ID</span><input type="text" name="clientId" placeholder="관리자에게 받은 클라이언트 ID" autocomplete="off"></label>'+
    '<label class="field"><span>영업용 폴더 링크</span><input type="text" name="sharePath" placeholder="관리자가 공유해 준 폴더 링크를 붙여 넣기" autocomplete="off"></label>'+
    '<p class="hint">관리자가 영업용 폴더를 내 구글 계정에 <b>보기 권한</b>으로 공유해야 합니다. 폴더 링크는 드라이브에서 폴더를 열었을 때 주소창에 있는 <code>https://drive.google.com/drive/folders/…</code> 입니다.</p>'+
    '<label class="check"><input type="checkbox" name="autoSync"> 열 때 최신본 자동 확인</label>'+
    '<div class="row"><button class="btn primary" type="button" data-act="pull">최신본 불러오기</button><button class="btn" type="button" data-act="signout">로그아웃</button></div>'+
    '<div class="status-line" data-status></div></div>'+
    '<div class="section"><h2>일반 모드로 돌아가기</h2><p class="hint">관리자 본인의 기기라면 일반 모드로 돌아갈 수 있습니다. 일반 모드의 데이터는 따로 보관돼 있습니다.</p><div class="row"><button class="btn danger" type="button" data-act="normal">일반 모드로 전환</button></div></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  const fill=()=>{for(const k of ['clientId','sharePath'])if(document.activeElement!==g(k))g(k).value=A.settings[k]||'';g('autoSync').checked=!!A.settings.autoSync};fill();
  el.querySelectorAll('[data-status]').forEach(s=>{s.textContent=A.statusText.msg||(A.sync.at?'마지막으로 불러온 시각: '+shortDT(A.sync.at):'최신본 불러오기를 누르면 구글 로그인 창이 뜹니다.');s.className='status-line'+(A.statusText.cls?' '+A.statusText.cls:'')});
  el.addEventListener('change',ev=>{const n=ev.target.name;
    if(n==='clientId'){A.settings.clientId=ev.target.value.trim();if(A.settings.clientId)A.loadGis().catch(()=>{})}
    if(n==='sharePath'&&ev.target.value.trim()!==A.settings.sharePath){A.settings.sharePath=ev.target.value.trim();A.settings.shareFolderId=null;A.settings.shareFolderFor=null}
    if(n==='autoSync')A.settings.autoSync=ev.target.checked;A.saveSettings()});
  el.addEventListener('click',ev=>{const b=ev.target.closest('[data-act]');if(!b)return;
    if(b.dataset.act==='pull')A.pull(true).then(ch=>{if(ch)A.toast('최신본을 불러왔습니다.')});
    else if(b.dataset.act==='signout'){A.signOut();A.status('로그아웃했습니다.')}
    else if(b.dataset.act==='normal'){if(!b.dataset.armed){b.dataset.armed='1';b.textContent='한 번 더 누르면 전환';return}A.settings.salesMode=false;A.saveSettings();location.reload()}});
  return{refresh:fill};
}
A.mountSettings=function(el){
  if(A.readOnly)return mountSalesSettings(el);
  el.innerHTML=
  '<div class="section"><h2>기본 설정</h2>'+
    '<label class="field"><span>입력자 이름</span><input type="text" name="userName" placeholder="단가를 기록할 때 이력에 남습니다" autocomplete="off"></label>'+
    '<p class="hint">구글에 처음 로그인하면 계정 이름이 자동으로 들어갑니다. 바꿔도 됩니다.</p>'+
    '<div class="form"><label class="field"><span>큰 변동 경고 기준 (%)</span><input type="number" name="warnPct" min="0" step="1"></label><label class="field"><span>저마진 기준 (판매가 대비 %)</span><input type="number" name="lowMargin" step="1"></label></div>'+
    '<p class="hint">단가를 입력할 때 기존보다 이 비율 이상 달라지면 한 번 더 확인합니다(0이면 끔). 마진율이 저마진 기준보다 낮으면 주황, 마이너스면 빨강으로 표시합니다.</p>'+
    '<label class="field"><span>거래처별 단가표에 넣을 회사명</span><input type="text" name="company" placeholder="예: OO산업"></label>'+
    '<label class="field"><span>단가표 안내 문구</span><input type="text" name="sheetNote" placeholder="예: 부가세 별도 / 문의 02-000-0000"></label></div>'+
  '<div class="section"><h2>구글 드라이브 연결</h2>'+
    '<label class="field"><span>구글 클라이언트 ID</span><input type="text" name="clientId" placeholder="123456-abc.apps.googleusercontent.com" autocomplete="off"></label>'+
    '<p class="hint">처음 한 번만 입력합니다. 만드는 방법은 저장소의 README에 있습니다. 이 기기 브라우저에만 저장됩니다.</p>'+
    '<label class="field"><span>저장 폴더 경로 또는 폴더 링크</span><input type="text" name="path" placeholder="예: 회사 폴더/단가관리"></label>'+
    '<p class="hint">\'내 드라이브\' 아래 경로를 <code>/</code>로 구분합니다. 없으면 저장할 때 만듭니다.</p>'+
    '<label class="field"><span>파일 이름</span><input type="text" name="prefix" placeholder="단가DB"></label>'+
    '<p class="hint"><code>단가DB.json</code> 하나를 계속 갱신하고, 하루 한 번 <code>백업</code> 폴더에 날짜별 사본을 남깁니다.</p>'+
    '<label class="check"><input type="checkbox" name="autoSync"> 열 때 드라이브 최신본 자동 확인</label>'+
    '<div class="row"><button class="btn primary" type="button" data-act="save">드라이브에 저장</button><button class="btn" type="button" data-act="pull">최신본 불러오기</button><button class="btn" type="button" data-act="check">경로 확인</button><button class="btn" type="button" data-act="signout">로그아웃</button></div>'+
    '<div class="status-line" data-status></div></div>'+
  '<div class="section"><h2>영업용 공유 (읽기 전용)</h2><p class="hint">영업 담당자에게 <b>매입가와 마진 없이 판매 단가만</b> 보여 주는 파일을 만듭니다. 파일이 들어 있는 폴더를 드라이브에서 담당자에게 <b>보기 권한</b>으로 공유하세요. 담당자는 이 앱을 열어 설정에서 "영업용으로 전환"을 하면 됩니다. (담당자 구글 계정은 테스트 사용자에도 추가해야 합니다.)</p>'+
    '<label class="field"><span>영업용 파일을 저장할 폴더</span><input type="text" name="sharePath" placeholder="예: 회사 폴더/단가관리/영업용"></label>'+
    '<label class="check"><input type="checkbox" name="autoShare"> 드라이브에 저장할 때 영업용 파일도 함께 갱신</label>'+
    '<div class="row"><button class="btn primary" type="button" data-act="share">영업용 파일 저장</button><button class="btn danger" type="button" data-act="salesmode">이 기기를 영업용으로 전환</button></div>'+
    '<p class="hint" data-sharestat></p></div>'+
  '<div class="section"><div class="row" style="justify-content:space-between"><h2>저장 기록 · 백업</h2><button class="btn small" type="button" data-act="list">목록 보기</button></div><div class="tbl-wrap" data-files><p class="hint">목록 보기를 누르면 저장 파일과 날짜별 백업이 나옵니다.</p></div></div>'+
  '<div class="section"><h2>단가 엑셀 가져오기</h2><p class="hint">처음 시작할 때 원본 <code>단가.xlsx</code>를 가져옵니다. 첫 시트의 양식(대구분·소구분·MODEL, D열부터 거래처)을 읽고, 시트 이름(예: 2026.04)을 기준일로 씁니다. 셀 메모의 EUR·RMB도 반영합니다.</p>'+
    '<div class="row"><label class="btn" for="xlsxIn">이 기기에서 파일 선택</label><input type="file" id="xlsxIn" accept=".xlsx,.xls" hidden></div>'+
    '<div class="row"><input type="text" name="xlsxLink" placeholder="또는 드라이브 파일 링크 붙여넣기" style="flex:1 1 220px"><button class="btn" type="button" data-act="xlsxDrive">드라이브에서 가져오기</button></div>'+
    '<div data-import></div></div>'+
  '<div class="section"><h2>백업 파일 · 초기화</h2><div class="row"><button class="btn" type="button" data-act="jsonOut">백업 파일 내려받기</button><label class="btn" for="jsonIn">백업 파일 불러오기</label><input type="file" id="jsonIn" accept=".json,application/json" hidden><button class="btn danger" type="button" data-act="wipe">이 기기 데이터 지우기</button></div><p class="hint" data-bstatus>이 기기에도 임시 보관되지만, 드라이브에 저장해야 다른 기기에서 볼 수 있습니다.</p></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  const fill=()=>{const ss=el.querySelector('[data-sharestat]');if(ss)ss.textContent=A.settings.lastShare?'마지막 영업용 파일 저장: '+shortDT(A.settings.lastShare):'아직 영업용 파일을 만들지 않았습니다.';g('autoShare').checked=!!A.settings.autoShare;for(const k of ['clientId','path','prefix','userName','warnPct','lowMargin','company','sheetNote','sharePath'])if(document.activeElement!==g(k))g(k).value=A.settings[k]||'';g('autoSync').checked=!!A.settings.autoSync};
  fill();
  el.querySelectorAll('[data-status]').forEach(s=>{s.textContent=A.statusText.msg||(A.settings.clientId?(A.hasToken()?'구글에 로그인돼 있습니다.':'저장하거나 불러오면 구글 로그인 창이 뜹니다.'):'구글 클라이언트 ID를 입력하면 드라이브에 저장할 수 있습니다.');s.className='status-line'+(A.statusText.cls?' '+A.statusText.cls:'')});
  el.addEventListener('change',ev=>{
    const n=ev.target.name;
    if(n==='path'&&ev.target.value.trim()!==A.settings.path){A.settings.path=ev.target.value.trim();A.settings.folderId=null;A.settings.folderFor=null}
    if(n==='prefix')A.settings.prefix=ev.target.value.trim()||'단가DB';
    if(n==='clientId'){A.settings.clientId=ev.target.value.trim();if(A.settings.clientId)A.loadGis().catch(()=>{})}
    if(n==='autoSync')A.settings.autoSync=ev.target.checked;
    if(n==='autoShare')A.settings.autoShare=ev.target.checked;
    if(n==='sharePath'&&ev.target.value.trim()!==A.settings.sharePath){A.settings.sharePath=ev.target.value.trim();A.settings.shareFolderId=null;A.settings.shareFolderFor=null}
    if(n==='userName'||n==='company'||n==='sheetNote')A.settings[n]=ev.target.value.trim();
    if(n==='warnPct'||n==='lowMargin'){const v=ev.target.value.trim();A.settings[n]=v===''?(n==='warnPct'?30:10):Number(v)}
    A.saveSettings();
  });
  let files=[];
  const showImport=res=>{const box=el.querySelector('[data-import]');
    box.innerHTML='<div class="status-line ok">'+esc(res.sheetName)+' 시트에서 품목 '+res.db.items.length+'개, 거래처 '+res.db.columns.length+'곳, 단가 '+res.count+'건을 읽었습니다. 기준일 '+esc(res.base)+'.</div><div class="row" style="margin-top:8px">'+
    (A.isEmpty()?'<button class="btn primary" type="button" data-imp="use">이 데이터로 시작</button>':'<button class="btn danger" type="button" data-imp="use">지금 데이터를 버리고 이 데이터로 교체</button>')+'<button class="btn" type="button" data-imp="cancel">취소</button></div>';
    box.querySelector('[data-imp="use"]').addEventListener('click',()=>{A.replaceDb(res.db,{dirty:1,overwrite:!A.isEmpty()||!!A.sync.remoteModified});box.innerHTML='<div class="status-line ok">가져왔습니다. 드라이브에 저장하면 다른 기기에서도 볼 수 있습니다.</div>';A.toast('단가 데이터를 가져왔습니다.')});
    box.querySelector('[data-imp="cancel"]').addEventListener('click',()=>box.innerHTML='');
  };
  el.querySelector('#xlsxIn').addEventListener('change',ev=>{const f=ev.target.files&&ev.target.files[0];if(!f)return;ev.target.value='';const rd=new FileReader();rd.onload=()=>A.importXlsxBuffer(rd.result).then(showImport).catch(e=>A.status(e.message,'err'));rd.readAsArrayBuffer(f)});
  el.querySelector('#jsonIn').addEventListener('change',ev=>{const f=ev.target.files&&ev.target.files[0];if(!f)return;ev.target.value='';const rd=new FileReader();rd.onload=()=>{const db=A.unpack(String(rd.result));const bs=el.querySelector('[data-bstatus]');if(!db){bs.textContent='단가 관리 백업 파일이 아닙니다.';return}A.replaceDb(db,{dirty:1,overwrite:true});bs.textContent=f.name+'을(를) 불러왔습니다. 저장하면 드라이브 파일을 이 내용으로 덮어씁니다.'};rd.readAsText(f)});
  el.addEventListener('click',ev=>{
    const b=ev.target.closest('[data-act]');
    if(!b){const lb=ev.target.closest('[data-load]');if(lb){const f=files[Number(lb.dataset.load)];if(!f)return;if(A.dirty>0&&!lb.dataset.armed){lb.dataset.armed='1';lb.textContent='변경 '+A.dirty+'건 버리고 불러오기';return}A.loadFile(f).then(()=>A.toast('불러왔습니다.')).catch(e=>A.status(e.message,'err'))}return}
    const act=b.dataset.act;
    if(act==='save')A.save();
    else if(act==='share')A.saveShare();
    else if(act==='salesmode'){if(!b.dataset.armed){b.dataset.armed='1';b.textContent='한 번 더 누르면 전환 (이 기기는 읽기 전용이 됩니다)';setTimeout(()=>{if(b.isConnected){delete b.dataset.armed;b.textContent='이 기기를 영업용으로 전환'}},6000);return}A.settings.salesMode=true;A.saveSettings();location.hash='';location.reload()}
    else if(act==='pull')A.pull(true).then(ch=>{if(ch)A.toast('드라이브 최신본을 불러왔습니다.')});
    else if(act==='check')A.checkPath();
    else if(act==='signout'){A.signOut();A.status('로그아웃했습니다.')}
    else if(act==='list')A.listSaves().then(fs=>{files=fs;const box=el.querySelector('[data-files]');box.innerHTML=fs.length?'<table class="list"><thead><tr><th>파일</th><th>수정</th><th></th></tr></thead><tbody>'+fs.map((f,k)=>'<tr><td class="num" style="word-break:break-all">'+(f.backup?'<span class="tag">백업</span> ':'')+esc(f.name)+(f.id===A.sync.fileId?' <span class="tag base">현재</span>':'')+'</td><td class="num hint">'+esc(localDT(f.modifiedTime))+'</td><td class="r" style="white-space:nowrap"><button class="btn small" type="button" data-load="'+k+'">불러오기</button> '+(f.webViewLink&&/^https:\/\//.test(f.webViewLink)?'<a href="'+esc(f.webViewLink)+'" target="_blank" rel="noopener">열기</a>':'')+'</td></tr>').join('')+'</tbody></table>':'<p class="hint">저장된 파일이 없습니다.</p>'}).catch(e=>A.status(e.message,'err'));
    else if(act==='xlsxDrive')A.importFromDrive(g('xlsxLink').value).then(showImport).catch(e=>A.status(e.message,'err'));
    else if(act==='jsonOut')A.downloadText((A.settings.prefix||'단가DB')+'_'+today().replace(/-/g,'')+'.json',A.pack());
    else if(act==='wipe'){if(!b.dataset.armed){b.dataset.armed='1';b.textContent='정말 지우기 (드라이브 파일은 그대로)';setTimeout(()=>{if(b.isConnected){delete b.dataset.armed;b.textContent='이 기기 데이터 지우기'}},5000);return}
      A.replaceDb(emptyDb());A.sync={fileId:null,remoteModified:null,at:null,title:null,overwrite:false};A.persist();A.emit();el.querySelector('[data-bstatus]').textContent='이 기기의 데이터를 지웠습니다. 드라이브에서 다시 불러올 수 있습니다.'}
  });
  return{refresh:fill};
};

/* 품목 · 거래처 추가 */
function openModal(html,onSubmit,size){
  let m=$('#modal');if(!m){m=document.createElement('div');m.id='modal';m.className='modal';m.innerHTML='<div class="modal-card" role="dialog" aria-modal="true"></div>';document.body.appendChild(m);m.addEventListener('click',ev=>{if(ev.target===m)A.closeModal()})}
  const card=m.firstChild;card.className='modal-card'+(size?' '+size:'');card.innerHTML=html;m.hidden=false;
  const f=card.querySelector('form');f.addEventListener('submit',ev=>{ev.preventDefault();if(onSubmit(f)!==false)A.closeModal()});
  card.querySelector('[data-cancel]').addEventListener('click',A.closeModal);const first=f.querySelector('input,select');if(first)first.focus();
}
A.closeModal=()=>{const m=$('#modal');if(m){m.hidden=true;m.firstChild.innerHTML=''}};
A.addItemDialog=function(preset){
  preset=preset||{};
  const NEW='__new__';
  const opt=(v,l,sel)=>'<option value="'+esc(v)+'"'+(sel?' selected':'')+'>'+esc(l)+'</option>';
  const buys=A.DB.columns.filter(c=>c.group==='buy'),sells=A.DB.columns.filter(c=>c.group!=='buy');
  const priceRow=c=>'<tr><td>'+esc(c.name)+'</td><td><input type="text" inputmode="decimal" data-pc="'+esc(c.key)+'" placeholder="-" style="text-align:right"></td><td><select data-pcur="'+esc(c.key)+'" class="cursel">'+A.curOptions('KRW')+'</select></td></tr>';
  openModal('<form class="additem" autocomplete="off"><h2>품목 추가</h2>'+
    '<div class="form"><label class="field"><span>대구분</span><select name="c1sel">'+A.cats().map(c=>opt(c,c,c===preset.cat1)).join('')+opt(NEW,'+ 새 대구분',!A.cats().length)+'</select></label>'+
    '<label class="field" data-new1 hidden><span>새 대구분 이름</span><input type="text" name="c1new"></label>'+
    '<label class="field"><span>소구분</span><select name="c2sel"></select></label>'+
    '<label class="field" data-new2 hidden><span>새 소구분 이름</span><input type="text" name="c2new"></label>'+
    '<label class="field full"><span>모델 (한 줄에 하나, 여러 개 가능)</span><textarea name="models" rows="3" placeholder="예:\n102 S 0600\n102 S 0610"></textarea></label></div>'+
    '<div class="section" style="padding:10px"><div class="row" style="justify-content:space-between"><h2>단가 (선택)</h2><label class="field" style="min-width:170px"><span>비슷한 품목 단가 복사</span><select name="copy"></select></label></div>'+
    '<div class="form"><label class="field"><span>적용일</span><input type="date" name="from"></label><label class="field"><span>사유</span><select name="reason">'+REASONS.map(r=>opt(r,r,r==='신규 거래')).join('')+'</select></label></div>'+
    '<p class="hint">비워 둔 거래처는 거래 없음으로 남습니다. 모델을 여러 개 적으면 같은 단가가 모두에 들어갑니다.</p>'+
    '<div class="tbl-wrap"><table class="list pricein"><thead><tr><th>거래처</th><th class="r">단가</th><th>통화</th></tr></thead><tbody>'+
    (buys.length?'<tr><td colspan="3" class="hint"><b>매입처</b></td></tr>'+buys.map(priceRow).join(''):'')+
    (sells.length?'<tr><td colspan="3" class="hint"><b>판매처</b></td></tr>'+sells.map(priceRow).join(''):'')+
    '</tbody></table></div></div>'+
    '<p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">추가</button></div></form>',f=>{
    const err=m=>{f.querySelector('[data-err]').textContent=m;return false};
    const c1=f.c1sel.value===NEW?f.c1new.value.trim():f.c1sel.value;
    const c2=f.c2sel.value===NEW?f.c2new.value.trim():f.c2sel.value;
    if(!c1)return err('대구분을 고르거나 새 이름을 입력하세요.');
    if(f.c2sel.value===NEW&&!c2)return err('새 소구분 이름을 입력하세요.');
    let models=f.models.value.split(/\n/).map(s=>s.replace(/\s+/g,' ').trim()).filter(Boolean);
    models=[...new Set(models)];
    if(!models.length){if(!c2)return err('모델을 입력하세요. (모델 없이 추가하려면 소구분이 있어야 합니다)');models=['']}
    const dup=models.filter(m=>A.DB.items.some(x=>x.cat1===c1&&x.cat2===c2&&x.model===m));
    if(dup.length)return err('이미 있는 품목입니다: '+(dup.join(', ')||'(모델 없음)'));
    const prices=[];let bad=null;
    f.querySelectorAll('[data-pc]').forEach(inp=>{const raw=inp.value.trim().replace(/,/g,'');if(!raw)return;const k=inp.dataset.pc;
      if(raw==='-'){prices.push({col:k,v:null,t:'-',cur:'KRW'});return}const v=Number(raw);if(isNaN(v)){bad=A.COL.get(k).name;return}
      prices.push({col:k,v,t:null,cur:f.querySelector('[data-pcur="'+k+'"]').value})});
    if(bad)return err(bad+' 단가를 숫자로 입력하세요.');
    if(prices.length&&!f.from.value)return err('적용일을 입력하세요.');
    let last=null;
    A.batch('품목 추가',()=>{for(const m of models){const it=A.addItem({cat1:c1,cat2:c2,model:m});last=it;
      for(const pr of prices)A.addEntry({item:it.id,col:pr.col,v:pr.v,t:pr.t,cur:pr.cur,from:f.from.value,reason:f.reason.value,note:''})}});
    A.toast((models.length>1?models.length+'개 품목':(models[0]||c2)+' 품목')+'을 추가했습니다'+(prices.length?' (단가 '+prices.length+'곳)':'')+'. 드라이브 저장 전입니다.');
    if(preset.onDone)preset.onDone(last,models.length);
  },'wide');
  const f=document.querySelector('#modal form.additem');
  f.from.value=preset.date||today();
  const fillC2=()=>{
    const isNew1=f.c1sel.value===NEW;f.querySelector('[data-new1]').hidden=!isNew1;
    const c2s=isNew1?[]:[...new Set(A.DB.items.filter(i=>i.cat1===f.c1sel.value).map(i=>i.cat2))];
    const want=preset.cat2!=null&&c2s.includes(preset.cat2)?preset.cat2:(c2s[0]!=null?c2s[0]:NEW);
    f.c2sel.innerHTML=c2s.map(c=>opt(c,c||'(소구분 없음)',c===want)).join('')+opt(NEW,'+ 새 소구분',want===NEW)+(c2s.includes('')?'':opt('','(소구분 없음)',false));
    fillC2b();
  };
  const fillC2b=()=>{
    f.querySelector('[data-new2]').hidden=f.c2sel.value!==NEW;
    const sib=f.c1sel.value===NEW?[]:A.DB.items.filter(i=>i.cat1===f.c1sel.value&&(f.c2sel.value===NEW||i.cat2===f.c2sel.value));
    f.copy.innerHTML='<option value="">복사 안 함</option>'+sib.map(i=>opt(i.id,A.itemName(i),false)).join('');
  };
  f.c1sel.addEventListener('change',()=>{preset.cat2=null;fillC2()});
  f.c2sel.addEventListener('change',fillC2b);
  f.copy.addEventListener('change',()=>{const id=f.copy.value;const d=f.from.value||today();
    f.querySelectorAll('[data-pc]').forEach(inp=>{const k=inp.dataset.pc;const e=id?A.entryAt(id,k,d):null;inp.value=e?(isNum(e)?String(e.v):(e.t||'')):'';A.setCur(f.querySelector('[data-pcur="'+k+'"]'),e&&e.cur?e.cur:'KRW')})});
  fillC2();
  if(f.c1sel.value===NEW)f.c1new.focus();else f.models.focus();
};
A.addColDialog=function(){
  openModal('<form><h2>거래처 추가</h2><label class="field"><span>거래처 이름</span><input type="text" name="nm" required></label><label class="field"><span>구분</span><select name="g"><option value="sell">판매처</option><option value="buy">매입처</option></select></label><p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">추가</button></div></form>',f=>{
    const name=f.nm.value.trim();if(A.DB.columns.some(c=>c.name===name)){f.querySelector('[data-err]').textContent='같은 이름의 거래처가 있습니다.';return false}
    A.addColumn({name,group:f.g.value});A.toast(name+' 거래처를 추가했습니다.');
  });
};

/* ---------- 엑셀 일괄 변경 ---------- */
const normName=s=>String(s==null?'':s).replace(/\s+/g,'').replace(/매입가$/,'');
const cellStr=e=>!e?null:isNum(e)?(e.cur&&e.cur!=='KRW'?e.v+' '+e.cur:e.v):(e.t||null);
/* parse "3.75", "3,750", "EUR 3.75", "3.75 EUR", "9,000원", "-" */
function parsePrice(raw){
  if(raw==null)return null;
  if(typeof raw==='number')return{v:raw,t:null,cur:null};
  let s=String(raw).trim();if(!s)return null;
  if(s==='-')return{v:null,t:'-',cur:'KRW'};
  const m=/^([A-Za-z]{3}|원|₩|€|\$|£)?\s*([+-]?[\d,]*\.?\d+)\s*([A-Za-z]{3}|원)?$/.exec(s.replace(/\s+/g,' '));
  if(!m)return{bad:true,raw:s};
  const tag=m[1]||m[3]||'';let cur=null;
  if(tag){cur=tag==='원'||tag==='₩'?'KRW':tag==='€'?'EUR':tag==='$'?'USD':tag==='£'?'GBP':A.normCur(tag);if(!cur)return{bad:true,raw:s}}
  return{v:Number(m[2].replace(/,/g,'')),t:null,cur};
}
A.exportTemplate=async function(d){
  const X=await A.loadXlsx();const cols=A.DB.columns;
  const head=['대구분','소구분','MODEL'].concat(cols.map(c=>c.name+(c.group==='buy'?' 매입가':'')),['품목ID (수정 금지)']);
  const rows=A.DB.items.map(it=>[it.cat1,it.cat2,it.model].concat(cols.map(c=>cellStr(A.entryAt(it.id,c.key,d))),[it.id]));
  const ws=X.utils.aoa_to_sheet([head].concat(rows));
  for(let r=1;r<=rows.length;r++)for(let c=3;c<3+cols.length;c++){const a=X.utils.encode_cell({r,c});if(ws[a]&&ws[a].t==='n')ws[a].z='#,##0.##'}
  ws['!cols']=[{wch:12},{wch:12},{wch:20}].concat(cols.map(()=>({wch:11})),[{wch:16}]);
  ws['!autofilter']={ref:X.utils.encode_range({s:{r:0,c:0},e:{r:rows.length,c:head.length-1}})};
  const guide=[['단가 일괄 변경 양식 사용법'],[],['기준일',d],['만든 시각',stampStr()],[],
    ['1','"단가입력" 시트에서 바꿀 단가만 고칩니다. 나머지 칸은 그대로 두세요.'],
    ['2','외화 단가는 "3.9 EUR", "1.3 RMB", "12 USD"처럼 통화를 함께 적습니다. 통화를 안 적으면 그 칸의 기존 통화(없으면 원)로 봅니다.'],
    ['3','새 품목은 맨 아래에 행을 추가하고 대구분·소구분·MODEL과 단가를 적습니다. 품목ID 칸은 비워 둡니다.'],
    ['4','빈칸은 "변경 없음"으로 처리됩니다. 거래를 끊으려면 앱의 미리보기에서 따로 선택하세요.'],
    ['5','행 순서를 바꾸거나 정렬·필터를 써도 됩니다. 품목ID로 품목을 찾습니다.'],
    ['6','거래처(열) 이름은 바꾸지 마세요. 새 거래처는 앱에서 먼저 추가한 뒤 양식을 다시 받으세요.'],
    ['7','앱의 "일괄 변경 → 양식 올리기"로 올리면 바뀐 칸만 미리 보여 줍니다. 적용일과 사유를 정하고 반영합니다.']];
  const gs=X.utils.aoa_to_sheet(guide);gs['!cols']=[{wch:10},{wch:100}];
  const wb=X.utils.book_new();X.utils.book_append_sheet(wb,ws,'단가입력');X.utils.book_append_sheet(wb,gs,'사용법');
  A.saveBlob('단가_일괄변경양식_'+d+'.xlsx',new Blob([X.write(wb,{type:'array',bookType:'xlsx'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
};
A.parseBulk=function(X,wb,d){
  const name=wb.SheetNames.includes('단가입력')?'단가입력':wb.SheetNames[0];const ws=wb.Sheets[name];
  if(!ws||!ws['!ref'])throw new Error('엑셀 시트가 비어 있습니다.');
  const aoa=X.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
  let hr=-1;for(let r=0;r<Math.min(aoa.length,10);r++){const row=aoa[r]||[];if(normName(row[2]).toUpperCase()==='MODEL'||normName(row[0])==='대구분'){hr=r;break}}
  if(hr<0)throw new Error('머리글 행(대구분 · 소구분 · MODEL)을 찾지 못했습니다. 앱에서 받은 양식을 쓰세요.');
  const head=aoa[hr];const byName=new Map(A.DB.columns.map(c=>[normName(c.name),c]));
  const colMap=[],unknownCols=[];let idCol=-1;
  for(let c=3;c<head.length;c++){const h=head[c];if(h==null||String(h).trim()==='')continue;
    if(/품목ID/.test(String(h))){idCol=c;continue}
    const col=byName.get(normName(h));if(col)colMap.push({c,col});else unknownCols.push(norm(h))}
  const changes=[],newItems=[],bad=[],blanks=[],seen=new Set();let p1='',p2='',same=0;
  for(let r=hr+1;r<aoa.length;r++){
    const row=aoa[r]||[];const a=row[0],b=row[1],md=row[2];
    const hasPrice=colMap.some(({c})=>row[c]!=null&&String(row[c]).trim()!=='');
    if(a==null&&b==null&&md==null&&!hasPrice)continue;
    const cat1=a!=null&&String(a).trim()?norm(a):p1;const cat2=b!=null&&String(b).trim()?norm(b):(a!=null&&String(a).trim()?'':p2);
    p1=cat1;p2=cat2;const model=md==null?'':norm(md);
    let it=null;const id=idCol>=0&&row[idCol]!=null?String(row[idCol]).trim():'';
    if(id&&A.ITEM.get(id))it=A.ITEM.get(id);
    if(!it)it=A.DB.items.find(x=>x.cat1===cat1&&x.cat2===cat2&&x.model===model)||null;
    if(it&&seen.has(it.id)){bad.push({row:r+1,col:'',raw:'같은 품목이 두 번 나옵니다: '+A.itemName(it)});continue}
    if(it)seen.add(it.id);
    if(!it){
      if(!cat1||(!model&&!cat2)){if(hasPrice)bad.push({row:r+1,col:'',raw:'대구분이나 MODEL이 없는 행'});continue}
      const prices=[];
      for(const {c,col} of colMap){const pv=parsePrice(row[c]);if(!pv)continue;if(pv.bad){bad.push({row:r+1,col:col.name,raw:pv.raw});continue}prices.push({col,v:pv.v,t:pv.t,cur:pv.cur||'KRW'})}
      newItems.push({cat1,cat2,model,prices,row:r+1});continue;
    }
    for(const {c,col} of colMap){
      const cur=A.entryAt(it.id,col.key,d);const pv=parsePrice(row[c]);
      if(!pv){if(A.hasVal(cur))blanks.push({it,col,cur});continue}
      if(pv.bad){if(cur&&!isNum(cur)&&cur.t&&norm(cur.t)===norm(pv.raw)){same++;continue}bad.push({row:r+1,col:col.name,raw:pv.raw});continue}
      const next={v:pv.v,t:pv.t,cur:pv.cur||(cur&&cur.cur)||'KRW'};
      if(cur&&A.hasVal(cur)&&A.sameVal(cur,next)){same++;continue}
      changes.push({it,col,cur:A.hasVal(cur)?cur:null,next});
    }
  }
  return{sheet:name,changes,newItems,bad,blanks,unknownCols,same};
};
A.bulkDialog=function(opts){
  opts=opts||{};const d=opts.date||today();
  openModal('<form class="bulk" autocomplete="off"><h2>엑셀로 단가 일괄 변경</h2>'+
    '<div class="section" style="padding:10px"><b>1. 양식 받기</b><p class="hint">'+esc(d)+' 기준 단가가 채워진 엑셀입니다. 바꿀 칸만 고치고, 새 품목은 맨 아래 행에 추가합니다. 외화는 "3.9 EUR"처럼 적습니다.</p><div class="row"><button class="btn" type="button" data-bk="tpl">양식 내려받기</button></div></div>'+
    '<div class="section" style="padding:10px"><b>2. 고친 양식 올리기</b><p class="hint">앱에서 받은 양식이나 "엑셀로 내보내기"로 받은 단가표를 올릴 수 있습니다.</p><div class="row"><label class="btn primary" for="bulkIn">파일 선택</label><input type="file" id="bulkIn" accept=".xlsx,.xls" hidden><span class="hint" data-fname></span></div></div>'+
    '<div data-preview></div><p class="hint up" data-err></p>'+
    '<div class="row-end"><button class="btn" type="button" data-cancel>닫기</button><button class="btn primary" type="submit" data-apply hidden>반영하기</button></div></form>',f=>{
    const err=m=>{f.querySelector('[data-err]').textContent=m;return false};
    if(!res)return false;
    const from=f.querySelector('[name=from]').value,reason=f.querySelector('[name=reason]').value,note=f.querySelector('[name=note]').value.trim();
    if(!from)return err('적용일을 입력하세요.');
    const pick=sel=>[...f.querySelectorAll(sel)].filter(x=>x.checked).map(x=>Number(x.value));
    const ch=pick('[data-ck="c"]'),ni=pick('[data-ck="n"]'),bl=pick('[data-ck="b"]');
    if(!ch.length&&!ni.length&&!bl.length)return err('반영할 항목을 하나 이상 선택하세요.');
    let n=0,k=0;
    A.batch('엑셀 일괄 변경',()=>{
    for(const i of ch){const c=res.changes[i];A.addEntry({item:c.it.id,col:c.col.key,v:c.next.v,t:c.next.t,cur:c.next.cur,from,reason,note:note||'엑셀 일괄 변경'});n++}
    for(const i of bl){const c=res.blanks[i];A.addEntry({item:c.it.id,col:c.col.key,v:null,t:null,cur:c.cur.cur||'KRW',from,reason:'거래 중단',note:note||'엑셀 일괄 변경'});n++}
    for(const i of ni){const x=res.newItems[i];if(A.DB.items.some(y=>y.cat1===x.cat1&&y.cat2===x.cat2&&y.model===x.model))continue;const it=A.addItem({cat1:x.cat1,cat2:x.cat2,model:x.model});k++;for(const p of x.prices){A.addEntry({item:it.id,col:p.col.key,v:p.v,t:p.t,cur:p.cur,from,reason:'신규 거래',note:note||'엑셀 일괄 변경'});n++}}
    });
    A.toast('단가 '+n+'건'+(k?', 새 품목 '+k+'개':'')+'를 반영했습니다. 드라이브에 저장하세요.');
  },'wide');
  const f=document.querySelector('#modal form.bulk');let res=null;
  f.querySelector('[data-bk="tpl"]').addEventListener('click',()=>A.exportTemplate(d).then(()=>A.toast('양식을 내려받았습니다.')).catch(e=>A.toast(e.message)));
  f.querySelector('#bulkIn').addEventListener('change',ev=>{
    const file=ev.target.files&&ev.target.files[0];if(!file)return;ev.target.value='';f.querySelector('[data-fname]').textContent=file.name;f.querySelector('[data-err]').textContent='';
    const rd=new FileReader();rd.onload=async()=>{try{const X=await A.loadXlsx();res=A.parseBulk(X,X.read(rd.result,{type:'array'}),d);renderPreview()}catch(e){res=null;f.querySelector('[data-preview]').innerHTML='';f.querySelector('[data-apply]').hidden=true;f.querySelector('[data-err]').textContent=e.message}};rd.readAsArrayBuffer(file);
  });
  function renderPreview(){
    const R=res;const up=R.changes.filter(c=>{const x=A.delta(c.cur,{v:c.next.v,cur:c.next.cur});return x&&x.d>0}).length,down=R.changes.filter(c=>{const x=A.delta(c.cur,{v:c.next.v,cur:c.next.cur});return x&&x.d<0}).length;
    const imp=A.impact(R.changes.map(c=>({item:c.it.id,col:c.col.key,v:c.next.v,t:c.next.t,cur:c.next.cur,from:d})),d);
    let h='<div class="section" style="padding:10px"><b>3. 미리보기</b><div class="stats"><span class="stat">바뀐 칸<b>'+R.changes.length+'</b></span><span class="stat">인상<b class="up">'+up+'</b></span><span class="stat">인하<b class="down">'+down+'</b></span><span class="stat">새 품목<b>'+R.newItems.length+'</b></span><span class="stat">그대로<b>'+R.same+'</b></span>'+(imp.neg1!==imp.neg0?'<span class="stat">역마진<b class="'+(imp.neg1>imp.neg0?'up':'')+'">'+imp.neg0+' → '+imp.neg1+'</b></span>':'')+'</div>';
    if(R.unknownCols.length)h+='<p class="hint up">모르는 거래처 열은 반영하지 않습니다: '+esc(R.unknownCols.join(', '))+'. 앱에서 거래처를 먼저 추가하세요.</p>';
    if(R.bad.length)h+='<p class="hint up">읽지 못한 칸 '+R.bad.length+'개 (반영 안 함): '+esc(R.bad.slice(0,8).map(b=>b.row+'행'+(b.col?' '+b.col:'')+' "'+b.raw+'"').join(', '))+(R.bad.length>8?' 외':'')+'</p>';
    const tbl=(rows,kind)=>'<div class="tbl-wrap" style="max-height:300px;overflow:auto"><table class="list"><thead><tr><th><input type="checkbox" data-all="'+kind+'" checked aria-label="전체 선택"></th><th>모델</th><th>거래처</th><th class="r">지금</th><th class="r">새 단가</th><th class="r">증감</th></tr></thead><tbody>'+rows+'</tbody></table></div>';
    if(R.changes.length)h+='<h2 style="margin-top:6px">바뀐 단가</h2>'+tbl(R.changes.map((c,i)=>{const nx={v:c.next.v,t:c.next.t,cur:c.next.cur};return '<tr><td><input type="checkbox" data-ck="c" value="'+i+'" checked></td><td class="num">'+esc(A.itemName(c.it))+'<br><span class="hint">'+esc(A.itemLabel(c.it))+'</span></td><td>'+esc(c.col.name)+'</td><td class="r num">'+esc(A.eText(c.cur))+'</td><td class="r num"><b>'+esc(A.eText(nx))+'</b></td><td class="r">'+A.deltaHtml(A.delta(c.cur,nx))+(A.isBig(c.cur,nx)?' <span class="tag warn">큰 변동</span>':'')+'</td></tr>'}).join(''),'c');
    if(R.newItems.length)h+='<h2 style="margin-top:6px">새 품목</h2><div class="tbl-wrap"><table class="list"><tbody>'+R.newItems.map((x,i)=>'<tr><td><input type="checkbox" data-ck="n" value="'+i+'" checked></td><td>'+esc([x.cat1,x.cat2].filter(Boolean).join(' › '))+' › <b class="num">'+esc(x.model||'(모델 없음)')+'</b><br><span class="hint">단가 '+x.prices.length+'곳: '+esc(x.prices.slice(0,6).map(p=>p.col.name+' '+A.eText(p)).join(', '))+(x.prices.length>6?' …':'')+'</span></td></tr>').join('')+'</tbody></table></div>';
    if(R.blanks.length)h+='<details style="margin-top:6px"><summary class="hint">비어 있는 칸 '+R.blanks.length+'개는 변경 없음으로 둡니다. 거래 없음으로 바꿀 칸이 있으면 펼쳐서 고르세요.</summary>'+tbl(R.blanks.map((c,i)=>'<tr><td><input type="checkbox" data-ck="b" value="'+i+'"></td><td class="num">'+esc(A.itemName(c.it))+'</td><td>'+esc(c.col.name)+'</td><td class="r num">'+esc(A.eText(c.cur))+'</td><td class="r">거래 없음</td><td></td></tr>').join(''),'b').replace('data-all="b" checked','data-all="b"')+'</details>';
    const any=R.changes.length||R.newItems.length||R.blanks.length;
    if(any)h+='<div class="form" style="margin-top:8px"><label class="field"><span>적용일</span><input type="date" name="from" value="'+esc(d)+'"></label><label class="field"><span>사유</span><select name="reason">'+REASONS.map(r=>'<option'+(r==='연간 단가 협상'?' selected':'')+'>'+esc(r)+'</option>').join('')+'</select></label><label class="field full"><span>비고</span><input type="text" name="note" placeholder="예: 2027년 단가 협상 결과"></label></div>';
    else h+='<p class="hint">기존 단가와 다른 칸이 없습니다.</p>';
    f.querySelector('[data-preview]').innerHTML=h+'</div>';
    f.querySelector('[data-apply]').hidden=!any;
    f.querySelectorAll('[data-all]').forEach(a=>a.addEventListener('change',()=>f.querySelectorAll('[data-ck="'+a.dataset.all+'"]').forEach(x=>x.checked=a.checked)));
  }
};

/* ---------- 품목 · 거래처 수정 / 삭제 / 복원 ---------- */
const NEWV='__new__';
const optH=(v,l,sel)=>'<option value="'+esc(v)+'"'+(sel?' selected':'')+'>'+esc(l)+'</option>';
const armBtn=(btn,label,fn)=>{if(btn.dataset.armed){fn();return}btn.dataset.armed='1';const o=btn.textContent;btn.textContent=label;setTimeout(()=>{if(btn.isConnected){delete btn.dataset.armed;btn.textContent=o}},5000)};
A.editItemDialog=function(id,opts){
  opts=opts||{};const it=A.ITEM.get(id);if(!it)return;
  const nPrices=A.DB.columns.filter(c=>A.hasVal(A.entryAt(it.id,c.key,today()))).length,nHist=A.DB.entries.filter(e=>e.item===id&&!e.base).length;
  openModal('<form class="edititem" autocomplete="off"><h2>품목 수정</h2><p class="hint">'+esc(itemPath(it))+'</p>'+
    '<div class="form"><label class="field"><span>대구분</span><select name="c1sel">'+A.cats().map(c=>optH(c,c,c===it.cat1)).join('')+optH(NEWV,'+ 새 대구분',false)+'</select></label>'+
    '<label class="field" data-new1 hidden><span>새 대구분 이름</span><input type="text" name="c1new"></label>'+
    '<label class="field"><span>소구분</span><select name="c2sel"></select></label>'+
    '<label class="field" data-new2 hidden><span>새 소구분 이름</span><input type="text" name="c2new"></label>'+
    '<label class="field full"><span>MODEL</span><input type="text" name="model" value="'+esc(it.model)+'"></label></div>'+
    '<p class="hint">대구분·소구분을 바꾸면 그 구분으로 옮겨집니다. 단가와 이력은 그대로 따라갑니다.</p>'+
    '<div class="section" style="padding:10px"><b>품목 삭제</b><p class="hint">단종 등으로 더 이상 쓰지 않는 품목을 단가표에서 뺍니다. 현재 단가 '+nPrices+'곳, 변경 이력 '+nHist+'건은 지우지 않고 보관하며, 관리 화면에서 복원할 수 있습니다.</p><div class="row"><button class="btn danger" type="button" data-del>이 품목 삭제</button></div></div>'+
    '<p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">저장</button></div></form>',f=>{
    const err=m=>{f.querySelector('[data-err]').textContent=m;return false};
    const c1=f.c1sel.value===NEWV?norm(f.c1new.value):f.c1sel.value,c2=f.c2sel.value===NEWV?norm(f.c2new.value):f.c2sel.value,md=norm(f.model.value);
    if(!c1)return err('대구분을 입력하세요.');if(f.c2sel.value===NEWV&&!c2)return err('새 소구분 이름을 입력하세요.');if(!md&&!c2)return err('MODEL을 입력하세요.');
    if(c1===it.cat1&&c2===it.cat2&&md===it.model)return;
    if(A.findDupItem(c1,c2,md,it.id))return err('같은 구분에 같은 MODEL이 이미 있습니다.');
    A.updateItem(id,{cat1:c1,cat2:c2,model:md});A.toast('품목을 수정했습니다. 드라이브 저장 전입니다.');if(opts.onDone)opts.onDone();
  },'wide');
  const f=document.querySelector('#modal form.edititem');
  const fill2=keep=>{const n1=f.c1sel.value===NEWV;f.querySelector('[data-new1]').hidden=!n1;
    const c2s=n1?[]:[...new Set(A.DB.items.filter(i=>i.cat1===f.c1sel.value).map(i=>i.cat2))];
    const want=keep&&c2s.includes(it.cat2)?it.cat2:(c2s.includes('')?'':(c2s[0]!=null?c2s[0]:NEWV));
    f.c2sel.innerHTML=c2s.map(c=>optH(c,c||'(소구분 없음)',c===want)).join('')+(c2s.includes('')?'':optH('','(소구분 없음)',want===''))+optH(NEWV,'+ 새 소구분',want===NEWV);
    f.querySelector('[data-new2]').hidden=f.c2sel.value!==NEWV};
  f.c1sel.addEventListener('change',()=>fill2(false));f.c2sel.addEventListener('change',()=>{f.querySelector('[data-new2]').hidden=f.c2sel.value!==NEWV});fill2(true);
  f.querySelector('[data-del]').addEventListener('click',ev=>armBtn(ev.currentTarget,'한 번 더 누르면 삭제',()=>{A.removeItem(id);A.closeModal();A.toast(A.itemName(it)+' 품목을 삭제했습니다. 관리 화면에서 복원할 수 있습니다.');if(opts.onDelete)opts.onDelete()}));
  f.model.focus();
};
A.editColDialog=function(key){
  const c=A.COL.get(key);if(!c)return;
  const n=A.DB.items.filter(it=>A.hasVal(A.entryAt(it.id,key,today()))).length;
  openModal('<form class="editcol" autocomplete="off"><h2>거래처 수정</h2>'+
    '<label class="field"><span>거래처 이름</span><input type="text" name="nm" value="'+esc(c.name)+'" required></label>'+
    '<label class="field"><span>구분</span><select name="g">'+optH('sell','판매처',c.group!=='buy')+optH('buy','매입처',c.group==='buy')+'</select></label>'+
    '<div class="section" style="padding:10px"><b>거래처 삭제</b><p class="hint">폐업이나 거래 종료로 없어진 거래처를 단가표, 양식, 마진에서 뺍니다. 지금 단가가 있는 품목 '+n+'개와 과거 기록은 지우지 않고 보관하며, 관리 화면에서 복원할 수 있습니다.</p><div class="row"><button class="btn danger" type="button" data-del>이 거래처 삭제</button></div></div>'+
    '<p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">저장</button></div></form>',f=>{
    const nm=norm(f.nm.value);if(!nm){f.querySelector('[data-err]').textContent='이름을 입력하세요.';return false}
    if(A.DB.columns.some(x=>x.key!==key&&x.name===nm)){f.querySelector('[data-err]').textContent='같은 이름의 거래처가 있습니다.';return false}
    if(nm===c.name&&f.g.value===c.group)return;
    A.updateColumn(key,{name:nm,group:f.g.value});A.toast('거래처를 수정했습니다. 드라이브 저장 전입니다.');
  });
  const f=document.querySelector('#modal form.editcol');
  f.querySelector('[data-del]').addEventListener('click',ev=>armBtn(ev.currentTarget,'한 번 더 누르면 삭제',()=>{A.removeColumn(key);A.closeModal();A.toast(c.name+' 거래처를 삭제했습니다. 관리 화면에서 복원할 수 있습니다.')}));
};
A.renameCatDialog=function(level,cat1,name){
  const cnt=A.DB.items.filter(it=>level===1?it.cat1===name:(it.cat1===cat1&&it.cat2===name)).length;
  openModal('<form class="renamecat" autocomplete="off"><h2>'+(level===1?'대구분':'소구분')+' 이름 변경</h2><p class="hint">'+esc(level===2?cat1+' › ':'')+esc(name||'(소구분 없음)')+' · 품목 '+cnt+'개가 함께 바뀝니다. 이미 있는 이름으로 바꾸면 그 구분으로 합쳐집니다.</p>'+
    '<label class="field"><span>새 이름</span><input type="text" name="nm" value="'+esc(name)+'"></label><p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">바꾸기</button></div></form>',f=>{
    const nm=norm(f.nm.value);if(nm===name)return;
    try{const n=A.renameCat(level,cat1,name,nm);A.toast((level===1?'대구분':'소구분')+' 이름을 바꿨습니다 (품목 '+n+'개).')}catch(e){f.querySelector('[data-err]').textContent=e.message;return false}
  });
};
/* 관리 화면: 거래처 목록, 구분 이름, 삭제된 항목 복원, 관리 기록 */
A.mountManage=function(el){
  function refresh(){
    const t=today();
    const cols=A.DB.columns.map(c=>{const n=A.DB.items.filter(it=>A.hasVal(A.entryAt(it.id,c.key,t))).length;return '<tr><td>'+esc(c.name)+'</td><td><span class="tag'+(c.group==='buy'?' base':'')+'">'+(c.group==='buy'?'매입':'판매')+'</span></td><td class="r num">'+n+'</td><td class="r"><button class="btn small" type="button" data-mc="'+esc(c.key)+'">수정 · 삭제</button></td></tr>'}).join('');
    let cats='';
    for(const c1 of A.cats()){const its=A.DB.items.filter(i=>i.cat1===c1);const c2s=[...new Set(its.map(i=>i.cat2))];
      cats+='<tr><td><b>'+esc(c1)+'</b></td><td class="r num">'+its.length+'</td><td class="r"><button class="btn small" type="button" data-mk="1" data-c1="'+esc(c1)+'">이름 변경</button></td></tr>';
      for(const c2 of c2s){if(!c2&&c2s.length===1)continue;cats+='<tr><td style="padding-left:22px">› '+esc(c2||'(소구분 없음)')+'</td><td class="r num">'+its.filter(i=>i.cat2===c2).length+'</td><td class="r">'+(c2?'<button class="btn small" type="button" data-mk="2" data-c1="'+esc(c1)+'" data-c2="'+esc(c2)+'">이름 변경</button>':'')+'</td></tr>'}}
    const ri=A.DB.removedItems,rc=A.DB.removedColumns;
    const rem=rc.map(c=>'<tr><td><span class="tag">거래처</span> '+esc(c.name)+'</td><td class="hint num">'+esc(shortDT(c.upd?localDT(c.upd):''))+'</td><td class="r"><button class="btn small" type="button" data-rc="'+esc(c.key)+'">복원</button></td></tr>').join('')+
      ri.map(it=>'<tr><td><span class="tag">품목</span> '+esc(itemPath(it))+'</td><td class="hint num">'+esc(it.upd?localDT(it.upd):'')+'</td><td class="r"><button class="btn small" type="button" data-ri="'+esc(it.id)+'">복원</button></td></tr>').join('');
    const log=A.DB.log.slice().sort((a,b)=>a.at<b.at?1:-1).slice(0,40);
    el.innerHTML='<div class="section"><div class="row" style="justify-content:space-between"><h2>거래처</h2><button class="btn small" type="button" data-addcol>거래처 추가</button></div><p class="hint">숫자는 오늘 단가가 있는 품목 수입니다. 이름 변경, 매입·판매 구분 변경, 삭제를 할 수 있습니다.</p><div class="tbl-wrap"><table class="list"><thead><tr><th>이름</th><th>구분</th><th class="r">품목</th><th></th></tr></thead><tbody>'+cols+'</tbody></table></div></div>'+
      '<div class="section"><h2>대구분 · 소구분</h2><p class="hint">이름을 바꾸면 그 구분의 품목이 모두 함께 바뀝니다. 품목 하나만 옮기려면 품목 수정에서 바꾸세요.</p><div class="tbl-wrap"><table class="list"><thead><tr><th>구분</th><th class="r">품목</th><th></th></tr></thead><tbody>'+(cats||'<tr><td colspan="3" class="hint">품목이 없습니다.</td></tr>')+'</tbody></table></div></div>'+
      '<div class="section"><h2>삭제된 거래처 · 품목</h2>'+(rem?'<p class="hint">복원하면 단가표에 다시 나오고, 보관된 단가와 이력도 그대로 돌아옵니다.</p><div class="tbl-wrap"><table class="list"><tbody>'+rem+'</tbody></table></div>':'<p class="hint">삭제된 항목이 없습니다.</p>')+'</div>'+
      '<div class="section"><h2>관리 기록</h2>'+(log.length?'<div class="tbl-wrap"><table class="list"><tbody>'+log.map(l=>'<tr><td class="num hint" style="white-space:nowrap">'+esc(shortDT(l.at))+'</td><td>'+esc(l.text)+'</td></tr>').join('')+'</tbody></table></div>':'<p class="hint">품목·거래처의 이름 변경, 삭제, 복원이 여기에 남습니다.</p>')+'</div>';
  }
  el.addEventListener('click',ev=>{
    const b=ev.target.closest('button');if(!b)return;
    try{
      if(b.dataset.mc)A.editColDialog(b.dataset.mc);
      else if(b.dataset.mk)A.renameCatDialog(Number(b.dataset.mk),b.dataset.c1,b.dataset.mk==='2'?b.dataset.c2:b.dataset.c1);
      else if(b.dataset.rc){A.restoreColumn(b.dataset.rc);A.toast('거래처를 복원했습니다.')}
      else if(b.dataset.ri){A.restoreItem(b.dataset.ri);A.toast('품목을 복원했습니다.')}
      else if(b.hasAttribute('data-addcol'))A.addColDialog();
    }catch(e){A.toast(e.message)}
  });
  refresh();return{refresh};
};

/* service worker for home-screen install & offline viewing */
if('serviceWorker' in navigator&&location.protocol==='https:')window.addEventListener('load',()=>navigator.serviceWorker.register('sw.js').catch(()=>{}));
window.addEventListener('beforeunload',ev=>{if(A.dirty>0&&A.settings.connected){ev.preventDefault();ev.returnValue=''}});

A.reindex();
})();

/* 단가 관리 — shared data model, Google Drive sync and shared views.
   No price data lives in this file or anywhere in the repository:
   data comes from importing 단가.xlsx and is kept in the user's Google Drive. */
(function(){
"use strict";
const A=window.Danga={};
const APP='danga-manager',SCHEMA=2;
const FOLDER='application/vnd.google-apps.folder';
const SCOPE='https://www.googleapis.com/auth/drive';
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
Object.assign(A,{$,esc,LS,today,addDays,nowIso,fmtN,shortDT,localDT,REASONS});

let toastT;
A.toast=function(msg){let t=$('#toast');if(!t){t=document.createElement('div');t.id='toast';t.className='toast';t.setAttribute('role','status');document.body.appendChild(t)}t.textContent=msg;t.hidden=false;clearTimeout(toastT);toastT=setTimeout(()=>t.hidden=true,3800)};

/* ---------- state ---------- */
const emptyDb=()=>({columns:[],items:[],entries:[],fx:[],deleted:[]});
const validDb=d=>d&&Array.isArray(d.columns)&&Array.isArray(d.items)&&Array.isArray(d.entries);
const normDb=d=>{if(!Array.isArray(d.fx))d.fx=[];if(!Array.isArray(d.deleted))d.deleted=[];return d};
A.settings=Object.assign({clientId:'',path:'회사 폴더/단가관리',prefix:'단가DB',autoSync:true,folderId:null,folderFor:null,connected:false,lastBackup:null},LS.get('danga.settings')||{});
A.saveSettings=()=>LS.set('danga.settings',A.settings);
const draft=LS.get('danga.draft');
A.DB=draft&&validDb(draft.db)?normDb(draft.db):emptyDb();
A.dirty=draft?draft.dirty||0:0;
A.sync=Object.assign({fileId:null,remoteModified:null,at:null,title:null,overwrite:false},draft&&draft.sync||{});
A.persist=()=>LS.set('danga.draft',{db:A.DB,dirty:A.dirty,sync:A.sync});
const listeners=[];
A.onChange=fn=>listeners.push(fn);
A.emit=()=>listeners.forEach(f=>{try{f()}catch(e){console.error(e)}});
A.touch=n=>{A.dirty+=n||1;A.persist();A.emit()};
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
A.eHtml=e=>{if(!e)return '';if(isNum(e))return (e.cur&&e.cur!=='KRW'?'<span class="cur">'+esc(e.cur)+'</span>':'')+esc(fmtN(e.v));return esc(e.t||'')};
A.eText=e=>{if(!e)return '없음';if(isNum(e))return fmtN(e.v)+(e.cur&&e.cur!=='KRW'?' '+e.cur:'원');return e.t?e.t:'없음'};
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
A.addEntry=e=>{e.id=uid('e');e.at=nowIso();A.DB.entries.push(e);A.reindex();A.touch()};
A.deleteEntry=id=>{A.DB.entries=A.DB.entries.filter(x=>x.id!==id);A.DB.deleted.push(id);A.reindex();A.touch()};
A.addItem=it=>{
  it.id=uid('i');let idx=-1;
  A.DB.items.forEach((x,k)=>{if(x.cat1===it.cat1&&x.cat2===it.cat2)idx=k});
  if(idx<0)A.DB.items.forEach((x,k)=>{if(x.cat1===it.cat1)idx=k});
  if(idx<0)A.DB.items.push(it);else A.DB.items.splice(idx+1,0,it);
  A.reindex();A.touch();return it;
};
A.addColumn=c=>{c.key=uid('c');if(c.group==='buy'){let k=0;A.DB.columns.forEach((x,i)=>{if(x.group==='buy')k=i+1});A.DB.columns.splice(k,0,c)}else A.DB.columns.push(c);A.reindex();A.touch();return c};

/* merge two versions: entries are immutable and carry unique ids, so a union minus tombstones is safe */
A.mergeDb=function(local,remote){
  local=normDb(clone(local));remote=normDb(clone(remote));
  const tomb=new Set([...local.deleted,...remote.deleted]);
  const byId=(mine,theirs)=>{const m=new Map();for(const o of theirs)m.set(o.id,o);for(const o of mine)m.set(o.id,o);return [...m.values()].filter(o=>!tomb.has(o.id))};
  const items=local.items.slice();const have=new Set(items.map(i=>i.id));
  for(const it of remote.items){if(have.has(it.id))continue;let idx=-1;items.forEach((x,k)=>{if(x.cat1===it.cat1&&x.cat2===it.cat2)idx=k});if(idx<0)items.forEach((x,k)=>{if(x.cat1===it.cat1)idx=k});if(idx<0)items.push(it);else items.splice(idx+1,0,it);have.add(it.id)}
  const cols=local.columns.slice();const ck=new Set(cols.map(c=>c.key));
  for(const c of remote.columns){if(ck.has(c.key))continue;if(c.group==='buy'){let k=0;cols.forEach((x,i)=>{if(x.group==='buy')k=i+1});cols.splice(k,0,c)}else cols.push(c);ck.add(c.key)}
  return{columns:cols,items:items.filter(i=>!tomb.has(i.id)),entries:byId(local.entries,remote.entries),fx:byId(local.fx,remote.fx),deleted:[...tomb]};
};
A.replaceDb=(db,opts)=>{A.DB=normDb(db);A.reindex();opts=opts||{};A.dirty=opts.dirty||0;if(opts.overwrite)A.sync.overwrite=true;A.persist();A.emit()};

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
  its.forEach((it,r)=>{const row=[it.cat1,it.cat2,it.model];cols.forEach((c,k)=>{const e=A.entryAt(it.id,c.key,d);row.push(e?(isNum(e)?e.v:(e.t||null)):null);if(isNum(e)&&e.cur&&e.cur!=='KRW')notes.push([r+1,k+3,e.cur])});aoa.push(row)});
  const merges=[];
  for(let r=0;r<its.length;){let n=1;while(r+n<its.length&&its[r+n].cat1===its[r].cat1)n++;if(n>1)merges.push({s:{r:r+1,c:0},e:{r:r+n,c:0}});r+=n}
  for(let r=0;r<its.length;){let n=1;while(r+n<its.length&&its[r+n].cat1===its[r].cat1&&its[r+n].cat2===its[r].cat2)n++;if(n>1)merges.push({s:{r:r+1,c:1},e:{r:r+n,c:1}});r+=n}
  const ws=X.utils.aoa_to_sheet(aoa);ws['!merges']=merges;
  notes.forEach(([r,c,t])=>{const a=X.utils.encode_cell({r,c});if(ws[a])ws[a].c=[{a:'단가 관리',t}]});
  for(let r=1;r<=its.length;r++)for(let c=3;c<3+cols.length;c++){const a=X.utils.encode_cell({r,c});if(ws[a]&&ws[a].t==='n')ws[a].z='#,##0.##'}
  ws['!cols']=[{wch:12},{wch:12},{wch:20}].concat(cols.map(()=>({wch:10})));
  const wb=X.utils.book_new();X.utils.book_append_sheet(wb,ws,d.slice(0,7).replace('-','.'));
  X.writeFile(wb,'단가표_'+d+'.xlsx');
};
A.downloadText=(name,text)=>{const b=new Blob([text],{type:'application/json'});const u=URL.createObjectURL(b);const a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),2000)};
A.pack=()=>JSON.stringify({app:APP,schema:SCHEMA,savedAt:nowIso(),db:A.DB});
A.unpack=text=>{let o;try{o=JSON.parse(text)}catch(e){return null}if(o&&o.app===APP&&validDb(o.db))return normDb(o.db);return null};

/* ---------- Google Drive ---------- */
const D={token:null,exp:0};
const gisReady=()=>!!(window.google&&google.accounts&&google.accounts.oauth2);
let gisP=null;
A.loadGis=()=>{if(gisReady())return Promise.resolve();if(!gisP)gisP=new Promise((res,rej)=>{const s=document.createElement('script');s.src='https://accounts.google.com/gsi/client';s.async=true;s.onload=()=>res();s.onerror=()=>{gisP=null;rej(new Error('구글 로그인 도구를 불러오지 못했습니다. 인터넷 연결을 확인하세요.'))};document.head.appendChild(s)});return gisP};
function cachedToken(){if(D.token&&Date.now()<D.exp-60000)return D.token;const s=LS.get('danga.token');if(s&&Date.now()<s.exp-60000){D.token=s.t;D.exp=s.exp;return D.token}return null}
A.hasToken=()=>!!cachedToken();
/* must be called directly from a click so the Google popup is allowed */
function getToken(interactive){
  const t=cachedToken();if(t)return Promise.resolve(t);
  if(!interactive)return Promise.reject(new Error('구글 로그인이 필요합니다.'));
  if(!A.settings.clientId)return Promise.reject(new Error('설정에서 구글 클라이언트 ID를 먼저 입력하세요. (README의 설정 방법 참고)'));
  if(!gisReady())return A.loadGis().then(()=>Promise.reject(new Error('구글 로그인 준비가 끝났습니다. 버튼을 한 번 더 눌러 주세요.')));
  return new Promise((res,rej)=>{
    const c=google.accounts.oauth2.initTokenClient({client_id:A.settings.clientId.trim(),scope:SCOPE,
      callback:r=>{if(r.error){rej(new Error('구글 로그인 실패: '+(r.error_description||r.error)));return}D.token=r.access_token;D.exp=Date.now()+(Number(r.expires_in)||3600)*1000;LS.set('danga.token',{t:D.token,exp:D.exp});res(D.token)},
      error_callback:e=>rej(new Error(e&&e.type==='popup_closed'?'구글 로그인 창이 닫혔습니다.':e&&e.type==='popup_failed_to_open'?'로그인 창이 차단됐습니다. 브라우저의 팝업 차단을 해제하세요.':'구글 로그인에 실패했습니다.'))});
    c.requestAccessToken({prompt:''});
  });
}
A.signOut=()=>{const t=D.token;D.token=null;D.exp=0;LS.del('danga.token');if(t&&gisReady())google.accounts.oauth2.revoke(t,()=>{});A.emit()};
async function api(method,url,body,headers,type){
  const t=cachedToken();if(!t)throw new Error('구글 로그인이 만료됐습니다. 다시 로그인하세요.');
  const r=await fetch(url,{method,headers:Object.assign({Authorization:'Bearer '+t},headers||{}),body});
  if(r.status===401){D.token=null;LS.del('danga.token');throw new Error('구글 로그인이 만료됐습니다. 다시 로그인하세요.')}
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
async function resolveFolder(create){
  const raw=(A.settings.path||'').trim();
  const m=raw.match(/folders\/([A-Za-z0-9_-]{10,})/);if(m)return{id:m[1],label:'폴더 링크'};
  if(/^[A-Za-z0-9_-]{25,}$/.test(raw))return{id:raw,label:'폴더 ID'};
  const segs=raw.split('/').map(s=>s.trim()).filter(Boolean);if(segs[0]==='내 드라이브'||segs[0]==='My Drive')segs.shift();
  let cur='root';
  for(let k=0;k<segs.length;k++){const id=await subFolder(cur,segs[k],create);if(!id)return{id:null,missing:segs.slice(k),found:segs.slice(0,k)};cur=id}
  return{id:cur,label:'내 드라이브'+(segs.length?' / '+segs.join(' / '):'')};
}
async function ensureFolder(create){
  if(A.settings.folderId&&A.settings.folderFor===A.settings.path)return A.settings.folderId;
  const r=await resolveFolder(create);if(!r.id)return null;
  A.settings.folderId=r.id;A.settings.folderFor=A.settings.path;A.saveSettings();return r.id;
}
const mainName=()=>(A.settings.prefix||'단가DB').trim()+'.json';
async function findMain(fid){
  const f=(await find("'"+qs(fid)+"' in parents and name = '"+qs(mainName())+"' and trashed = false"))[0];if(f)return{file:f,main:true};
  const pre=(A.settings.prefix||'단가DB').trim();
  const l=(await find("'"+qs(fid)+"' in parents and name contains '"+qs(pre)+"' and trashed = false")).filter(x=>/\.json$/i.test(x.name)&&x.mimeType!==FOLDER);
  return l[0]?{file:l[0],main:false}:null;
}
async function readDb(id){const t=await api('GET',DRV+'/'+id+'?alt=media&supportsAllDrives=true',null,null,'text');const db=A.unpack(t);if(!db)throw new Error('드라이브 파일을 단가 데이터로 읽지 못했습니다.');return db}

let busy=false;
A.isBusy=()=>busy;
async function run(fn){if(busy)throw new Error('이전 작업이 진행 중입니다.');busy=true;A.status('작업 중…');A.emit();try{return await fn()}finally{busy=false;A.emit()}}
A.statusText={msg:'',cls:''};
A.status=(msg,cls)=>{A.statusText={msg,cls:cls||''};document.querySelectorAll('[data-status]').forEach(el=>{el.textContent=msg;el.className='status-line'+(cls?' '+cls:'')})};

/* save: merge with whatever another device saved, then write the main file; one dated backup per day */
A.save=function(){
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
    A.settings.connected=true;A.saveSettings();A.persist();
    const day=today();
    if(A.settings.lastBackup!==day){
      try{const bf=await subFolder(fid,'백업',true);const bn=(A.settings.prefix||'단가DB').trim()+'_'+day.replace(/-/g,'')+'.json';const ex=(await find("'"+qs(bf)+"' in parents and name = '"+qs(bn)+"' and trashed = false"))[0];await upload(ex?{}:{name:bn,parents:[bf],mimeType:'application/json'},content,ex&&ex.id);A.settings.lastBackup=day;A.saveSettings()}catch(e){console.warn('backup',e)}
    }
    A.status('저장했습니다: '+r.name+' ('+localDT(r.modifiedTime)+')','ok');
    A.toast(merged?'다른 기기의 변경과 합쳐서 저장했습니다.':'드라이브에 저장했습니다.');
    A.emit();
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
    A.settings.connected=true;A.saveSettings();A.persist();
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
    A.persist();A.status(isMain?'불러왔습니다: '+f.name:f.name+'을(를) 불러왔습니다. 저장하면 이 내용으로 현재 파일을 덮어씁니다.','ok');A.emit();
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
  let b='<div><div class="hint">오늘 적용 중인 단가</div><div class="now"><span class="big">'+esc(cur?A.eText(cur):'없음')+'</span>'+(cur?'<span class="tag">'+esc(cur.from)+'부터</span>':'')+'</div></div>';
  const nums=h.filter(isNum).filter(e=>(e.cur||'KRW')===lastCur);
  if(nums.length>=2)b+='<div class="chart">'+stepChart(nums)+'</div>';
  b+='<div><h2 style="margin-bottom:6px">변경 이력 <span class="hint">('+h.length+'건)</span></h2>';
  if(!h.length)b+='<p class="hint">기록이 없습니다. 아래에서 첫 단가를 입력하세요.</p>';
  else{b+='<ol class="timeline">';for(const e of [...h].reverse()){const p=A.prevOf(e);
    b+='<li class="'+(e.from>t?'future':'')+'"><span class="d">'+esc(e.from)+(e.from>t?'<br>적용 예정':'')+'</span><span><span class="val">'+esc(A.eText(e))+'</span> '+(p?'<span class="hint">← '+esc(A.eText(p))+'</span> ':'')+A.deltaHtml(A.delta(p,e))+'</span>'+
    '<span class="meta">'+(e.base?'<span class="tag base">기준 단가</span>':'')+(e.reason&&!e.base?'<span class="tag">'+esc(e.reason)+'</span>':'')+(e.note?'<span>'+esc(e.note)+'</span>':'')+'<span>입력 '+esc(shortDT(e.at))+'</span>'+(e.base?'':'<button class="btn small danger" type="button" data-del="'+esc(e.id)+'">삭제</button>')+'</span></li>'}
    b+='</ol>'}
  b+='</div><form class="section" autocomplete="off" data-cellform><h2>단가 변경 기록</h2><div class="form">'+
    '<label class="field"><span>새 단가</span><input type="text" inputmode="decimal" name="v" placeholder="숫자"></label>'+
    '<label class="field"><span>통화</span><select name="cur"><option value="KRW">원</option><option value="EUR">EUR</option><option value="RMB">RMB</option><option value="USD">USD</option></select></label>'+
    '<label class="field"><span>적용일</span><input type="date" name="from" required></label>'+
    '<label class="field"><span>사유</span><select name="reason">'+REASONS.map(r=>'<option>'+esc(r)+'</option>').join('')+'</select></label>'+
    '<label class="field full"><span>비고</span><input type="text" name="note" placeholder="예: 5월 출고분부터, 견적서 번호"></label>'+
    '<label class="check full"><input type="checkbox" name="none"> 거래 없음으로 변경 (칸 비우기)</label></div>'+
    '<p class="hint" data-msg>적용일 이전 날짜로 보면 이전 단가가, 이후 날짜로 보면 새 단가가 나옵니다.</p><div class="row-end"><button class="btn primary" type="submit">기록하기</button></div></form>';
  el.innerHTML=b;
  const f=el.querySelector('[data-cellform]');f.cur.value=lastCur;const d=opts.date||t;f.from.value=d>t?d:t;
  f.addEventListener('submit',ev=>{ev.preventDefault();const msg=f.querySelector('[data-msg]');
    const none=f.none.checked,raw=f.v.value.trim().replace(/,/g,''),from=f.from.value;
    if(!from){msg.textContent='적용일을 입력하세요.';msg.className='hint up';return}
    let v=null,tx=null;
    if(!none){if(raw==='-')tx='-';else{v=Number(raw);if(raw===''||isNaN(v)){msg.textContent='새 단가를 숫자로 입력하세요. 거래를 끊는 경우 "거래 없음"을 체크하세요.';msg.className='hint up';return}}}
    const e={item:i,col:c,v,t:tx,cur:f.cur.value,from,reason:f.reason.value,note:f.note.value.trim()};
    A.addEntry(e);A.toast(from+'부터 '+A.eText(e)+'(으)로 기록했습니다. 드라이브 저장 전입니다.');A.renderCell(el,i,c,opts);
  });
  el.querySelectorAll('[data-del]').forEach(btn=>btn.addEventListener('click',()=>{
    if(btn.dataset.armed){A.deleteEntry(btn.dataset.del);A.toast('기록을 삭제했습니다.');A.renderCell(el,i,c,opts)}
    else{btn.dataset.armed='1';btn.textContent='한 번 더 누르면 삭제';setTimeout(()=>{if(btn.isConnected){delete btn.dataset.armed;btn.textContent='삭제'}},4000)}
  }));
};

function catOptions(){return '<option value="">전체</option>'+A.cats().map(c=>'<option>'+esc(c)+'</option>').join('')}
function colOptions(){return '<option value="">전체</option>'+A.DB.columns.map(c=>'<option value="'+esc(c.key)+'">'+esc(c.name)+(c.group==='buy'?' (매입)':'')+'</option>').join('')}
function keepVal(sel,html){const v=sel.value;sel.innerHTML=html;sel.value=v}

/* 이력 검색 */
A.mountHistory=function(el,opts){
  opts=opts||{};
  el.innerHTML='<div class="filters"><label class="field"><span>모델 · 구분</span><input type="search" name="q" placeholder="모델명 일부"></label><label class="field"><span>거래처</span><select name="col"></select></label><label class="field"><span>대구분</span><select name="cat"></select></label><label class="field"><span>적용일 시작</span><input type="date" name="from"></label><label class="field"><span>적용일 끝</span><input type="date" name="to"></label><label class="field"><span>사유 · 비고</span><input type="search" name="reason" placeholder="예: 원자재"></label><label class="check"><input type="checkbox" name="base"> 기준 단가 포함</label></div><div class="stats" data-stats></div><div data-out></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  function refresh(){
    keepVal(g('col'),colOptions());keepVal(g('cat'),catOptions());
    const q=g('q').value.trim(),col=g('col').value,cat=g('cat').value,f=g('from').value,t=g('to').value,rq=g('reason').value.trim().toLowerCase(),base=g('base').checked;
    const match=(e,o)=>{if(!o.base&&e.base)return false;const it=A.ITEM.get(e.item);if(!it||!A.COL.get(e.col))return false;if(o.col&&e.col!==o.col)return false;if(!A.matchItem(it,o.q,o.cat))return false;if(o.f&&e.from<o.f)return false;if(o.t&&e.from>o.t)return false;if(o.rq&&!((e.reason||'')+' '+(e.note||'')).toLowerCase().includes(o.rq))return false;return true};
    const cond={q,col,cat,f,t,rq,base};
    const list=A.DB.entries.filter(e=>match(e,cond)).sort((a,b)=>-cmpE(a,b));
    let up=0,down=0;const rows=list.map(e=>{const p=A.prevOf(e),x=A.delta(p,e);if(x&&x.d>0)up++;if(x&&x.d<0)down++;return{e,p,x}});
    const lastCh=A.DB.entries.filter(e=>!e.base).sort(cmpE).pop();
    el.querySelector('[data-stats]').innerHTML='<span class="stat">결과<b>'+list.length+'</b></span><span class="stat">인상<b class="up">'+up+'</b></span><span class="stat">인하<b class="down">'+down+'</b></span><span class="stat">마지막 변경<b>'+esc(lastCh?lastCh.from:'없음')+'</b></span>';
    const out=el.querySelector('[data-out]');
    if(!list.length){
      if(A.isEmpty()){out.innerHTML='<div class="empty"><b>아직 단가 데이터가 없습니다.</b>설정에서 단가.xlsx를 가져오거나 드라이브에서 불러오세요.</div>';return}
      const any=A.DB.entries.some(e=>!e.base);
      const active=[];const cn=A.COL.get(col);
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
    const shown=rows.slice(0,600);
    if(opts.cards){
      out.innerHTML='<div class="cards">'+shown.map(({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return '<div class="card click" data-i="'+esc(e.item)+'" data-c="'+esc(e.col)+'"><div class="top"><span class="model">'+esc(A.itemName(it))+'</span><span class="num hint">'+esc(e.from)+(e.from>today()?' <span class="tag fut">예정</span>':'')+'</span></div><div class="sub">'+esc(c.name)+' · '+esc(A.itemLabel(it))+'</div><div class="vals">'+(p?'<span class="hint">'+esc(A.eText(p))+' →</span>':'')+'<b>'+esc(A.eText(e))+'</b>'+A.deltaHtml(x)+'</div><div class="sub">'+(e.base?'<span class="tag base">기준 단가</span>':'<span class="tag">'+esc(e.reason||'')+'</span>')+(e.note?' '+esc(e.note):'')+'</div></div>'}).join('')+'</div>';
    }else{
      out.innerHTML='<div class="tbl-wrap section" style="padding:6px"><table class="list"><thead><tr><th>적용일</th><th>구분</th><th>모델</th><th>거래처</th><th class="r">이전</th><th class="r">변경</th><th class="r">증감</th><th>사유</th><th>비고</th><th>입력</th></tr></thead><tbody>'+
      shown.map(({e,p,x})=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col);return '<tr class="click" data-i="'+esc(e.item)+'" data-c="'+esc(e.col)+'"><td class="num">'+esc(e.from)+(e.from>today()?' <span class="tag fut">예정</span>':'')+'</td><td>'+esc(A.itemLabel(it))+'</td><td class="num">'+esc(it.model)+'</td><td>'+esc(c.name)+'</td><td class="r num">'+(p?esc(A.eText(p)):'<span class="hint">-</span>')+'</td><td class="r num"><b>'+esc(A.eText(e))+'</b></td><td class="r">'+A.deltaHtml(x)+'</td><td>'+(e.base?'<span class="tag base">기준 단가</span>':'<span class="tag">'+esc(e.reason||'')+'</span>')+'</td><td>'+esc(e.note||'')+'</td><td class="num hint">'+esc(shortDT(e.at))+'</td></tr>'}).join('')+'</tbody></table></div>';
    }
    if(rows.length>600)out.insertAdjacentHTML('beforeend','<p class="hint">처음 600건만 표시합니다. 조건을 좁혀 주세요.</p>');
  }
  el.addEventListener('input',refresh);
  el.addEventListener('click',ev=>{const b=ev.target.closest('[data-hx]');if(!b)return;const k=b.dataset.hx;
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
  el.innerHTML='<div class="section"><div class="filters"><h2 style="align-self:center">두 날짜 비교</h2><label class="field"><span>이전 날짜</span><input type="date" name="a"></label><label class="field"><span>이후 날짜</span><input type="date" name="b"></label><label class="field"><span>거래처</span><select name="col"></select></label></div><div class="stats" data-stats></div><div data-diff></div></div>'+
    '<div class="section"><h2>월별 변경 건수</h2><p class="hint">적용일 기준 최근 12개월, 기준 단가 제외.</p><div class="chart" data-monthly></div></div>'+
    '<div class="section"><h2>거래처별 변경 요약</h2><div class="tbl-wrap" data-bycol></div></div>'+
    '<div class="section"><h2>변경이 잦은 품목</h2><div class="tbl-wrap" data-byitem></div></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  function refresh(){
    keepVal(g('col'),colOptions());
    if(!g('a').value){const bases=A.DB.entries.filter(e=>e.base).map(e=>e.from).sort();g('a').value=bases[0]||addDays(today(),-30)}
    if(!g('b').value)g('b').value=today();
    const Ad=g('a').value,Bd=g('b').value,colF=g('col').value;const diffs=[];
    for(const it of A.DB.items)for(const c of A.DB.columns){if(colF&&c.key!==colF)continue;const a=A.entryAt(it.id,c.key,Ad),b=A.entryAt(it.id,c.key,Bd);if(!A.sameVal(a&&A.hasVal(a)?a:null,b&&A.hasVal(b)?b:null))diffs.push({it,c,a,b,x:A.delta(a,b)})}
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
  el.addEventListener('input',refresh);refresh();return{refresh};
};

/* 환율 */
A.mountFx=function(el){
  function refresh(){
    const fx=[...A.DB.fx].sort((a,b)=>a.date<b.date?1:-1);
    const inp=(f,k)=>'<input type="number" step="0.01" data-k="'+k+'" value="'+esc(f[k]==null?'':f[k])+'" style="width:100px;text-align:right">';
    el.innerHTML='<div class="row" style="justify-content:space-between"><h2>환율 (외화 매입가 환산용)</h2><button class="btn small" type="button" data-fxadd>환율 추가</button></div><p class="hint">적용일부터 다음 환율 전까지 쓰입니다. 1 EUR, 1 RMB가 몇 원인지 입력하세요.</p>'+
      (fx.length?'<div class="tbl-wrap"><table class="list"><thead><tr><th>적용일</th><th class="r">EUR</th><th class="r">RMB</th><th class="r">USD</th><th></th></tr></thead><tbody>'+fx.map(f=>'<tr data-fx="'+esc(f.id)+'"><td><input type="date" data-k="date" value="'+esc(f.date)+'" style="width:140px"></td><td class="r">'+inp(f,'EUR')+'</td><td class="r">'+inp(f,'RMB')+'</td><td class="r">'+inp(f,'USD')+'</td><td class="r"><button class="btn small danger" type="button" data-fxdel="'+esc(f.id)+'">삭제</button></td></tr>').join('')+'</tbody></table></div>'
      :'<div class="empty"><b>환율이 아직 없습니다.</b>LOCK SEAL(EUR), STOPPER(RMB) 마진을 보려면 환율을 추가하세요.</div>');
  }
  el.addEventListener('click',ev=>{
    if(ev.target.closest('[data-fxadd]')){A.DB.fx.push({id:uid('fx'),date:today(),EUR:null,RMB:null,USD:null});A.touch();refresh();return}
    const b=ev.target.closest('[data-fxdel]');if(!b)return;if(b.dataset.armed){A.DB.fx=A.DB.fx.filter(x=>x.id!==b.dataset.fxdel);A.DB.deleted.push(b.dataset.fxdel);A.touch();refresh()}else{b.dataset.armed='1';b.textContent='한 번 더'}
  });
  el.addEventListener('change',ev=>{const i=ev.target.closest('input[data-k]');if(!i)return;const f=A.DB.fx.find(x=>x.id===i.closest('tr').dataset.fx);if(!f)return;const k=i.dataset.k;f[k]=k==='date'?i.value:(i.value===''?null:Number(i.value));A.touch()});
  refresh();return{refresh};
};

/* 설정 · 저장 · 가져오기 */
A.mountSettings=function(el){
  el.innerHTML=
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
  '<div class="section"><div class="row" style="justify-content:space-between"><h2>저장 기록 · 백업</h2><button class="btn small" type="button" data-act="list">목록 보기</button></div><div class="tbl-wrap" data-files><p class="hint">목록 보기를 누르면 저장 파일과 날짜별 백업이 나옵니다.</p></div></div>'+
  '<div class="section"><h2>단가 엑셀 가져오기</h2><p class="hint">처음 시작할 때 원본 <code>단가.xlsx</code>를 가져옵니다. 첫 시트의 양식(대구분·소구분·MODEL, D열부터 거래처)을 읽고, 시트 이름(예: 2026.04)을 기준일로 씁니다. 셀 메모의 EUR·RMB도 반영합니다.</p>'+
    '<div class="row"><label class="btn" for="xlsxIn">이 기기에서 파일 선택</label><input type="file" id="xlsxIn" accept=".xlsx,.xls" hidden></div>'+
    '<div class="row"><input type="text" name="xlsxLink" placeholder="또는 드라이브 파일 링크 붙여넣기" style="flex:1 1 220px"><button class="btn" type="button" data-act="xlsxDrive">드라이브에서 가져오기</button></div>'+
    '<div data-import></div></div>'+
  '<div class="section"><h2>백업 파일 · 초기화</h2><div class="row"><button class="btn" type="button" data-act="jsonOut">백업 파일 내려받기</button><label class="btn" for="jsonIn">백업 파일 불러오기</label><input type="file" id="jsonIn" accept=".json,application/json" hidden><button class="btn danger" type="button" data-act="wipe">이 기기 데이터 지우기</button></div><p class="hint" data-bstatus>이 기기에도 임시 보관되지만, 드라이브에 저장해야 다른 기기에서 볼 수 있습니다.</p></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  const fill=()=>{for(const k of ['clientId','path','prefix'])if(document.activeElement!==g(k))g(k).value=A.settings[k]||'';g('autoSync').checked=!!A.settings.autoSync};
  fill();
  el.querySelectorAll('[data-status]').forEach(s=>{s.textContent=A.statusText.msg||(A.settings.clientId?(A.hasToken()?'구글에 로그인돼 있습니다.':'저장하거나 불러오면 구글 로그인 창이 뜹니다.'):'구글 클라이언트 ID를 입력하면 드라이브에 저장할 수 있습니다.');s.className='status-line'+(A.statusText.cls?' '+A.statusText.cls:'')});
  el.addEventListener('change',ev=>{
    const n=ev.target.name;
    if(n==='path'&&ev.target.value.trim()!==A.settings.path){A.settings.path=ev.target.value.trim();A.settings.folderId=null;A.settings.folderFor=null}
    if(n==='prefix')A.settings.prefix=ev.target.value.trim()||'단가DB';
    if(n==='clientId'){A.settings.clientId=ev.target.value.trim();if(A.settings.clientId)A.loadGis().catch(()=>{})}
    if(n==='autoSync')A.settings.autoSync=ev.target.checked;
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
function openModal(html,onSubmit){
  let m=$('#modal');if(!m){m=document.createElement('div');m.id='modal';m.className='modal';m.innerHTML='<div class="modal-card" role="dialog" aria-modal="true"></div>';document.body.appendChild(m);m.addEventListener('click',ev=>{if(ev.target===m)A.closeModal()})}
  const card=m.firstChild;card.innerHTML=html;m.hidden=false;
  const f=card.querySelector('form');f.addEventListener('submit',ev=>{ev.preventDefault();if(onSubmit(f)!==false)A.closeModal()});
  card.querySelector('[data-cancel]').addEventListener('click',A.closeModal);const first=f.querySelector('input,select');if(first)first.focus();
}
A.closeModal=()=>{const m=$('#modal');if(m){m.hidden=true;m.firstChild.innerHTML=''}};
A.addItemDialog=function(){
  openModal('<form><h2>품목 추가</h2><datalist id="dl1">'+A.cats().map(c=>'<option value="'+esc(c)+'">').join('')+'</datalist><datalist id="dl2">'+[...new Set(A.DB.items.map(i=>i.cat2).filter(Boolean))].map(c=>'<option value="'+esc(c)+'">').join('')+'</datalist>'+
    '<label class="field"><span>대구분</span><input type="text" name="c1" list="dl1" required></label><label class="field"><span>소구분 (없으면 비움)</span><input type="text" name="c2" list="dl2"></label><label class="field"><span>MODEL</span><input type="text" name="md" required></label>'+
    '<p class="hint">같은 대구분·소구분의 마지막 줄 아래에 추가됩니다. 단가는 추가한 뒤 입력하세요.</p><p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">추가</button></div></form>',f=>{
    const it={cat1:f.c1.value.trim(),cat2:f.c2.value.trim(),model:f.md.value.trim()};
    if(A.DB.items.some(x=>x.cat1===it.cat1&&x.cat2===it.cat2&&x.model===it.model)){f.querySelector('[data-err]').textContent='같은 품목이 이미 있습니다.';return false}
    A.addItem(it);A.toast(it.model+' 품목을 추가했습니다.');
  });
};
A.addColDialog=function(){
  openModal('<form><h2>거래처 추가</h2><label class="field"><span>거래처 이름</span><input type="text" name="nm" required></label><label class="field"><span>구분</span><select name="g"><option value="sell">판매처</option><option value="buy">매입처</option></select></label><p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>취소</button><button class="btn primary" type="submit">추가</button></div></form>',f=>{
    const name=f.nm.value.trim();if(A.DB.columns.some(c=>c.name===name)){f.querySelector('[data-err]').textContent='같은 이름의 거래처가 있습니다.';return false}
    A.addColumn({name,group:f.g.value});A.toast(name+' 거래처를 추가했습니다.');
  });
};

/* service worker for home-screen install & offline viewing */
if('serviceWorker' in navigator&&location.protocol==='https:')window.addEventListener('load',()=>navigator.serviceWorker.register('sw.js').catch(()=>{}));
window.addEventListener('beforeunload',ev=>{if(A.dirty>0&&A.settings.connected){ev.preventDefault();ev.returnValue=''}});

A.reindex();
})();

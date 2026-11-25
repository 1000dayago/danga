/* 단가 관리 — features built on core.js: home, tools hub, % change, customer price sheets,
   comparison, simulation, upcoming prices, tabs, shortcuts. No price data lives here. */
(function(){
"use strict";
const A=window.Danga;if(!A)return;
const {$,esc,today,addDays,fmtN,shortDT,REASONS}=A;
const isNum=A.isNum;
const pad=n=>String(n).padStart(2,'0');

/* ---------- icons ---------- */
const P={home:'M3 11l9-8 9 8M5 10v10h5v-6h4v6h5V10',table:'M3 5h18v14H3zM3 10h18M9 10v9',clock:'M12 3a9 9 0 100 18 9 9 0 000-18zM12 7v5l3 2',chart:'M4 20V10M10 20V4M16 20v-7M22 20H2',percent:'M19 5L5 19M7 5a2 2 0 100 4 2 2 0 000-4zM17 15a2 2 0 100 4 2 2 0 000-4z',tools:'M14 6l4 4-9 9H5v-4zM13 7l4 4',folder:'M3 6h6l2 2h10v11H3z',gear:'M12 8a4 4 0 100 8 4 4 0 000-8zM12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M5 19l2-2M17 7l2-2',more:'M5 12h.01M12 12h.01M19 12h.01',undo:'M9 14L4 9l5-5M4 9h10a6 6 0 010 12h-3',star:'M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z',plus:'M12 5v14M5 12h14',file:'M6 3h9l4 4v14H6zM14 3v5h5',calendar:'M4 6h16v14H4zM4 10h16M8 3v4M16 3v4',search:'M11 4a7 7 0 100 14 7 7 0 000-14zM21 21l-5-5',save:'M5 4h12l3 3v13H5zM8 4v5h7V4M8 20v-6h8v6',alert:'M12 3l10 18H2zM12 10v5M12 18h.01',export:'M12 3v12M7 8l5-5 5 5M5 21h14',users:'M9 11a3 3 0 100-6 3 3 0 000 6zM3 20c0-3 3-5 6-5s6 2 6 5M17 11a3 3 0 100-6M21 20c0-2-1-4-3-4.5',flask:'M9 3h6M10 3v6l-5 9a2 2 0 002 3h10a2 2 0 002-3l-5-9V3'};
A.icon=(n,sz)=>'<svg class="ico" viewBox="0 0 24 24" width="'+(sz||20)+'" height="'+(sz||20)+'" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="'+(P[n]||P.more)+'"/></svg>';

const noData=()=>{if(A.isEmpty()){A.toast('단가 데이터가 없습니다. 설정에서 엑셀을 가져오거나 드라이브에서 불러오세요.');return true}return false};
const keepVal=(sel,html)=>{const v=sel.value;sel.innerHTML=html;sel.value=v};
const catOpts=()=>'<option value="">전체</option>'+A.cats().map(c=>'<option>'+esc(c)+'</option>').join('');
const num=el=>{const v=Number(el.value);return isNaN(v)?0:v};

/* ---------- tabs (segmented control) ---------- */
A.mountTabs=function(el,tabs,key){
  let cur=A.LS.get('danga.tab.'+key);if(!tabs.some(t=>t.id===cur))cur=tabs[0].id;
  el.innerHTML='<div class="seg" role="tablist"></div><div class="seg-body"></div>';
  const bar=el.querySelector('.seg'),body=el.querySelector('.seg-body'),mounted={},panels={};
  bar.innerHTML=tabs.map(t=>'<button type="button" role="tab" data-t="'+t.id+'">'+esc(t.label)+'</button>').join('');
  tabs.forEach(t=>{const p=document.createElement('div');p.className='seg-panel';body.appendChild(p);panels[t.id]=p});
  function show(id){cur=id;A.LS.set('danga.tab.'+key,id);
    bar.querySelectorAll('button').forEach(b=>b.setAttribute('aria-selected',String(b.dataset.t===id)));
    for(const t of tabs)panels[t.id].hidden=t.id!==id;
    if(!mounted[id])mounted[id]=tabs.find(t=>t.id===id).mount(panels[id])||{};else if(mounted[id].refresh)mounted[id].refresh()}
  bar.addEventListener('click',ev=>{const b=ev.target.closest('button[data-t]');if(b)show(b.dataset.t)});
  show(cur);
  return{refresh(){const m=mounted[cur];if(m&&m.refresh)m.refresh()},show};
};

/* ---------- 적용 예정 단가 ---------- */
A.mountUpcoming=function(el){
  function refresh(){
    const t=today();
    const list=A.DB.entries.filter(e=>e.from>t&&A.ITEM.get(e.item)&&A.COL.get(e.col)).sort(A.cmpE);
    if(!list.length){el.innerHTML='<div class="empty"><b>적용 예정인 단가가 없습니다.</b>단가를 입력할 때 적용일을 미래 날짜로 정하면 여기에 모입니다. 예: 다음 달 1일부터 인상.</div>';return}
    const days=s=>Math.round((Date.parse(s)-Date.parse(t))/864e5);
    const groups=[];for(const e of list){let g=groups[groups.length-1];if(!g||g.d!==e.from){g={d:e.from,rows:[]};groups.push(g)}g.rows.push(e)}
    el.innerHTML='<p class="hint">오늘('+esc(t)+') 이후에 적용되는 단가입니다. 적용일이 되면 단가표에 자동으로 반영됩니다.</p>'+groups.map(g=>
      '<div class="section" style="margin-bottom:10px"><div class="row" style="justify-content:space-between"><h2>'+esc(g.d)+'</h2><span class="tag fut">D-'+days(g.d)+' · '+g.rows.length+'건</span></div><div class="tbl-wrap"><table class="list"><thead><tr><th>구분</th><th>모델</th><th>거래처</th><th class="r">지금</th><th class="r">예정</th><th class="r">증감</th><th>사유</th><th></th></tr></thead><tbody>'+
      g.rows.map(e=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col),now=A.entryAt(e.item,e.col,t);return '<tr class="click" data-i="'+esc(it.id)+'" data-c="'+esc(c.key)+'"><td>'+esc(A.itemLabel(it))+'</td><td class="num">'+esc(it.model)+'</td><td>'+esc(c.name)+'</td><td class="r num">'+esc(A.eText(now))+'</td><td class="r num"><b>'+esc(A.eText(e))+'</b></td><td class="r">'+A.deltaHtml(A.delta(now,e))+'</td><td>'+esc(e.reason||'')+'</td><td class="r"><button class="btn small danger" type="button" data-cancel-entry="'+esc(e.id)+'">취소</button></td></tr>'}).join('')+'</tbody></table></div></div>').join('');
  }
  el.addEventListener('click',ev=>{const b=ev.target.closest('[data-cancel-entry]');if(!b)return;ev.stopPropagation();
    if(b.dataset.armed){A.deleteEntry(b.dataset.cancelEntry);A.toast('예정 단가를 취소했습니다.',{action:{label:'되돌리기',fn:()=>A.undo()}})}
    else{b.dataset.armed='1';b.textContent='한 번 더';setTimeout(()=>{if(b.isConnected){delete b.dataset.armed;b.textContent='취소'}},4000)}});
  refresh();return{refresh};
};

/* ---------- 거래처 · 매입처 비교 ---------- */
A.mountCompare=function(el){
  let sel=A.LS.get('danga.cmp');
  el.innerHTML='<div class="section"><label class="field"><span>품목 찾기</span><input type="search" data-q placeholder="모델명 일부 (예: 0580)"></label><div class="pick" data-list></div></div><div data-detail style="display:flex;flex-direction:column;gap:14px"></div><div class="section" data-spread></div>';
  const qEl=el.querySelector('[data-q]'),listEl=el.querySelector('[data-list]'),det=el.querySelector('[data-detail]'),spr=el.querySelector('[data-spread]');
  const krw=(e,d)=>{if(!isNum(e))return null;const r=A.fxAt(d,e.cur||'KRW');return r?e.v*r:null};
  function refresh(){
    const d=today();
    if(A.isEmpty()){listEl.innerHTML='';det.innerHTML='<div class="empty"><b>단가 데이터가 없습니다.</b></div>';spr.innerHTML='';return}
    if(!A.ITEM.get(sel))sel=null;
    const q=qEl.value.trim(),list=A.DB.items.filter(it=>A.matchItem(it,q,'')).slice(0,q?12:8);
    listEl.innerHTML=list.map(it=>'<button type="button" class="pickbtn'+(it.id===sel?' on':'')+'" data-pick="'+esc(it.id)+'"><b class="num">'+esc(A.itemName(it))+'</b><span>'+esc(A.itemLabel(it))+'</span></button>').join('')||'<p class="hint">검색 결과가 없습니다.</p>';
    const it=A.ITEM.get(sel);
    if(!it)det.innerHTML='<div class="empty"><b>비교할 품목을 골라 주세요.</b>위 목록에서 품목을 누르면 거래처별 단가와 마진을 비교합니다.</div>';
    else{
      const buys=A.DB.columns.filter(c=>c.group==='buy').map(c=>{const e=A.entryAt(it.id,c.key,d);return isNum(e)?{c,e,k:krw(e,d)}:null}).filter(Boolean);
      const sells=A.DB.columns.filter(c=>c.group!=='buy').map(c=>{const e=A.entryAt(it.id,c.key,d);return isNum(e)?{c,e,k:krw(e,d),r:A.marginOf(it,c,d,A.basis())}:null}).filter(Boolean).sort((a,b)=>(b.k||0)-(a.k||0));
      const ks=sells.map(x=>x.k).filter(v=>v!=null),mn=ks.length?Math.min(...ks):null,mx=ks.length?Math.max(...ks):null,avg=ks.length?ks.reduce((a,b)=>a+b,0)/ks.length:null;
      const bk=buys.map(x=>x.k).filter(v=>v!=null),bmin=bk.length?Math.min(...bk):null;
      const lm=A.lowMargin();
      let h='<div class="section"><div><div class="hint">'+esc(A.itemLabel(it))+'</div><h2 class="num" style="font-size:20px">'+esc(A.itemName(it))+'</h2></div>';
      h+='<div class="stats"><span class="stat">판매처<b>'+sells.length+'곳</b></span>'+(avg!=null?'<span class="stat">최저<b>'+fmtN(Math.round(mn))+'</b></span><span class="stat">평균<b>'+fmtN(Math.round(avg))+'</b></span><span class="stat">최고<b>'+fmtN(Math.round(mx))+'</b></span>'+(ks.length>1&&mn?'<span class="stat">편차<b>'+((mx/mn-1)*100).toFixed(1)+'%</b></span>':''):'')+'</div>';
      h+=sells.length?'<div class="tbl-wrap"><table class="list"><thead><tr><th>판매처</th><th class="r">단가</th><th class="r">평균 대비</th><th class="r">마진</th><th></th></tr></thead><tbody>'+sells.map(x=>{
        const vs=x.k!=null&&avg&&ks.length>1?(x.k/avg-1)*100:null;
        const mg=x.r&&!x.r.unknown?'<span class="'+(x.r.m<0?'up':x.r.p<lm?'warn-t':'')+'">'+esc(fmtN(Math.round(x.r.m*100)/100))+'원 · '+x.r.p.toFixed(1)+'%</span>':'<span class="hint">-</span>';
        const tags=(x.k!=null&&ks.length>1&&x.k===mn?'<span class="tag base">최저</span> ':'')+(x.k!=null&&ks.length>1&&x.k===mx?'<span class="tag warn">최고</span>':'');
        return '<tr class="click" data-i="'+esc(it.id)+'" data-c="'+esc(x.c.key)+'"><td>'+esc(x.c.name)+'</td><td class="r num">'+A.eHtml(x.e,d)+'</td><td class="r num '+(vs>0?'up':vs<0?'down':'')+'">'+(vs==null?'':(vs>0?'+':'')+vs.toFixed(1)+'%')+'</td><td class="r num">'+mg+'</td><td>'+tags+'</td></tr>'}).join('')+'</tbody></table></div>':'<p class="hint">이 품목의 판매 단가가 없습니다.</p>';
      h+='</div>';
      h+='<div class="section"><h2>매입처 비교</h2>'+(buys.length?'<div class="tbl-wrap"><table class="list"><thead><tr><th>매입처</th><th class="r">매입가</th><th class="r">원화 환산</th><th></th></tr></thead><tbody>'+buys.map(x=>'<tr class="click" data-i="'+esc(it.id)+'" data-c="'+esc(x.c.key)+'"><td>'+esc(x.c.name)+'</td><td class="r num">'+A.eHtml(x.e,d)+'</td><td class="r num">'+(x.k!=null?esc(fmtN(Math.round(x.k*100)/100))+'원':'<span class="hint">환율 필요</span>')+'</td><td>'+(x.k!=null&&bk.length>1&&x.k===bmin?'<span class="tag base">가장 저렴</span>':'')+'</td></tr>').join('')+'</tbody></table></div>'+(bk.length>1?'<p class="hint">매입처 간 차이: '+esc(fmtN(Math.round((Math.max(...bk)-bmin)*100)/100))+'원 ('+((Math.max(...bk)/bmin-1)*100).toFixed(1)+'%)</p>':''):'<p class="hint">이 품목의 매입가가 없습니다.</p>')+'</div>';
      det.innerHTML=h;
    }
    // spread ranking
    const sellCols=A.DB.columns.filter(c=>c.group!=='buy'),rows=[];
    for(const x of A.DB.items){const vs=[];for(const c of sellCols){const e=A.entryAt(x.id,c.key,d);const k=krw(e,d);if(k!=null)vs.push({c,k})}
      if(vs.length>=2){vs.sort((a,b)=>a.k-b.k);const lo=vs[0],hi=vs[vs.length-1];if(hi.k>lo.k)rows.push({it:x,n:vs.length,lo,hi,sp:hi.k/lo.k-1})}}
    rows.sort((a,b)=>b.sp-a.sp);
    spr.innerHTML='<h2>거래처 간 단가 편차가 큰 품목</h2><p class="hint">같은 모델인데 거래처마다 단가가 얼마나 다른지 보여 줍니다. 누르면 위에서 비교합니다.</p>'+(rows.length?'<div class="tbl-wrap"><table class="list"><thead><tr><th>모델</th><th class="r">판매처</th><th>최저</th><th>최고</th><th class="r">편차</th></tr></thead><tbody>'+rows.slice(0,10).map(r=>'<tr class="click" data-sp="'+esc(r.it.id)+'"><td class="num">'+esc(A.itemName(r.it))+'</td><td class="r num">'+r.n+'</td><td>'+esc(r.lo.c.name)+' <span class="num">'+esc(fmtN(Math.round(r.lo.k)))+'</span></td><td>'+esc(r.hi.c.name)+' <span class="num">'+esc(fmtN(Math.round(r.hi.k)))+'</span></td><td class="r num up">'+(r.sp*100).toFixed(1)+'%</td></tr>').join('')+'</tbody></table></div>':'<p class="hint">판매처가 둘 이상인 품목이 없습니다.</p>');
  }
  el.addEventListener('input',ev=>{if(ev.target.matches('[data-q]'))refresh()});
  el.addEventListener('click',ev=>{const b=ev.target.closest('[data-pick],[data-sp]');if(!b)return;ev.stopPropagation();sel=b.dataset.pick||b.dataset.sp;A.LS.set('danga.cmp',sel);refresh();if(b.dataset.sp)det.scrollIntoView({block:'start',behavior:'smooth'})});
  refresh();return{refresh};
};

/* ---------- 원가 변동 시뮬레이션 ---------- */
A.mountSim=function(el){
  el.innerHTML='<div class="section"><h2>원가 변동 시뮬레이션</h2><p class="hint">저장하지 않고 미리 계산합니다. 매입 원가나 환율, 판매가가 바뀌면 마진이 어떻게 되는지 확인하세요. 원가는 설정한 매입처 우선순위를 따릅니다.</p>'+
    '<div class="filters"><label class="field"><span>매입 원가 변동 (%)</span><input type="number" step="0.1" name="cp" value="5"></label><label class="field"><span>외화 환율 변동 (%)</span><input type="number" step="0.1" name="fp" value="0"></label><label class="field"><span>판매가 변동 (%)</span><input type="number" step="0.1" name="sp" value="0"></label><label class="field"><span>대구분</span><select name="cat"></select></label><label class="field"><span>기준일</span><input type="date" name="d"></label><label class="field"><span>원가 기준</span><select name="basis"><option value="hana">하나테크 우선</option><option value="dukwoo">덕우 우선</option></select></label></div><div class="stats" data-stats></div></div><div class="section" data-out></div>';
  const g=n=>el.querySelector('[name="'+n+'"]');
  function refresh(){
    keepVal(g('cat'),catOpts());if(!g('d').value)g('d').value=today();g('basis').value=A.basis();
    const d=g('d').value,cp=num(g('cp')),fp=num(g('fp')),sp=num(g('sp')),cat=g('cat').value,basis=g('basis').value,lm=A.lowMargin();
    const out=g('d').closest('.section').nextElementSibling;
    if(A.isEmpty()){el.querySelector('[data-stats]').innerHTML='';out.innerHTML='<div class="empty"><b>단가 데이터가 없습니다.</b></div>';return}
    const rows=[];
    for(const it of A.DB.items){if(cat&&it.cat1!==cat)continue;const cost=A.costAt(it,d,basis);if(!cost||cost.needFx)continue;
      const fc=(cost.e.cur||'KRW')!=='KRW',costN=cost.krw*(1+cp/100)*(fc?1+fp/100:1);
      for(const c of A.DB.columns){if(c.group==='buy')continue;const s=A.entryAt(it.id,c.key,d);if(!isNum(s))continue;const sr=A.fxAt(d,s.cur||'KRW');if(!sr)continue;
        const sk=s.v*sr,fs=(s.cur||'KRW')!=='KRW',skN=sk*(1+sp/100)*(fs?1+fp/100:1);
        const m0=sk-cost.krw,m1=skN-costN,p0=sk?m0/sk*100:0,p1=skN?m1/skN*100:0;rows.push({it,c,m0,m1,p0,p1})}}
    if(!rows.length){el.querySelector('[data-stats]').innerHTML='';out.innerHTML='<div class="empty"><b>계산할 칸이 없습니다.</b>원가와 판매가가 모두 있는 품목이 필요합니다. 외화 단가는 환율을 먼저 입력하세요.</div>';return}
    const avg=a=>a.reduce((x,y)=>x+y,0)/a.length,neg0=rows.filter(r=>r.m0<0).length,neg1=rows.filter(r=>r.m1<0).length,low0=rows.filter(r=>r.m0>=0&&r.p0<lm).length,low1=rows.filter(r=>r.m1>=0&&r.p1<lm).length;
    const arrow=(a,b,badUp)=>'<b>'+a+' → '+b+'</b>';
    el.querySelector('[data-stats]').innerHTML='<span class="stat">비교 칸<b>'+rows.length+'</b></span><span class="stat">평균 마진율<b>'+avg(rows.map(r=>r.p0)).toFixed(1)+'% → '+avg(rows.map(r=>r.p1)).toFixed(1)+'%</b></span><span class="stat">역마진<b class="'+(neg1>neg0?'up':'')+'">'+neg0+' → '+neg1+'</b></span><span class="stat">저마진<b>'+low0+' → '+low1+'</b></span>';
    const worst=rows.slice().sort((a,b)=>(a.p1-a.p0)-(b.p1-b.p0)).slice(0,20);
    out.innerHTML='<h2>마진이 가장 크게 줄어드는 칸</h2><div class="tbl-wrap"><table class="list"><thead><tr><th>모델</th><th>거래처</th><th class="r">지금 마진</th><th class="r">변경 후</th><th class="r">변화</th><th></th></tr></thead><tbody>'+worst.map(r=>{
      const tag=r.m0>=0&&r.m1<0?'<span class="tag warn">역마진 전환</span>':(r.p0>=lm&&r.p1<lm?'<span class="tag warn">저마진 전환</span>':'');
      return '<tr class="click" data-i="'+esc(r.it.id)+'" data-c="'+esc(r.c.key)+'"><td class="num">'+esc(A.itemName(r.it))+'</td><td>'+esc(r.c.name)+'</td><td class="r num">'+esc(fmtN(Math.round(r.m0)))+'원 · '+r.p0.toFixed(1)+'%</td><td class="r num '+(r.m1<0?'up':'')+'">'+esc(fmtN(Math.round(r.m1)))+'원 · '+r.p1.toFixed(1)+'%</td><td class="r num down">'+(r.p1-r.p0).toFixed(1)+'%p</td><td>'+tag+'</td></tr>'}).join('')+'</tbody></table></div>';
  }
  el.addEventListener('input',ev=>{if(ev.target.name==='basis')A.setBasis(ev.target.value);refresh()});
  refresh();return{refresh};
};

/* ---------- shared preview table for change sets ---------- */
A.previewTable=function(ch){
  return '<div class="tbl-wrap" style="max-height:320px;overflow:auto"><table class="list"><thead><tr><th><input type="checkbox" data-all checked aria-label="전체 선택"></th><th>모델</th><th>거래처</th><th class="r">지금</th><th class="r">새 단가</th><th class="r">증감</th></tr></thead><tbody>'+
    ch.map((c,i)=>'<tr><td><input type="checkbox" data-ck value="'+i+'" checked></td><td class="num">'+esc(A.itemName(c.it))+'<br><span class="hint">'+esc(A.itemLabel(c.it))+'</span></td><td>'+esc(c.col.name)+'</td><td class="r num">'+esc(A.eText(c.cur))+'</td><td class="r num"><b>'+esc(A.eText(c.next))+'</b></td><td class="r">'+A.deltaHtml(A.delta(c.cur,c.next))+(A.isBig(c.cur,c.next)?' <span class="tag warn">큰 변동</span>':'')+'</td></tr>').join('')+'</tbody></table></div>';
};

/* ---------- 비율 일괄 인상 · 인하 ---------- */
A.pctDialog=function(opts){
  opts=opts||{};if(noData())return;
  const buy=A.DB.columns.filter(c=>c.group==='buy'),sell=A.DB.columns.filter(c=>c.group!=='buy');
  const chk=c=>'<label class="check"><input type="checkbox" data-pc="'+esc(c.key)+'"> '+esc(c.name)+'</label>';
  const curs=['KRW',...[...A.usedCurs()]];
  A.openModal('<form class="pctform" autocomplete="off"><h2>비율로 일괄 인상 · 인하</h2><p class="hint">고른 거래처의 단가를 같은 비율로 바꿉니다. 반영 전에 바뀌는 칸을 모두 보여 드립니다.</p>'+
    '<div class="section" style="padding:10px"><b>1. 대상</b><div class="row"><button type="button" class="btn small" data-sel="sell">판매처 전체</button><button type="button" class="btn small" data-sel="buy">매입처 전체</button><button type="button" class="btn small" data-sel="none">해제</button></div>'+
    (sell.length?'<div class="hint">판매처</div><div class="chkgrid">'+sell.map(chk).join('')+'</div>':'')+(buy.length?'<div class="hint">매입처</div><div class="chkgrid">'+buy.map(chk).join('')+'</div>':'')+
    '<div class="form"><label class="field"><span>대구분</span><select name="cat">'+catOpts()+'</select></label><label class="field"><span>통화</span><select name="curf"><option value="">전체</option>'+curs.map(c=>'<option value="'+c+'">'+esc(A.curLabel(c))+'</option>').join('')+'</select></label></div></div>'+
    '<div class="section" style="padding:10px"><b>2. 변경률</b><div class="form"><label class="field"><span>변경률 (%) · 인하는 마이너스</span><input type="number" step="0.1" name="pct" placeholder="예: 3 또는 -5"></label><label class="field"><span>원화 반올림</span><select name="unit"><option value="1">1원 단위</option><option value="10">10원 단위</option><option value="100">100원 단위</option><option value="0.01">소수 둘째 자리</option></select></label></div><p class="hint">외화 단가는 항상 소수 둘째 자리로 계산합니다.</p></div>'+
    '<div data-preview></div><p class="hint up" data-err></p>'+
    '<div class="row-end"><button class="btn" type="button" data-cancel>닫기</button><button class="btn primary" type="submit" data-apply hidden>반영하기</button></div></form>',f=>{
    const err=m=>{f.querySelector('[data-err]').textContent=m;return false};
    if(!list||!list.length)return err('반영할 변경이 없습니다.');
    const picked=[...f.querySelectorAll('[data-ck]')].filter(x=>x.checked).map(x=>Number(x.value));
    if(!picked.length)return err('반영할 칸을 하나 이상 선택하세요.');
    const from=f.from.value;if(!from)return err('적용일을 입력하세요.');
    const pct=Number(f.pct.value),note=f.note.value.trim()||('비율 일괄 변경 '+(pct>0?'+':'')+pct+'%');
    const before=A.alerts(from).neg.length;
    A.batch('비율 일괄 변경',()=>{for(const i of picked){const c=list[i];A.addEntry({item:c.it.id,col:c.col.key,v:c.next.v,t:null,cur:c.next.cur,from,reason:f.reason.value,note})}});
    const after=A.alerts(from).neg.length;
    A.toast(picked.length+'칸을 '+(pct>0?'+':'')+pct+'% 반영했습니다.'+(after>before?' 역마진 칸이 '+(after-before)+'개 늘었습니다.':''),{action:{label:'되돌리기',fn:()=>A.undo()}});
  },'wide');
  const f=document.querySelector('#modal form.pctform');let list=[];
  const box=f.querySelector('[data-preview]');
  function compute(){
    const keys=[...f.querySelectorAll('[data-pc]')].filter(x=>x.checked).map(x=>x.dataset.pc),raw=f.pct.value,pct=Number(raw);
    const from=(f.from&&f.from.value)||opts.date||today();
    list=[];
    if(!keys.length||raw===''||isNaN(pct)||pct===0)return;
    const unit=Number(f.unit.value),cat=f.cat.value,cf=f.curf.value;
    for(const it of A.DB.items){if(cat&&it.cat1!==cat)continue;for(const k of keys){const col=A.COL.get(k),e=A.entryAt(it.id,k,from);if(!isNum(e))continue;const c=e.cur||'KRW';if(cf&&c!==cf)continue;
      const u=c==='KRW'?unit:0.01;let nv=Math.round(e.v*(1+pct/100)/u)*u;nv=Math.round(nv*100)/100;if(nv===e.v||nv<0)continue;list.push({it,col,cur:e,next:{v:nv,t:null,cur:c}})}}
  }
  function render(){
    const keep=f.from?{from:f.from.value,reason:f.reason.value,note:f.note.value}:null;
    compute();
    const pct=Number(f.pct.value),th=Number(A.settings.warnPct);
    if(!list.length){box.innerHTML='<p class="hint">'+(f.querySelector('[data-pc]:checked')&&f.pct.value!==''?'바뀌는 칸이 없습니다.':'거래처를 고르고 변경률을 입력하면 바뀌는 칸이 여기에 나옵니다.')+'</p>';f.querySelector('[data-apply]').hidden=true;return}
    box.innerHTML='<div class="section" style="padding:10px"><b>3. 미리보기</b><div class="stats"><span class="stat">바뀌는 칸<b>'+list.length+'</b></span><span class="stat">변경률<b class="'+(pct>0?'up':'down')+'">'+(pct>0?'+':'')+pct+'%</b></span></div>'+(th>0&&Math.abs(pct)>=th?'<p class="hint up">⚠ 변경률이 큰 변동 경고 기준('+th+'%) 이상입니다. 맞는지 확인하세요.</p>':'')+A.previewTable(list)+
      '<div class="form" style="margin-top:8px"><label class="field"><span>적용일</span><input type="date" name="from" value="'+esc(keep?keep.from:(opts.date||today()))+'"></label><label class="field"><span>사유</span><select name="reason">'+REASONS.map(r=>'<option'+((keep?keep.reason:'연간 단가 협상')===r?' selected':'')+'>'+esc(r)+'</option>').join('')+'</select></label><label class="field full"><span>비고</span><input type="text" name="note" value="'+esc(keep?keep.note:'')+'" placeholder="예: 2027년 단가 협상 결과"></label></div></div>';
    f.querySelector('[data-apply]').hidden=false;
  }
  f.addEventListener('input',ev=>{if(ev.target.matches('[data-all]')){f.querySelectorAll('[data-ck]').forEach(x=>x.checked=ev.target.checked);return}if(ev.target.matches('[data-ck],[name=note],[name=reason]'))return;render()});
  f.addEventListener('click',ev=>{const b=ev.target.closest('[data-sel]');if(!b)return;const s=b.dataset.sel;f.querySelectorAll('[data-pc]').forEach(x=>{const c=A.COL.get(x.dataset.pc);x.checked=s==='none'?false:s==='buy'?c.group==='buy':c.group!=='buy'});render()});
  render();
};

/* ---------- 거래처별 단가표 ---------- */
const sheetName=(s,used)=>{let n=String(s).replace(/[:\\/?*\[\]]/g,' ').trim().slice(0,28)||'단가표',k=n,i=2;while(used.has(k))k=n.slice(0,25)+' '+(i++);used.add(k);return k};
function disp(e,d,krw){if(!isNum(e))return{num:null,txt:(e&&e.t)||'',cur:'KRW'};let cur=e.cur||'KRW',v=e.v;if(krw&&cur!=='KRW'){const r=A.fxAt(d,cur);if(r){v=Math.round(v*r);cur='KRW'}}return{num:v,txt:fmtN(v),cur}}
function buildSheet(key,d,scopeDays){
  const col=A.COL.get(key),since=scopeDays?addDays(d,-scopeDays):null,rows=[];
  for(const it of A.DB.items){const e=A.entryAt(it.id,key,d);if(!A.hasVal(e))continue;if(since&&!(!e.base&&e.from>=since&&!A.sameVal(A.prevOf(e),e)))continue;rows.push({it,e,prev:A.prevOf(e)})}
  return{col,rows};
}
A.customerSheetDialog=function(opts){
  opts=opts||{};if(noData())return;
  const d0=opts.date||today(),sells=A.DB.columns.filter(c=>c.group!=='buy');
  if(!sells.length){A.toast('판매처가 없습니다.');return}
  A.openModal('<form class="csform" autocomplete="off"><h2>거래처별 단가표 만들기</h2><p class="hint">거래처에 보낼 단가표입니다. 매입가와 마진 같은 내부 정보는 들어가지 않습니다.</p>'+
    '<div class="section" style="padding:10px"><b>1. 거래처 (여러 곳 선택 가능)</b><div class="row"><button type="button" class="btn small" data-sel="all">전체</button><button type="button" class="btn small" data-sel="none">해제</button></div><div class="chkgrid" data-cols></div></div>'+
    '<div class="section" style="padding:10px"><b>2. 내용</b><div class="form"><label class="field"><span>기준일</span><input type="date" name="d" value="'+esc(d0)+'"></label><label class="field"><span>포함할 품목</span><select name="scope"><option value="">단가가 있는 전체 품목</option><option value="30">최근 30일 변경분만</option><option value="90">최근 90일 변경분만</option><option value="365">최근 1년 변경분만</option></select></label><label class="field"><span>외화 단가 표시</span><select name="krw"><option value="">원래 통화</option><option value="1"'+(A.view.krw?' selected':'')+'>원화로 환산</option></select></label></div><div class="row"><label class="check"><input type="checkbox" name="showFrom" checked> 적용일 열</label><label class="check"><input type="checkbox" name="showPrev"> 변경 전 단가 열</label></div></div>'+
    '<div class="section" style="padding:10px"><b>3. 머리글</b><div class="form"><label class="field"><span>회사명</span><input type="text" name="company" value="'+esc(A.settings.company||'')+'"></label><label class="field"><span>안내 문구</span><input type="text" name="note" value="'+esc(A.settings.sheetNote||'')+'" placeholder="예: 부가세 별도"></label></div></div>'+
    '<p class="hint up" data-err></p><div class="row-end"><button class="btn" type="button" data-cancel>닫기</button><button class="btn" type="button" data-go="xlsx">엑셀로 저장</button><button class="btn primary" type="button" data-go="pdf">PDF로 저장</button></div></form>',()=>false,'wide');
  const f=document.querySelector('#modal form.csform'),colsEl=f.querySelector('[data-cols]');
  const preSel=new Set(opts.keys||[]);
  function drawCols(){const d=f.d.value||d0,keep=new Set([...colsEl.querySelectorAll('input:checked')].map(x=>x.value));
    colsEl.innerHTML=sells.map(c=>{const n=A.DB.items.filter(it=>A.hasVal(A.entryAt(it.id,c.key,d))).length;return '<label class="check"><input type="checkbox" value="'+esc(c.key)+'"'+((keep.has(c.key)||(!keep.size&&preSel.has(c.key)))?' checked':'')+'> '+esc(c.name)+' <span class="hint">('+n+')</span></label>'}).join('')}
  drawCols();f.d.addEventListener('change',drawCols);
  f.addEventListener('click',ev=>{
    const s=ev.target.closest('[data-sel]');if(s){colsEl.querySelectorAll('input').forEach(x=>x.checked=s.dataset.sel==='all');return}
    const go=ev.target.closest('[data-go]');if(!go)return;
    const err=m=>{f.querySelector('[data-err]').textContent=m};err('');
    const keys=[...colsEl.querySelectorAll('input:checked')].map(x=>x.value);if(!keys.length){err('거래처를 하나 이상 선택하세요.');return}
    const d=f.d.value||d0,scope=Number(f.scope.value)||0,krw=!!f.krw.value,showFrom=f.showFrom.checked,showPrev=f.showPrev.checked,company=f.company.value.trim(),note=f.note.value.trim();
    A.settings.company=company;A.settings.sheetNote=note;A.saveSettings();
    const sheets=keys.map(k=>buildSheet(k,d,scope)).filter(s=>s.rows.length);
    if(!sheets.length){err('조건에 맞는 품목이 없습니다. 기준일이나 포함할 품목을 바꿔 보세요.');return}
    const scopeTxt=scope?' · 최근 '+scope+'일 변경분':'';
    if(go.dataset.go==='xlsx'){
      const used=new Set();
      A.exportBook('거래처별단가표_'+d+'.xlsx',sheets.map(s=>{const anyF=s.rows.some(r=>disp(r.e,d,krw).cur!=='KRW');
        const head=['대구분','소구분','MODEL','단가',anyF?'통화':null,showPrev?'변경 전':null,showFrom?'적용일':null].filter(Boolean);
        const rows=s.rows.map(r=>{const x=disp(r.e,d,krw);return [r.it.cat1,r.it.cat2,r.it.model,x.num!=null?x.num:x.txt].concat(anyF?[A.curLabel(x.cur)]:[],showPrev?[r.prev&&A.hasVal(r.prev)?disp(r.prev,d,krw).num:'']:[],showFrom?[r.e.base?'':r.e.from]:[])});
        return{name:sheetName(s.col.name,used),title:(company?company+' · ':'')+s.col.name+' 단가표',subtitle:d+' 기준'+scopeTxt+(note?' · '+note:''),head,rows,widths:[12,12,20,11,6,11,11].slice(0,head.length),num:[3,5]}})).then(()=>A.toast('엑셀 파일을 내려받았습니다.')).catch(e=>A.toast(e.message));
    }else{
      const tbl=s=>{const anyF=s.rows.some(r=>disp(r.e,d,krw).cur!=='KRW'),cols=2+(anyF?1:0)+(showPrev?1:0)+(showFrom?1:0);
        let h='<table class="pr cs"><thead><tr><th>MODEL</th><th class="r">단가'+(anyF?'':' (원)')+'</th>'+(anyF?'<th>통화</th>':'')+(showPrev?'<th class="r">변경 전</th>':'')+(showFrom?'<th>적용일</th>':'')+'</tr></thead><tbody>';
        let last='';for(const r of s.rows){const g=A.itemLabel(r.it);if(g!==last){h+='<tr class="grp"><td colspan="'+cols+'">'+esc(g)+'</td></tr>';last=g}
          const x=disp(r.e,d,krw);h+='<tr><td class="md">'+esc(r.it.model)+'</td><td class="r">'+esc(x.txt)+'</td>'+(anyF?'<td>'+esc(A.curLabel(x.cur))+'</td>':'')+(showPrev?'<td class="r">'+(r.prev&&A.hasVal(r.prev)?esc(disp(r.prev,d,krw).txt):'')+'</td>':'')+(showFrom?'<td>'+(r.e.base?'':esc(r.e.from))+'</td>':'')+'</tr>'}
        return h+'</tbody></table>'+(note?'<p class="pr-note">'+esc(note)+'</p>':'')};
      if(sheets.length===1)A.printReport((company?company+' · ':'')+sheets[0].col.name+' 단가표',d+' 기준 · 품목 '+sheets[0].rows.length+'개'+scopeTxt,tbl(sheets[0]));
      else A.printReport((company?company+' · ':'')+'거래처별 단가표',d+' 기준 · 거래처 '+sheets.length+'곳'+scopeTxt,sheets.map((s,i)=>'<section class="'+(i?'pr-break':'')+'"><h2 class="pr-h2">'+esc(s.col.name)+' <small>품목 '+s.rows.length+'개</small></h2>'+tbl(s)+'</section>').join(''));
    }
  });
};

/* ---------- 내보내기 ---------- */
A.exportDialog=function(opts){
  if(noData())return;const d0=(opts&&opts.date)||today();
  A.openModal('<form autocomplete="off"><h2>내보내기</h2><label class="field"><span>기준일</span><input type="date" name="d" value="'+esc(d0)+'"></label><label class="check"><input type="checkbox" name="krw"'+(A.view.krw?' checked':'')+'> 외화 단가를 원화로 환산해서 내보내기</label>'+
    '<div class="grid-btns"><button type="button" class="btn" data-x="px">단가표 엑셀</button><button type="button" class="btn" data-x="pp">단가표 PDF</button><button type="button" class="btn" data-x="mx">마진표 엑셀</button><button type="button" class="btn" data-x="mp">마진표 PDF</button></div><p class="hint">변경 이력과 두 날짜 비교는 각 화면에서 내보냅니다. 거래처에 보낼 단가표는 "거래처별 단가표"를 이용하세요.</p><div class="row-end"><button class="btn" type="button" data-cancel>닫기</button></div></form>',()=>false);
  const f=document.querySelector('#modal form');
  f.addEventListener('click',ev=>{const b=ev.target.closest('[data-x]');if(!b)return;const d=f.d.value||d0;if(f.krw.checked!==A.view.krw)A.setKrw(f.krw.checked);
    const x=b.dataset.x;if(x==='px')A.exportXlsx(d).then(()=>A.toast('엑셀 파일을 내려받았습니다.')).catch(e=>A.toast(e.message));else if(x==='pp')A.printPriceTable(d);else if(x==='mx')A.exportMarginXlsx(d,A.basis()).then(()=>A.toast('엑셀 파일을 내려받았습니다.')).catch(e=>A.toast(e.message));else A.printMargin(d,A.basis())});
};

/* ---------- tools hub ---------- */
A.TOOLS=[
  {id:'find',icon:'search',title:'단가 찾아서 바꾸기',desc:'모델을 검색해 단가를 보고 바로 변경합니다.'},
  {id:'add',icon:'plus',title:'품목 추가',desc:'대구분·소구분·모델과 단가를 한 번에 추가합니다.'},
  {id:'bulk',icon:'file',title:'엑셀 일괄 변경',desc:'양식을 받아 고친 뒤 올리면 바뀐 칸만 반영합니다.'},
  {id:'pct',icon:'percent',title:'비율 인상 · 인하',desc:'거래처별로 3% 인상처럼 한 번에 바꿉니다.'},
  {id:'sheet',icon:'users',title:'거래처별 단가표',desc:'거래처에 보낼 단가표를 PDF나 엑셀로 만듭니다.'},
  {id:'export',icon:'export',title:'내보내기',desc:'단가표와 마진표를 엑셀이나 PDF로 저장합니다.'}
];
A.runTool=function(id,ctx){
  ctx=ctx||{};const d=(ctx.date&&ctx.date())||today();
  if(id==='find'){if(ctx.find)ctx.find();return}
  if(id==='add'){A.addItemDialog({date:d,cat1:ctx.cat&&ctx.cat()||undefined});return}
  if(id==='bulk'){if(!noData())A.bulkDialog({date:d});return}
  if(id==='pct'){A.pctDialog({date:d});return}
  if(id==='sheet'){A.customerSheetDialog({date:d});return}
  if(id==='export'){A.exportDialog({date:d});return}
};
A.mountTools=function(el,ctx){
  el.innerHTML='<div class="tiles">'+A.TOOLS.map(t=>'<button type="button" class="tile" data-tool="'+t.id+'"><span class="tile-ico">'+A.icon(t.icon,22)+'</span><span class="tile-t">'+esc(t.title)+'</span><span class="tile-d">'+esc(t.desc)+'</span></button>').join('')+'</div>';
  el.addEventListener('click',ev=>{const b=ev.target.closest('[data-tool]');if(b)A.runTool(b.dataset.tool,ctx)});
  return{refresh(){}};
};

/* ---------- home ---------- */
A.mountHome=function(el,ctx){
  let openKey=null;
  function listRows(rows,kind){return '<div class="tbl-wrap"><table class="list"><tbody>'+rows.slice(0,15).map(x=>'<tr class="click" data-i="'+esc(x.it.id)+'" data-c="'+esc(x.c.key)+'"><td class="num">'+esc(A.itemName(x.it))+'</td><td>'+esc(x.c.name)+'</td><td class="r num '+(kind==='neg'?'up':'warn-t')+'">'+esc(fmtN(Math.round(x.r.m*100)/100))+'원 · '+x.r.p.toFixed(1)+'%</td></tr>').join('')+'</tbody></table></div>'+(rows.length>15?'<p class="hint">'+(rows.length-15)+'칸 더 있습니다. 마진 화면에서 전체를 확인하세요.</p>':'')}
  function refresh(){
    const d=today();
    if(A.isEmpty()){
      const s=A.settings,st=[[!!s.clientId,'구글 클라이언트 ID 입력','드라이브에 저장하려면 한 번만 입력합니다. (README의 설정 방법 참고)'],[false,'단가 엑셀 가져오기','처음이면 원본 단가.xlsx를 가져옵니다.'],[!!A.sync.at,'드라이브에 저장','저장하면 다른 기기에서도 볼 수 있습니다.']];
      el.innerHTML='<div class="onboard"><h2>시작하기</h2><p class="hint">처음 쓰신다면 아래 순서대로 진행하세요. 다른 기기에서 이미 저장한 적이 있다면 ID만 입력하고 "최신본 불러오기"를 누르면 됩니다.</p><ol class="steps">'+st.map((x,i)=>'<li class="'+(x[0]?'done':'')+'"><span class="stepn">'+(x[0]?'✓':i+1)+'</span><div><b>'+esc(x[1])+'</b><div class="hint">'+esc(x[2])+'</div></div></li>').join('')+'</ol><div class="row"><button class="btn primary" type="button" data-home="settings">설정 열기</button><button class="btn" type="button" data-home="pull">최신본 불러오기</button></div></div>';return}
    const al=A.alerts(d),up=A.DB.entries.filter(e=>e.from>d&&A.ITEM.get(e.item)&&A.COL.get(e.col)).length;
    const recent=A.DB.entries.filter(e=>!e.base&&A.ITEM.get(e.item)&&A.COL.get(e.col)).sort((a,b)=>a.at<b.at?1:-1).slice(0,8);
    const card=(key,label,val,cls,sub)=>'<button type="button" class="statcard '+(cls||'')+(openKey===key?' on':'')+'" data-card="'+key+'"><span class="sc-l">'+label+'</span><span class="sc-v num">'+val+'</span>'+(sub?'<span class="sc-s">'+sub+'</span>':'')+'</button>';
    let h='<div class="home-head"><div><h2>오늘 '+esc(d)+'</h2><p class="hint">품목 '+A.DB.items.length+'개 · 거래처 '+A.DB.columns.length+'곳'+(A.sync.at?' · 마지막 저장 '+esc(shortDT(A.sync.at)):'')+'</p></div></div>';
    h+='<div class="statrow">'+
      card('save','저장 상태',A.dirty>0?A.dirty+'건':'완료',A.dirty>0?'warn':'ok',A.dirty>0?'저장 안 된 변경 · 누르면 저장':'모두 저장됨')+
      card('neg','역마진',al.neg.length+'칸',al.neg.length?'bad':'ok',al.neg.length?'누르면 목록':'없음')+
      card('low','저마진 ('+A.lowMargin()+'% 미만)',al.low.length+'칸',al.low.length?'warn':'ok',al.low.length?'누르면 목록':'없음')+
      card('up','적용 예정 단가',up+'건','',up?'누르면 예정 목록':'없음')+
      (al.fxCells?card('fx','환율 필요',al.fxCells+'칸','warn',al.fxCur.join(', ')+' 환율 입력'):'')+'</div>';
    if(openKey==='neg'&&al.neg.length)h+='<div class="section">'+'<h2>역마진 칸</h2>'+listRows(al.neg,'neg')+'</div>';
    if(openKey==='low'&&al.low.length)h+='<div class="section"><h2>저마진 칸</h2>'+listRows(al.low,'low')+'</div>';
    h+='<h3 class="sec-t">자주 쓰는 작업</h3><div class="tiles">'+A.TOOLS.map(t=>'<button type="button" class="tile" data-tool="'+t.id+'"><span class="tile-ico">'+A.icon(t.icon,22)+'</span><span class="tile-t">'+esc(t.title)+'</span><span class="tile-d">'+esc(t.desc)+'</span></button>').join('')+'</div>';
    h+='<h3 class="sec-t">최근 변경</h3>'+(recent.length?'<div class="reclist">'+recent.map(e=>{const it=A.ITEM.get(e.item),c=A.COL.get(e.col),p=A.prevOf(e);return '<button type="button" class="rec" data-i="'+esc(it.id)+'" data-c="'+esc(c.key)+'"><span class="rec-d num">'+esc(e.from)+'</span><span class="rec-m"><b class="num">'+esc(A.itemName(it))+'</b> · '+esc(c.name)+'</span><span class="rec-v num">'+(p?esc(A.eText(p))+' → ':'')+'<b>'+esc(A.eText(e))+'</b> '+A.deltaHtml(A.delta(p,e))+'</span></button>'}).join('')+'</div>':'<div class="empty"><b>아직 변경 기록이 없습니다.</b>단가표에서 칸을 눌러 변경하면 여기에 나옵니다.</div>');
    el.innerHTML=h;
  }
  el.addEventListener('click',ev=>{
    const c=ev.target.closest('[data-card]');
    if(c){const k=c.dataset.card;if(k==='save'){if(A.dirty>0)A.save();else A.toast('저장할 변경이 없습니다.');return}if(k==='up'){ctx.goto('upcoming');return}if(k==='fx'){ctx.goto('fx');return}openKey=openKey===k?null:k;refresh();return}
    const t=ev.target.closest('[data-tool]');if(t){A.runTool(t.dataset.tool,ctx);return}
    const h=ev.target.closest('[data-home]');if(h){if(h.dataset.home==='settings')ctx.goto('settings');else A.pull(true).then(ch=>{if(ch)A.toast('드라이브 최신본을 불러왔습니다.')});return}
  });
  refresh();return{refresh};
};

/* ---------- undo button + keyboard shortcuts ---------- */
A.bindUndo=function(btn){
  const upd=()=>{const n=A.undoStack.length;btn.disabled=!n;btn.title=n?'되돌리기: '+A.undoStack[n-1].label+' (Ctrl+Z)':'되돌릴 변경이 없습니다. 저장 전 변경만 되돌릴 수 있습니다.'};
  btn.addEventListener('click',()=>{const l=A.undo();if(l)A.toast('되돌렸습니다: '+l)});
  A.onChange(upd);upd();
};
A.bindShortcuts=function(h){
  document.addEventListener('keydown',ev=>{
    const typing=/^(INPUT|TEXTAREA|SELECT)$/.test(ev.target.tagName||'')||ev.target.isContentEditable,mod=ev.ctrlKey||ev.metaKey,k=(ev.key||'').toLowerCase();
    if(mod&&k==='s'){ev.preventDefault();A.save();return}
    if(mod&&k==='z'&&!typing){ev.preventDefault();const l=A.undo();if(l)A.toast('되돌렸습니다: '+l);return}
    if(ev.key==='/'&&!typing&&h&&h.search){ev.preventDefault();h.search()}
  });
};
})();

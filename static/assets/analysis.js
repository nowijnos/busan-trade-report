'use strict';
// Browser-only investigation UI; cannot claim server-side K-Stat auto-query.
const $=id=>document.getElementById(id);
const state={rows:[],issues:[],sources:[],articles:[],catalog:[],countries:[],visible:[],findings:''};
const COUNTRY_CODES={"미국":"US","벨기에":"BE","중국":"CN","일본":"JP","독일":"DE","프랑스":"FR","네덜란드":"NL","영국":"GB","베트남":"VN","인도":"IN","대만":"TW","태국":"TH","인도네시아":"ID","호주":"AU","멕시코":"MX","캐나다":"CA","폴란드":"PL","스페인":"ES","브라질":"BR","싱가포르":"SG","말레이시아":"MY","필리핀":"PH","튀르키예":"TR","아랍에미리트":"AE","사우디아라비아":"SA","이탈리아":"IT","러시아":"RU","홍콩":"HK"};
const tradeName=()=> $('trade').value==='export'?'수출':'수입';
const amt=v=> v==null?'—':Math.round(Number(v)).toLocaleString('ko-KR');
const money=v=> v==null?'—':`${(Number(v)/1000000).toLocaleString('ko-KR',{maximumFractionDigits:1})}백만 달러`;
function text(el,t){el.textContent=String(t??'');}
function show(step){document.querySelectorAll('.step').forEach(n=>n.classList.toggle('active',n.id==='step-'+step));document.querySelectorAll('[data-step]').forEach(n=>n.classList.toggle('current',n.dataset.step===step));window.scrollTo({top:0,behavior:'smooth'});}
document.querySelectorAll('[data-step]').forEach(b=>b.addEventListener('click',()=>show(b.dataset.step)));
document.querySelectorAll('[data-next]').forEach(b=>b.addEventListener('click',()=>show(b.dataset.next)));
// Native <select> and <input type=month> are keyboard and mouse accessible.
const countries=$('countries');for(const c of Object.keys(COUNTRY_CODES)){const op=document.createElement('option');op.value=c;countries.append(op);}
function paintCountries(){const box=$('selectedCountries');box.replaceChildren();for(const c of state.countries){const tag=document.createElement('span');tag.className='tag';tag.append(document.createTextNode(c));const b=document.createElement('button');b.type='button';b.textContent='×';b.setAttribute('aria-label',c+' 삭제');b.onclick=()=>{state.countries=state.countries.filter(x=>x!==c);paintCountries();};tag.append(b);box.append(tag);}}
$('addCountry').onclick=()=>{const v=$('countrySearch').value.trim();if(!v)return;if(!COUNTRY_CODES[v]&&!/^[A-Z]{2}$/.test(v)){alert('국가 이름은 자동완성 목록에서 선택하거나 2자리 국가코드(US, BE 등)를 입력해 주세요.');return;}if(!state.countries.includes(v))state.countries.push(v);$('countrySearch').value='';paintCountries();};
$('countrySearch').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();$('addCountry').click();}});
// K-Stat-derived partial catalog; never represent this as a complete HS/MTI dictionary.
async function loadCatalog(){try{const r=await fetch('/analysis/catalog');if(!r.ok)throw Error();state.catalog=await r.json();}catch(_){state.catalog=[];}}
loadCatalog();
function matchesCatalog(){const q=$('keyword').value.toLowerCase().trim();return state.catalog.filter(x=>!q||[x.name,x.code].some(v=>String(v).toLowerCase().includes(q))).slice(0,30);}
$('searchCatalog').onclick=()=>{const q=$('keyword').value.trim(),r=$('productResults');r.replaceChildren();const list=matchesCatalog();if(!q){text($('productHint'),'품목 키워드를 입력하세요.');return;}
 for(const c of list){const b=document.createElement('button');b.className='secondary';b.type='button';b.style.cssText='padding:5px 9px;margin:3px;font-size:12px';b.textContent=`${c.family} ${c.level}단위 · ${c.name} (${c.code})`;b.onclick=()=>{$('keyword').value=c.name;$('family').value=c.family;$('level').value=String(c.level);text($('productHint'),`${c.name} · ${c.family}${c.level} ${c.code} 선택`);r.replaceChildren();};r.append(b);}
 if(!list.length){const p=document.createElement('div');p.className='warning';p.textContent='이 자료 기반 후보 목록에는 없습니다. K-Stat 공식 품목 키워드 검색창에서 HS/MTI 코드를 확인해 주세요.';r.append(p);}else text($('productHint'),`일부 조회자료에서 ${list.length}개 후보 검색됨 (전체 품목 사전 아님)`);
};
function condition(){return {region:$('region').selectedOptions[0].textContent,regionCode:$('region').value,trade:tradeName(),tradeCode:$('trade').value,product:$('keyword').value.trim()||'전체 품목',family:$('family').value,level:$('level').value,interval:$('interval').value,period:$('startMonth').value+' ~ '+$('endMonth').value,countries:state.countries.slice(),comparison:$('comparison').value};}
// Only a user click triggers K-Stat lookup. The extension bridge handles scoped messaging.
$('applyKstat').onclick=()=>{
 const c=condition();if(!/^20\d\d-\d\d$/.test($('endMonth').value)){text($('applyStatus'),'조회 연월을 선택해 주세요.');return;}
 if(c.regionCode==='manual'){text($('applyStatus'),'부산 외 지역의 K-Stat 내부 코드는 확인되지 않았습니다. 공식 사이트의 지역 선택 팝업에서 직접 지정해야 합니다.');window.open('https://stat.kita.net/stat/kts/prod/ProdItemImpExpList.screen','_blank','noopener');return;}
 if(c.family!=='MTI'){text($('applyStatus'),'HS/HSK 품목은 K-Stat에서 직접 선택해야 합니다. 공식 검색 화면을 엽니다.');window.open('https://stat.kita.net/stat/kts/prod/ProdItemImpExpList.screen','_blank','noopener');return;}
 if(state.countries.length>1){text($('applyStatus'),'국가별 조회는 한 번에 한 나라씩 가능합니다. 첫 국가로 조회조건을 설정하므로 나머지 국가는 순서대로 조회 후 파일을 추가해 주세요.');}
 else text($('applyStatus'),'크롬 도우미에 조회조건 적용을 요청했습니다. 탭에서 실제 조회조건과 품목 선택을 확인해 주세요.');
 const payload={requestId:Math.random().toString(36).slice(2),region:c.regionCode||'',country:COUNTRY_CODES[state.countries[0]]||state.countries[0]||'',year:Number($('endMonth').value.slice(0,4)),month:Number($('endMonth').value.slice(5,7)),trade:c.tradeCode,interval:c.interval,level:Number(c.level),family:c.family,keyword:c.product};
 if(!document.documentElement.dataset.analysisBridge){
   text($('applyStatus'),'확장 도우미가 이 페이지에 연결됐는지 확인하지 못했습니다. v42 도우미가 설치되어 있는지 확인하거나 K-Stat에서 직접 조회해 주세요.');
 }
 window.postMessage({type:'BUSAN_ANALYSIS_REQUEST_V42',payload},location.origin);
};
window.addEventListener('message',e=>{if(e.source!==window||e.origin!==location.origin||e.data?.type!=='BUSAN_ANALYSIS_RESULT_V42')return;
 const d=e.data?.payload||{};text($('applyStatus'),d.ok?'K-Stat 조회조건 입력 완료. 실제 지역명·국가명·품목 코드·결과표를 확인한 뒤 조회/다운로드하세요. '+(d.message||''):'조회조건 적용 확인: '+(d.message||'도우미가 설치되지 않았거나 K-Stat 조건 일부를 변경하지 못했습니다.'));
});
$('parseFiles').onclick=async()=>{const files=$('sourceFiles').files;if(!files.length){text($('parseStatus'),'엑셀 또는 CSV를 선택해 주세요.');return;}if(files.length>35){text($('parseStatus'),'최대 35개 파일만 허용됩니다.');return;}
 const btn=$('parseFiles');btn.disabled=true;text($('parseStatus'),`${files.length}개 통계 파일 구조 확인 중…`);try{
 const fd=new FormData();for(const f of files)fd.append('sources',f,f.name);const r=await fetch('/analysis/parse',{method:'POST',body:fd});const ob=await r.json();if(!r.ok)throw Error(ob.error||'검증 실패');state.rows=ob.rows||[];state.sources=ob.sources||[];state.issues=ob.issues||[];
 text($('parseStatus'),`통계 ${ob.count}행 인식 / 파일 ${state.sources.length}개 검사. 검토 경고 ${state.issues.length}건.`);
 const issue=$('sourceIssues');issue.replaceChildren();for(const s of state.sources){const p=document.createElement('div');p.className='muted';p.textContent=`${s.status==='ok'?'✓':'✕'} ${s.file}: ${s.rows}행`;issue.append(p);}if(state.issues.length){const box=document.createElement('details');box.open=true;const title=document.createElement('summary');title.textContent=`오류·주의사항 ${state.issues.length}건`;box.append(title);for(const msg of state.issues){const p=document.createElement('p');p.className='warning';p.textContent=msg;box.append(p);}issue.append(box);}
 if(ob.count){calculate();show('analyze');}
 }catch(e){text($('parseStatus'),'파일 확인 오류: '+e.message);}finally{btn.disabled=false;}};
function normalizedCountry(s){return String(s||'').replace(/\s+/g,'').toLowerCase();}
function filtered(){const q=$('keyword').value.toLowerCase().trim(),period1=$('startMonth').value,period2=$('endMonth').value,countries=state.countries,interval=$('interval').value,region=$('region').selectedOptions[0].textContent;
 return state.rows.filter(r=>r.period>=period1&&r.period<=period2&&r.interval===interval
 && (!q||`${r.item} ${r.code}`.toLowerCase().includes(q))
 && (!countries.length||countries.some(c=>normalizedCountry(c)===normalizedCountry(r.country)||COUNTRY_CODES[c]===r.country))
 && (normalizedCountry(r.region)===normalizedCountry(region)||r.region==='미확인'||r.region==='전국'));
}
function createSVG(tag,attrs={},t=''){const el=document.createElementNS('http://www.w3.org/2000/svg',tag);for(const [a,v] of Object.entries(attrs))el.setAttribute(a,v);if(t)el.textContent=t;return el;}
function chart(rows){const s=$('trendChart');s.replaceChildren();const key=$('trade').value,series=rows.filter(x=>x.interval==='당월'&&x[key]!=null);
 const months=[...new Set(series.map(x=>x.period))].sort();if(months.length<2){s.append(createSVG('text',{x:30,y:135,'font-size':15,fill:'#667c8e'},'서로 다른 월의 당월 실적이 2개월 이상 있어야 추이 그래프를 볼 수 있습니다.'));return;}
 const grouped=new Map();for(const r of series){const id=r.country+'|'+r.code+'|'+r.family;const a=grouped.get(id)||{name:r.country+' · '+r.item,points:new Map()};const prev=a.points.get(r.period);a.points.set(r.period,prev==null?r[key]:prev+r[key]);grouped.set(id,a);}
 const sets=[...grouped.values()].slice(0,6),values=sets.flatMap(x=>[...x.points.values()]);const max=Math.max(1,...values);const X=i=>70+(i*(785/Math.max(1,months.length-1))),Y=v=>218-170*(v/max);
 for(let i=0;i<4;i++){const y=218-i*50;s.append(createSVG('line',{x1:70,y1:y,x2:855,y2:y,stroke:'#d9e5ed'}));s.append(createSVG('text',{x:5,y:y+4,'font-size':11,fill:'#5f7488'},((max*i/3)/1000000).toFixed(1)+'M'));}
 const colors=['#3577b8','#df8535','#2d977d','#8869bc','#c45766','#627588'];sets.forEach((set,ix)=>{const vals=months.map((m,i)=>set.points.has(m)?`${X(i)},${Y(set.points.get(m))}`:null);let segments=[];for(const p of vals){if(p)segments.push(p);else if(segments.length){s.append(createSVG('polyline',{points:segments.join(' '),fill:'none',stroke:colors[ix],'stroke-width':2.5}));segments=[];}}if(segments.length)s.append(createSVG('polyline',{points:segments.join(' '),fill:'none',stroke:colors[ix],'stroke-width':2.5}));
 s.append(createSVG('text',{x:82+(ix%3)*260,y:15+Math.floor(ix/3)*15,'font-size':11,fill:colors[ix]},set.name.slice(0,28)));});months.forEach((m,i)=>s.append(createSVG('text',{x:X(i),y:245,'text-anchor':'middle','font-size':10,fill:'#58748a'},m.slice(2))));
}
function prevMonth(period,delta){const [yy,mm]=period.split('-').map(Number);if(!yy||!mm)return '';const d=new Date(Date.UTC(yy,mm-1-delta,1));return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0');}
function comparable(r,key){const c=$('comparison').value;
 if(c==='전년 동월')return r['prior_'+key]??null;
 if(c==='동일 기간 누계')return r.interval==='누계'?(r['prior_'+key]??null):null;
 if(r.interval!=='당월')return null;
 const periods=c==='전월'?[prevMonth(r.period,1)]:[1,2,3].map(i=>prevMonth(r.period,i));
 const comps=[];
 for(const period of periods){
  const matches=state.rows.filter(z=>z.period===period&&z.interval==='당월'&&z.region===r.region&&z.country===r.country&&z.family===r.family&&z.code===r.code&&z[key]!=null);
  if(matches.length!==1)return null;comps.push(matches[0][key]);
 }
 return comps.reduce((a,b)=>a+b,0)/comps.length;
}
function calculate(){state.visible=filtered();const rows=state.visible;let key=$('trade').value;const allInPeriod=state.rows.length, singleCountry=rows.filter(r=>r.country!=='전체 국가');const dates=[...new Set(rows.map(r=>r.period))].sort();const newest=dates[dates.length-1];const last=rows.filter(r=>r.period===newest&&r[key]!=null);const combined=last.reduce((s,r)=>s+r[key],0);
 const picked=last.filter(r=>r.country!=='전체 국가');const differentSources=[...new Set(rows.map(r=>r.source))];
 $('dataNotice').textContent=`불러온 전체 관측 ${allInPeriod}행 · 조건에 일치한 ${rows.length}행 · 파일 ${differentSources.length}개. ${rows.some(r=>r.region==='전국')?'원본의 지역 표기가 전국인 파일이 있어 부산 실적인지는 조회 화면·합계 검증이 필요합니다.':''}`;
 const stats=$('stats');stats.replaceChildren();for(const [name,value] of [['일치하는 관측값',rows.length+'행'],['가장 최근 관측 월',newest||'자료 없음'],['최근 월 관측 품목 수',String(last.length)+'건']]){const box=document.createElement('div');box.className='stat';const a=document.createElement('span');a.textContent=name;const b=document.createElement('strong');b.textContent=value;box.append(a,b);stats.append(box);}
 const body=$('resultRows');body.replaceChildren();const listed=rows.slice().sort((a,b)=>b.period.localeCompare(a.period)||(b[key]||0)-(a[key]||0)).slice(0,200);
 text($('compareHeader'),$('comparison').value==='최근 3개월'?'직전 3개월 평균':$('comparison').value);
 for(const r of listed){const tr=document.createElement('tr');const comparator=comparable(r,key);for(const s of [r.period,r.country,r.item+' ('+r.family+' '+r.code+')',r.interval,amt(r[key]),amt(comparator),r[key]!=null&&comparator!=null?amt(r[key]-comparator):'—',r.source]){const td=document.createElement('td');td.textContent=s;tr.append(td);}body.append(tr);}
 let out=[];
 if(!rows.length)out.push('현재 조건에 일치하는 자료가 없습니다. 조회 기간·품목·국가·당월/누계를 확인하세요.');
 if(singleCountry.length===0&&rows.length)out.push('현재 자료는 국가별 분해가 없는 품목 합계입니다. 미국·벨기에를 비교하려면 각 국가를 선택한 K-Stat 자료나 국가별 CSV가 필요합니다.');
 const uniqueKeys=new Set(rows.map(r=>r.region+'|'+r.period+'|'+r.country+'|'+r.family+'|'+r.code));if(uniqueKeys.size<rows.length)out.push('동일 조회 범위에서 복수 정렬 파일이 있을 수 있습니다. 단순 합산으로 전체 규모를 추정하지 마세요.');
 if(['전월','최근 3개월','동일 기간 누계'].includes($('comparison').value)&&rows.length&&rows.every(r=>comparable(r,key)==null))out.push('선택한 비교 기준의 실적 자료가 충분하지 않거나 당월/누계가 일치하지 않아 비교 금액을 계산하지 않았습니다.');
 if(last.length)out.push(`선택 파일에 ${newest} ${tradeName()} 실적이 ${last.length}건 기록되어 있습니다. MTI 상위·하위 품목, 국가별 중복, 정렬별 파일이 섞일 수 있으므로 총액으로 단순 합산하지 않습니다.`);
 if(picked.length){const ranked=picked.slice().sort((a,b)=>(b[key]||0)-(a[key]||0));out.push(`최근 월에서 선택한 품목·국가 관측 중 가장 큰 항목: ${ranked[0].country} / ${ranked[0].item} / ${money(ranked[0][key])}. 국가별 비중은 전체 국가를 빠짐없이 확보했을 때만 계산해야 합니다.`);}
 out.push('국가 A의 감소와 국가 B의 증가는 물량 이전의 증거가 아닙니다. 관련 HS/MTI 분류·업체·물류·정책기사를 별도로 확인해야 합니다.');state.findings=out.join('\n');text($('analysisFindings'),state.findings);$('analysisFindings').style.whiteSpace='pre-line';chart(rows);
}
$('compareButton').onclick=calculate;document.querySelectorAll('#trade,#startMonth,#endMonth,#interval,#comparison,#keyword,#region').forEach(n=>n.addEventListener('change',()=>{if(state.rows.length)calculate();}));
function newsTerms(){const product=$('keyword').value.trim()||'수출입',cs=state.countries.length?state.countries:['미국','벨기에'];const base=tradeName();return [ `${product} ${cs[0]} ${base} 증가 감소`,cs[1]?`${product} ${cs[1]} ${base} 동향`: `${product} ${base} 대상국 변화`,`${product} ${base} 물류 통관 관세 정책`,`${product} ${base} 기업 생산공장 신규 수요`];}
$('makeNewsQuery').onclick=()=>{const root=$('newsQueries');root.replaceChildren();newsTerms().forEach(term=>{const r=document.createElement('div');r.className='news-row';const title=document.createElement('span');title.textContent=term;r.append(title);for(const [label,base] of [['네이버','https://search.naver.com/search.naver?where=news&query='],['구글 뉴스','https://news.google.com/search?q=']]){const b=document.createElement('button');b.className='secondary';b.style.padding='6px 10px';b.textContent=label;b.onclick=()=>window.open(base+encodeURIComponent(term),'_blank','noopener');r.append(b);}root.append(r);});};
$('saveArticle').onclick=()=>{const title=$('articleTitle').value.trim(),url=$('articleUrl').value.trim();if(!title||!/^https?:\/\/[^/]+/i.test(url)){alert('실제 확인한 기사 제목과 http(s) 원문 링크를 입력해 주세요.');return;}
 state.articles.push({title,url,date:$('articleDate').value,summary:$('articleSummary').value.trim()});for(const id of ['articleTitle','articleDate','articleUrl','articleSummary'])$(id).value='';paintArticles();};
function paintArticles(){const b=$('articleList');b.replaceChildren();state.articles.forEach((a,i)=>{const r=document.createElement('div');r.className='record';const h=document.createElement('strong');h.textContent=`${i+1}. ${a.title}`;r.append(h);const p=document.createElement('p');p.className='muted';p.textContent=`${a.date||'일자 미확인'} · ${a.url}`;r.append(p);const del=document.createElement('button');del.className='secondary';del.textContent='삭제';del.onclick=()=>{state.articles.splice(i,1);paintArticles();};r.append(del);b.append(r);});}
function reportPayload(){const c=condition();return {title:$('reportTitle').value.trim(),condition:c,analysis:($('analysisText').value.trim()||state.findings),hypothesis:$('hypothesisText').value.trim(),articles:state.articles,rows:state.visible.slice(0,400),sources:state.sources.filter(s=>s.status==='ok').map(s=>`${s.file} · 읽은 행 ${s.rows}개`).concat(state.issues)};}
$('downloadDocx').onclick=async()=>{const btn=$('downloadDocx');btn.disabled=true;text($('reportStatus'),'수치·출처와 메모를 편집 가능한 Word 보고서로 생성 중…');try{const r=await fetch('/analysis/report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(reportPayload())});if(!r.ok){let err=await r.json();throw Error(err.error||'보고서 생성 실패');}const blob=await r.blob();const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='부산_수출입동향_특별분석보고서.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);text($('reportStatus'),'Word 분석보고서 다운로드를 요청했습니다. 뉴스 요지는 직접 등록한 내용만 반영했어요.');}catch(e){text($('reportStatus'),'보고서 오류: '+e.message);}finally{btn.disabled=false;}};
$('downloadJson').onclick=()=>{const blob=new Blob([JSON.stringify(reportPayload(),null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='분석실_조사메모.json';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);};
const htmlEsc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
$('printReport').onclick=()=>{const p=reportPayload(),rows=p.rows.slice(0,40);const table='<table border="1" cellspacing="0" cellpadding="5"><tr><th>연월</th><th>국가</th><th>품목</th><th>수출</th><th>수입</th><th>출처</th></tr>'+rows.map(r=>'<tr>'+[r.period,r.country,r.item,amt(r.export),amt(r.import),r.source].map(t=>`<td>${htmlEsc(t)}</td>`).join('')+'</tr>').join('')+'</table>';const articles=p.articles.map(a=>`<p><b>${htmlEsc(a.title)}</b> (${htmlEsc(a.date)})<br>${htmlEsc(a.url)}<br>${htmlEsc(a.summary)}</p>`).join('')||'<p>검증한 기사 없음</p>';const body=`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${htmlEsc(p.title)}</title><style>body{font:13px/1.7 "Malgun Gothic",sans-serif;color:#1b303f;margin:44px}h1{font-size:24px}h2{margin-top:28px;border-bottom:1px solid #ddd}table{width:100%;font-size:10px;border-collapse:collapse}td,th{border:1px solid #ddd;word-break:break-all}pre{white-space:pre-wrap;font:inherit}@media print{body{margin:14mm}}</style></head><body><h1>${htmlEsc(p.title)}</h1><p>작성일: ${new Date().toLocaleDateString('ko-KR')}</p><h2>조회범위</h2><pre>${htmlEsc(JSON.stringify(p.condition,null,2))}</pre><h2>통계 분석</h2><pre>${htmlEsc(p.analysis)}</pre>${table}<h2>기사·근거</h2>${articles}<h2>추가 확인 및 가설</h2><pre>${htmlEsc(p.hypothesis)}</pre><h2>통계 출처</h2><pre>${htmlEsc(p.sources.join('\n'))}</pre><p>통계와 뉴스 내용은 작성자 최종 확인 필요.</p></body></html>`;const b=new Blob([body],{type:'text/html'});const u=URL.createObjectURL(b);const w=window.open(u,'_blank');if(!w){alert('팝업 차단을 해제한 뒤 다시 눌러 주세요. 새 탭에서 Ctrl+P → PDF로 저장할 수 있습니다.');}setTimeout(()=>URL.revokeObjectURL(u),60000);};

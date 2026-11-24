'use strict';
const $=id=>document.getElementById(id);
const value=id=>$(id).value.trim();
const state={articles:[],queries:[]};
function text(id,s){$(id).textContent=String(s||'');}
function situation(){return {region:value('region'),country:value('country'),product:value('product'),trade:value('trade'),change:value('change'),period:value('period'),topic:value('topic'),extra:value('extra')};}
function makeTerms(){
 const s=situation();
 if(!s.product){text('newsQueries','품목명 또는 키워드를 입력해 주세요.');state.queries=[];return;}
 const basic=[s.country,s.product,s.trade,s.change].filter(Boolean).join(' ');
 const period=s.period?`${s.period.slice(0,4)}년 ${Number(s.period.slice(5))}월`:'';
 const list=[basic,
    [s.region,s.country,s.product,s.trade,s.change,period].filter(Boolean).join(' '),
    [s.country,s.product,s.trade,s.topic||'시장 동향'].filter(Boolean).join(' '),
    [s.product,s.country,s.trade,s.extra||'기업 정책 물류'].filter(Boolean).join(' ')];
 state.queries=[...new Set(list.filter(Boolean))];
 const target=$('newsQueries');target.replaceChildren();
 for(const query of state.queries){
  const box=document.createElement('div');box.className='query';
  const name=document.createElement('strong');name.textContent=query;
  const buttons=document.createElement('div');buttons.className='buttons';
  for(const [label,prefix] of [['네이버 뉴스','https://search.naver.com/search.naver?where=news&query='],['구글 뉴스','https://news.google.com/search?q=']]){
   const a=document.createElement('a');a.href=prefix+encodeURIComponent(query);a.target='_blank';a.rel='noopener noreferrer';a.className='linkbtn secondary';a.textContent=label;buttons.append(a);
  }
  const copy=document.createElement('button');copy.type='button';copy.className='secondary';copy.textContent='검색어 복사';copy.onclick=()=>navigator.clipboard.writeText(query).then(()=>{copy.textContent='복사됨';}).catch(()=>{window.prompt('검색어 복사',query);});buttons.append(copy);
  box.append(name,buttons);target.append(box);
 }
}
$('makeNews').addEventListener('click',makeTerms);
$('clearNews').addEventListener('click',()=>{for(const id of ['country','product','period','extra'])$(id).value='';for(const [id,v] of [['trade','수출'],['change','증가'],['topic',''],['region','부산']])$(id).value=v;state.queries=[];$('newsQueries').replaceChildren();});
for(const id of ['country','product','extra'])$(id).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();makeTerms();}});
function articleList(){const list=$('articleList');list.replaceChildren();for(const [i,a] of state.articles.entries()){
 const el=document.createElement('div');el.className='article';const info=document.createElement('div');
 const title=document.createElement('strong');title.textContent=a.title;
 const link=document.createElement('a');link.href=a.url;link.target='_blank';link.rel='noopener noreferrer';link.textContent=a.url;
 const detail=document.createElement('small');detail.textContent=[a.date||'발행일 미기록',a.summary].filter(Boolean).join(' · ');
 const del=document.createElement('button');del.type='button';del.className='secondary';del.textContent='삭제';del.onclick=()=>{state.articles.splice(i,1);articleList();};
 info.append(title,link,detail);el.append(info,del);list.append(el);
}}
$('saveArticle').addEventListener('click',()=>{
 const title=value('articleTitle'),url=value('articleUrl');
 try{if(!title||!['http:','https:'].includes(new URL(url).protocol))throw Error('missing');}catch(_){alert('기사 제목과 http(s)로 시작하는 기사 원문 주소를 입력해 주세요.');return;}
 state.articles.push({title,url,date:value('articleDate'),summary:value('articleSummary')});
 for(const id of ['articleTitle','articleUrl','articleDate','articleSummary'])$(id).value='';articleList();
});
function payload(){const s=situation();return {mode:'news_only',title:value('reportTitle'),condition:{region:s.region,countries:s.country?[s.country]:[],product:s.product,trade:s.trade,period:s.period,change:s.change,topic:s.topic,extra:s.extra},queries:state.queries,analysis:value('analysisText'),hypothesis:value('hypothesisText'),articles:state.articles,sources:[],rows:[]};}
$('downloadDocx').addEventListener('click',async()=>{
 const btn=$('downloadDocx');btn.disabled=true;text('reportStatus','Word 보고서 생성 중…');
 try{const r=await fetch('/analysis/report',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload())});
 if(!r.ok){let j={};try{j=await r.json();}catch(_){}throw Error(j.error||'보고서 생성 실패');}
 const u=URL.createObjectURL(await r.blob());const a=document.createElement('a');a.href=u;a.download='수출입동향_관련뉴스_조사보고서.docx';a.click();setTimeout(()=>URL.revokeObjectURL(u),12000);text('reportStatus','Word 보고서 다운로드를 요청했어. 기사 내용과 통계 사실을 최종 확인해 줘.');}
 catch(e){text('reportStatus','보고서 생성 오류: '+e.message);}finally{btn.disabled=false;}
});
$('downloadJson').addEventListener('click',()=>{const u=URL.createObjectURL(new Blob([JSON.stringify(payload(),null,2)],{type:'application/json'}));const a=document.createElement('a');a.href=u;a.download='수출입동향_뉴스조사_메모.json';a.click();setTimeout(()=>URL.revokeObjectURL(u),12000);});
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
$('printReport').addEventListener('click',()=>{
 const p=payload(),s=p.condition;const items=p.articles.map(a=>`<div class="article"><b>${esc(a.title)}</b> (${esc(a.date||'발행일 미확인')})<p>${esc(a.summary)}</p><small>${esc(a.url)}</small></div>`).join('')||'<p>기사 기록 없음</p>';
 const content=`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${esc(p.title)}</title><style>body{font:12px/1.7 'Malgun Gothic',sans-serif;color:#243d51;margin:14mm}h1{font-size:23px}h2{font-size:16px;border-bottom:1px solid #ccc;padding-bottom:5px;margin-top:20px}pre{font:inherit;white-space:pre-wrap}.article{page-break-inside:avoid;margin:12px 0}small{overflow-wrap:anywhere}@media print{body{margin:12mm}}</style></head><body><h1>${esc(p.title)}</h1><p>품목 ${esc(s.product)} · 국가 ${esc(s.countries.join(', '))} · ${esc(s.trade)} · ${esc(s.period)}</p><h2>관련 뉴스</h2>${items}<h2>통계 확인 및 조사 메모</h2><pre>${esc(p.analysis||'미작성')}</pre><h2>해석 및 추가 확인사항</h2><pre>${esc(p.hypothesis||'미작성')}</pre><p><small>통계와 기사 내용을 작성자가 확인해야 하며 뉴스 검색 결과 자체는 사실 검증이 아닙니다.</small></p></body></html>`;
 const u=URL.createObjectURL(new Blob([content],{type:'text/html'}));if(!window.open(u,'_blank'))alert('팝업이 차단됐습니다. 팝업을 허용한 뒤 다시 눌러주세요.');setTimeout(()=>URL.revokeObjectURL(u),60000);
});

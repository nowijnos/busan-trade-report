 'use strict';
const $=id=>document.getElementById(id);
const KSTAT_URL='https://stat.kita.net/stat/kts/prod/ProdItemImpExpList.screen';
function clickExcelInPageWorld(){
  const normalize=s=>String(s||'').replace(/\s+/g,'').trim();
  const isShown=e=>!!(e.getClientRects().length && e.getBoundingClientRect().width && e.getBoundingClientRect().height);
  const label=e=>normalize(e.getAttribute('aria-label')||e.getAttribute('title')||e.getAttribute('alt')||e.getAttribute('value')||e.textContent);
  if(location.origin!=='https://stat.kita.net'||location.pathname!=='/stat/kts/prod/ProdItemImpExpList.screen')
    return {ok:false,reason:'올바른 K-Stat 조회 화면이 아닙니다.'};
  const els=[...document.querySelectorAll('a,button,input[type="button"],input[type="submit"],[role="button"]')]
    .filter(e=>isShown(e)&&!e.disabled&&e.getAttribute('aria-disabled')!=='true');
  const excel=els.filter(e=>/^(엑셀다운|엑셀다운로드|ExcelDown|ExcelDownload)$/i.test(label(e)));
  const visibleLink=excel.filter(e=>e.tagName==='A');
  const found=visibleLink.length?visibleLink:excel;
  if(!found.length)return {ok:false,reason:'화면에서 엑셀다운 버튼을 찾지 못했습니다. 표 오른쪽 위 [엑셀다운]을 직접 눌러 보세요.'};
  const countControl=document.getElementById('listCount')||document.querySelector('select[name="listCount"]');
  const r=countControl?.getBoundingClientRect();
  let target;
  if(r){
    const near=found.map(e=>{
      const b=e.getBoundingClientRect();
      const sameRow=Math.abs((b.top+b.bottom)/2-(r.top+r.bottom)/2);
      const isRight=b.left>=r.left-25;
      const score=sameRow+(isRight?0:1000)+Math.max(0,b.left-r.right)*0.005;
      return {e,sameRow,isRight,score};
    }).filter(x=>x.sameRow<=120&&x.isRight).sort((a,b)=>a.score-b.score);
    if(near.length===1||(near.length>1&&near[1].score-near[0].score>10))target=near[0].e;
  }
  if(!target&&found.length===1)target=found[0];
  if(!target)return {ok:false,reason:'엑셀다운 후보가 여러 개여서 자동 클릭을 중단했습니다. 표 오른쪽 위 버튼을 직접 클릭해 주세요.',candidates:found.length};
  try{
    target.click();
    return {ok:true,clicked:true,tag:target.tagName.toLowerCase(),nearListCount:!!r,reason:'실제 엑셀다운 요소 클릭 이벤트를 실행했습니다. 다운로드 발생 여부는 별도 확인합니다.'};
  }catch(err){return {ok:false,reason:'엑셀다운 클릭 오류: '+String(err?.message||err).slice(0,160)};}
}
const KEYS=['mti2_exp_month','mti2_exp_current','mti2_exp_prior','mti2_imp_month','mti2_imp_current','mti2_imp_prior','china_month','china_current','china_prior'];
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
let stateCache=null;
let visibleStepToken='';
function updateDownloadEnabled(){
 const s=stateCache;
 $('download').disabled=!(s&&['verify','ready'].includes(s.phase)&&$('verified').checked);
}
function formatCheckItems(key,period){
 const china=key.startsWith('china'),importTrade=china||key.startsWith('mti2_imp_'),prior=key.endsWith('prior'),monthly=key.endsWith('month');
 const y=prior?period.year-1:period.year,m=Number(period.month);
 const scope=china?'중국(CN)':'전체 국가';
 return [
  `지역이 부산(코드 60), 국가가 ${scope}인지`,
  `품목이 MTI ${china?'3':'2'}단위인지`,
  `조회 기간이 ${y}년 ${m}월 · ${monthly?'당월':'1~'+m+'월 누계'}인지`,
  `금액(US$, %) 기준, ${importTrade?'수입금액':'수출금액'} 내림차순인지`,
  '조회 결과표에 해당 품목과 실제 수치가 표시됐는지 (전국/이전 단계 결과가 아닌지)'
 ];
}
function updateVerifyPanel(s){
 const active=!!s.period&&s.step<KEYS.length&&['verify','ready'].includes(s.phase);
 $('verifyBox').hidden=!active;
 if(!active)return;
 const key=KEYS[s.step];
 const token=`${s.period.year}-${s.period.month}:${s.step}:${key}`;
 if(visibleStepToken!==token){visibleStepToken=token;$('verified').checked=false;}
 $('verifyTitle').textContent=`${s.step+1}/9 · ${key.startsWith('china')?'중국 수입 MTI3':key.startsWith('mti2_imp_')?'부산 수입 MTI2':'부산 수출 MTI2'} 조회 결과 확인`;
 const list=$('verifyItems');list.replaceChildren();
 for(const t of formatCheckItems(key,s.period)){
  const li=document.createElement('li');li.textContent=t;list.appendChild(li);
 }
}
$('verified').addEventListener('change',updateDownloadEnabled);
async function req(msg){let r=await chrome.runtime.sendMessage(msg);if(!r?.ok)throw Error(r?.error||'실행 실패');return r;}
function status(s){$('status').textContent=s;}
function keyDesc(key,s){
 const p=s.period;return `${key.startsWith('china')?'부산 · 중국 · MTI3':key.startsWith('mti2_imp_')?'부산 · 국가전체 · 수입 MTI2':'부산 · 국가전체 · 수출 MTI2'} | ${key.endsWith('prior')?p.year-1:p.year}년 ${p.month}월 | ${key.endsWith('month')?'당월':'누계'}`;
}
async function refresh(){
 const s=await req({type:'STATE'});stateCache=s;
 if(s.period){$('year').value=s.period.year;$('month').value=s.period.month;}
 const remaining=KEYS.length-Object.keys(s.collected).filter(k=>KEYS.includes(k)).length;
 $('steps').replaceChildren();s.steps.forEach((label,i)=>{
   const li=document.createElement('li'),done=!!s.collected[KEYS[i]];
   li.textContent=`${done?'✓':i===s.step?'▶':'○'} ${i+1}. ${label}`+(done?' · 완료':'');
   if(done)li.classList.add('ok');if(i===s.step&&!done)li.classList.add('current');$('steps').append(li);
 });
 updateVerifyPanel(s);
 updateDownloadEnabled();
 $('download').textContent=s.phase==='done'?'9개 다운로드 완료':`엑셀다운 (${Math.min(s.step+1,9)}/9)`;
 $('transfer').disabled=s.phase!=='done'||remaining!==0;
 $('retry').disabled=!s.period||s.phase==='downloading'||s.phase==='done';
 $('start').disabled=['preparing','querying','downloading'].includes(s.phase);
 const labels={idle:'연월을 입력하고 새 작업을 시작하세요.',preparing:'다음 조건을 자동으로 입력하는 중…',querying:'K-Stat 조회 실행 중…',verify:'자동 조회를 실행했어요. 아래 확인 목록을 확인한 뒤 체크하고 [엑셀다운]을 눌러 주세요.',ready:'화면 확인을 완료했어요. [엑셀다운]을 눌러 주세요.',downloading:'엑셀 다운로드 완료를 기다리는 중…',error:'자동 조회 중단. 아래 오류를 확인한 후 [현재 단계 다시 조회]를 누르세요.',done:'9개 엑셀 다운로드가 모두 기록됐어요. [보고서 사이트로 가져오기]를 누르세요.'};
 const detail=s.period&&s.step<KEYS.length?'\n'+keyDesc(KEYS[s.step],s):'';
 status((labels[s.phase]||labels.idle)+detail+(s.error?'\n'+s.error:''));
}
async function kstatTab(){
 const tabs=await chrome.tabs.query({url:'https://stat.kita.net/*'});
 const t=tabs.find(x=>x.url?.includes('/ProdItemImpExpList.screen')&&!x.discarded);
 if(!t)throw Error('K-Stat 품목별 지자체 수출입 화면을 열어 주세요.');
 return t;
}
async function inject(tab){await chrome.scripting.executeScript({target:{tabId:tab.id},files:['kstat_helper.js']});}
$('open').addEventListener('click',async()=>{
 const tabs=await chrome.tabs.query({url:'https://stat.kita.net/*'});const t=tabs.find(x=>x.url?.includes('/ProdItemImpExpList.screen'));
 if(t)await chrome.tabs.update(t.id,{active:true});else await chrome.tabs.create({url:KSTAT_URL,active:true});
});
$('start').addEventListener('click',async()=>{
 try{
   status('첫 단계 조건 설정·조회 중…');
   $('verified').checked=false;visibleStepToken='';
   await req({type:'START',year:$('year').value,month:$('month').value});await refresh();
 }catch(e){await refresh();status('첫 조회 준비 실패: '+e.message+'\n[현재 단계 다시 조회]로 재시도할 수 있어요.');}
});
$('download').addEventListener('click',async()=>{
 try{
   const s=await req({type:'STATE'});if(!['verify','ready'].includes(s.phase))throw Error('아직 현재 단계의 조회가 준비되지 않았습니다.');
   if(!$('verified').checked)throw Error('조회 결과를 직접 확인한 후 체크박스를 체크해 주세요.');
   const key=KEYS[s.step],tab=await kstatTab();await inject(tab);
   const v=await chrome.tabs.sendMessage(tab.id,{type:'KSTAT_CHECK',key,period:s.period});
   if(!v?.ok)throw Error('현재 화면의 K-Stat 조건 불일치: '+(v?.missing||[]).join(', '));
   if(s.phase==='verify')await req({type:'CONFIRM_RESULT',tabId:tab.id,confirmed:true});
   await req({type:'ARM',tabId:tab.id});$('verified').checked=false;await refresh();
   const click=await chrome.scripting.executeScript({target:{tabId:tab.id},world:'MAIN',func:clickExcelInPageWorld});
   const r=click?.[0]?.result;if(!r?.ok){await req({type:'CANCEL_ARM'});throw Error(r?.reason||'엑셀다운 클릭 실패');}
   status(`(${s.step+1}/9) 엑셀다운 클릭했어. 파일 다운로드가 완료되면 다음 단계 조건과 조회를 자동으로 진행해.\n(다운로드 완료 전에는 연속 클릭하지 마세요.)`);
   // The popup may close. Background onChanged always advances the workflow.
 }catch(e){await refresh();status('엑셀다운 오류: '+e.message+'\nK-Stat 화면을 확인하고 다시 눌러 주세요.');}
});
$('retry').addEventListener('click',async()=>{
 try{status('현재 단계 조건·조회 재시도 중…');$('verified').checked=false;updateDownloadEnabled();await req({type:'RETRY'});await refresh();}
 catch(e){await refresh();status('재조회 실패: '+e.message);}
});
$('inspect').addEventListener('click',async()=>{
 try{
 const tab=await kstatTab();await inject(tab);
 const res=await chrome.tabs.sendMessage(tab.id,{type:'KSTAT_DIAG'});
 if(res?.error)throw Error(res.error);
 const url=URL.createObjectURL(new Blob([JSON.stringify(res,null,2)],{type:'application/json'}));
 const a=document.createElement('a');a.href=url;a.download='KSTAT_입력칸_진단_개인정보제외.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
 status('진단 JSON 저장 완료. 로그인값·쿠키·통계 수치는 저장되지 않습니다.');
 }catch(e){status('진단 오류: '+e.message);}
});
$('transfer').addEventListener('click',()=>chrome.tabs.create({url:chrome.runtime.getURL('transfer.html')}));
chrome.storage.onChanged.addListener(ch=>{if(ch.phase||ch.step||ch.collected||ch.error)refresh().catch(()=>{});});
refresh().catch(e=>status(e.message));

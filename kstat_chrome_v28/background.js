 'use strict';
// K-Stat v25. Uses the user's already-authenticated browser tab. Never reads or
// stores cookies, tokens, account data, or trade results. Changes/query only
// on user's explicit job start and upon verified completion of each XLS download.
const KEYS=['mti2_exp_month','mti2_exp_current','mti2_exp_prior','mti2_imp_month','mti2_imp_current','mti2_imp_prior','china_month','china_current','china_prior'];
const STEP_LABELS=['MTI2 수출 · 당월','MTI2 수출 · 금년 누계','MTI2 수출 · 전년 누계','MTI2 수입 · 당월','MTI2 수입 · 금년 누계','MTI2 수입 · 전년 누계','중국 MTI3 수입 · 당월','중국 MTI3 수입 · 금년 누계','중국 MTI3 수입 · 전년 누계'];
const KSTAT_URL='https://stat.kita.net/stat/kts/prod/ProdItemImpExpList.screen';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function getState(){
 const d=await chrome.storage.local.get(['period','step','collected','armed','phase','error','resultSeen','tabId']);
 return {period:d.period||null,step:Number.isInteger(d.step)?d.step:0,collected:d.collected||{},armed:d.armed||null,
 phase:d.phase||'idle',error:d.error||'',resultSeen:!!d.resultSeen,tabId:Number.isInteger(d.tabId)?d.tabId:null};
}
async function setState(obj){await chrome.storage.local.set(obj);}
async function readyTab(){
 const tabs=await chrome.tabs.query({url:'https://stat.kita.net/*'});
 let t=tabs.find(x=>x.url?.includes('/ProdItemImpExpList.screen')&&!x.discarded);
 if(!t)t=await chrome.tabs.create({url:KSTAT_URL,active:true});
 for(let i=0;i<35;i++){
   const now=await chrome.tabs.get(t.id);
   if(now.status==='complete'&&now.url?.includes('/ProdItemImpExpList.screen'))return now;
   await sleep(250);
 }
 throw Error('K-Stat 화면이 준비되지 않았습니다. 페이지를 열고 다시 조회하세요.');
}
async function inject(tabId){await chrome.scripting.executeScript({target:{tabId},files:['kstat_helper.js']});}
async function helper(tabId,msg){return await chrome.tabs.sendMessage(tabId,msg);}
function clickQueryInPageWorld(){
  const compact=x=>String(x||'').replace(/\s+/g,'').trim();
  const isVisible=e=>!!(e.offsetWidth||e.offsetHeight||e.getClientRects().length);
  if(location.origin!=='https://stat.kita.net'||location.pathname!=='/stat/kts/prod/ProdItemImpExpList.screen')
    return {ok:false,reason:'K-Stat 조회 화면이 아닙니다.'};
  const nodes=[...document.querySelectorAll('img[alt],img[title],input[type="image"],input[type="button"],input[type="submit"],button,a,[role="button"]')];
  const label=e=>compact(e.getAttribute('alt')||e.getAttribute('aria-label')||e.getAttribute('title')||e.getAttribute('value')||e.textContent);
  // The diagnosed site identifies 조회 as an IMG. Click the IMG itself: its
  // click bubbles up to its parent A and also honors an IMG-specific handler.
  const imgs=nodes.filter(e=>isVisible(e)&&e.tagName==='IMG'&&/^(조회|검색)$/.test(label(e)));
  const buttons=nodes.filter(e=>isVisible(e)&&e.tagName!=='IMG'&&/^(조회|검색)$/.test(label(e)));
  const items=imgs.length?imgs:buttons;
  if(items.length!==1)return {ok:false,reason:`조회 대상이 ${items.length}개여서 자동 클릭하지 않았습니다.`,candidates:items.length};
  const e=items[0];
  const kind=e.tagName.toLowerCase();
  const hasClickHandler=!!(e.hasAttribute('onclick')||e.onclick||e.closest('[onclick]'));
  try{e.click();return {ok:true,dispatched:true,tag:kind,hasClickHandler};}
  catch(error){return {ok:false,reason:'조회 클릭 실행 오류: '+String(error?.message||error).slice(0,160)};}
}
async function prepareAndQuery(expectedStep=null){
 let s=await getState();
 if(!s.period)throw Error('먼저 보고서 연월을 지정하세요.');
 if(expectedStep!==null&&s.step!==expectedStep)return {skipped:true};
 if(s.step>=KEYS.length){await setState({phase:'done',error:''});return {done:true};}
 const index=s.step,key=KEYS[index];
 await setState({phase:'preparing',error:'',resultSeen:false});
 try{
   let tab=await readyTab();
   await setState({tabId:tab.id});
   await inject(tab.id);
   const prep=await helper(tab.id,{type:'KSTAT_PREPARE',key,period:s.period});
   if(!prep?.ok)throw Error('조건 설정 실패: '+(prep?.missing||[prep?.message||'오류']).join(', '));
   const ok=await helper(tab.id,{type:'KSTAT_CHECK',key,period:s.period});
   if(!ok?.ok)throw Error('조건 검증 실패: '+(ok?.missing||[]).join(', '));
   await setState({phase:'querying'});
   const result=await chrome.scripting.executeScript({target:{tabId:tab.id},world:'MAIN',func:clickQueryInPageWorld});
   if(!result?.[0]?.result?.ok)throw Error('조회 클릭 실패: '+(result?.[0]?.result?.reason||'버튼을 찾지 못함'));
   // Do not try to infer that statistics refreshed from DOM text or row
   // counts: K-Stat can reuse the same rows/counter for different periods.
   // The user must inspect the visible query results for *every* step.
   const same=await getState();
   if(same.step!==index)return {skipped:true};
   await setState({phase:'verify',error:'K-Stat 조회 버튼을 실행했습니다. 화면의 지역·국가·기간·정렬·결과표를 직접 확인한 뒤 체크박스를 선택해 주세요.',resultSeen:false,tabId:tab.id});
   return {ok:true,index,key,manualVerificationRequired:true};
 }catch(e){
   const current=await getState();
   if(current.step===index)await setState({phase:'error',error:String(e.message||e).slice(0,350)});
   throw e;
 }
}
chrome.runtime.onMessage.addListener((msg,sender,reply)=>{
 (async()=>{
   if(msg.type==='START'){
     const y=Number(msg.year),m=Number(msg.month);
     if(!Number.isInteger(y)||y<2000||y>2100||!Number.isInteger(m)||m<1||m>12)throw Error('연월을 확인하세요.');
     await setState({period:{year:y,month:m},step:0,collected:{},armed:null,phase:'preparing',error:'',resultSeen:false,tabId:null,pendingDownloads:{}});
     await prepareAndQuery(0);
     return {ok:true};
   }
   if(msg.type==='STATE')return {ok:true,...await getState(),steps:STEP_LABELS};
   if(msg.type==='RETRY'){
     const s=await getState();
     if(s.phase==='downloading')throw Error('다운로드 완료를 기다리는 중입니다. Chrome 다운로드 목록을 확인하세요.');
     await setState({armed:null});await prepareAndQuery(s.step);return {ok:true};
   }
   if(msg.type==='CONFIRM_RESULT'){
     if(msg.confirmed!==true)throw Error('화면을 직접 확인한 후 체크해 주세요.');
     const s=await getState();
     if(s.phase!=='verify'||s.step>=KEYS.length)throw Error('화면 확인이 필요한 단계가 아닙니다.');
     if(Number(msg.tabId)!==s.tabId)throw Error('K-Stat 화면이 일치하지 않습니다.');
     await inject(s.tabId);
     const v=await helper(s.tabId,{type:'KSTAT_CHECK',key:KEYS[s.step],period:s.period});
     if(!v?.ok)throw Error('현재 조회조건 확인 실패: '+(v?.missing||[]).join(', '));
     await setState({phase:'ready',error:'',resultSeen:true});
     return {ok:true};
   }
   if(msg.type==='ARM'){
     const s=await getState();
     if(s.phase!=='ready'||!s.resultSeen||s.step>=KEYS.length)throw Error('현재 단계 조회가 완료되지 않았습니다.');
     if(!Number.isInteger(msg.tabId)||s.tabId!==msg.tabId)throw Error('조회 중인 K-Stat 탭이 아닙니다.');
     await setState({phase:'downloading',armed:{key:KEYS[s.step],index:s.step,tabId:msg.tabId,started:Date.now()}});
     return {ok:true,key:KEYS[s.step],index:s.step};
   }
   if(msg.type==='CANCEL_ARM'){
     const s=await getState();
     if(s.phase==='downloading')await setState({phase:'verify',resultSeen:false,armed:null});
     return {ok:true};
   }
   if(msg.type==='CLEAR'){
     await setState({period:null,step:0,collected:{},armed:null,phase:'idle',error:'',resultSeen:false,tabId:null,pendingDownloads:{}});
     return {ok:true};
   }
   throw Error('알 수 없는 요청입니다.');
 })().then(reply).catch(e=>reply({ok:false,error:String(e.message||e)}));
 return true;
});
// Capture download metadata only, never read downloaded file content here.
chrome.downloads.onCreated.addListener(async(item)=>{
 try{
   const s=await getState();
   if(!s.armed||s.phase!=='downloading'||Date.now()-s.armed.started>8*60*1000)return;
   if(s.armed.tabId!==null&&Number.isInteger(item.tabId)&&item.tabId>=0&&item.tabId!==s.armed.tabId)return;
   const url=item.finalUrl||item.url||'';
   const fromKstat=/^https:\/\/(?:[^/]+\.)?kita\.net(?:\/|$)/i.test(url);
   const sameTab=item.tabId===s.armed.tabId;
   const blob=/^blob:https:\/\/(?:[^/]+\.)?kita\.net\//i.test(url);
   if(!fromKstat&&!(sameTab&&(blob||url.startsWith('data:'))))return;
   const old=(await chrome.storage.local.get('pendingDownloads')).pendingDownloads||{};
   old[item.id]={key:s.armed.key,index:s.armed.index,created:Date.now()};
   await setState({pendingDownloads:old});
 }catch(_e){}
});
chrome.downloads.onChanged.addListener(async(delta)=>{
 if(!delta.state||!['complete','interrupted'].includes(delta.state.current))return;
 let old,record;
 for(let i=0;i<12;i++){
   old=(await chrome.storage.local.get('pendingDownloads')).pendingDownloads||{};
   record=old[delta.id];if(record)break;
   await sleep(100);
 }
 if(!record)return;
 delete old[delta.id];await setState({pendingDownloads:old});
 const s=await getState();
 if(s.step!==record.index||s.armed?.key!==record.key)return; // stale job
 if(delta.state.current!=='complete'){
   await setState({armed:null,phase:'verify',resultSeen:false,error:'Chrome 다운로드가 중단됐습니다. 수치를 다시 확인하고 체크해 주세요.'});return;
 }
 try{
   const [item]=await chrome.downloads.search({id:delta.id});
   if(!item?.filename||!/\.xlsx?$/i.test(item.filename))throw Error('XLS/XLSX 파일이 아닙니다.');
   const col={...s.collected};
   col[record.key]={basename:item.filename.split(/[\\/]/).pop(),key:record.key,downloadId:delta.id,recordedAt:new Date().toISOString()};
   const next=record.index+1;
   await setState({collected:col,armed:null,step:next,phase:next===KEYS.length?'done':'preparing',error:'',resultSeen:false});
   if(next<KEYS.length){
     // This continues when the popup is closed; the extension service worker
     // owns the transition, not popup timers or user clicking a 'Next' button.
     try{await prepareAndQuery(next);}catch(_e){/* error recorded to storage; Retry button available */}
   }
 }catch(e){await setState({armed:null,phase:'verify',resultSeen:false,error:'파일 기록 실패: '+e.message});}
});

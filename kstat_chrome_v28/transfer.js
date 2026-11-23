'use strict';
// User explicitly selects six files. No folder picker, folder permission or IndexedDB.
const REPORT='https://busan-trade-report.onrender.com/report';
const $=id=>document.getElementById(id);
function status(s){$('status').textContent=s;}
$('open').addEventListener('click',()=>chrome.tabs.create({url:REPORT}));
async function readFileBase64(file){
 if(file.size>8*1024*1024)throw Error(`${file.name}: 파일이 8MB를 초과했습니다.`);
 return new Promise((ok,fail)=>{const fr=new FileReader();fr.onload=()=>ok(String(fr.result).split(',')[1]);fr.onerror=()=>fail(fr.error);fr.readAsDataURL(file);});
}
async function validateFile(file){
 if(!/\.xlsx?$/i.test(file.name))throw Error(`${file.name}: XLS/XLSX 파일만 지원합니다.`);
 if(!file.size)throw Error(`${file.name}: 빈 파일입니다.`);
 const sig=new Uint8Array(await file.slice(0,8).arrayBuffer());
 const xls=sig.length>=8&&[0xD0,0xCF,0x11,0xE0].every((x,i)=>sig[i]===x);
 const xlsx=sig.length>=4&&[0x50,0x4B,0x03,0x04].every((x,i)=>sig[i]===x);
 if(!xls&&!xlsx)throw Error(`${file.name}: XLS/XLSX 파일 형식인지 확인할 수 없습니다.`);
}
async function waitTabReady(tab){
 const limit=Date.now()+90000;
 while(Date.now()<limit){
   try{
     const t=await chrome.tabs.get(tab.id);
     if(t?.status==='complete'&&t.url?.startsWith(REPORT)){
       const probe=await chrome.tabs.sendMessage(tab.id,{type:'BUSAN_BRIDGE_PING'});
       if(probe?.ok&&probe.version>=28)return t;
     }
   }catch(_err){}
   await new Promise(r=>setTimeout(r,450));
 }
 throw Error('보고서 작성 화면과 연결되지 않았습니다. 웹사이트 v28을 GitHub에 적용하고 사이트 로그인 후 다시 시도하세요.');
}
async function reportTab(){
 const tabs=await chrome.tabs.query({url:'https://busan-trade-report.onrender.com/*'});
 let tab=tabs.find(t=>t.url?.startsWith(REPORT)&&!t.discarded);
 if(tab){await chrome.tabs.update(tab.id,{active:true});}
 else{
   tab=tabs.find(t=>!t.discarded);
   if(tab)tab=await chrome.tabs.update(tab.id,{url:REPORT,active:true});
   else tab=await chrome.tabs.create({url:REPORT,active:true});
 }
 return await waitTabReady(tab);
}
async function sendToReport(files){
 if(files.length!==9)throw Error('K-Stat 엑셀 파일은 정확히 9개 선택해 주세요.');
 status('보고서 작성 페이지를 여는 중... (처음 접속이라면 서버가 시작되는 시간이 걸릴 수 있습니다.)');
 const tab=await reportTab();
 let result;
 try{result=await chrome.tabs.sendMessage(tab.id,{type:'BUSAN_IMPORT_KSTAT_FILES',files});}
 catch(_err){throw Error('파일 전달 연결이 실패했습니다. 보고서 사이트에 로그인하고 웹사이트 v28 배포 여부를 확인하세요.');}
 if(!result?.ok)throw Error(result?.error||'파일 전달 실패');
 await chrome.tabs.update(tab.id,{active:true});
 status(`완료: K-Stat 엑셀 ${result.count}개가 보고서 작성 화면에 선택됐습니다.\n본부 엑셀 1개를 추가한 다음 [한글 보고서 생성하기]를 누르세요.`);
}

const pick=$('manualfiles'),submit=$('manualTransfer');
pick.addEventListener('change',()=>{
 const n=pick.files.length;
 $('selectedCount').textContent=`선택된 파일: ${n} / 9개`;
 submit.disabled=n!==9;
 status(n===9?'9개 파일이 선택됐습니다. [사이트로 가져오기]를 눌러 주세요.':(n?'정확히 9개 파일을 선택해 주세요.':'9개 엑셀을 선택해 주세요.'));
});
submit.addEventListener('click',async()=>{
 submit.disabled=true;
 try{
  const list=[...pick.files];
  if(list.length!==9)throw Error('K-Stat 엑셀 9개를 선택하세요.');
  const files=[];
  for(const f of list){await validateFile(f);files.push({name:f.name,base64:await readFileBase64(f)});}
  status('파일 9개를 확인했습니다. 보고서 사이트로 전달 중...');
  await sendToReport(files);
 }catch(e){status('가져오기 중단: '+e.message);}
 finally{submit.disabled=pick.files.length!==9;}
});

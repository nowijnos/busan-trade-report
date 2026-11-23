'use strict';
// Runs on the report website only; never forwards cookies or SSO data.
// The user explicitly selects six files; they are placed in a File input and
// uploaded only after the user presses the site's "한글 보고서 생성하기" button.
chrome.runtime.onMessage.addListener((msg,_sender,reply)=>{
  if(msg?.type==='BUSAN_BRIDGE_PING'){
    reply({ok:location.origin==='https://busan-trade-report.onrender.com'&&location.pathname==='/report',version:28});return;
  }
  if(msg?.type!=='BUSAN_IMPORT_KSTAT_FILES')return;
  try{
    if(location.origin!=='https://busan-trade-report.onrender.com'||location.pathname!=='/report')
      throw Error('보고서 작성 페이지가 아닙니다. /report 페이지로 이동해 주세요.');
    const input=document.querySelector('form#upload input[type="file"][name="kstat_bulk"]');
    if(!input?.multiple)throw Error('새 웹사이트의 K-Stat 일괄 업로드 칸이 없습니다. GitHub 배포 버전을 업데이트해 주세요.');
    if(!Array.isArray(msg.files)||msg.files.length!==9)throw Error('엑셀 파일을 정확히 9개 선택해 주세요.');
    // If the server has an optional approved auto-connector selected, use manual bulk mode.
    const manual=document.querySelector('input[name="kstat_mode"][value="manual"]');
    if(manual&&!manual.checked){manual.checked=true;manual.dispatchEvent(new Event('change',{bubbles:true}));}
    const transfer=new DataTransfer();
    let bytesTotal=0;
    for(const item of msg.files){
      if(typeof item.name!=='string'||!/\.xlsx?$/i.test(item.name)||typeof item.base64!=='string'||item.base64.length>12*1024*1024)
        throw Error('잘못된 엑셀 파일이 있습니다.');
      const raw=atob(item.base64),content=new Uint8Array(raw.length);
      bytesTotal+=content.byteLength;
      if(bytesTotal>45*1024*1024)throw Error('업로드 용량이 너무 큽니다.');
      for(let i=0;i<raw.length;i++)content[i]=raw.charCodeAt(i);
      transfer.items.add(new File([content],item.name,{type:'application/octet-stream'}));
    }
    input.files=transfer.files;
    input.dispatchEvent(new Event('change',{bubbles:true}));
    if(input.files.length!==9)throw Error('파일 9개가 입력칸에 반영되지 않았습니다.');
    const notice=document.getElementById('chrome-import-notice');
    if(notice){notice.hidden=false;notice.textContent='✓ K-Stat 엑셀 9개 파일 선택 완료! 아래에서 본부 엑셀 1개를 선택하고 [한글 보고서 생성하기]를 누르세요.';}
    document.getElementById('kstat-source')?.scrollIntoView({behavior:'auto',block:'center'});
    reply({ok:true,count:input.files.length,version:28});
  }catch(err){reply({ok:false,error:String(err?.message||err)});}
});

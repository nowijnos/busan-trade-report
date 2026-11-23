'use strict';
// Runs only in user's visible K-Stat page. No reading/storing of credentials, cookies or sessions.
// Source page access is limited to the known K-Stat product screen by manifest host_permissions.
if (!globalThis.__busanKstatHelperV28Ready) {
  globalThis.__busanKstatHelperV28Ready = true;
  const allowedKeys = new Set(['mti2_exp_month','mti2_exp_current','mti2_exp_prior','mti2_imp_month','mti2_imp_current','mti2_imp_prior','china_month','china_current','china_prior']);
  const trim = s => String(s || '').replace(/\s+/g,' ').trim().slice(0,80);
  const visible = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const byName = name => document.querySelector(`select[name="${name}"]`);
  const inputByName = name => document.querySelector(`input[name="${name}"]`);
  const selectOptions = el => Array.from(el?.options || []).map(o => ({text:trim(o.textContent),value:o.value}));
  const hasValue = (el,v) => selectOptions(el).some(o => o.value === String(v));
  const onPage = () => location.origin === 'https://stat.kita.net' && location.pathname === '/stat/kts/prod/ProdItemImpExpList.screen';
  function hint(el) {
    if(el.id) {const lab=document.querySelector(`label[for="${CSS.escape(el.id)}"]`);if(lab)return trim(lab.textContent);}
    const lab=el.closest('label');if(lab)return trim(lab.textContent);
    const row=el.closest('tr');if(row){const heads=Array.from(row.querySelectorAll('th')).map(x=>trim(x.textContent)).filter(Boolean);if(heads.length)return heads.join(' / ').slice(0,80);}
    const prev=el.previousElementSibling;if(prev&&trim(prev.textContent).length<45)return trim(prev.textContent);
    return trim(el.parentElement?.parentElement?.querySelector('label, th')?.textContent);
  }
  function actionLabel(e){if(!e)return '';return trim(e.getAttribute('aria-label')||e.getAttribute('alt')||e.getAttribute('title')||e.getAttribute('value')||e.textContent);}
  function diagnostic() {
    if(!onPage())throw Error('K-Stat 품목별 지자체 수출입 화면에서 실행하세요.');
    const controls = Array.from(document.querySelectorAll('select,input:not([type=hidden]),button'))
      .filter(visible).filter(e=>!['password','email','tel'].includes(e.type)).slice(0,120).map((e,index)=>({
        index,tag:e.tagName.toLowerCase(),id:e.id||'',name:e.name||'',hint:hint(e),
        options:e.tagName==='SELECT'?selectOptions(e).slice(0,120):[],
        inputType:e.getAttribute('type')||'',readOnly:!!e.readOnly
      }));
    const actions = Array.from(document.querySelectorAll('button,input[type=button],input[type=submit],a,[role=button],img[alt],img[title]'))
      .filter(visible).map(e=>({tag:e.tagName.toLowerCase(),id:(e.id||'').slice(0,65),name:(e.name||'').slice(0,65),
        label:actionLabel(e),type:(e.getAttribute('type')||'').slice(0,25)}))
      .filter(x=>/조회|검색|다운로드|엑셀|excel|출력|지역|국가|선택|저장/i.test(x.label)).slice(0,40);
    return {source:'K-Stat 품목별 지자체 수출입',page:location.pathname,
      counts:{select:document.querySelectorAll('select').length,inputs:document.querySelectorAll('input').length},
      controls,actions,actionStructure:uiProbe().buttons,mti3OptionVisible:hasValue(byName('s_cond_unit'),'3'),
      note:'개인정보 제외: 로그인·사용자 입력값·선택 지역/국가·세션·쿠키·요청 헤더·통계 결과·URL 매개변수는 수집하지 않습니다.'};
  }
  function validateMsg(msg) {
    if(!onPage())throw Error('K-Stat 품목별 지자체 수출입 화면에서만 가능합니다.');
    if(!allowedKeys.has(msg?.key))throw Error('정의되지 않은 수집 단계입니다.');
    const year=Number(msg?.period?.year),month=Number(msg?.period?.month);
    if(!Number.isInteger(year)||year<2000||year>2100||!Number.isInteger(month)||month<1||month>12)throw Error('보고서 기준연월 오류');
    const prior=msg.key.endsWith('prior'),monthly=msg.key.endsWith('month'),china=msg.key.startsWith('china');
    const target = {
      s_cond_gb:'MTI',s_cond_unit:china?'3':'2',s_year:String(prior?year-1:year),
      s_month:String(month).padStart(2,'0'),s_monthsum_gb:monthly?'1':'2',s_field:'AMT',s_measure:'1',
      s_sort:(china||msg.key.startsWith('mti2_imp_'))?'THIS_IMP_AMT':'THIS_EXP_AMT'
    };
    // Codes are user-confirmed K-Stat region/country codes; not login credentials.
    // MTI2: country=all (blank), China MTI3: country=CN.
    return {target,china,scope:{region:'60',country:china?'CN':''}};
  }
  function showVerifiedCode(scope){
    const region=inputByName('s_prod_code'),country=inputByName('s_ctr_code');
    if(!region||region.value.trim()!=='60'||!country||country.value.trim()!==scope.country)return;
    let banner=document.getElementById('busan-kstat-code-indicator-v27');
    if(!banner){
      banner=document.createElement('div');banner.id='busan-kstat-code-indicator-v27';
      banner.setAttribute('role','status');
    }
    // Always update and move the badge: a former v27 injected element can
    // survive an extension upgrade in the same K-Stat browser tab.
    banner.style.cssText='position:fixed;bottom:12px;right:12px;z-index:2147483000;pointer-events:none;display:block;max-width:300px;padding:6px 9px;background:#eef7ff;box-shadow:0 1px 7px #23394e40;border:1px solid #2d78a8;border-radius:7px;color:#17476c;font:700 11px/1.45 Malgun Gothic,sans-serif;white-space:normal;';
    document.body.appendChild(banner);
    banner.textContent='지역코드 부산(60) · '+(scope.country==='CN'?'중국(CN)':'국가 전체')+' / 실제 표 확인';
  }
  function setInputCode(name, value, done, missing, label){
    const el=inputByName(name);
    if(!el || el.disabled || el.readOnly || el.type==='password'){
      missing.push(label+' 입력칸을 수정할 수 없음');return false;
    }
    if(el.value.trim() !== value){
      // Native setter and input/change simulate ordinary form entry, never bypass authentication.
      const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set;
      if(setter) setter.call(el,value); else el.value=value;
      el.dispatchEvent(new Event('input',{bubbles:true}));
      el.dispatchEvent(new Event('change',{bubbles:true}));
    }
    if(el.value.trim()!==value){
      missing.push(label+' 코드가 적용되지 않음 (팝업에서 직접 선택 필요)');return false;
    }
    done.push(label+': '+(value||'전체 (빈 코드)')+' [입력칸 값 확인]');
    return true;
  }
  function checkScope(scope,missing){
    for(const [field,want,label] of [
      ['s_prod_code',scope.region,'부산 지역코드'],
      ['s_ctr_code',scope.country,'국가코드']
    ]){
      const el=inputByName(field);
      if(!el || el.disabled || el.value.trim()!==want){
        missing.push(label+' '+(want||'전체')+' 확인 안 됨 (선택 팝업에서 직접 지정 가능)');
      }
    }
  }
  function setSelect(name,value,done,missing,desc,optional=false){
    const el=byName(name);
    if(!el||!hasValue(el,value)){if(!optional)missing.push(`${desc}: '${value}' 선택항목 없음`);return false;}
    if(el.value!==String(value)){
      el.value=String(value);
      el.dispatchEvent(new Event('input',{bubbles:true}));
      el.dispatchEvent(new Event('change',{bubbles:true}));
    }
    if(el.value!==String(value)){if(!optional)missing.push(`${desc}: 값이 유지되지 않음`);return false;}
    done.push(`${desc}: ${el.options[el.selectedIndex]?.textContent.trim()||value}`);
    return true;
  }
  async function waitForValue(name,v,ms=2400){const start=Date.now();while(Date.now()-start<ms){if(hasValue(byName(name),v))return true;await new Promise(r=>setTimeout(r,120));}return false;}
  function setSortDescending(done,missing){
    const radio=document.getElementById('s_sort_val_desc') ||
      Array.from(document.querySelectorAll('input[type="radio"][name="s_sort_val"]')).find(x=>/desc/i.test(x.id||x.value||''));
    if(!radio || radio.disabled){missing.push('정렬: 내림차순 선택 버튼 없음');return false;}
    if(!radio.checked){
      const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'checked')?.set;
      if(setter)setter.call(radio,true);else radio.checked=true;
      radio.dispatchEvent(new Event('input',{bubbles:true}));
      radio.dispatchEvent(new Event('change',{bubbles:true}));
    }
    if(!radio.checked){missing.push('정렬 내림차순이 적용되지 않음');return false;}
    done.push('정렬 방향: 내림차순');return true;
  }
  async function prepare(msg){
    const {target,china,scope}=validateMsg(msg),done=[],missing=[];
    setSelect('s_cond_gb',target.s_cond_gb,done,missing,'품목');
    await waitForValue('s_cond_unit',target.s_cond_unit);
    const ds={s_cond_unit:'MTI단위',s_year:'연도',s_month:'월',s_monthsum_gb:'당월/누계',s_field:'화면',s_measure:'금액 단위',s_sort:(china||msg.key.startsWith('mti2_imp_'))?'수입금액 정렬':'수출금액 정렬'};
    for(const [k,name] of Object.entries(ds))setSelect(k,target[k],done,missing,name);
    setSortDescending(done,missing);
    // The 50-row dropdown controls page presentation, not guaranteed Excel export scope.
    // Keep it optional, so a changed K-Stat UI does not break the six reports.
    setSelect('listCount','100',done,missing,'목록 표시: 100개씩 보기',true);
    // Set codes last; some K-Stat handlers reset region/country when category or month changes.
    setInputCode('s_prod_code',scope.region,done,missing,'지역(부산)');
    setInputCode('s_ctr_code',scope.country,done,missing,china?'국가(중국)':'국가(전체)');
    showVerifiedCode(scope);
    return {ok:!missing.length,done,missing,
      message:`${done.join('\n')}\n${missing.length?'\n설정 불가: '+missing.join(', '):''}\n\n주의: 코드 입력만으로 K-Stat의 지역·국가 선택 팝업 상태까지 반영되는지는 아직 확인되지 않았습니다. 코드 60이더라도 원래 지역 설명에 '전국' 글씨가 남을 수 있으니 실제 통계표와 적용코드를 확인해 주세요.\n검증 후 체크박스를 선택해야 조회할 수 있습니다.`,
      mti3OptionVisible:hasValue(byName('s_cond_unit'),'3')};
  }
  function checkCurrent(msg) {
    const {target,china,scope}=validateMsg(msg);
    const missing=[];
    for(const [name,want] of Object.entries(target)){
      const el=byName(name);
      if(!el||el.value!==want)missing.push(`${name} (${name==='s_cond_unit'?'MTI 단위':name==='s_monthsum_gb'?'당월/누계':'조회 조건'})`);
    }
    checkScope(scope,missing);
    const desc=document.getElementById('s_sort_val_desc');
    if(!desc || !desc.checked)missing.push('정렬 기준: 내림차순 미설정');
    return {ok:missing.length===0,missing,china,scopeVerifiedByInput:missing.length===0};
  }
  // Prefer the actual event-bearing ancestor over a decorative IMG. The v20 helper
  // clicked the IMG itself and reported success without verifying a page response.
  function clickableAncestor(e){
    if(!e)return null;
    let n=e;
    for(let i=0;i<5 && n && n!==document.body;i++,n=n.parentElement){
      if(n.matches?.('a,button,input[type="button"],input[type="submit"],[role="button"],[onclick]') ||
         typeof n.onclick==='function')return n;
    }
    return e;
  }
  function candidates(kind){
    const elements=Array.from(document.querySelectorAll('a,button,input[type="button"],input[type="submit"],img[alt],img[title]'))
      .filter(visible);
    if(kind==='query')return [...new Set(elements.filter(e=>/^(조회|검색)$/.test(actionLabel(e))).map(clickableAncestor))];
    // Some K-Stat pages display an outer "다운로드" image and the actual
    // "엑셀다운" link in a popup/menu. Handle them as separate stages.
    if(kind==='excel')return [...new Set(elements.filter(e=>/^엑셀\s*다운$/.test(actionLabel(e))).map(clickableAncestor))];
    if(kind==='downloadMenu')return [...new Set(elements.filter(e=>/^다운로드$/.test(actionLabel(e))).map(clickableAncestor))];
    return [];
  }
  function statusOfResults(){
    // Never send statistics or user-specific values to the extension.
    // Only the number of records, if an explicit result counter is available.
    const texts=[];
    const nodes=document.querySelectorAll('span,div,td,th,p,strong,b');
    for(const n of nodes){
      if(!visible(n)||n.children.length>3)continue;
      const t=trim(n.textContent);
      if(/전체\s*[0-9,]+\s*건/.test(t) || /^[0-9,]+\s*건$/.test(t))texts.push(t);
      if(texts.length>=25)break;
    }
    let count=null;
    for(const t of texts){
      const m=t.match(/(?:전체\s*)?([0-9,]+)\s*건/);
      if(m){const v=Number(m[1].replace(/,/g,''));if(Number.isFinite(v)&&v>=0){count=v;break;}}
    }
    // The live K-Stat screen may render the result grid without the textual
    // "전체 N건" counter. Detect a visible table with several numeric result
    // cells as a secondary signal; no cell contents leave the page.
    let tableRows=0;
    for(const table of document.querySelectorAll('table')){
      if(!visible(table))continue;
      const rows=[...table.querySelectorAll('tr')].filter(visible);
      if(rows.length<3)continue;
      const header=trim(table.querySelector('thead')?.textContent||rows[0]?.textContent);
      if(!/수출|수입|품목|금액/.test(header))continue;
      const data=rows.slice(1);
      const hasNumbers=data.some(row=>[...row.querySelectorAll('td')].some(td=>/\d[\d,.]*/.test(td.textContent||'')));
      if(hasNumbers)tableRows=Math.max(tableRows,data.length);
    }
    return {detected:count!==null||tableRows>0,hasRecords:(count!==null&&count>0)||tableRows>0,count:count??tableRows,tableDetected:tableRows>0};
  }
  function elementMeta(el){
    return {tag:el.tagName.toLowerCase(),label:actionLabel(el)||actionLabel(el.querySelector?.('img')),
      // Structural diagnostics only. Do not include href, onclick source or sensitive data.
      hasNativeHandler:typeof el.onclick==='function'||el.hasAttribute?.('onclick'),
      parentChain:[el.parentElement?.tagName||'',el.parentElement?.parentElement?.tagName||''].filter(Boolean)};
  }
  function queryElementStructure(){
    // Structural details only: no URL, inline script, field values, user data.
    const nodes=Array.from(document.querySelectorAll('img[alt],img[title],button,input[type=button],a'))
      .filter(visible).filter(e=>/^(조회|검색)$/.test(actionLabel(e))).slice(0,5);
    return nodes.map(e=>({targetTag:e.tagName.toLowerCase(),chain:[e,e.parentElement,e.parentElement?.parentElement].filter(Boolean).map(n=>({
      tag:n.tagName.toLowerCase(),hasOnclick:!!(n.hasAttribute?.('onclick')||typeof n.onclick==='function'),
      anchor:n.tagName==='A',hasHref:!!n.hasAttribute?.('href'),
      role:(n.getAttribute?.('role')||'').slice(0,20),
      classNames:trim(n.className?.baseVal||n.className||'').slice(0,80)
    }))}));
  }
  function uiProbe(){
    const kinds=['query','downloadMenu','excel'];
    return {result:statusOfResults(),queryStructure:queryElementStructure(),buttons:Object.fromEntries(kinds.map(k=>[k,candidates(k).slice(0,5).map(elementMeta)])),
      note:'조회 결과 건수와 버튼 구조만 진단합니다. 조회 내용, 주소, 쿠키, 세션, 인증 정보는 포함하지 않습니다.'};
  }
  function queryAction(msg){
    const v=checkCurrent(msg);
    if(!v.ok)return {ok:false,message:'조건 불일치: '+v.missing.join(', ')};
    if(msg?.confirmed!==true)return {ok:false,message:'부산/국가 선택 확인이 필요합니다.'};
    return {ok:false,message:'이전 조회 방식은 사용하지 않습니다. 페이지 안에서 직접 실행하는 v22 조회 버튼을 사용하세요.'};
  }
  function excelAction(msg){
    // Excel download runs in MAIN world via popup.js, never from this isolated content script.
    const v=checkCurrent(msg);
    if(!v.ok)return {ok:false,message:'조건 불일치: '+v.missing.join(', ')};
    return {ok:false,message:'v23의 [④ 엑셀다운 실행]에서 페이지 내부 실행 방식을 사용해 주세요.'};
  }
  chrome.runtime.onMessage.addListener((msg,_sender,respond)=>{
    if(msg?.type==='KSTAT_DIAG'){try{respond(diagnostic());}catch(e){respond({error:e.message});}return;}
    if(msg?.type==='KSTAT_PREPARE'){
      prepare(msg).then(respond).catch(e=>respond({ok:false,message:e.message,missing:[e.message],done:[]}));return true;
    }
    if(msg?.type==='KSTAT_RESULT_STATUS') {try{respond({ok:true,...statusOfResults()});}catch(e){respond({ok:false,message:e.message});}return;}
    if(msg?.type==='KSTAT_BUTTON_PROBE') {try{respond({ok:true,...uiProbe()});}catch(e){respond({ok:false,message:e.message});}return;}
    if(msg?.type==='KSTAT_CHECK') {try{respond(checkCurrent(msg));}catch(e){respond({ok:false,message:e.message});}return;}
    if(msg?.type==='KSTAT_CLICK_QUERY') {try{respond(queryAction(msg));}catch(e){respond({ok:false,message:e.message});}return;}
    if(msg?.type==='KSTAT_CLICK_EXCEL') {try{respond(excelAction(msg));}catch(e){respond({ok:false,message:e.message});}return;}
  });
}

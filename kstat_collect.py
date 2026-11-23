"""K-Stat supplementary data intake and optional *approved* server-to-server bridge.

The official K-Stat web page exposes an interactive export UI, but its public
server-side API contract is not documented. Do not guess private endpoints or
send browser cookies to this service. An organisation-approved connector may
be configured with KSTAT_CONNECTOR_URL; until then bulk XLS intake is usable.

The connector receives {year, month, region, requests} and returns a ZIP of
K-Stat XLS/XLSX exports. Source metadata and amounts are verified independently
of the ZIP member names before being handed to the existing report engine.
"""
from __future__ import annotations
from collections import defaultdict
from io import BytesIO
from pathlib import Path
from urllib import request, parse, error
from zipfile import ZipFile, BadZipFile
import json, os, socket, hashlib
from engine import Reject, read_kstat, make_month_info, main_country_totals
from xlreader import load_excel, number

# The reporter consumes six logical period groups; MTI2 now needs two
# different sorted downloads per period to cover both export and import items.
EXPECTED_KEYS=('mti2_month','mti2_current','mti2_prior',
               'china_month','china_current','china_prior')
REQUIRED_SOURCES=(
 'mti2_exp_month','mti2_exp_current','mti2_exp_prior',
 'mti2_imp_month','mti2_imp_current','mti2_imp_prior',
 'china_month','china_current','china_prior',
)
SOURCE_TO_ENGINE={k:('mti2_'+k.rsplit('_',1)[1] if k.startswith('mti2_') else k) for k in REQUIRED_SOURCES}
SORT_FOR_SOURCE={k:('수출금액' if k.startswith('mti2_exp_') else '수입금액') for k in REQUIRED_SOURCES}
MAX_BRIDGE_DOWNLOAD=12*1024*1024
MAX_MEMBER=8*1024*1024
MAX_MEMBERS=30

def period_for_main(main_data: bytes, main_name: str):
    book=load_excel(main_data,main_name)
    year,month=make_month_info(book)
    return book,year,month

def requirements(year,month):
    return {
        'mti2_exp_month':(2,False,year,month,'당월'),
        'mti2_exp_current':(2,False,year,month,'누계'),
        'mti2_exp_prior':(2,False,year-1,month,'누계'),
        'mti2_imp_month':(2,False,year,month,'당월'),
        'mti2_imp_current':(2,False,year,month,'누계'),
        'mti2_imp_prior':(2,False,year-1,month,'누계'),
        'china_month':(3,True,year,month,'당월'),
        'china_current':(3,True,year,month,'누계'),
        'china_prior':(3,True,year-1,month,'누계'),
    }

def _expected_totals(main,key):
    s=main.sheet('1_지역총괄')
    if key.startswith('mti2'):
        col={'month':'E','current':'K','prior':'G'}[key.rsplit('_',1)[1]]
        return number(s.val(6,col)),number(s.val(8,col))
    return None, main_country_totals(main)[{'china_month':1,'china_prior':2,'china_current':3}[key]]

def _safe_file_payloads(named_files):
    for name,data in named_files:
        name=Path(name).name
        if not name.lower().endswith(('.xls','.xlsx')):
            continue
        if not data or len(data)>MAX_MEMBER:
            raise Reject(f'K-Stat 파일 용량 오류: {name} (8MB 이하)')
        yield name,data

def unpack_zip(data:bytes):
    if len(data)>MAX_BRIDGE_DOWNLOAD:raise Reject('K-Stat ZIP 용량은 12MB 이하여야 합니다.')
    try:
        with ZipFile(BytesIO(data)) as z:
            members=[i for i in z.infolist() if not i.is_dir() and i.filename.lower().endswith(('.xls','.xlsx'))]
            if len(members)>MAX_MEMBERS or sum(i.file_size for i in members)>24*1024*1024:
                raise Reject('K-Stat ZIP 압축 해제 용량 또는 파일 수 제한 초과')
            if z.testzip():raise Reject('K-Stat ZIP이 손상되어 있습니다.')
            files=[]
            for info in z.infolist():
                if info.is_dir():continue
                if not info.filename.lower().endswith(('.xls','.xlsx')):continue
                if info.file_size>MAX_MEMBER:raise Reject('K-Stat ZIP에 과도하게 큰 파일이 있습니다.')
                if len(files)>=MAX_MEMBERS:raise Reject('K-Stat ZIP의 파일 수는 30개 이하입니다.')
                files.append((Path(info.filename).name,z.read(info)))
    except (BadZipFile,ValueError) as e:
        raise Reject('K-Stat ZIP 형식을 확인할 수 없습니다.') from e
    if not files:raise Reject('ZIP 안에 K-Stat 엑셀 자료가 없습니다.')
    return files

def classify(named_files,main_data,main_name,require_complete=False):
    """Detect 6 data types from XLS A2 metadata, not file names.

    Reject period/region/MTI mismatches and misleading China metadata with
    totals exceeding the trusted main workbook. Full report generation uses
    the existing engine's additional strict crosschecks.
    """
    main,year,month=period_for_main(main_data,main_name)
    expected=requirements(year,month)
    # Metadata also contains a sort label, essential to avoid mistaking the
    # export top-N list for the import top-N list of the same period.
    by_meta={(tuple(v), SORT_FOR_SOURCE[key]):key for key,v in expected.items()}
    result=defaultdict(list)
    warnings=[]
    fingerprints=set()
    for name,data in _safe_file_payloads(named_files):
        if hashlib.sha256(data).digest() in fingerprints:
            continue
        fingerprints.add(hashlib.sha256(data).digest())
        try:
            stat=read_kstat(data,name)
        except (Reject,ValueError,KeyError) as exc:
            warnings.append(f'{name}: K-Stat 원본 판독 불가 ({str(exc)[:140]})')
            continue
        marker=(stat.level,stat.is_china,stat.year,stat.month,stat.interval)
        sort_match=__import__('re').search(r'정렬\s*:\s*(수출금액|수입금액)',stat.meta)
        source_sort=sort_match.group(1) if sort_match else None
        key=by_meta.get((marker,source_sort))
        if not key:
            warnings.append(f'{name}: 보고기간·MTI·국가·정렬 조건이 다르므로 제외 (수출·수입 정렬 확인)')
            continue
        exp_export,exp_import=_expected_totals(main,key)
        got_export=stat.total.get('cy_export')
        got_import=stat.total.get('cy_import')
        if key.startswith('mti2'):
            # Require BOTH exact totals even if K-Stat export displays '전국'.
            if any(got is None or expect is None or abs(got-expect)>1
                   for got,expect in ((got_export,exp_export),(got_import,exp_import))):
                warnings.append(f'{name}: 부산 수출·수입 총계가 기본 엑셀과 일치하지 않아 제외')
                continue
        elif stat.region_label=='전국' and (got_import is None or exp_import is None or abs(got_import-exp_import)>1):
            warnings.append(f'{name}: 지역 표시가 전국이고 본부 엑셀의 중국 수입 총계와 일치하지 않아 제외')
            continue
        elif got_import is not None and exp_import is not None:
            if got_import>exp_import+1:
                warnings.append(f'{name}: 중국 수입 총계보다 커서 제외 (이전 오류 유형)')
                continue
            if got_import<exp_import-1 and got_import>exp_import*.25:
                warnings.append(f'{name}: 중국 수입 총계가 일치하지 않아 제외')
                continue
        result[key].append((name,data))
    missing=[x for x in REQUIRED_SOURCES if not result[x]]
    if require_complete:
        for key in REQUIRED_SOURCES:
            rows=result[key]
            if not rows or not key.startswith('china_'):continue
            _e,expected_import=_expected_totals(main,key)
            if not any(read_kstat(data,name).total.get('cy_import') is not None and
                       expected_import is not None and
                       abs(read_kstat(data,name).total['cy_import']-expected_import)<=1
                       for name,data in rows):
                missing.append(key+'(중국 전체 총계 원본)')
        if missing:
            raise Reject('K-Stat 필수 자료 9종 중 누락 또는 합계 불일치: '+', '.join(dict.fromkeys(missing))+
                         '. 수출·수입 정렬과 조회기간을 확인하세요.')
    # The established report engine expects 6 logical groups. Export/import
    # MTI2 files from the same period are merged BY CODE by that engine.
    combined=defaultdict(list)
    for source_key,files in result.items():
        combined[SOURCE_TO_ENGINE[source_key]].extend(files)
    return dict(combined),warnings,missing,(year,month)

def integration_status():
    url=os.environ.get('KSTAT_CONNECTOR_URL','').strip()
    if not url:return {'available':False,'reason':'K-Stat 공식 조회 연계 주소가 아직 설정되지 않았습니다.',
                       'mode':'bulk_upload'}
    p=parse.urlparse(url)
    allowed_http=os.environ.get('KSTAT_DEV_ALLOW_HTTP_LOCAL')=='1' and p.hostname in ('localhost','127.0.0.1')
    if (p.scheme!='https' and not allowed_http) or not p.netloc or p.username or p.password:
        return {'available':False,'reason':'승인된 HTTPS 연결 주소만 사용할 수 있습니다.','mode':'bulk_upload'}
    # Never promise that endpoint works until a live call is verified.
    return {'available':True,'reason':'관리자가 설정한 기관 승인 데이터 연동을 시도할 수 있습니다.',
            'mode':'configured_connector'}

def fetch_approved_connector(year:int,month:int):
    status=integration_status()
    if not status['available']:raise Reject(status['reason']+' K-Stat 일괄 업로드를 이용해 주세요.')
    url=os.environ['KSTAT_CONNECTOR_URL'].strip()
    body=json.dumps({'year':year,'month':month,'region':'부산',
                     'requested_datasets':list(REQUIRED_SOURCES)},ensure_ascii=False).encode('utf-8')
    headers={'Content-Type':'application/json; charset=utf-8','Accept':'application/zip',
             'User-Agent':'BusanTradeReporter/16'}
    token=os.environ.get('KSTAT_CONNECTOR_TOKEN','').strip()
    if token:headers['Authorization']='Bearer '+token
    req=request.Request(url,data=body,headers=headers,method='POST')
    # No redirects to unexpected hosts, and never give credentials to a redirect target.
    class NoRedirect(request.HTTPRedirectHandler):
        def redirect_request(self,*args,**kwargs):return None
    opener=request.build_opener(NoRedirect)
    try:
        with opener.open(req,timeout=45) as resp:
            if resp.status!=200:raise Reject('K-Stat 연계 서버 오류: HTTP '+str(resp.status))
            payload=resp.read(MAX_BRIDGE_DOWNLOAD+1)
            if len(payload)>MAX_BRIDGE_DOWNLOAD:raise Reject('연계 서버 응답 크기 초과')
    except (error.URLError,socket.timeout,TimeoutError,OSError) as exc:
        raise Reject('연계 서버 연결 실패. 기존 K-Stat 자료 업로드를 사용해 주세요.') from exc
    return unpack_zip(payload)

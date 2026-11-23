"""Password-protected cloud/web deployment for Busan trade report generator.

Use HTTPS on the hosting platform. APP_PASSWORD is mandatory for cloud mode.
Uploads and output are processed in memory; no workbook is written to disk.
"""
from __future__ import annotations
import base64
from email.parser import BytesParser
from email import policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse, quote
from pathlib import Path
import json, os, time, threading, secrets, hmac
from engine import build_report, Reject, concat_kstats, require_core_kstat_rows, validate_main_schema, read_kstat
from xlreader import number
from kstat_collect import integration_status, fetch_approved_connector, classify, unpack_zip, period_for_main, EXPECTED_KEYS, REQUIRED_SOURCES, requirements, SORT_FOR_SOURCE, _expected_totals

ROOT=Path(__file__).resolve().parent
TEMPLATE=(ROOT/'templates'/'report_template.hwpx').read_bytes()
MAX_UPLOAD=50*1024*1024
CACHE={}
LOCK=threading.Lock()
CACHE_LIFETIME=15*60
PASSWORD=os.environ.get('APP_PASSWORD','')
USERNAME=os.environ.get('APP_USERNAME','report')
CLOUD=os.environ.get('CLOUD_MODE','0')=='1' or bool(os.environ.get('PORT'))


def clean_cache():
    now=time.monotonic()
    with LOCK:
        for key, entry in list(CACHE.items()):
            if now-entry['created']>CACHE_LIFETIME:
                del CACHE[key]
        while len(CACHE)>15:
            CACHE.pop(next(iter(CACHE)))

class Handler(BaseHTTPRequestHandler):
    server_version='TradeReportWeb/1.1'
    def log_message(self,fmt,*args):
        # Avoid writing filenames, workbook contents, or private statistics into logs.
        print('[WEB] '+fmt%args)
    def sendbytes(self,status,content,ctype='application/octet-stream',headers=None):
        self.send_response(status)
        self.send_header('Content-Type',ctype)
        self.send_header('Content-Length',str(len(content)))
        self.send_header('Cache-Control','no-store, max-age=0')
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('X-Frame-Options','DENY')
        self.send_header('Referrer-Policy','no-referrer')
        for k,v in (headers or {}).items(): self.send_header(k,v)
        self.end_headers()
        self.wfile.write(content)
    def json(self,status,ob,headers=None):
        self.sendbytes(status,json.dumps(ob,ensure_ascii=False).encode('utf-8'),'application/json; charset=utf-8',headers)
    def authorized(self):
        if not PASSWORD: return not CLOUD
        auth=self.headers.get('Authorization','')
        if not auth.startswith('Basic '): return False
        try:
            combined=base64.b64decode(auth[6:],validate=True).decode('utf-8')
            user,passwd=combined.split(':',1)
        except (ValueError,UnicodeError): return False
        return hmac.compare_digest(user,USERNAME) and hmac.compare_digest(passwd,PASSWORD)
    def check_auth(self):
        if self.path=='/health': return True
        if self.authorized(): return True
        self.sendbytes(401,b'Authentication required', 'text/plain; charset=utf-8',
                       {'WWW-Authenticate':'Basic realm="Busan-Trade-Report", charset="UTF-8"'})
        return False
    def do_GET(self):
        if not self.check_auth(): return
        path=urlparse(self.path)
        if path.path=='/kstat/status':
            self.json(200,integration_status())
        elif path.path=='/health':
            self.json(200,{'status':'ok','version':'1.2'})
        elif path.path in ('/','/guide','/report'):
            pages={'/':'index.html','/guide':'guide.html','/report':'report.html'}
            self.sendbytes(200,(ROOT/'static'/pages[path.path]).read_bytes(),'text/html; charset=utf-8')
        elif path.path.startswith('/assets/') or path.path.startswith('/downloads/'):
            # Serve only bounded, release-owned static assets; never expose source or uploaded data.
            base=(ROOT/'static').resolve()
            item=(base/path.path.lstrip('/')).resolve()
            if not item.is_file() or base not in item.parents:
                self.json(404,{'error':'파일을 찾을 수 없습니다.'});return
            ext=item.suffix.lower()
            mime={'.svg':'image/svg+xml','.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp','.zip':'application/zip','.css':'text/css','.js':'text/javascript'}.get(ext)
            if not mime:
                self.json(404,{'error':'지원하지 않는 파일입니다.'});return
            extra={'Content-Disposition':"attachment; filename=busan-trade-report-helper-v32.zip"} if ext=='.zip' else None
            self.sendbytes(200,item.read_bytes(),mime,extra)
        elif path.path=='/download':
            clean_cache()
            q=parse_qs(path.query);token=q.get('token',[''])[0]
            with LOCK:result=CACHE.get(token)
            if result is None:
                self.json(404,{'error':'다운로드 링크가 만료됐습니다. 다시 생성해주세요.'});return
            kind=q.get('type',['report'])[0]
            if kind=='audit':
                name='검증_확인사항.txt'; data=result['audit'].encode('utf-8')
                content_type='text/plain; charset=utf-8'
            else:
                name=result['name'];data=result['data'];content_type='application/octet-stream'
            self.sendbytes(200,data,content_type,{'Content-Disposition':"attachment; filename*=UTF-8''"+quote(name)})
        else:self.json(404,{'error':'페이지를 찾을 수 없습니다'})
    def do_POST(self):
        if not self.check_auth(): return
        endpoint=urlparse(self.path).path
        if endpoint not in ('/generate','/inspect'):
            self.json(404,{'error':'잘못된 요청'});return
        try:size=int(self.headers.get('Content-Length','0') or 0)
        except ValueError:size=0
        if size<=0 or size>MAX_UPLOAD:
            self.json(413,{'error':'업로드 용량 제한: 전체 50MB 이하입니다.'});return
        ctype=self.headers.get('Content-Type','')
        if 'multipart/form-data' not in ctype:
            self.json(400,{'error':'파일 업로드 형식이 아닙니다.'});return
        try:
            raw=self.rfile.read(size)
            msg=BytesParser(policy=policy.default).parsebytes(
                b'MIME-Version: 1.0\r\nContent-Type: '+ctype.encode()+b'\r\n\r\n'+raw)
            uploads={}
            controls={}
            for part in msg.iter_parts():
                key=part.get_param('name',header='content-disposition')
                filename=part.get_filename()
                if key and not filename:
                    controls[key]=str(part.get_content() or '').strip()
                if key and filename:
                    data=part.get_payload(decode=True)
                    if len(data)>20*1024*1024: raise Reject('개별 업로드 파일은 20MB 이하여야 합니다.')
                    uploads.setdefault(key,[]).append((Path(filename).name,data))
            main=uploads.pop('main',None)
            if not main:raise Reject('기본 보고서 엑셀(.xlsx)은 필수입니다.')
            filename,contents=main[0]
            if not filename.lower().endswith('.xlsx'):raise Reject('기본 보고서 엑셀은 .xlsx 형식이어야 합니다.')
            mode=controls.get('kstat_mode','manual')
            if mode not in ('manual','auto'):raise Reject('K-Stat 수집 방식이 올바르지 않습니다.')
            warnings=[]
            if mode=='auto':
                if any(uploads.get(key) for key in EXPECTED_KEYS) or uploads.get('kstat_bulk'):
                    raise Reject('자동 수집 모드에서는 개별 K-Stat 업로드를 동시에 선택하지 마세요.')
                book,year,month=period_for_main(contents,filename)
                raw=fetch_approved_connector(year,month)
                supplements,warnings,_missing,_period=classify(raw,contents,filename,require_complete=True)
                uploads.update(supplements)
            else:
                extra=uploads.pop('kstat_bulk',[])+uploads.pop('kstat_extra',[])
                raw=[]
                for extra_name,extra_bytes in extra:
                    if extra_name.lower().endswith('.zip'):
                        raw.extend(unpack_zip(extra_bytes))
                    else:
                        raw.append((extra_name,extra_bytes))
                if raw:
                    bulk_map,warnings,_missing,_period=classify(raw,contents,filename,require_complete=(endpoint!='/inspect'))
                    for key,files in bulk_map.items():
                        uploads.setdefault(key,[]).extend(files)
            if endpoint=='/inspect':
                # Same validation as real generation, without creating a report.
                # Order-independent: source keys come exclusively from K-Stat A2 metadata.
                book,year,month=period_for_main(contents,filename)
                validate_main_schema(book)
                expected=requirements(year,month)
                labels={
                    'mti2_exp_month':'수출 MTI2 · 당월',
                    'mti2_exp_current':'수출 MTI2 · 금년 누계',
                    'mti2_exp_prior':'수출 MTI2 · 전년 누계',
                    'mti2_imp_month':'수입 MTI2 · 당월',
                    'mti2_imp_current':'수입 MTI2 · 금년 누계',
                    'mti2_imp_prior':'수입 MTI2 · 전년 누계',
                    'china_month':'중국 수입 MTI3 · 당월',
                    'china_current':'중국 수입 MTI3 · 금년 누계',
                    'china_prior':'중국 수입 MTI3 · 전년 누계',
                }
                # Display each recognized group and its original filename.
                checks=[]
                for key,expect in expected.items():
                    logical=key if key.startswith('china_') else 'mti2_'+key.rsplit('_',1)[1]
                    paths=[]
                    for orig,blob in uploads.get(logical,[]):
                        try:
                            k=read_kstat(blob,orig)
                            marker=(k.level,k.is_china,k.year,k.month,k.interval)
                            import re
                            matched=re.search(r'정렬\s*:\s*(수출금액|수입금액)',k.meta)
                            if marker==tuple(expect) and matched and matched.group(1)==SORT_FOR_SOURCE[key]:
                                paths.append({'name':orig,'rows':len(k.rows)})
                        except (Reject,ValueError,KeyError):pass
                    checks.append({'key':key,'label':labels[key], 'files':paths, 'ok':bool(paths)})
                missing=[c['label'] for c in checks if not c['ok']]
                notes=list(warnings)
                if not missing:
                    try:
                        # Preserve the strict core-code and China full-total checks.
                        # Any failure is reported as a concrete explanation, not a vague error.
                        # 'uploads' contains only the validated grouped sources now.
                        from engine import Audit, main_country_totals
                        aug=Audit()
                        regional=book.sheet('1_지역총괄')
                        import_expected={k:{'export':number(regional.val(6,col)), 'import':number(regional.val(8,col))}
                                         for k,col in [('mti2_month','E'),('mti2_current','K'),('mti2_prior','G')]}
                        parsed=concat_kstats(uploads,year,month,aug,main_country_totals(book),import_expected)
                        require_core_kstat_rows(book,parsed)
                        if aug.issues:
                            notes.extend(aug.issues)
                    except Reject as ex:
                        notes.append(str(ex))
                # Check full-China totals, not just metadata, matching generation policy.
                if not missing:
                    try:
                        _combined,_w,_miss,_period=classify(raw,contents,filename,require_complete=True) if mode=='manual' and raw else (uploads,[],[],(year,month))
                    except Reject as ex:notes.append(str(ex))
                ready=not missing and not notes
                self.json(200,{'ok':True,'ready':ready,'year':year,'month':month,
                               'checks':checks,'missing':missing,'issues':notes,
                               'message':('9종 파일이 모두 인식되었고 필수 품목도 확인됐습니다.' if ready else '확인이 필요한 자료가 있습니다. 아래 항목을 확인하세요.')})
                return
            output,outname,audit,stat=build_report(contents,filename,uploads,TEMPLATE)
            if mode=='auto' and stat.empty>0:
                raise Reject(f'자동수집 자료에 필요한 데이터 {stat.empty}칸이 누락되어 생성하지 않았습니다. 연계 서버의 MTI 코드별 보충 자료를 확인하세요.')
            for message in warnings:
                stat.issue('K-Stat 자료 자동분류: '+message)
            if warnings:
                audit+='\n[일괄 업로드 자동분류에서 제외된 파일]\n'+'\n'.join('· '+w for w in warnings)+'\n'
            token=secrets.token_urlsafe(24)
            clean_cache()
            with LOCK:CACHE[token]={'created':time.monotonic(),'name':outname,'data':output,'audit':audit}
            self.json(200,{'ok':True,'token':token,'filename':outname,
                           'issues':stat.issues,'notes':stat.notes,
                           'filled':stat.filled,'empty':stat.empty})
        except Reject as exc:
            self.json(422,{'error':str(exc)})
        except Exception as exc:
            print('[WEB] report generation error:',type(exc).__name__)
            self.json(422,{'error':'보고서 생성 오류가 발생했습니다. 엑셀 형식과 조회 조건을 검토하세요.'})

def run():
    if CLOUD and not PASSWORD:
        raise SystemExit('보안을 위해 APP_PASSWORD 환경변수를 먼저 설정하세요.')
    port=int(os.environ.get('PORT','8765'))
    host='0.0.0.0' if CLOUD else '127.0.0.1'
    print('Busan Trade Report web service on',host,port,flush=True)
    srv=ThreadingHTTPServer((host,port),Handler)
    try:srv.serve_forever()
    except KeyboardInterrupt:pass
    finally:srv.server_close()

if __name__=='__main__':run()

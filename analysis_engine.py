"""Independent, evidence-first K-Stat investigation workspace.

The report automation is deliberately not imported here.  No synthetic K-Stat
numbers or undocumented K-Stat API calls are made. User-supplied exports are
parsed in memory and each observation carries provenance and its coverage.
"""
from __future__ import annotations
import csv, io, re, json, datetime
from pathlib import Path
from collections import defaultdict
from xlreader import load_excel, number

MAX_FILES=35
MAX_ROWS=4000


def _s(v):return str(v if v is not None else '').strip()

def _get(sheet,row,col):return _s(sheet.val(row,col))

def _yyyymm(y,m):
    try:
        y=int(y);m=int(m)
        if 2000<=y<=2100 and 1<=m<=12:return f'{y:04d}-{m:02d}'
    except (ValueError,TypeError):pass
    return ''

def _month(meta):
    m=re.search(r'(?:년월|조회기간|기간)\s*[:：]?\s*(20\d\d)\s*년?\s*[.\-/년\s]*\s*(\d{1,2})\s*월?',meta)
    return _yyyymm(*m.groups()) if m else ''

def _amount(v):
    n=number(v)
    return float(n) if n is not None else None

def _normalize(s):return re.sub(r'\s+','',_s(s))


def parse_kstat_excel(data,name):
    book=load_excel(data,name)
    sh=next(iter(book.sheets.values()))
    meta=_get(sh,2,'A')
    if not ('품목' in meta and ('년월' in meta or '기간' in meta) and ('국가' in meta or '지역' in meta)):
        raise ValueError('K-Stat 조회조건(A2)을 확인할 수 없습니다. 공식 K-Stat 품목 조회 엑셀인지 확인하세요.')
    period=_month(meta)
    if not period: raise ValueError('조회 연월을 확인할 수 없습니다.')
    group=re.search(r'품목\s*[:：]\s*(MTI|HSK?|SITC)\s*(\d{1,2})\s*단위',meta,re.I)
    family=(group.group(1).upper() if group else '')
    digits=int(group.group(2)) if group else 0
    region_match=re.search(r'지역\s*[:：]\s*([^/|\n]+?)(?=\s{2,}|\s*(?:국가|품목|년월|화면|정렬|당월/누계)\s*[:：]|$)',meta)
    country_match=re.search(r'국가\s*[:：]\s*([^/|\n]+?)(?=\s{2,}|\s*(?:지역|품목|년월|화면|정렬|당월/누계)\s*[:：]|$)',meta)
    region=_s(region_match.group(1)) if region_match else '미확인'
    country=_s(country_match.group(1)) if country_match else ''
    if country in ('','전체','전세계','전국','전체국가','--') or '전체' in country:country='전체 국가'
    interval_match=re.search(r'당월/누계\s*[:：]\s*(당월|누계)',meta)
    interval=interval_match.group(1) if interval_match else '미확인'
    # The known K-Stat export format: B code, C product name, D/F previous export/import,
    # I/K current export/import. Verify the structure, not merely file extension.
    header=' '.join(_get(sh,r,'C')+' '+_get(sh,r,'D')+' '+_get(sh,r,'I') for r in range(3,6))
    items=[]
    maxrow=max(sh.rows.keys(),default=0)
    for r in range(6,min(maxrow+1,MAX_ROWS+6)):
        code=_get(sh,r,'B').removesuffix('.0')
        if not re.fullmatch(r'[0-9]{2,10}',code):continue
        if digits and len(code)!=digits:continue
        item=_get(sh,r,'C')
        if not item:continue
        exp,imp=_amount(sh.val(r,'I')),_amount(sh.val(r,'K'))
        prev_exp,prev_imp=_amount(sh.val(r,'D')),_amount(sh.val(r,'F'))
        if exp is None and imp is None:continue
        items.append({'period':period,'interval':interval,'region':region,'country':country,
                      'family':family or 'MTI/HS 미확인','code':code,'item':item,
                      'export':exp,'import':imp,'prior_export':prev_exp,'prior_import':prev_imp,
                      'source':name,'source_kind':'kstat_xls','scope':'국가별 실적' if country!='전체 국가' else '품목 합계'})
    if not items:
        raise ValueError('A2 조회조건은 있지만 유효한 품목코드·금액 행이 없습니다. 빈 결과표 또는 품목 단위 불일치일 수 있습니다.')
    warning=[]
    if region in ('전국','미확인'):
        warning.append('엑셀 조회조건의 지역이 '+region+'으로 표시됩니다. 부산 실적인지는 K-Stat 화면과 합계를 따로 확인하세요.')
    if country=='전체 국가':
        warning.append('전체 국가 품목표입니다. 개별 국가별 비교에는 국가별 조회 파일이 별도로 필요합니다.')
    if interval=='누계':
        warning.append('누계 자료는 월별 당월 추이로 사용하지 않습니다.')
    return items,warning


def parse_csv(data,name):
    decoded=None
    for enc in ('utf-8-sig','cp949','utf-16'):
        try:decoded=data.decode(enc);break
        except UnicodeError:continue
    if decoded is None:raise ValueError('CSV 인코딩을 확인해 주세요 (UTF-8 또는 CP949).')
    rows=csv.DictReader(io.StringIO(decoded))
    required={'연월','지역','국가','품목코드','품목명','분류','수출액','수입액'}
    if not required.issubset(rows.fieldnames or []):
        raise ValueError('국가별 분석 CSV 양식과 열 이름이 다릅니다. [CSV 양식 받기]를 이용하세요.')
    res=[]
    for ix,row in enumerate(rows):
        if ix>=MAX_ROWS:raise ValueError('CSV 최대 4,000행까지만 분석할 수 있습니다.')
        period=re.sub(r'[./]','-',_s(row['연월']))
        if not re.fullmatch(r'20\d\d-(0[1-9]|1[0-2])',period):continue
        code=_s(row['품목코드']);country=_s(row['국가']);product=_s(row['품목명'])
        if not (code and country and product):continue
        exp,imp=_amount(row['수출액']),_amount(row['수입액'])
        if exp is None and imp is None:continue
        res.append({'period':period,'interval':_s(row.get('당월누계','당월')) or '당월','region':_s(row['지역']),
                    'country':country,'family':_s(row['분류']) or 'MTI','code':code,'item':product,
                    'export':exp,'import':imp,'prior_export':_amount(row.get('전년수출액')),
                    'prior_import':_amount(row.get('전년수입액')),
                    'source':name,'source_kind':'user_csv','scope':'사용자 정리 데이터'})
    if not res:raise ValueError('조회 가능한 국가·품목·금액 데이터가 없습니다.')
    return res,['CSV는 사용자 정리 자료입니다. 출처 원본·코드 정의와 숫자를 반드시 대조하세요.']


def parse_files(named):
    if not named or len(named)>MAX_FILES:raise ValueError('통계 파일은 1~35개 선택하세요.')
    rows,issues,sources=[],[],[]
    keys=set()
    for name,data in named:
        name=Path(name).name
        try:
            if not data or len(data)>9*1024*1024:raise ValueError('파일이 비었거나 9MB를 초과합니다.')
            if name.lower().endswith('.csv'):out,warnings=parse_csv(data,name)
            elif name.lower().endswith(('.xls','.xlsx')):out,warnings=parse_kstat_excel(data,name)
            else:raise ValueError('XLS/XLSX/CSV만 지원합니다.')
            added=0
            for r in out:
                key=tuple(r.get(x) for x in ('period','interval','region','country','family','code','source_kind'))
                if key in keys:
                    # Never add duplicate country/period rows and inflate totals.
                    continue
                keys.add(key);rows.append(r);added+=1
            issues.extend(f'{name}: {w}' for w in warnings)
            sources.append({'file':name,'rows':added,'status':'ok'})
        except (ValueError,KeyError,TypeError,IndexError) as e:
            issues.append(f'{name}: {str(e)[:220]}')
            sources.append({'file':name,'rows':0,'status':'error'})
    if len(rows)>16000:raise ValueError('분석자료는 최대 16,000행까지 허용합니다.')
    return {'ok':bool(rows),'rows':rows,'issues':issues,'sources':sources,'count':len(rows)}


def make_docx(payload):
    """Generate editable source-traceable report, never claim automatic fact verification."""
    from docx import Document
    from docx.shared import Pt,Cm,RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(2);sec.bottom_margin=Cm(2);sec.left_margin=Cm(2.25);sec.right_margin=Cm(2.25)
    normal=doc.styles['Normal'];normal.font.name='Malgun Gothic';normal.font.size=Pt(9)
    normal._element.rPr.rFonts.set(qn('w:eastAsia'),'맑은 고딕')
    normal.paragraph_format.space_after=Pt(6)
    title=str(payload.get('title') or '수출입동향 특별 분석보고서').strip()[:120]
    doc.add_heading(title,0)
    doc.add_paragraph('작성일: '+datetime.datetime.now().strftime('%Y-%m-%d')+'  |  데이터: 작성자가 직접 확인한 K-Stat·뉴스·분석 메모')
    doc.add_paragraph('주의: 통계 사실, 기사 내용, 추정은 구분하여 기재했습니다. 자료의 기간·품목 분류·국가·출처를 최종 확인하세요.')
    condition=payload.get('condition') if isinstance(payload.get('condition'),dict) else {}
    doc.add_heading('1. 분석 범위',level=1)
    for label,key in [('지역','region'),('수출/수입','trade'),('품목','product'),('국가','countries'),('조회기간','period'),('비교기준','comparison')]:
        v=condition.get(key,'')
        if isinstance(v,list):v=', '.join(str(z)[:50] for z in v)
        doc.add_paragraph(f'{label}: {str(v)[:250]}')
    doc.add_heading('2. 통계 분석',level=1)
    notes=str(payload.get('analysis') or '').strip()[:12000]
    doc.add_paragraph(notes or '통계 분석 메모 미작성. 조회한 원자료를 검증한 후 작성 바랍니다.')
    rows=payload.get('rows',[])
    if isinstance(rows,list) and rows:
        doc.add_heading('조회 결과 요약 (실적 원본)',level=2)
        headers=['연월','국가','품목','수출액(US$)','수입액(US$)','출처']
        t=doc.add_table(rows=1,cols=len(headers));t.style='Table Grid'
        for c,h in zip(t.rows[0].cells,headers):c.text=h
        for r in rows[:35]:
            if not isinstance(r,dict):continue
            cells=t.add_row().cells
            items=[str(r.get('period','')),str(r.get('country',''))[:18],str(r.get('item',''))[:23],
                   (f"{r.get('export'):,.0f}" if isinstance(r.get('export'),(int,float)) else '—'),
                   (f"{r.get('import'):,.0f}" if isinstance(r.get('import'),(int,float)) else '—'),str(r.get('source',''))[:40]]
            for c,val in zip(cells,items):c.text=val
        if len(rows)>35:doc.add_paragraph(f'표는 상위 35행만 수록. 전체 {len(rows)}행 중 나머지는 원본 파일 참조.')
    doc.add_heading('3. 관련 뉴스 및 근거',level=1)
    articles=payload.get('articles',[])
    if isinstance(articles,list) and articles:
        for a in articles[:25]:
            if not isinstance(a,dict):continue
            name=str(a.get('title','미제목'))[:180]
            p=doc.add_paragraph(style='List Bullet');p.add_run(name).bold=True
            date=str(a.get('date',''))[:40];url=str(a.get('url',''))[:300]
            doc.add_paragraph(f'기사일자: {date or "미확인"} | 링크: {url or "미입력"}')
            doc.add_paragraph('기사 요지(사용자 메모): '+str(a.get('summary',''))[:700])
    else:doc.add_paragraph('관련 기사를 직접 검색하여 확인한 후 등록하세요. 뉴스는 자동으로 작성·검증되지 않습니다.')
    doc.add_heading('4. 해석 및 확인 필요 사항',level=1)
    doc.add_paragraph(str(payload.get('hypothesis') or '원인 및 인과관계는 아직 확인되지 않았습니다.')[:10000])
    doc.add_heading('5. 자료 출처·검토 기준',level=1)
    sources=payload.get('sources',[])
    for s in (sources[:40] if isinstance(sources,list) else []):
        doc.add_paragraph(str(s)[:220],style='List Bullet')
    doc.add_paragraph('국가별 물량이 실제 이동했다는 결론은 국가별 수출액 동시 변화만으로 확정할 수 없습니다. HS/MTI 분류, 업체 물류, 관세·통계 기록, 관련 기사로 추가 검증하세요.')
    footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run('부산 수출입동향 분석실 | 분석 보조자료 · 최종 검토 필요')
    b=io.BytesIO();doc.save(b);return b.getvalue()

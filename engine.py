"""Busan import/export HWPX reporter. Accuracy-first, read-only source parser.

The template is a user-approved HWPX. Tables are rewritten in place; never
reconstructed, expanded, or resized. Unsupported commentary is left blank.
"""
from __future__ import annotations
import datetime as dt
from collections import defaultdict
from copy import deepcopy
from decimal import Decimal,ROUND_HALF_UP
from dataclasses import dataclass,field
import io,re,zipfile
from pathlib import Path
from lxml import etree
from xlreader import load_excel,number

HP='http://www.hancom.co.kr/hwpml/2011/paragraph'
HH='http://www.hancom.co.kr/hwpml/2011/head'
N={'hp':HP,'hh':HH}
Q=lambda s:'{%s}%s'%(HP,s)
CONTINENTS=['아시아','북미','유럽','중남미','중동','아프리카','대양주','기타지역']
COUNTRY_GROUPS={
 '아시아':'중국 일본 베트남 싱가포르 인도(인디아) 인도 인도네시아 말레이시아 대만 태국 필리핀 홍콩 파키스탄 방글라데시 캄보디아 미얀마 스리랑카 네팔 몽골 라오스 브루나이',
 '북미':'미국 캐나다',
 '유럽':'벨기에 영국 독일 프랑스 이탈리아 네덜란드 스페인 폴란드 러시아 스위스 오스트리아 노르웨이 스웨덴 핀란드 덴마크 체코 헝가리 아일랜드 튀르키예 루마니아 그리스 포르투갈 우크라이나',
 '중남미':'멕시코 브라질 칠레 페루 콜롬비아 아르헨티나 에콰도르 파나마 과테말라 코스타리카',
 '중동':'사우디아라비아 아랍에미리트 이란 이라크 이스라엘 카타르 쿠웨이트 오만 바레인 요르단',
 '아프리카':'라이베리아 남아프리카공화국 이집트 모로코 나이지리아 가나 케냐 알제리',
 '대양주':'호주 뉴질랜드 파푸아뉴기니 피지',
}
GROUP_OF={c:g for g,names in COUNTRY_GROUPS.items() for c in names.split()}
GROUP_OF.update({'인도(인디아)':'아시아','호주':'대양주','인디아':'아시아','그리스':'유럽','러시아':'유럽'})
REGIONS={'경기남부':'경기(남부)','경기북부':'경기(북부)','전남광주':'전남(광주)'}

class Reject(ValueError):pass

@dataclass
class Audit:
    filled:int=0
    empty:int=0
    notes:list=field(default_factory=list)
    issues:list=field(default_factory=list)
    def note(self,msg):self.notes.append(msg)
    def issue(self,msg):
        if msg not in self.issues:self.issues.append(msg)


def fmoney(v):
    n=number(v)
    if n is None:return ''
    rounded=int((Decimal(str(n))/Decimal('1000000')).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    return f'{rounded:,}'

def fpct(v):
    n=number(v)
    if n is None:return ''
    return f'{Decimal(str(n)).quantize(Decimal("0.1"),rounding=ROUND_HALF_UP):,.1f}'

def set_text(cell,value):
    """Change text only. Do not alter widths, rows, paragraph/character styles."""
    value=str(value or '')
    texts=cell.xpath('.//hp:t',namespaces=N)
    if not texts:
        para=cell.find('.//'+Q('p'))
        if para is None:return
        run=para.find(Q('run'))
        if run is None:
            run=etree.SubElement(para,Q('run'))
        texts=[etree.SubElement(run,Q('t'))]
    # Drop obsolete inline revision/highlight controls and their tails.
    # Otherwise a template's highlight marker can leak stale punctuation after rewritten text.
    for t in texts:
        for child in list(t):t.remove(child)
        t.text=''
    texts[0].text=value

def get_text(element):return ''.join(element.xpath('.//hp:t/text()',namespaces=N)).strip()

def cells(row):return row.xpath('./hp:tc',namespaces=N)

def set_row(row,name,metrics,rank=None,audit=None):
    cc=cells(row); is_ranked=(len(cc)==14)
    label_idx=1 if is_ranked else 0
    if is_ranked:set_text(cc[0],str(rank) if rank is not None else '')
    set_text(cc[label_idx],name)
    for j in range(12):
        value=metrics[j] if metrics and j<len(metrics) else ''
        dest=cc[label_idx+1+j] if label_idx+1+j<len(cc) else None
        if dest is not None:
            set_text(dest,value)
            if audit:
                if value=='' or value is None:audit.empty+=1
                else:audit.filled+=1


# Month-to-month category allocation: an MTI2 parent consumes one fixed table
# row, each MTI4 item consumes one. Optimise *exact* available row capacity.
# This prevents the silent blank row caused by dropping the last category.
def select_mti_sections(items, capacity, trade='수출'):
    candidates=[x for x in items if x['rank']<=20 and x['current']>0 and x['group_code']]
    grouped=defaultdict(list)
    for it in candidates:grouped[it['group_code']].append(it)
    def importance(x):
        # Monthly amount rank dominates; large material changes are rewarded.
        r=x['rank']; delta=abs(x['current']-x['previous'])
        return (110.0/(1+.25*(r-1))
                +min(35.0,delta/1e7*5)
                + (min(15,abs(x['rate'])/35) if x['rate'] is not None and x['current']>=1e7 else 0)
                + (35 if r<=10 and x['rate'] is not None and abs(x['rate'])>=25 and x['current']>=2e7 else 0))
    groups=sorted(grouped.items(),key=lambda kv:min(x['rank'] for x in kv[1]))
    dp={0:(0.0,[])}
    for code,choices in groups:
        sorted_choices=sorted(choices,key=lambda x:(-importance(x),x['rank']))
        nextdp=dict(dp)
        for used,(value,selected) in dp.items():
            for n in range(1,len(sorted_choices)+1):
                cost=n+1
                if used+cost>capacity:break
                chosen=sorted_choices[:n]
                score=value+sum(importance(x) for x in chosen)
                previous=nextdp.get(used+cost)
                if previous is None or score>previous[0]:
                    nextdp[used+cost]=(score,selected+chosen)
        dp=nextdp
    if capacity not in dp:
        # If source top-20 simply cannot fill capacity, do not invent data.
        best=max(dp)
        selected=dp[best][1]
    else:selected=dp[capacity][1]
    output=defaultdict(list)
    for x in selected:output[x['group_code']].append(x)
    return [(g,sorted(its,key=lambda z:z['rank'])) for g,its in sorted(output.items(),key=lambda kv:min(z['rank'] for z in kv[1]))]

# Format style helper: keep the original border geometry and only recolour
# using a copy of each unique border/fill style.
def create_style_context(header):
    hh='{http://www.hancom.co.kr/hwpml/2011/head}'
    hc='{http://www.hancom.co.kr/hwpml/2011/core}'
    bf=header.find('.//'+hh+'borderFills')
    cp=header.find('.//'+hh+'charProperties')
    if bf is None or cp is None:raise Reject('HWPX 서식 목록을 찾지 못했습니다.')
    return {'bf':bf,'cp':cp,'hh':hh,'hc':hc,'bf_cache':{},'cp_cache':{},
            'bf_next':max(int(x.get('id')) for x in bf)+1,
            'cp_next':max(int(x.get('id')) for x in cp)+1}

def set_cell_fill_color(cell, color, styles):
    src_id=cell.get('borderFillIDRef')
    cache_key=(src_id,color)
    if cache_key not in styles['bf_cache']:
        bf=styles['bf'];hh=styles['hh'];hc=styles['hc']
        old=next((x for x in bf if x.get('id')==src_id),None)
        if old is None:raise Reject(f'셀 테두리 서식 번호 누락: {src_id}')
        clone=deepcopy(old)
        clone.set('id',str(styles['bf_next']));styles['bf_next']+=1
        brush=clone.find(hc+'fillBrush')
        if brush is None:brush=etree.SubElement(clone,hc+'fillBrush')
        win=brush.find(hc+'winBrush')
        if win is None:win=etree.SubElement(brush,hc+'winBrush')
        win.set('faceColor',color)
        win.set('hatchColor','#000000')
        win.set('alpha','0')
        bf.append(clone);bf.set('itemCnt',str(len(bf)))
        styles['bf_cache'][cache_key]=clone.get('id')
    cell.set('borderFillIDRef',styles['bf_cache'][cache_key])

def set_row_band(row, is_group, styles):
    shade='#F9EDF5' if is_group else '#FFFFFF'
    for c in cells(row):set_cell_fill_color(c,shade,styles)

def set_cell_safe_spacing(cell, value, styles):
    # This is an isolated CHAR SPACING change. Never resize a cell or shrink text.
    length=len(str(value))
    target=-18 if length>=8 else (-10 if length>=7 else None)
    if target is None:return
    for run in cell.xpath('.//hp:run[hp:t]',namespaces=N):
        old_id=run.get('charPrIDRef')
        if not old_id:continue
        key=(old_id,target)
        if key not in styles['cp_cache']:
            cp=styles['cp'];hh=styles['hh']
            old=next((x for x in cp if x.get('id')==old_id),None)
            if old is None:continue
            clone=deepcopy(old);clone.set('id',str(styles['cp_next']));styles['cp_next']+=1
            spacing=clone.find(hh+'spacing')
            if spacing is None:spacing=etree.SubElement(clone,hh+'spacing')
            for k in list(spacing.attrib):spacing.set(k,str(max(-30,min(int(spacing.get(k,'0')),target))))
            cp.append(clone);cp.set('itemCnt',str(len(cp)))
            styles['cp_cache'][key]=clone.get('id')
        if key in styles['cp_cache']:run.set('charPrIDRef',styles['cp_cache'][key])


def fit_table6_long_product_labels(header, root1, audit=None):
    """Keep two MTI4 names on one line without touching table geometry.

    Hancom table 6's name column is about 82pt wide. The two known longer
    labels are 11pt in the original template and can wrap to two lines,
    pushing the entire Table 7 notes below the page break. Clone ONLY these
    cell-run charPr styles and use a readable 10pt font with restrained
    negative tracking. Never shorten names, edit numerical values, or resize
    any row/cell. Apply after MTI4 rows have been populated each month.
    """
    H='{http://www.hancom.co.kr/hwpml/2011/head}'
    chars=header.find('.//'+H+'charProperties')
    if chars is None:raise Reject('표6 품목 서식: 글자 속성 목록이 없습니다.')
    tables=root1.xpath('//hp:tbl',namespaces=N)
    if len(tables)!=6:raise Reject('표6 품목 서식: 표 구성 변경 감지')
    rows=tables[4].xpath('./hp:tr',namespaces=N)
    targets={'회전기기부분품':-18,'전동축및기어':-12}
    found=[]
    style_cache={}
    next_id=max(int(x.get('id')) for x in chars)+1
    for row in rows[3:]:
        cc=cells(row)
        if not cc:continue
        label=cc[1]
        text=get_text(label)
        norm=re.sub(r'\(\d+\)$','',text).replace(' ','')
        if norm not in targets:continue
        ps=label.xpath('.//hp:p',namespaces=N)
        if len(ps)!=1:raise Reject(f'표6 {text}: 셀에 복수 문단이 발견되어 자동 조정 중단')
        tracking=targets[norm]
        for run in ps[0].xpath('./hp:run[hp:t]',namespaces=N):
            old_id=run.get('charPrIDRef')
            if old_id is None:raise Reject('표6 품목명 글자 속성 참조가 없습니다.')
            key=(old_id,tracking)
            if key not in style_cache:
                old=next((x for x in chars if x.get('id')==old_id),None)
                if old is None:raise Reject(f'글자 서식 ID {old_id} 누락')
                clone=deepcopy(old)
                clone.set('id',str(next_id));next_id+=1
                clone.set('height','1000')  # 10 pt, original is 11pt
                spacing=clone.find(H+'spacing')
                if spacing is None:spacing=etree.SubElement(clone,H+'spacing')
                for lang in ('hangul','latin','hanja','japanese','other','symbol','user'):
                    spacing.set(lang,str(tracking))
                chars.append(clone)
                style_cache[key]=clone.get('id')
            run.set('charPrIDRef',style_cache[key])
        found.append((text,tracking))
    chars.set('itemCnt',str(len(chars)))
    if audit is not None:
        if found:
            audit.note('표6 장문 품목명 1줄 표시: '+', '.join(f'{name} (10pt, 자간 {tracking}%)' for name,tracking in found)+'; 표 높이·내용 불변')
        else:
            audit.note('표6 장문 품목명 1줄 맞춤 대상 미등장: 기존 품목명 서식 유지')
    return found

def use_word_safe_left_align(header,root,positions):
    # Only the requested long continent descriptions are affected.
    # We clone the paragraph style rather than changing unrelated paragraphs.
    hh='{http://www.hancom.co.kr/hwpml/2011/head}'
    container=header.find('.//'+hh+'paraProperties')
    if container is None:return
    nextid=max(int(x.get('id')) for x in container)+1
    cache={}
    paragraphs=root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    for idx in positions:
        if idx>=len(paragraphs):continue
        p=paragraphs[idx];source=p.get('paraPrIDRef')
        if source not in cache:
            old=next((x for x in container if x.get('id')==source),None)
            if old is None:continue
            clone=deepcopy(old);clone.set('id',str(nextid));nextid+=1
            align=clone.find(hh+'align')
            if align is not None:align.set('horizontal','LEFT')
            br=clone.find(hh+'breakSetting')
            if br is not None:
                br.set('breakNonLatinWord','KEEP_WORD')
                br.set('breakLatinWord','KEEP_WORD')
            container.append(clone)
            cache[source]=clone.get('id')
        if source in cache:p.set('paraPrIDRef',cache[source])
    container.set('itemCnt',str(len(container)))

def normalise_generated_body_style(root, positions):
    # No template-specific red run (74), bold first run (58), exaggerated
    # tracking or artificial highlighting can leak into generated prose.
    pp=root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    for idx in positions:
        if idx>=len(pp):continue
        p=pp[idx]
        for run in p.xpath('./hp:run[hp:t]',namespaces=N):
            run.set('charPrIDRef','47')  # 14pt regular black, 0 tracking

def _line_page_style(header, paragraph, *, line_pct=None, keep_lines=False):
    """Only local paragraph spacing & keep-lines, on a clone of original style.

    The template carries two HWPX compatibility line-spacing representations;
    update both. Paragraph/node count and font point sizes remain unchanged.
    """
    H='{http://www.hancom.co.kr/hwpml/2011/head}'
    props=header.find('.//'+H+'paraProperties')
    if props is None:return
    oldid=paragraph.get('paraPrIDRef')
    old=next((pr for pr in props if pr.get('id')==oldid),None)
    if old is None:return
    clone=deepcopy(old)
    clone.set('id',str(max(int(pr.get('id')) for pr in props)+1))
    if line_pct is not None:
        for el in clone.iter(H+'lineSpacing'):
            if el.get('type')=='PERCENT':el.set('value',str(line_pct))
    if keep_lines:
        br=clone.find(H+'breakSetting')
        if br is not None:
            br.set('keepLines','1')
            br.set('widowOrphan','1')
    props.append(clone);props.set('itemCnt',str(len(props)))
    paragraph.set('paraPrIDRef',clone.get('id'))

def _local_tracking(header, paragraph, percent):
    """Compress only glyph tracking of designated country prose; no font-size change."""
    H='{http://www.hancom.co.kr/hwpml/2011/head}'
    chars=header.find('.//'+H+'charProperties')
    if chars is None:return
    ids={r.get('charPrIDRef') for r in paragraph.xpath('./hp:run[hp:t]',namespaces=N)}
    remap={}
    for oldid in ids:
        old=next((cp for cp in chars if cp.get('id')==oldid),None)
        if old is None:continue
        clone=deepcopy(old)
        clone.set('id',str(max(int(cp.get('id')) for cp in chars)+1))
        spacing=clone.find(H+'spacing')
        if spacing is not None:
            for k in spacing.attrib:spacing.set(k,str(percent))
        chars.append(clone);remap[oldid]=clone.get('id')
    chars.set('itemCnt',str(len(chars)))
    for run in paragraph.xpath('./hp:run[hp:t]',namespaces=N):
        if run.get('charPrIDRef') in remap:run.set('charPrIDRef',remap[run.get('charPrIDRef')])

def _blank_paragraph_height(header, paragraph, height):
    """Only change the size of a truly empty Enter paragraph (in 1/100 pt).

    A line-spacing percentage by itself does not shrink an empty HWPX line:
    the HWP engine also takes its run's font height into account. Clone the
    run style locally, leaving all text runs and all table styles untouched.
    """
    if get_text(paragraph) or paragraph.xpath('.//hp:tbl', namespaces=N):
        return False
    H='{http://www.hancom.co.kr/hwpml/2011/head}'
    properties=header.find('.//'+H+'charProperties')
    if properties is None:return False
    first_id=max(int(x.get('id')) for x in properties)+1
    cache={}
    for run in paragraph.xpath('./hp:run',namespaces=N):
        old_id=run.get('charPrIDRef')
        if old_id is None:continue
        if old_id not in cache:
            original=next((x for x in properties if x.get('id')==old_id),None)
            if original is None:continue
            clone=deepcopy(original)
            clone.set('id',str(first_id));first_id+=1
            clone.set('height',str(height))
            properties.append(clone)
            cache[old_id]=clone.get('id')
        if old_id in cache:run.set('charPrIDRef',cache[old_id])
    properties.set('itemCnt',str(len(properties)))
    return bool(cache)


def layout_compact_pages(header,root0,root1,audit):
    """V8: preserve v7 page compaction and shorten two Table 6 wrapped labels.

    The overview stays at baseline geometry so Table 1 and its notes fit.
    Continent narrative + Table 3, and Table 6 + Table 7, get tighter blank
    Enter lines and regular body line spacing to reduce overflow. This cannot
    prove a final page count without Hancom's proprietary pagination engine.
    """
    p0=root0.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    p1=root1.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    if (len(p0),len(p1))!=(84,58):
        raise Reject('기준 한글 문단 수가 달라 페이지 배치 보정을 중단했습니다.')
    for idx in (24,53,65):
        if p0[idx].get('pageBreak')!='1':raise Reject('기존 페이지 경계가 변경되어 중단합니다.')
    for idx in (16,31):
        if p1[idx].get('pageBreak')!='1':raise Reject('기존 페이지 경계가 변경되어 중단합니다.')
    # The two import product tables follow the import chart and processing table.
    # Preserve the explicit start of the final report page.
    p1[43].set('pageBreak','1')

    # 1. Overview / Table 1: restore the PROVEN v4 spacing profile.
    #    V5 enlarged eight blank paragraphs before Table1 to 12pt /145%,
    #    sending the table onto the next page. The v4 baseline instead uses
    #    84%-line spacing on just six existing blank separators. Keep 35 and 45
    #    untouched and keep the table/caption/notes completely unchanged.
    for idx in (29,31,33,39,41,43):
        if get_text(p0[idx]) or p0[idx].xpath('.//hp:tbl',namespaces=N):
            raise Reject('개황 빈 문단 구조가 달라 표1 위치 복원 중단')
        _line_page_style(header,p0[idx],line_pct=84)

    # 2. Continent/three-country commentary+표3. Tighten empty Enter space,
    #    not the table cells. Do not turn Hangul into break-inside-word mode.
    for idx in (1,3,5,7,9,11):
        if not get_text(p1[idx]):
            _blank_paragraph_height(header,p1[idx],450)   # 4.5 pt
            _line_page_style(header,p1[idx],line_pct=78)
    for idx in (2,4,6,8):
        _line_page_style(header,p1[idx],line_pct=108,keep_lines=True)
    for idx in (4,6,8):_local_tracking(header,p1[idx],-13)
    # Caption and notes: local vertical spacing, wording unchanged.
    for idx in (10,12,14,15):
        _line_page_style(header,p1[idx],line_pct=94)

    # 3. Last page/표6+표7: v7 micro-compaction. In v6 the final
    #    '주2 : 증감률은 ...' moved to the next printed page in Hancom.
    #    Reserve approximately one extra full note line by reducing ONLY
    #    existing paragraph line/blank spacings on the final report page.
    #    Do not modify fonts, text, table height, rows, columns or cell margins.
    for idx in (44,50,52):
        if get_text(p1[idx]) or p1[idx].xpath('.//hp:tbl',namespaces=N):
            raise Reject('표7 주석 간격 조정용 빈 문단 구조 변경 감지: 안전하게 중단')
        _blank_paragraph_height(header,p1[idx],140)  # 1.4 pt (v6: 4.5 pt)
        _line_page_style(header,p1[idx],line_pct=68)
    for idx in (43,51):
        _line_page_style(header,p1[idx],line_pct=96,keep_lines=True)
    for idx in (45,46,48,49,53,54,56,57):
        _line_page_style(header,p1[idx],line_pct=86)

    audit.note('v7 마지막 페이지: 표6·표7 사이 기존 빈 문단 3곳의 높이와 마지막 페이지 문단 줄간격만 국소 축소, 주2가 표7과 같은 페이지에 머물도록 공간 확보 (본문·표·수치·글자크기 변경 없음)')
    audit.note('1페이지 개황은 v6에서 복원한 위치와 서식을 그대로 유지하고, 표3도 v6 설정 유지')
    audit.note('HWPX XML과 ZIP을 검증하더라도 한컴오피스의 실제 페이지 나눔은 직접 열어 확인해야 합니다.')

def from_main(sheet,row):
    # Row layout: previous month CDE, current month FGH, previous YTD IJK,
    # current YTD OPQ (previous ANNUAL LMN is intentionally ignored).
    out=[]
    for cols in [('C','D','E'),('F','G','H'),('I','J','K'),('O','P','Q')]:
        out.extend([fmoney(sheet.val(row,cols[0])),fpct(sheet.val(row,cols[1])),fpct(sheet.val(row,cols[2]))])
    return out

def from_mti4(sheet,row):
    # Important: another 3-column layout than region/country sheet.
    out=[]
    for cols in [('D','E','F'),('G','H','I'),('J','K','L'),('P','Q','R')]:
        out.extend([fmoney(sheet.val(row,cols[0])),fpct(sheet.val(row,cols[1])),fpct(sheet.val(row,cols[2]))])
    return out

def get_mti4(sheet,trade):
    start,end=(5,24) if trade=='수출' else (31,50)
    records=[]
    for row in range(start,end+1):
        code=str(sheet.val(row,'B') or '').strip()
        if not code:continue
        rank=int(number(sheet.val(row,'A')) or (row-start+1))
        data={'code':code.zfill(4),'group_code':code.zfill(4)[:2],
              'name':str(sheet.val(row,'C') or ''),'rank':rank,
              'current':number(sheet.val(row,'G')) or 0,
              'previous':number(sheet.val(row,'D')) or 0,
              'rate':number(sheet.val(row,'H')),
              'metrics':from_mti4(sheet,row)}
        records.append(data)
    return sorted(records,key=lambda x:x['rank'])

def make_month_info(book):
    sh=book.sheet('1_지역총괄')
    header=str(sh.val(1,'A') or '')
    m=re.search(r'(20\d\d)\s*년\s*(\d{1,2})\s*월',header)
    if not m:raise Reject('기본 보고서 엑셀에서 기준 연월을 찾지 못했습니다 (1_지역총괄 A1).')
    year,month=map(int,m.groups())
    if not 1<=month<=12:raise Reject('월 범위가 잘못되었습니다.')
    return year,month

@dataclass
class Kstat:
    meta:str
    year:int
    month:int
    interval:str
    level:int
    is_china:bool
    rows:dict
    total:dict
    label:str

def read_kstat(data,name):
    book=load_excel(data,name)
    sh=next(iter(book.sheets.values()))
    meta=str(sh.val(2,'A') or '')
    m=re.search(r'년월\s*:\s*(20\d\d)년\s*(\d{1,2})월',meta)
    lev=re.search(r'품목\s*:\s*MTI\s*(\d)단위',meta)
    intr=re.search(r'당월/누계\s*:\s*(당월|누계)',meta)
    if not (m and lev and intr and '지역 : 부산' in meta):
        raise Reject(f'{name}: K-Stat 조회 조건을 인식하지 못했거나 지역이 부산이 아닙니다.')
    yy,mm=map(int,m.groups());level=int(lev.group(1));interval=intr.group(1)
    result={}
    for r in range(6,max(sh.rows.keys() or [5])+1):
        code=str(sh.val(r,'B') or '').strip()
        if not re.fullmatch(r'\d{2,4}',code):continue
        result[code]={
          'name':str(sh.val(r,'C') or '').strip(),
          'py_export':number(sh.val(r,'D')),'py_export_pct':number(sh.val(r,'E')),
          'py_import':number(sh.val(r,'F')),'py_import_pct':number(sh.val(r,'G')),
          'cy_export':number(sh.val(r,'I')),'cy_export_pct':number(sh.val(r,'J')),
          'cy_import':number(sh.val(r,'K')),'cy_import_pct':number(sh.val(r,'L')),
        }
    tot={}
    for side,pre_col,cur_col in [('export','D','I'),('import','F','K')]:
        tot['py_'+side]=number(sh.val(5,pre_col))
        tot['cy_'+side]=number(sh.val(5,cur_col))
    return Kstat(meta,yy,mm,interval,level,'국가 : 중국' in meta,result,tot,name)

def concat_kstats(inputs,year,month,audit,china_expected_totals=None):
    # named groups with explicit validation; labels do not control interpretation.
    groups={}
    expected={
      'mti2_month':(2,False,year,month,'당월'),
      'mti2_prior':(2,False,year-1,month,'누계'),
      'mti2_current':(2,False,year,month,'누계'),
      'china_month':(3,True,year,month,'당월'),
      'china_prior':(3,True,year-1,month,'누계'),
      'china_current':(3,True,year,month,'누계'),
    }
    for key,expect in expected.items():
        files=inputs.get(key,[])
        parts=[]
        for name,data in files:
            if not data:continue
            try:k=read_kstat(data,name)
            except (ValueError,Reject) as e:
                audit.issue(f'{key}/{name}: {e}');continue
            found=(k.level,k.is_china,k.year,k.month,k.interval)
            if found!=expect:
                audit.issue(f'{key}/{name}: 조회 조건 불일치. 기대={expect} / 실제={found} → 미반영')
                continue
            # The K-Stat export may display '국가: 중국' in the header yet
            # contain Busan ALL-country trade. The user's 2025-10 old (6).xls
            # is exactly this case. Validate the actual total, not only labels.
            if key.startswith('china_') and china_expected_totals is not None:
                idx={'china_month':1,'china_prior':2,'china_current':3}[key]
                expected_amount=china_expected_totals[idx]
                got=k.total.get('cy_import')
                if expected_amount is not None and got is not None:
                    if got > expected_amount + 1:
                        audit.issue(f'{key}/{name}: 중국 수입 합계 불일치 ({got:,.0f}달러 > 확인된 중국 총계 {expected_amount:,.0f}달러). 잘못 조회된 자료 → 제외')
                        continue
                    # Smaller subtotals are permitted only as *additional*
                    # targeted MTI code lookups, never as a replacement for
                    # the full-China cumulative source.
                    if got < expected_amount - 1:
                        if got > expected_amount * 0.25:
                            audit.issue(f'{key}/{name}: 중국 수입 소계 ({got:,.0f}달러)가 전체 실적 ({expected_amount:,.0f}달러)과 크게 달라 확인 필요 → 제외')
                            continue
                        audit.note(f'{key}/{name}: 중국 수입 전체 조회가 아닌 부분 품목 보충 조회로 처리 (소계 {got:,.0f}달러)')
            parts.append(k)
        if key.startswith('china_') and china_expected_totals is not None and parts:
            idx={'china_month':1,'china_prior':2,'china_current':3}[key]
            want=china_expected_totals[idx]
            if want is not None:
                full=[k for k in parts if k.total.get('cy_import') is not None and abs(k.total['cy_import']-want)<=1]
                if full:
                    full_names={id(k) for k in full}
                    parts=full+[k for k in parts if id(k) not in full_names]
                else:
                    audit.issue(f'{key}: 중국 전체 총계가 일치하는 기본 K-Stat 조회 파일 없음; 추가 조회 값만으로 확인되는 코드에 한해 작성')
        merged={};total=None
        for k in parts:
            if total is None:total=k.total
            for code,stats in k.rows.items():
                if code in merged and any((merged[code].get(c)!=stats.get(c)) for c in stats):
                    audit.issue(f'{key}: MTI {code} 값이 파일마다 달라서 해당 품목 미반영')
                    merged[code]=None
                elif code not in merged:merged[code]=stats
        groups[key]={'items':merged,'total':total,'sources':[k.label for k in parts]}
        if not parts:audit.issue(f'{key}: 유효한 K-Stat 원본이 없습니다. 해당 항목 공란')
    return groups

def amount_metrics(rec,code,group,trade,kstats,totals):
    prefix='export' if trade=='수출' else 'import'
    fields=[('mti2_month','py'),('mti2_month','cy'),('mti2_prior','cy'),('mti2_current','cy')] if group=='mti2' else [
      ('china_month','py'),('china_month','cy'),('china_prior','cy'),('china_current','cy')]
    result=[]
    for idx,(key,side) in enumerate(fields):
        st=kstats[key]['items'].get(code)
        if not st:
            result.extend(['','','']);continue
        a=st.get(f'{side}_{prefix}')
        pct=st.get(f'{side}_{prefix}_pct')
        base=totals[idx]
        share=(a/base*100) if a is not None and base else None
        result.extend([fmoney(a),fpct(pct),fpct(share)])
    return result

def calc_total_month(sh,trade):
    row,change=(6,7) if trade=='수출' else (8,9)
    return [v for period in [('C',),('E',),('G',),('K',)] for v in (sh.val(row,period[0]),sh.val(change,period[0]))]

def rate_line(val,kind):
    n=number(val)
    if n is None:return ''
    return f'{abs(n):.1f}% {"증가" if n>=0 else "감소"}'

def won_dollar_phrase(n):
    """US dollar integer to Korean 억·만 달러, rounded to nearest 10k."""
    if n is None:return ''
    m=int((Decimal(str(n))/Decimal('10000')).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
    big,small=divmod(m,10000)
    if big and small:return f'{big:,}억 {small:,}만 달러'
    if big:return f'{big:,}억 달러'
    return f'{small:,}만 달러'

def table_signature(root):
    sig=[]
    for tbl in root.xpath('//hp:tbl',namespaces=N):
        sig.append((len(tbl.xpath('./hp:tr',namespaces=N)),tbl.get('colCnt'),len(tbl.xpath('.//hp:tc',namespaces=N))))
    return sig

def blank_top_level_facts(root,year,month,fmt,audit):
    """Overwrite any old-month commentary in top-level body; leave uncertain facts blank."""
    # Text replacement occurs in explicitly chosen paragraphs only, not table values.
    for p in root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N):
        t=get_text(p)
        if not t or not ('ㅇ' in t or t.lstrip().startswith('-')):continue
        # Text contains numeric commentary; will be overwritten by stable rule below,
        # or erased to prevent stale facts from previous report leaking into new month.
        if any(x in t for x in ['전년 동월','전년당월','수출 금액','수출액','수입액','수출은','수입은','수출이','수입이','무역수지','對','대륙별','금년 월별','수출실적','품목','금액']):
            # Do not erase special section headings; only narrative paragraphs.
            if t.lstrip().startswith(('ㅇ','-')):set_text(p,'')

def set_p_by_prefix(root,prefix,value,audit):
    matches=[]
    for p in root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N):
        if get_text(p).startswith(prefix):matches.append(p)
    if len(matches)==1:set_text(matches[0],value);audit.filled+=1
    elif matches:audit.note(f'본문 {prefix} 위치가 {len(matches)}개라 자동 입력하지 않았습니다.')


def choose_import_commentary(imports):
    """Choose fact-supported import product mentions in the historical report style.

    Historical monthly originals (Mar, Apr, May, Sep 2026) first discuss current
    MTI4 ranking 1..5 where meaningful, then occasionally a 6..10 ranked item
    with an unusually large *absolute* rate and material dollar movement.
    The same item never appears in both increased and decreased sets.
    This selection is for prose ONLY; table row allocation remains unchanged.
    """
    eligible=[x for x in imports if x['rank']<=20 and x['current']>0 and x['rate'] is not None]
    def material(x):
        return x['current']>=8_000_000 and abs(x['current']-x['previous'])>=3_000_000
    top5=[x for x in eligible if x['rank']<=5]
    special=[x for x in eligible if 5<x['rank']<=10 and material(x) and abs(x['rate'])>=150]
    union=top5+sorted(special,key=lambda x:(-abs(x['current']-x['previous']),x['rank']))
    # Do not call a 0.1% fluctuation a major import driver.
    rises=[x for x in union if x['rate']>=5.0]
    falls=[x for x in union if x['rate']<-.05]
    # If five leading ranks are one-sided, include the most material opposite
    # change from the top 10 only if supported by meaningful absolute movement.
    if not falls:
        extra=next((x for x in eligible if x['rank']<=10 and x['rate']<-5 and material(x)),None)
        if extra and extra not in union:falls.append(extra)
    if not rises:
        extra=next((x for x in eligible if x['rank']<=10 and x['rate']>5 and material(x)),None)
        if extra and extra not in union:rises.append(extra)
    # Keep material 100%+ growth within the five leading MTI4 ranks even
    # when three smaller-change items precede it. Example: October imports
    # rank 5 batteries (+748.5%) rather than rank 3 car parts (+39.7%).
    priority=[x for x in rises if x['rank']<=5 and x['rate']>=100 and material(x)]
    priority=sorted(priority,key=lambda x:x['rank'])
    candidates=priority+[x for x in sorted(rises,key=lambda x:x['rank']) if x not in priority]
    rises=sorted(candidates[:3],key=lambda x:x['rank'])
    falls=sorted(falls,key=lambda x:x['rank'])[:3]
    # A single standout change can be mentioned in the overview, but don't
    # label rank-high / tiny-dollar percent outliers as the main driver.
    highlights=[x for x in rises if material(x)]
    highlight=max(highlights,key=lambda x:abs(x['current']-x['previous']),default=None)
    if highlight is None and rises:highlight=rises[0]
    return {'rises':rises,'falls':falls,'highlight':highlight,
            'new':[x for x in top5 if x['previous']==0 and x['current']>0]}


def make_import_mti_sentence(items, author):
    """Exactly monthly import MTI4 ranks 1-5; rises before falls.

    Preserve the historical `item(rank)(growth) ...의 수입은 증가, ...의
    수입은 감소` grammar. Never include rank 6+ just to invent an opposite
    direction, and do not silently classify missing data as a fall.
    """
    top5=sorted([x for x in items if 1<=x['rank']<=5],key=lambda x:x['rank'])
    if len(top5)!=5 or sorted(x['rank'] for x in top5)!=[1,2,3,4,5]:
        return 'ㅇ '+author('수입금액 상위 1~5위 품목 확인 요망')
    if any(x['rate'] is None for x in top5):
        return 'ㅇ '+author('수입금액 상위 1~5위 증감률 확인 요망')
    def tag(x):return f'{x["name"]}{x["rank"]}위({fpct(x["rate"])}%)'
    def clause(records,verb):
        if not records:return ''
        return ', '.join(tag(x) for x in records)+f'의 수입은 {verb}'
    inc=[x for x in top5 if x['rate']>0]
    dec=[x for x in top5 if x['rate']<0]
    flat=[x for x in top5 if x['rate']==0 and x['previous']>0]
    new_items=[x for x in top5 if x['rate']==0 and x['previous']==0 and x['current']>0]
    # The source may encode the previous-month zero denominator as 0.0%;
    # this does NOT mean 'unchanged'. Describe the new occurrence accurately.
    new_clause=', '.join(f'{x["name"]}{x["rank"]}위' for x in new_items)+'의 신규 수입 발생' if new_items else ''
    parts=[clause(inc,'증가'),clause(dec,'감소'),clause(flat,'변동 없음'),new_clause]
    return 'ㅇ '+', '.join(x for x in parts if x)


def format_import_rank_superscript(header, root1, audit):
    """Apply original Sep HWPX 'Alt+Shift+P' upper-script to only N위 tokens.

    The reference report uses charPr=69 for text (14pt, -10% tracking)
    and charPr=72 for N위 (<hh:supscript/>). Keep the existing paragraph,
    table geometry, and pagination. No manual line breaks inside words.
    """
    ps=root1.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    if len(ps)!=58:raise Reject('기존 수입 해설 문단 수가 변경되어 위첨자 서식을 적용하지 않았습니다.')
    p=ps[43]
    text=get_text(p)
    if not text.startswith('ㅇ ') or '의 수입은 ' not in text:
        # When source figures were not verified, no rank should be faked.
        if '[작성자 입력' in text:return
        raise Reject('수입 품목 설명 문장의 구조가 기존 양식과 다릅니다.')
    tokens=re.split(r'(\d{1,2}위)(?=\(|의|,)' ,text)
    if ''.join(tokens)!=text:raise Reject('수입 품목 순위 위첨자 분리 중 문자 불일치')
    charprops=header.find('.//{http://www.hancom.co.kr/hwpml/2011/head}charProperties')
    if charprops is None:raise Reject('글자 서식 목록이 없습니다.')
    styles={x.get('id'):x for x in charprops}
    if '69' not in styles or '72' not in styles or styles['72'].find('{http://www.hancom.co.kr/hwpml/2011/head}supscript') is None:
        raise Reject('기존 보고서의 순위 위첨자 서식이 없습니다.')
    # The template has 13 existing runs with one hp:t each. Reuse those runs,
    # preserve all structural siblings and avoid adding new empty paragraphs.
    runs=p.xpath('./hp:run[hp:t]',namespaces=N)
    if len(runs)<len(tokens):raise Reject('수입 설명 글자 구간 수가 부족합니다.')
    for i,run in enumerate(runs):
        ts=run.xpath('./hp:t',namespaces=N)
        if not ts:continue
        chunk=tokens[i] if i<len(tokens) else ''
        for t in ts:
            for child in list(t):t.remove(child)
            t.text=''
        ts[0].text=chunk
        run.set('charPrIDRef','72' if i<len(tokens) and re.fullmatch(r'\d{1,2}위',chunk) else '69')
    if get_text(p)!=text:raise Reject('수입 품목 순위 서식 변경으로 문장 손실')
    sup=[r for r in runs if r.get('charPrIDRef')=='72' and get_text(r)]
    if len(sup)!=len(re.findall(r'\d{1,2}위(?=\(|의|,)' ,text)):
        raise Reject('순위 위첨자 글자 구간 수 불일치')
    audit.note(f'표6 상단 설명: 당월 MTI4 수입액 상위 1~5위, 증가 먼저·감소 이후. 순위 위첨자 {len(sup)}곳은 기존 9월 완성본과 동일한 글자 서식(72) 적용')



def format_destination_rank_superscript(header, root1, audit):
    """Style *only* MTI4 rank tokens in US/China/Japan export prose.

    Historical September reports use genuine HWPX <hh:supscript/> for '1위',
    '2위', etc. Keep each country's original 14pt base font/tracking,
    paragraph spacing, existing runs, text, and table geometry untouched.
    Only rewrite run boundaries to isolate rank tokens and add cloned
    superscript character styles. This works for dynamically changing items.
    """
    ps=root1.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)
    if len(ps)!=58:
        raise Reject('대륙별 수출동향 문단 수가 원본과 달라 위첨자를 적용하지 않았습니다.')
    charprops=header.find('.//{%s}charProperties' % HH)
    if charprops is None:
        raise Reject('HWPX 글자 서식 목록이 없어 순위 위첨자를 적용하지 않았습니다.')
    props_by_id={cp.get('id'):cp for cp in charprops}
    next_id=max(int(i) for i in props_by_id if i is not None and i.isdigit())+1
    total=0
    for index,country in [(4,'미국'),(6,'중국'),(8,'일본')]:
        p=ps[index]
        source=get_text(p)
        if not source:
            continue
        if not source.startswith(f'- [{country}] 對{country} 수출은'):
            raise Reject(f'{country} 수출 해설 문장 위치가 기존 양식과 달라 적용을 중단했습니다.')
        chunks=re.split(r'(\d{1,2}위)(?=\(|,|\s|의|$)',source)
        if ''.join(chunks)!=source:
            raise Reject(f'{country} 설명 순위 분리 중 문자 일치 실패')
        count=sum(1 for token in chunks if re.fullmatch(r'\d{1,2}위',token))
        if count==0:
            continue
        runs=p.xpath('./hp:run[hp:t]',namespaces=N)
        if len(runs)<len(chunks):
            raise Reject(f'{country} 수출 해설의 기존 글자 구간이 부족해 적용하지 않았습니다.')
        base_id=runs[0].get('charPrIDRef')
        base=props_by_id.get(base_id)
        if base is None:
            raise Reject(f'{country} 수출 해설의 기본 글자 서식을 찾을 수 없습니다.')
        sup_style=deepcopy(base)
        sup_style.set('id',str(next_id))
        if sup_style.find('{%s}supscript' % HH) is None:
            etree.SubElement(sup_style,'{%s}supscript' % HH)
        charprops.append(sup_style)
        sup_id=str(next_id)
        next_id+=1
        # Only use existing hp:run elements; preserve hp:p and its children count.
        for i,run in enumerate(runs):
            tt=run.xpath('./hp:t',namespaces=N)
            if not tt:
                continue
            for t in tt:
                for child in list(t):t.remove(child)
                t.text=''
            token=chunks[i] if i<len(chunks) else ''
            tt[0].text=token
            run.set('charPrIDRef',sup_id if re.fullmatch(r'\d{1,2}위',token) else base_id)
        if get_text(p)!=source:
            raise Reject(f'{country} 수출 설명이 위첨자 처리로 변경되었습니다.')
        actual=sum(1 for run in runs if run.get('charPrIDRef')==sup_id and get_text(run))
        if actual!=count:
            raise Reject(f'{country} 수출 품목 순위 위첨자 수가 일치하지 않습니다.')
        total+=actual
    charprops.set('itemCnt',str(len(charprops)))
    audit.note(f'대륙별 수출동향의 미국·중국·일본 품목 순위 {total}곳을 위첨자(Alt+Shift+P)로 표시. 기존 문장·수치·서식 간격·표 구조 유지')



def _narrative_visual_width(text):
    """Conservative, font-independent width guard for short Hangul prose.

    Full-width Korean syllables take one unit, ASCII punctuation/digits roughly
    half. A deliberate line break is added only at an editorial clause boundary,
    never between syllables of a Korean word. This is an estimate, not a Hancom
    render of the page.
    """
    import unicodedata
    return sum(1 if unicodedata.east_asian_width(ch) in ('W','F') else .55 for ch in text)


def _overview_continent_lines(regions):
    """Follow the Mar/Apr/May/Sep overview style, using two concise clauses.

    The old v14 glued 4+4 destinations with middle dots into a very long
    sentence, which Hancom split inside '아프리카'. Keep comma+space-separated
    proper names and line-break between complete *clauses*. No data is added
    or inferred. Choose a minority-contrast sentence for 5+/3- months.
    """
    if len(regions)!=8 or any(v is None for v in regions.values()):return None
    rising=[n for n,v in regions.items() if v>0]
    falling=[n for n,v in regions.items() if v<0]
    flat=[n for n,v in regions.items() if v==0]
    if flat:
        # An exact 0.0% is neither growth nor decline. Keep the summary
        # short rather than listing too many locations and splitting a name.
        return (f'- 대륙별 수출은 {len(rising)}개 지역 증가, {len(falling)}개 지역 감소,',
                f'{len(flat)}개 지역은 전년 동월과 같은 수준')
    if rising and falling and len(rising)<=4 and len(falling)<=4:
        # The strongest positive-region YoY rate, when it fits, adds the
        # concrete highlight commonly used in finished historical reports.
        best=max(rising,key=lambda nm:regions[nm])
        highlighted=[f'{n}({fpct(regions[n])}%)' if n==best else n for n in rising]
        high='- 대륙별로 '+', '.join(highlighted)+' 수출 증가,'
        simple='- 대륙별로 '+', '.join(rising)+' 수출 증가,'
        first=high if _narrative_visual_width(high)<=30.5 else simple
        second=', '.join(falling)+' 수출 감소'
        return first,second
    if falling and len(falling)<=3 and rising:
        # The historical '...을 제외하고' sentence can conceal whether the
        # excluded locations fell or stayed flat. Here the sign is explicit.
        first='- 대륙별로 '+', '.join(falling)+' 수출은 감소,'
        second='그 외 지역에서는 수출 증가'
        return first,second
    if rising and len(rising)<=3 and falling:
        first='- 대륙별로 '+', '.join(rising)+' 수출은 증가,'
        second='그 외 지역에서는 수출 감소'
        return first,second
    if rising:
        best=max(rising,key=lambda nm:regions[nm])
        return ('- 대륙별로 모든 지역에서 수출 증가,',
                f'{best}({fpct(regions[best])}%)의 증가율이 가장 높음')
    if falling:
        return ('- 대륙별로 모든 지역에서 수출 감소', '')
    return ('- 대륙별 수출 증감 없음', '')


def _set_safe_overview_lines(paragraph, lines):
    """Use a real HWPX soft-line break after a complete phrase.

    Only the existing overview paragraph is changed; its paragraph style,
    run styles, paragraph count, table contents and page-break flags remain.
    HWPX newlines must be hp:lineBreak, not a literal '\n' in hp:t.
    """
    first,second=lines
    if not first or _narrative_visual_width(first)>32.5:
        raise Reject('개황 대륙별 첫 문장이 너무 길어 자동 줄 배치를 중단합니다.')
    if second and _narrative_visual_width(second)>32.5:
        raise Reject('개황 대륙별 둘째 문장이 너무 길어 자동 줄 배치를 중단합니다.')
    set_text(paragraph,first + (' '+second if second else ''))
    if second:
        elem=paragraph.xpath('./hp:run/hp:t',namespaces=N)[0]
        elem.text=first
        line_break=etree.SubElement(elem,Q('lineBreak'))
        line_break.tail=' '+second
    # Ensure the generated textual words did not silently disappear.
    plain=get_text(paragraph)
    assert plain==first + (' '+second if second else ''),plain
    return plain

def fill_headline_paras(root0,root1,main,year,month,audit):
    """Rewrite *existing* paragraphs with verified facts or conspicuous author prompts.

    Neither paragraph count nor table dimensions change. Do not leave stale 2026
    narrative and do not manufacture economic causes from a percentage series.
    """
    paras=[]
    for root in (root0,root1):
        pp=[p for p in root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N)]
        paras.append({i:p for i,p in enumerate(pp)})
    mainsh=main.sheet('1_지역총괄')
    expo,impo=(number(mainsh.val(row,'E')) for row in (6,8))
    exp_rate,imp_rate=(number(mainsh.val(row,'E')) for row in (7,9))
    nat_rate=number(mainsh.val(7,'F'))
    trade=number(mainsh.val(10,'E'))
    mti=main.sheet('6_품목별(MTI4)')
    exports=get_mti4(mti,'수출'); imports=get_mti4(mti,'수입')
    by_name={v['name']:v for v in exports}
    # Explanatory commentary must describe the same dynamically selected MTI4 rows.
    top_exp=[x for _,entries in select_mti_sections(exports,16,'수출') for x in entries]
    top_imp=[x for _,entries in select_mti_sections(imports,17,'수입') for x in entries]

    def pct(n):return (fpct(n)+'%') if n is not None else None
    def change(n):
        return f'{pct(abs(n))} '+('증가' if n>=0 else '감소') if n is not None else None
    def author(label):return f'[작성자 입력: {label}]'
    def vplace(side,index,value):
        # If quantitative basis is absent, leave an explicit editorial marker,
        # not an empty paragraph or an inferred figure.
        target=paras[side].get(index)
        if target is None:raise Reject(f'템플릿 본문 위치가 달라졌습니다: {side}/{index}')
        if (side,index) in ((0,32),(1,2)) and isinstance(value,tuple):
            value=_set_safe_overview_lines(target,value)
            location='개황' if side==0 else '대륙별 수출동향'
            audit.note(f'{location}: 주요 수치 보존, 쉼표 구분 및 HWPX 실제 줄바꿈 1곳으로 어절 중간 잘림 방지 (기존 두 줄 배치)')
        else:
            value=value or f'ㅇ {author("통계 조회 및 서술 확인")}'
            set_text(target,value)
        if '[작성자 입력' in value:
            audit.issue(f'본문 {"수출/개황" if side==0 else "수출입 분석"} 문단 {index}: 작성자 검토 문구 포함')
        if len(value)>145:
            audit.note(f'본문 문단 {index}: 설명문 길이 {len(value)}자, 한컴 페이지/줄바꿈 확인 필요')

    export_head=(f'ㅇ {year}년 {month}월 부산지역 수출은 전년 동월 대비 {change(exp_rate)}한 {won_dollar_phrase(expo)} 기록' if expo is not None and exp_rate is not None else '')
    import_head=(f'ㅇ {year}년 {month}월 부산지역 수입은 전년 동월 대비 {change(imp_rate)}한 {won_dollar_phrase(impo)} 기록' if impo is not None and imp_rate is not None else '')
    trade_head=(f'ㅇ {year}년 {month}월 부산지역 무역수지는 {won_dollar_phrase(abs(trade))} {"흑자" if trade>0 else "적자" if trade<0 else "균형"} 기록' if trade is not None else '')

    def noteworthy(records,sign,n=2):
        found=[x for x in records if x['rate'] is not None and (x['rate']>0 if sign>0 else x['rate']<0) and x['current']>1000000]
        found.sort(key=lambda x: abs(x['current']-x['previous']),reverse=True)
        return found[:n]
    def item_tag(x):return f'{x["name"]}({x["rank"]}위, {pct(x["rate"])})'
    up,down=noteworthy(top_exp,1,2),noteworthy(top_exp,-1,1)
    summary=('·'.join(item_tag(x) for x in up)+(' 수출 증가' if up else ''))
    if down:summary+=(', ' if summary else '')+'·'.join(item_tag(x) for x in down)+' 수출 감소'
    vplace(0,28,export_head)
    vplace(0,30,('- '+summary if summary else '- '+author('주요 수출 품목 변동')))

    cont=main.sheet('9_대륙')
    regions={str(cont.val(r,'B') or '').strip():number(cont.val(r,'G')) for r in range(5,13)}
    if len(regions)==8 and all(v is not None for v in regions.values()):
        positive=[f'{nm}({pct(val)})' for nm,val in regions.items() if val>=0]
        negative=[f'{nm}({pct(val)})' for nm,val in regions.items() if val<0]
        # Original editorial style: concise group names, strongest verified rate.
        increasing=[nm for nm,val in regions.items() if val>=0]
        decreasing=[nm for nm,val in regions.items() if val<0]
        biggest=max(regions.items(),key=lambda item:abs(item[1]))
        geo='- 대륙별로 '+ '·'.join(increasing)+'에서 수출 증가'
        if decreasing:geo+=', '+ '·'.join(decreasing)+'에서는 감소'
        geo+=f'({biggest[0]} {change(biggest[1])})'
    else:geo='- '+author('대륙별 수출 증감률 확인')
    if len(regions)==8 and all(v is not None for v in regions.values()):
        # A maximum of two short editor-constructed clauses replaces a
        # glued 8-continent enumeration. Explicit break occurs only BETWEEN
        # phrases, not inside any individual Korean word such as 아프리카.
        overview_lines=_overview_continent_lines(regions)
        vplace(0,32,overview_lines)
    else:vplace(0,32,geo)
    if exp_rate is not None and nat_rate is not None:
        compared='높은' if exp_rate>nat_rate else '낮은' if exp_rate<nat_rate else '동일한'
        diff=abs(exp_rate-nat_rate)
        text=f'ㅇ 부산 수출 증감률은 {pct(exp_rate)}로 전국 평균({pct(nat_rate)})보다 {pct(diff)}p {compared} 수준' if compared!='동일한' else f'ㅇ 부산 수출 증감률은 전국 평균과 동일한 {pct(exp_rate)} 기록'
    else:text=''
    vplace(0,34,text)
    vplace(0,38,import_head)
    st=main.sheet('10_가공단계')
    stage=[(str(st.val(r,'B') or '').strip(),number(st.val(r,'G'))) for r in range(16,21)]
    # Import overview: match the older March/April/May/September finished reports.
    # Mention actual up/down MTI4 items, not the processing-stage percentages.
    imp_story = choose_import_commentary(imports)
    if imp_story['rises'] or imp_story['falls']:
        named_up=', '.join(x['name'] for x in imp_story['rises'][:2])
        named_down=', '.join(x['name'] for x in imp_story['falls'][:2])
        if named_up and named_down:
            overview_import=f'- {named_up} 등의 수입은 증가한 반면, {named_down} 등의 수입은 감소'
        elif named_up:
            overview_import=f'- {named_up} 등 주요 품목의 수입 증가'
        else:
            overview_import=f'- {named_down} 등 주요 품목의 수입 감소'
        vplace(0,40,overview_import)
        key_import=imp_story['highlight']
        if key_import and key_import['rate'] is not None:
            movement=change(key_import['rate'])
            vplace(0,42,f'ㅇ {key_import["name"]} 수입은 전년 동월 대비 {movement}하며 주요 수입품목 중 {key_import["rank"]}위 기록')
        else:
            vplace(0,42,'ㅇ '+author('주요 수입 품목의 증감 확인'))
    else:
        vplace(0,40,'- '+author('주요 수입품목의 증가·감소 확인'))
        vplace(0,42,'ㅇ '+author('주요 수입품목 변동 확인'))
    vplace(0,46,trade_head)

    # Section0: export trend, month-on-month, notable MTI4 product descriptions.
    vplace(0,55,export_head)
    mon=main.sheet('2_3년간월별수출입')
    months=[]
    for r in range(5,41):
        ym=str(mon.val(r,'B') or '')
        if re.fullmatch(r'20\d\d\.\d\d',ym):
            current=number(mon.val(r,'C')); yoy=number(mon.val(r,'D'))
            if current is not None and yoy is not None:months.append((ym,current,yoy))
    months=sorted(months,key=lambda x:x[0],reverse=True)
    if len(months)>=3:
        prior3=list(reversed(months[:3]))
        if all(r>0 for _,_,r in prior3):
            trend=f'- {int(prior3[0][0][5:])}~{int(prior3[-1][0][5:])}월 전년 동월 대비 수출 증가세 지속'
        elif all(r<0 for _,_,r in prior3):
            trend=f'- 최근 3개월 전년 동월 대비 수출 감소세 지속'
        else:
            trend='- 월별 수출 증감률의 등락이 이어지는 가운데 '+f'{month}월 전년 동월 대비 {change(prior3[-1][2])}'
        vplace(0,57,trend)
    else:vplace(0,57,'- '+author('최근 월별 수출 추세 확인'))
    if len(months)>=2 and months[0][1] and months[1][1]:
        mom=(months[0][1]/months[1][1]-1)*100
        vplace(0,59,f'ㅇ 전월 대비 {month}월 부산지역 수출액은 {change(mom)}')
    else:vplace(0,59,'ㅇ '+author('전월 대비 수출 증감률'))
    up=noteworthy(top_exp,1,3)
    down=noteworthy(top_exp,-1,2)
    inc_text='·'.join(x['name'] for x in up[:3])
    dec_text='·'.join(x['name'] for x in down[:2])
    if up and down:
        intro=f'ㅇ 품목별로는 {inc_text} 등의 수출이 증가한 반면, {dec_text} 수출은 감소'
    elif up:
        intro=f'ㅇ 품목별로는 {inc_text} 등 주요 품목을 중심으로 수출 증가'
    elif down:
        intro=f'ㅇ 품목별로는 {dec_text} 등의 수출이 감소'
    else:intro='ㅇ '+author('주요 품목별 수출 동향')
    vplace(0,67,intro)
    # Choose four material changes, with category diversity to mimic Mar-Apr-May.
    ranked=sorted(top_exp,key=lambda x:(abs(x['current']-x['previous']),abs(x['rate'] or 0)),reverse=True)
    featured=[];seen_groups=set()
    # Reflect major negative and positive shocks even if in the same MTI2 group.
    biggest_fall=next((x for x in ranked if x['rate'] is not None and x['rate']<0),None)
    biggest_rise=next((x for x in ranked if x['rate'] is not None and x['rate']>0),None)
    for rec in (biggest_fall,biggest_rise):
        if rec and rec not in featured:
            featured.append(rec);seen_groups.add(rec['group_code'])
    for rec in ranked:
        if len(featured)==4:break
        if rec not in featured and rec['group_code'] not in seen_groups:
            featured.append(rec);seen_groups.add(rec['group_code'])
    for rec in ranked:
        if len(featured)==4:break
        if rec not in featured:featured.append(rec)
    for pos,rec in zip((69,71,73,75),featured):
        clause=f'- [{rec["name"]}] 전년 동월 대비 {change(rec["rate"])}하며 당월 수출금액 {rec["rank"]}위 기록' if rec['rate'] is not None else f'- [{rec["name"]}] '+author('증감률 확인')
        vplace(0,pos,clause+' [작성자 입력 : 상세내용 확인 요망]')
    for pos in (69,71,73,75)[len(featured):]:vplace(0,pos,'- '+author('주요 수출 품목 분석'))

    # Section1: continents, top destinations, national region comparison, imports.
    if len(regions)==8 and all(v is not None for v in regions.values()):
        geo_lines=_overview_continent_lines(regions)
        vplace(1,2,('ㅇ '+geo_lines[0].lstrip('- '),geo_lines[1]))
    else:
        vplace(1,2,'ㅇ '+geo.lstrip('- '))
    destination=main.sheet('8_국가')
    destrows={str(destination.val(r,'B') or '').strip():r for r in range(5,25)}
    top_by_dest=main.sheet('13_국가별상위10대수출품목(MTI4)')
    for pos,country in [(4,'미국'),(6,'중국'),(8,'일본')]:
        r=destrows.get(country)
        rate=number(destination.val(r,'G')) if r is not None else None
        goods=[]
        for rr in range(5,60):
            if str(top_by_dest.val(rr,'C') or '').strip()!=country:continue
            amt=number(top_by_dest.val(rr,'H'))
            if amt is None:continue
            goods.append((amt,str(top_by_dest.val(rr,'D') or '').strip(),number(top_by_dest.val(rr,'I'))))
        goods.sort(reverse=True)
        if rate is not None and len(goods)>=3:
            ranks=[]
            for i,(amount,nm,pc) in enumerate(goods[:5],1):
                prior_row=next((rr for rr in range(5,60) if
                               str(top_by_dest.val(rr,'C') or '').strip()==country and
                               str(top_by_dest.val(rr,'D') or '').strip()==nm),None)
                previous=number(top_by_dest.val(prior_row,'E')) if prior_row is not None else None
                if previous in (None,0) and amount>0:
                    ranks.append((f'{nm}{i}위(전년 동월 실적 없음)',1))
                elif pc is not None:
                    ranks.append((f'{nm}{i}위({pct(pc)})',1 if pc>=0 else -1))
                else:
                    ranks.append((f'{nm}{i}위',0))
            inc=[label for label,sgn in ranks if sgn>0]
            dec=[label for label,sgn in ranks if sgn<0]
            # Preserve the original report's [국가] subject -> concrete
            # MTI4 ranking -> up/down contrast -> total YoY rate structure.
            if inc and dec:
                if rate>=0:
                    desc=(f'- [{country}] 對{country} 수출은 '+', '.join(dec[:2])+
                          ' 등의 감소에도 '+', '.join(inc[:4])+
                          f' 등을 중심으로 전년 동월 대비 {change(rate)}')
                else:
                    desc=(f'- [{country}] 對{country} 수출은 '+', '.join(inc[:3])+
                          ' 등의 증가에도 '+', '.join(dec[:3])+
                          f' 등의 감소로 전년 동월 대비 {change(rate)}')
            elif inc:
                desc=(f'- [{country}] 對{country} 수출은 '+', '.join(inc[:5])+
                      f' 등을 중심으로 전년 동월 대비 {change(rate)}')
            else:
                desc=(f'- [{country}] 對{country} 수출은 '+', '.join(dec[:5])+
                      f' 등의 영향으로 전년 동월 대비 {change(rate)}')
        else:desc='- '+author(f'對{country} 주요 수출품목 및 변동 원인')
        vplace(1,pos,desc)

    reg=main.sheet('4_전국지자체')
    rdata=[]
    for r in range(5,22):
        name=str(reg.val(r,'B') or '').strip()
        amount=number(reg.val(r,'F')); growth=number(reg.val(r,'G'))
        if name and amount is not None and growth is not None:rdata.append((name,amount,growth))
    bus=next((x for x in rdata if x[0]=='부산'),None)
    if bus and len(rdata)==17:
        rank=1+sum(x[1]>bus[1] for x in rdata)
        growthrank=1+sum(x[2]>bus[2] for x in rdata)
        rtop=sorted(rdata,key=lambda x:x[2],reverse=True)[:3]
        nat_amt=number(mainsh.val(6,'F'))
        share=bus[1]/nat_amt*100 if nat_amt else None
        vplace(1,18,f'ㅇ {year}년 {month}월 부산지역 수출액은 전국 17개 광역지자체 중 {rank}위 기록')
        vplace(1,20,f'ㅇ 수출 증감률({pct(bus[2])}) 기준 전국 17개 광역지자체 중 {growthrank}위 기록')
        vplace(1,22,'- '+', '.join(f'{REGIONS.get(nm,nm).replace("경기(남부)","경기 남부").replace("경기(북부)","경기 북부")}({pct(rate)})' for nm,_,rate in rtop)+f' 등에서 수출이 증가하며 전국 수출 증감률은 {pct(nat_rate)} 기록')
        # Finished reports end the regional section with the YTD rank, not
        # an abrupt one-line national-share observation.
        totals_ytd=[(str(reg.val(r,'B') or '').strip(),number(reg.val(r,'O'))) for r in range(5,22)]
        bus_ytd=next((amt for nm,amt in totals_ytd if nm=='부산'),None)
        if bus_ytd is not None and all(nm and amt is not None for nm,amt in totals_ytd):
            ytd_rank=1+sum(amt>bus_ytd for nm,amt in totals_ytd)
            vplace(1,24,f'ㅇ {year}년 1~{month}월 부산지역 수출액은 전국 광역지자체 중 {ytd_rank}위 기록')
        else:
            vplace(1,24,'ㅇ '+author('부산지역 누계 수출액 기준 전국 순위 확인'))
    else:
        for pos in (18,20,22,24):vplace(1,pos,'ㅇ '+author('전국 지자체 수출 순위 및 비중'))
    vplace(1,33,import_head)
    # Ⅲ 수입 동향: source reports follow the fixed sequence
    # 총액/그림2 → 가공단계별 증감/표5 → MTI4 증감/표6 → 대중국 수입/표7.
    # Use no arbitrary bold/red, no inferred market causes, no extra paragraphs.
    if len(stage)==5 and all(nm and r is not None for nm,r in stage):
        positive=[f'{nm}({pct(r)})' for nm,r in stage if r>=0]
        negative=[f'{nm}({pct(r)})' for nm,r in stage if r<0]
        if positive and negative:
            stage_sentence='ㅇ 가공단계별 수입은 '+', '.join(positive)+'는 증가하였으나, '+', '.join(negative)+' 수입은 감소함'
        elif positive:
            stage_sentence='ㅇ 가공단계별 수입은 '+', '.join(positive)+' 등 모든 부문에서 증가함'
        else:
            stage_sentence='ㅇ 가공단계별 수입은 '+', '.join(negative)+' 등 모든 부문에서 감소함'
        vplace(1,37,stage_sentence)
    else:
        vplace(1,37,'ㅇ '+author('가공단계별 수입 증감 확인'))
    vplace(1,43,make_import_mti_sentence(imports,author))
    china_row=next((r for r in range(31,51) if str(destination.val(r,'B') or '').strip()=='중국'),None)
    china_rate=number(destination.val(china_row,'G')) if china_row is not None else None
    vplace(1,51,f'ㅇ 對중국 수입은 전년 동월 대비 {change(china_rate)}' if china_rate is not None else 'ㅇ '+author('對중국 수입 증감률 확인'))
    audit.note('서술형 분석은 확인 가능한 통계 문장으로 작성하고, 외부 확인이 필요한 원인·시장 해석은 [작성자 입력: ...]으로 표시합니다.')

def rewrite_dates(root,year,month):
    prev_year=year-1
    next_yy,next_mm=(year+1,1) if month==12 else (year,month+1)
    for t in root.xpath('//hp:t',namespaces=N):
        s=t.text
        if not s:continue
        # Template is 2026-10. All occurrences of previous dates are rewritten,
        # including the header period and caption labels.
        s=s.replace('2025년 1∼10월',f'{prev_year}년 1∼{month}월')
        s=s.replace('2026년 1∼10월',f'{year}년 1∼{month}월')
        s=s.replace('2025년 10월',f'{prev_year}년 {month}월')
        s=s.replace('2026년 10월',f'{year}년 {month}월')
        s=s.replace('2026. 11.',f'{next_yy}. {next_mm}.')
        s=s.replace('10월',f'{month}월')
        s=s.replace('2026년 당월',f'{year}년 당월')
        t.text=s


def derive_period_total(sh,trade):
    row,inc=(6,7) if trade=='수출' else (8,9)
    return [item for c in ('C','E','G','K') for item in (fmoney(sh.val(row,c)),fpct(sh.val(inc,c)),'100.0')]

def main_country_totals(main):
    sh=main.sheet('8_국가');row=31
    # China import top row is usually #1. Avoid row-position assumption.
    for r in range(31,51):
        if str(sh.val(r,'B') or '').strip()=='중국':row=r;break
    else:return [None]*4
    return [number(sh.val(row,c)) for c in ('C','F','I','O')]


def build_tables(root0,root1,main,kstats,audit,styles):
    sh=main.sheet('1_지역총괄')
    sec0=root0.xpath('//hp:tbl',namespaces=N)
    sec1=root1.xpath('//hp:tbl',namespaces=N)
    if len(sec0)!=5 or len(sec1)!=6:raise Reject('기존 템플릿의 표 위치가 예상과 달라 중단했습니다.')
    # 표1: nationwide vs Busan, PY month / CY month / PY YTD / CY YTD.
    tb=sec0[2]; rr=tb.xpath('./hp:tr',namespaces=N)
    for report_row,src_row,valcol in [(2,6,True),(3,7,False),(4,8,True),(5,9,False),(6,10,True)]:
        cc=cells(rr[report_row]);values=[]
        for c_nat,c_bus in [('D','C'),('F','E'),('H','G'),('L','K')]:
            for c in (c_nat,c_bus):
                v=sh.val(src_row,c)
                values.append(fmoney(v) if valcol else fpct(v))
        for i,v in enumerate(values):set_text(cc[i+(1 if valcol else 0)],v if valcol else f'({v})' if v else '')
        audit.filled+=len(values)

    # 표2/표6: detail MTI4 from main; category MTI2 from K-Stat.
    for t,trade,slots in [(sec0[4],'수출',16),(sec1[4],'수입',17)]:
        rows=t.xpath('./hp:tr',namespaces=N)
        item_source=main.sheet('6_품목별(MTI4)')
        items=get_mti4(item_source,trade)
        # IMPORTANT: exact fixed row capacity, no stale names, no guessed MTI2
        # amount. Category placement is dynamic, not tied to last month's slots.
        ordered=select_mti_sections(items,slots,trade)
        chosen=[it for _,its in ordered for it in its]
        if sum(1+len(entries) for _,entries in ordered)!=slots:
            audit.issue(f'표{"2" if trade=="수출" else "6"}: 원본 상위20개로 고정 행 {slots}개를 전부 채울 수 없음 → 남는 칸 공란')
        stat_rec=kstats['mti2_month']['items']
        all_rows=[]
        btotal=[number(sh.val(6 if trade=='수출' else 8,c)) for c in ('C','E','G','K')]
        for group_code,entries in ordered:
            stats=stat_rec.get(group_code)
            nm=stats['name'] if stats else ''
            metrics=amount_metrics(None,group_code,'mti2',trade,kstats,btotal)
            all_rows.append((nm,metrics,None))
            if any(metrics[j]=='' for j in (0,3,6,9)):
                audit.issue(f'표{"2" if trade=="수출" else "6"}: MTI2 {group_code}의 일부 기간 통계 없음 → 공란')
            if not stats:audit.issue(f'표{"2" if trade=="수출" else "6"}: MTI2 {group_code} 항목명·통계 자료 없음')
            for item in sorted(entries,key=lambda x:x['rank']):
                all_rows.append((f'{item["name"]}({item["rank"]})',item['metrics'],None))
        totalrow=rows[2]
        set_row(totalrow,'총 수 출' if trade=='수출' else '합  계',derive_period_total(sh,trade),audit=audit)
        datarows=rows[3:]
        if len(datarows)!=slots:raise Reject('표2/표6의 행 수가 템플릿과 일치하지 않습니다')
        for i,row in enumerate(datarows):
            nm,metrics,_=all_rows[i] if i<len(all_rows) else ('', ['']*12,None)
            set_row(row,nm,metrics,audit=audit)
            # First-level MTI2 category rows are always pink; MTI4 white.
            is_group=(i<len(all_rows) and not bool(re.search(r'\(\d+\)$',nm)))
            set_row_band(row,is_group,styles)
            for c,val in zip(cells(row)[2:],metrics):
                set_cell_safe_spacing(c,val,styles)
        audit.note(f'표{"2" if trade=="수출" else "6"}: MTI4 당월 순위·변동액 고려 자동 후보; 최종 품목 구성 확인 권장')

    # 표3: 8 continents, all top-10 countries, dynamically grouped.
    src_country=main.sheet('8_국가');src_group=main.sheet('9_대륙')
    countries=[]
    for r in range(5,15):
        name=str(src_country.val(r,'B') or '').strip()
        if not name:continue
        group=GROUP_OF.get(name)
        if group is None:
            audit.issue(f'표3: 국가 {name}의 대륙 분류를 확인할 수 없어 공란')
            continue
        countries.append({'name':name,'rank':int(number(src_country.val(r,'A')) or r-4),'group':group,
                          'metrics':from_main(src_country,r)})
    all_continents=[]
    for group in CONTINENTS:
        rr=next((r for r in range(5,13) if str(src_group.val(r,'B') or '')==group),None)
        if rr is None:audit.issue(f'표3: {group} 합계 구역 누락')
        all_continents.append((group,from_main(src_group,rr) if rr else ['']*12,None))
        for c in sorted((x for x in countries if x['group']==group),key=lambda x:x['rank']):
            all_continents.append((c['name']+f'({c["rank"]})',c['metrics'],c['rank']))
    rows=sec1[0].xpath('./hp:tr',namespaces=N)
    set_row(rows[2],'총 수 출',derive_period_total(sh,'수출'),audit=audit)
    if len(all_continents)>len(rows)-3:raise Reject('표3 기존 행 수보다 국가가 많습니다; 자동 수정 중단')
    for i,row in enumerate(rows[3:]):
        name,metrics,rank=all_continents[i] if i<len(all_continents) else ('',['']*12,None)
        set_row(row,name,metrics,audit=audit)
    if len(countries)!=10:audit.issue(f'표3: 상위 10개국 중 대륙 확인 가능한 국가 {len(countries)}개')

    # 표4: national regions ranked by current-month export value.
    regions=main.sheet('4_전국지자체');rows=sec1[1].xpath('./hp:tr',namespaces=N)
    totals_from_main=[
      v for amount,rate in [(sh.val(6,'D'),sh.val(7,'D')), (sh.val(6,'F'),sh.val(7,'F')),
                            (sh.val(6,'H'),sh.val(7,'H')),(sh.val(6,'L'),sh.val(7,'L'))]
      for v in (fmoney(amount),fpct(rate),'100.0')]
    set_row(rows[2],'총계',totals_from_main,audit=audit)
    special_para=None
    for existing in rows[3:]:
        cc0=cells(existing)
        if len(cc0)>1 and ('남부' in get_text(cc0[1]) or '광주' in get_text(cc0[1])):
            paragraphs=cc0[1].xpath('./hp:subList/hp:p',namespaces=N)
            if len(paragraphs)>=2:
                special_para=[deepcopy(p) for p in paragraphs[:2]]
                break
    for i,row in enumerate(rows[3:]):
        sr=i+5
        nm=str(regions.val(sr,'B') or '').strip()
        fmt=REGIONS.get(nm,nm)
        set_row(row,fmt,from_main(regions,sr),rank=i+1,audit=audit)
        if '(' in fmt and ')' in fmt and special_para:
            cc0=cells(row);target=cc0[1];sub=target.find(Q('subList'))
            if sub is not None:
                for p in list(sub):
                    if p.tag==Q('p'):sub.remove(p)
                for p in special_para:sub.append(deepcopy(p))
                main_label,bracket=fmt.split('(',1)
                ps=sub.xpath('./hp:p',namespaces=N)
                set_text(ps[0],main_label);set_text(ps[1],'('+bracket)

    # 표5: choose import SECTION, never export section of same sheet.
    stages=main.sheet('10_가공단계');rows=sec1[3].xpath('./hp:tr',namespaces=N)
    set_row(rows[2],'총계',derive_period_total(sh,'수입'),audit=audit)
    stage_order=['중간재','소비재','자본재','1차산품','기타']
    stage_lookup={str(stages.val(r,'B') or ''):r for r in range(16,21)}
    for i,row in enumerate(rows[3:]):
        nm=stage_order[i]
        sr=stage_lookup.get(nm)
        set_row(row,nm,from_main(stages,sr) if sr else ['']*12,audit=audit)

    # 표7: China MTI3, current-month rank is fixed, YTD rows joined by code.
    rows=sec1[5].xpath('./hp:tr',namespaces=N)
    china_tot=main_country_totals(main)
    busan_tot=[number(sh.val(8,c)) for c in ('C','E','G','K')]
    country=main.sheet('8_국가')
    china_row=next((r for r in range(31,51) if str(country.val(r,'B'))=='중국'),None)
    if china_row:
        allc=from_main(country,china_row)
        # Table7 total shares are ratios of China trade to Busan import.
        set_row(rows[2],'총계',allc,audit=audit)
    else:set_row(rows[2],'총계',['']*12,audit=audit)
    current=kstats['china_month']['items']
    top=sorted([(code,st) for code,st in current.items() if st],key=lambda kv:-(kv[1].get('cy_import') or -1))[:5]
    for i,row in enumerate(rows[3:]):
        if i>=len(top):set_row(row,'',['']*12,rank=None,audit=audit);continue
        code,stats=top[i]
        metrics=amount_metrics(None,code,'china','수입',kstats,china_tot)
        if any(metrics[j]=='' for j in (0,3,6,9)):
            audit.issue(f'표7: MTI3 {code} 일부 기간의 중국 수입 자료 없음 → 공란')
        name=stats['name']
        # Small, readable two-line existing cell style for long ship structure label.
        if name.replace(' ','')=='선박해양구조물및부품':name='선박해양구조물 및 부품'
        set_row(row,name,metrics,rank=i+1,audit=audit)
        if name=='선박해양구조물 및 부품':
            label=cells(row)[1]
            pars=label.xpath('./hp:subList/hp:p',namespaces=N)
            if len(pars)>=2:
                set_text(pars[0],'선박해양구조물')
                set_text(pars[1],' 및 부품')
        if len(name)>11:audit.note(f'표7 {name}: 긴 품목명은 한글에서 두 줄/자간 검수 필요')


def replace_charts(parts,main,year,month,audit):
    """Draw readable 13-month bar/line graphs at the exact template image aspect.

    The first monthly data block is rows 5..40, in descending month order.
    Never read the second chart/calculation block (the old version did).
    """
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        from matplotlib.ticker import FuncFormatter
        from PIL import Image
        font_candidates=['Malgun Gothic','NanumGothic','Noto Sans CJK KR','NanumSquare','AppleGothic','UnDotum']
        families={f.name for f in font_manager.fontManager.ttflist}
        found=next((f for f in font_candidates if f in families),None)
        if found:plt.rcParams['font.family']=found
        else:plt.rcParams['font.family']='DejaVu Sans'
        plt.rcParams['axes.unicode_minus']=False
        sh=main.sheet('2_3년간월별수출입')
        points=[]
        for r in range(5,41):
            ym=str(sh.val(r,'B') or '').strip()
            if not re.fullmatch(r'20\d\d\.\d\d',ym):continue
            v=(number(sh.val(r,'C')),number(sh.val(r,'D')),number(sh.val(r,'E')),number(sh.val(r,'F')))
            if any(n is None for n in v):continue
            points.append((ym,*v))
        points=sorted(points,key=lambda x:x[0])[-13:]
        if len(points)!=13:raise Reject('월별 그래프: 연속 13개월 데이터가 없어 정확한 그래프 작성 불가')
        import datetime
        ym_values=[datetime.datetime.strptime(p[0],'%Y.%m') for p in points]
        if any(((b.year-a.year)*12+b.month-a.month)!=1 for a,b in zip(ym_values,ym_values[1:])):
            raise Reject('월별 그래프의 월이 연속되지 않습니다')
        for kind,imagekey,amount_ix,rate_ix in [('수출','BinData/image2.bmp',1,2),('수입','BinData/image3.bmp',3,4)]:
            values=[p[amount_ix]/1000000 for p in points]
            rates=[p[rate_ix] for p in points]
            labels=[p[0] for p in points]
            # Match original image size and aspect to avoid layout reflow in HWPX.
            original=Image.open(io.BytesIO(parts[imagekey]))
            width,height=original.size
            fig,ax=plt.subplots(figsize=(width/95,height/95),dpi=170)
            fig.patch.set_facecolor('white');ax.set_facecolor('white')
            bars=ax.bar(range(13),values,width=.59,color='#4776BD',zorder=2)
            vmax=max(values)
            ax.set_ylim(0,vmax*1.21)
            ax.grid(axis='y',color='#D9DEE6',linewidth=.8,zorder=0)
            ax.spines[['top','right','left']].set_visible(False)
            ax.spines['bottom'].set_color('#B6C0CB')
            ax.tick_params(axis='y',labelsize=8,length=0)
            ax.yaxis.set_major_formatter(FuncFormatter(lambda y,pos:f'{y:,.0f}'))
            ax.set_xticks(range(13));ax.set_xticklabels(labels,rotation=38,ha='right',fontsize=7.4)
            ax.tick_params(axis='x',length=0,pad=5)
            # Each blue bar has a value label, like the reference HWPX charts.
            for i,value in enumerate(values):
                ax.annotate(f'{value:,.0f}',(i,value),xytext=(0,3),textcoords='offset points',fontsize=7.4,ha='center',color='#202832')
            ax2=ax.twinx()
            lo,hi=min(rates),max(rates)
            spread=max(hi-lo,8)
            ax2.set_ylim(lo-spread*.26,hi+spread*.29)
            ax2.spines[['top','left','right']].set_visible(False)
            ax2.tick_params(axis='y',length=0,labelsize=8)
            ax2.plot(range(13),rates,color='#EE7D28',marker='s',markersize=3.6,linewidth=2,zorder=3)
            # Label selected points to reduce overlap, but always mark current and prior month.
            for i,r in enumerate(rates):
                if i in (0,3,6,9,11,12):
                    ax2.annotate(f'{r:.1f}',(i,r),textcoords='offset points',xytext=(0,-14 if i%2 else 6),
                                 ha='center',fontsize=7.1,color='#B24D12',
                                 bbox={'facecolor':'white','edgecolor':'none','alpha':.8,'pad':.6})
            from matplotlib.lines import Line2D
            from matplotlib.patches import Patch
            lab1=f'{kind}액(좌측, 백만 달러)' if found else f'{"Exports" if kind=="수출" else "Imports"} (left, USD mn)'
            lab2='전년동월비 증감률(우측)' if found else 'YoY growth (right, %)'
            fig.legend([Patch(color='#4776BD'),Line2D([0],[0],color='#EE7D28',marker='s',linewidth=2)],
                       [lab1,lab2],loc='upper center',ncol=2,frameon=False,fontsize=8.8,
                       bbox_to_anchor=(.5,.995))
            fig.subplots_adjust(top=.87,bottom=.22,left=.092,right=.91)
            # No duplicate chart title: document already prints its figure caption.
            buf=io.BytesIO();fig.savefig(buf,format='png',dpi=170,facecolor='white');plt.close(fig)
            img=Image.open(io.BytesIO(buf.getvalue())).convert('RGB')
            img=img.resize((width*2,height*2),Image.Resampling.LANCZOS)
            bmp=io.BytesIO();img.save(bmp,format='BMP');parts[imagekey]=bmp.getvalue()
            audit.note(f'{imagekey}: 연속 13개월의 {kind} 금액·전년 동월 대비율 차트 생성 (원본 이미지 비율 유지, 라벨 표시, 한글 지원 시 한글 폰트 사용)')
    except (ImportError,Reject) as e:
        # Never retain a graph from the prior reporting month if chart building fails.
        raise Reject(f'월별 그래프를 검증된 자료로 작성할 수 없습니다: {e}')

def validate_main_schema(main):
    # Fail closed if the association changes a table layout/column meaning.
    expected=[
        ('1_지역총괄',4,'E','금년당월'),
        ('1_지역총괄',5,'C','부산지역'),
        ('1_지역총괄',6,'A','수출'),
        ('1_지역총괄',8,'A','수입'),
        ('6_품목별(MTI4)',3,'B','품목코드'),
        ('6_품목별(MTI4)',4,'D','수출금액'),
        ('6_품목별(MTI4)',30,'D','수입금액'),
        ('8_국가',4,'C','수출금액'),
        ('8_국가',30,'C','수입금액'),
        ('9_대륙',4,'C','수출금액'),
        ('10_가공단계',15,'C','수입금액'),
    ]
    for title,row,col,kw in expected:
        got=str(main.sheet(title).val(row,col) or '')
        if kw not in got:
            raise Reject(f'엑셀 구조 확인 실패: {title}!{col}{row}은(는) {kw}여야 하는데 {got!r}입니다. 임의로 다른 열을 읽지 않습니다.')
    sh=main.sheet('1_지역총괄')
    import_stage=main.sheet('10_가공단계')
    e0=number(sh.val(8,'E'))
    e1=number(import_stage.val(21,'F'))
    if e0 is not None and e1 is not None and abs(e0-e1)>1:
        raise Reject('가공단계 수입 합계와 지역총괄 수입 합계가 다릅니다. 서식/시트 변경 확인이 필요합니다.')

def build_report(main_data,main_name,uploads,template_data):
    audit=Audit()
    main=load_excel(main_data,main_name)
    validate_main_schema(main)
    year,month=make_month_info(main)
    if not (2018<=year<=2040):raise Reject('허용되지 않은 보고서 연도')
    # Strict source authentication for China import: compare K-Stat export
    # amounts against the China row in the main report, per matched period.
    # Metadata alone is insufficient (incorrect K-Stat exports can retain the
    # old '국가 : 중국' header after a query is changed).
    china_expected_totals=main_country_totals(main)
    ks=concat_kstats(uploads,year,month,audit,china_expected_totals)
    totals=main.sheet('1_지역총괄')
    for key,field,valcol in [
        ('mti2_month','cy_export','E'),('mti2_current','cy_export','K'),('mti2_prior','cy_export','G'),
        ('china_month','cy_import',None),('china_current','cy_import',None),('china_prior','cy_import',None)
    ]:
        t=ks[key]['total'];got=t.get(field) if t else None
        expected=number(totals.val(6,valcol)) if valcol else None
        if got is not None and expected is not None and abs(got-expected)>1:
            raise Reject(f'{key}: K-Stat 총계({got:,.0f})와 기본 엑셀({expected:,.0f}) 불일치. 잘못된 자료를 입력하지 않도록 생성 중단.')
    # Both import and export totals are checked in all three MTI2 periods.
    for key,py_col,cy_col in [('mti2_month','C','E'),('mti2_prior',None,'G'),('mti2_current',None,'K')]:
        tot=ks[key]['total']
        if tot is None:continue
        for trade,main_row in [('export',6),('import',8)]:
            for side,col in [('py',py_col),('cy',cy_col)]:
                if col is None:continue
                got=tot.get(f'{side}_{trade}')
                want=number(totals.val(main_row,col))
                if got is not None and want is not None and abs(got-want)>1:
                    raise Reject(f'{key}: {trade}/{side} 총계가 기본 엑셀과 달라 생성을 중단했습니다. ({got:,.0f} ≠ {want:,.0f})')
    # Crosscheck China totals against country section.
    ch=main_country_totals(main)
    for key,idx in [('china_month',1),('china_prior',2),('china_current',3)]:
        t=ks[key]['total'];got=t.get('cy_import') if t else None
        if got is not None and ch[idx] is not None and abs(got-ch[idx])>1:
            # Allow supplemental focused query with smaller 'total' but forbid mixing it into
            # top-ranked current dataset if it has inconsistent country total.
            if key=='china_prior' and len(ks[key]['sources'])>1:
                audit.issue(f'{key}: 일부 필터 조회파일의 소계가 전체 중국 수입액과 다릅니다. 품목별 코드로만 통합했고 총계는 기본 엑셀 기준입니다.')
            else:raise Reject(f'{key}: 중국 수입 총계가 기본 엑셀과 일치하지 않아 생성 중단')
    template=io.BytesIO(template_data)
    with zipfile.ZipFile(template) as z:
        if z.testzip():raise Reject('한글 원본 템플릿이 손상되어 있습니다')
        parts={name:z.read(name) for name in z.namelist()}
        infos=z.infolist()
    roots={name:etree.fromstring(parts[name]) for name in ['Contents/section0.xml','Contents/section1.xml']}
    before=[table_signature(roots[n]) for n in roots]
    header=etree.fromstring(parts['Contents/header.xml'])
    styles=create_style_context(header)
    build_tables(roots['Contents/section0.xml'],roots['Contents/section1.xml'],main,ks,audit,styles)
    fit_table6_long_product_labels(header,roots['Contents/section1.xml'],audit)
    fill_headline_paras(roots['Contents/section0.xml'],roots['Contents/section1.xml'],main,year,month,audit)
    normalise_generated_body_style(roots['Contents/section0.xml'],(28,30,32,34,38,40,42,46,55,57,59,67,69,71,73,75))
    normalise_generated_body_style(roots['Contents/section1.xml'],(2,4,6,8,18,20,22,24,33,37,43,51))
    use_word_safe_left_align(header,roots['Contents/section0.xml'],(32,))
    use_word_safe_left_align(header,roots['Contents/section1.xml'],(2,))
    layout_compact_pages(header,roots['Contents/section0.xml'],roots['Contents/section1.xml'],audit)
    format_import_rank_superscript(header,roots['Contents/section1.xml'],audit)
    format_destination_rank_superscript(header,roots['Contents/section1.xml'],audit)
    for root in roots.values():rewrite_dates(root,year,month)
    if [table_signature(roots[n]) for n in roots]!=before:
        raise Reject('표 구조 변경 감지: 출력하지 않았습니다.')
    for name,root in roots.items():parts[name]=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
    parts['Contents/header.xml']=etree.tostring(header,encoding='UTF-8',xml_declaration=True,standalone=True)
    replace_charts(parts,main,year,month,audit)
    if 'Preview/PrvText.txt' in parts:
        preview=parts['Preview/PrvText.txt'].decode('utf-8',errors='replace')
        next_y,next_m=(year+1,1) if month==12 else (year,month+1)
        preview=preview.replace('2025년 10월',f'{year-1}년 {month}월')
        preview=preview.replace('2026년 10월',f'{year}년 {month}월')
        preview=preview.replace('2026. 11.',f'{next_y}. {next_m}.')
        parts['Preview/PrvText.txt']=preview.encode('utf-8')
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w') as z:
        for info in infos:z.writestr(info,parts[info.filename])
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as z:
        if z.testzip():raise Reject('출력 HWPX ZIP 무결성 검사 실패')
        for n in ('Contents/section0.xml','Contents/section1.xml','Contents/header.xml'):
            etree.fromstring(z.read(n))
        if z.getinfo('mimetype').compress_type!=zipfile.ZIP_STORED:raise Reject('HWPX mimetype 저장 형식 오류')
    audit.note(f'보고기간: {year}년 {month}월 / 표의 행·열·셀 수 불변 / HWPX XML·ZIP 검사 통과')
    name=f'{year}년_{month}월_부산지역_수출입동향_자동작성_검토본.hwpx'
    report='부산지역 수출입동향 자동작성 검증 및 확인 목록\n'+'='*46+'\n\n'
    report+='[검증 내용]\n'+'\n'.join('· '+x for x in audit.notes)+'\n\n'
    report+='[확인 필요/미작성 항목]\n'+ ('\n'.join('· '+x for x in audit.issues) if audit.issues else '· 추가 데이터 파일 검증 경고 없음')+'\n\n'
    report+='[작성 원칙]\n· 출처가 불명확한 수치는 공란으로 처리\n· 전년 당월/금년 당월/전년 1~당월 누계/금년 1~당월 누계 엄격히 구분\n· 수출과 수입 별도 검증\n· 기존 HWPX 표 구조 유지\n· 페이지 수와 가독성은 한컴오피스에서 최종 확인 필요\n'
    return buf.getvalue(),name,report,audit

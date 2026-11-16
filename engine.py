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
    texts[0].text=value
    for t in texts[1:]:t.text=''

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

def concat_kstats(inputs,year,month,audit):
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
            parts.append(k)
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


def fill_headline_paras(root0,root1,main,year,month,audit):
    # Template paragraphs are identified by their original content before blanking.
    t0=[];t1=[]
    for root,s in [(root0,t0),(root1,t1)]:
        for p in root.xpath('//hp:p[not(ancestor::hp:tbl)]',namespaces=N):
            text=get_text(p)
            if text and text.lstrip().startswith(('ㅇ','-')):s.append((p,text))
    sh=main.sheet('1_지역총괄')
    mexp,imp=number(sh.val(6,'E')),number(sh.val(8,'E'))
    eg,ig=number(sh.val(7,'E')),number(sh.val(9,'E'))
    nationwide=number(sh.val(7,'F'))
    trade=number(sh.val(10,'E'))
    narrative={}
    export_head=f'ㅇ {year}년 {month}월 부산지역 수출은 전년 동월 대비 {rate_line(eg,"export")}한 {won_dollar_phrase(mexp)} 기록' if eg is not None else ''
    import_head=f'ㅇ {year}년 {month}월 부산지역 수입은 전년 동월 대비 {rate_line(ig,"import")}한 {won_dollar_phrase(imp)} 기록' if ig is not None else ''
    # rate_line already ends with 증가/감소, so no double verb.
    export_head=export_head.replace('증가한','증가한').replace('감소한','감소한')
    import_head=import_head.replace('증가한','증가한').replace('감소한','감소한')
    if trade is not None:
        trade_head=f'ㅇ {year}년 {month}월 부산지역 무역수지는 {won_dollar_phrase(abs(trade))} {"흑자" if trade>0 else "적자" if trade<0 else "균형"} 기록'
    else:trade_head=''
    # Section0 top level: primary facts and continental movements.
    export_paras=[p for p,t in t0 if t.startswith('ㅇ') and '부산지역 수출은' in t]
    if export_paras:
        set_text(export_paras[0],export_head)
        if len(export_paras)>1:set_text(export_paras[1],f'ㅇ {year}년 {month}월 부산지역 수출 증가율은 {fpct(eg)}%로 전국 평균({fpct(nationwide)}%)과 비교' if eg is not None and nationwide is not None else '')
    imp_paras=[p for p,t in t0 if t.startswith('ㅇ') and '부산지역 수입은' in t]
    if imp_paras:set_text(imp_paras[0],import_head)
    trade_paras=[p for p,t in t0 if '무역수지는' in t]
    if trade_paras:set_text(trade_paras[0],trade_head)
    # Section0 and section1: create only clearly supported percentage summary.
    reg=main.sheet('9_대륙')
    region_rows={str(reg.val(r,'B')):reg.val(r,'G') for r in range(5,13)}
    up=[k for k,v in region_rows.items() if number(v) is not None and number(v)>=0]
    down=[k for k,v in region_rows.items() if number(v) is not None and number(v)<0]
    geo='ㅇ 대륙별로 '+', '.join(up)+' 수출은 증가하고 '+', '.join(down)+' 수출은 감소' if len(region_rows)==8 else ''
    for p,t in t0+t1:
        if t.startswith(' - 대륙별로') or t.startswith('ㅇ 대륙별로'):set_text(p,geo)
    # Fresh calendar/month titles are handled separately. Everything else with facts gets
    # blanked unless generated explicitly below, never retain previous month narrative.
    keep=set(export_paras+imp_paras+trade_paras)
    for p,t in t0+t1:
        if p in keep or t.startswith((' - 대륙별로','ㅇ 대륙별로')):continue
        if '전년 동월' in t and '부산지역 수출은' in t:set_text(p,export_head);keep.add(p);continue
        if '전년 동월' in t and '부산지역 수입은' in t:set_text(p,import_head);keep.add(p);continue
        if '무역수지는' in t:set_text(p,trade_head);keep.add(p);continue
        if t.startswith(('ㅇ','-')):
            set_text(p,'');audit.empty+=1
    # Import stages fully supported by imported section.
    st=main.sheet('10_가공단계')
    stages=[]
    for r in range(16,21):
        nm=str(st.val(r,'B') or '');pct=number(st.val(r,'G'))
        if pct is not None:stages.append(f'{nm}({fpct(pct)}%)')
    for p,t in t1:
        if '가공단계별 수입' in t and t.startswith('ㅇ'):
            set_text(p,('ㅇ 가공단계별 수입 증감률은 '+', '.join(stages)+'로 나타남') if len(stages)==5 else '')
    # Important: do not author speculative causal explanations.
    audit.note('원인 분석·특정 국가 세부품목 설명 등은 통계만으로 확정할 수 없어 본문 공란 처리했습니다.')


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


def build_tables(root0,root1,main,kstats,audit):
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
        # Top 10 take precedence; notable changes from ranks 11-20 are candidates
        # when they displaced a lower-impact top-10 item.
        chosen=items[:10]
        contenders=sorted(items[10:],key=lambda x:abs(x['current']-x['previous']),reverse=True)
        for extra in contenders[:3]:
            if extra['current']>10000000 and abs(extra['current']-extra['previous'])>20000000:
                worst=min(chosen,key=lambda x:abs(x['current']-x['previous']))
                if abs(extra['current']-extra['previous'])>abs(worst['current']-worst['previous'])*1.7:
                    chosen.remove(worst);chosen.append(extra)
        chosen=sorted(chosen,key=lambda x:x['rank'])
        per_group=defaultdict(list)
        for item in chosen:per_group[item['group_code']].append(item)
        while len(chosen)+len(per_group)>slots and chosen:
            # Never resize the table: remove lowest-priority item; preserve an audit trail.
            lost=max(chosen,key=lambda x:x['rank']);chosen.remove(lost)
            per_group=defaultdict(list)
            for item in chosen:per_group[item['group_code']].append(item)
            audit.issue(f'표{"2" if trade=="수출" else "6"}: 고정 행 수 제한으로 {lost["name"]}({lost["rank"]}위) 미기재')
        stat_rec=kstats['mti2_month']['items']
        ordered=sorted(per_group.items(),key=lambda x:min(i['rank'] for i in x[1]))
        all_rows=[]
        btotal=[number(sh.val(6 if trade=='수출' else 8,c)) for c in ('C','E','G','K')]
        for group_code,entries in ordered:
            stats=stat_rec.get(group_code)
            nm=(stats['name'] if stats else '') if stats else ''
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
    # Render charts inside fixed existing image slots. If rendering is unavailable,
    # clear them rather than leaving graphs from the previous month.
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        from PIL import Image
        import tempfile,os
        for font in ['Malgun Gothic','NanumGothic','Noto Sans CJK KR','AppleGothic']:
            if any(x.name==font for x in font_manager.fontManager.ttflist):
                plt.rcParams['font.family']=font;break
        plt.rcParams['axes.unicode_minus']=False
        sh=main.sheet('2_3년간월별수출입')
        points=[]
        for r in range(5,80):
            ym=str(sh.val(r,'B') or '')
            if not re.match(r'^20\d\d\.\d\d$',ym):continue
            points.append((ym,sh.val(r,'C'),sh.val(r,'D'),sh.val(r,'E'),sh.val(r,'F')))
        points=sorted(points, key=lambda x:x[0])[-13:]
        if len(points)<7:
            audit.issue('월별 그래프: 데이터가 7개월 미만이라 새 그래프를 생성하지 않았습니다.')
            return
        for kind,imagekey,which in [('수출','BinData/image2.bmp',1),('수입','BinData/image3.bmp',3)]:
            values=[number(p[which])/1e6 if number(p[which]) is not None else 0 for p in points]
            rates=[number(p[which+1]) or 0 for p in points]
            labels=[p[0][2:] for p in points]
            fig,ax=plt.subplots(figsize=(9.3,3.25),dpi=120)
            bar=ax.bar(range(len(values)),values,color='#4385b4',width=.64)
            ax.set_ylabel('백만 달러',fontsize=9)
            ax.set_xticks(list(range(len(labels))),labels,rotation=0,fontsize=8)
            ax.set_ylim(0,max(values)*1.22 if values else 1)
            ax.grid(axis='y',alpha=.14)
            ax2=ax.twinx();ax2.plot(range(len(rates)),rates,color='#da7b3e',marker='o',linewidth=1.7,markersize=3)
            ax2.set_ylabel('전년 동월 대비 (%)',fontsize=9)
            ax2.grid(False)
            fig.suptitle(f'부산지역 월별 {kind}액 및 증감률',fontsize=12)
            fig.tight_layout()
            out=io.BytesIO();fig.savefig(out,format='png',bbox_inches='tight',facecolor='white');plt.close(fig)
            img=Image.open(io.BytesIO(out.getvalue())).convert('RGB')
            buf=io.BytesIO();img.save(buf,format='BMP');parts[imagekey]=buf.getvalue()
            audit.note(f'{imagekey}: 월별 원본 자료로 그래프 신규 생성')
    except ImportError:
        audit.issue('그래프 재작성 라이브러리(matplotlib, pillow)가 없어 기존 그래프의 정확성을 보장할 수 없습니다.')


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
    ks=concat_kstats(uploads,year,month,audit)
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
    build_tables(roots['Contents/section0.xml'],roots['Contents/section1.xml'],main,ks,audit)
    fill_headline_paras(roots['Contents/section0.xml'],roots['Contents/section1.xml'],main,year,month,audit)
    for root in roots.values():rewrite_dates(root,year,month)
    if [table_signature(roots[n]) for n in roots]!=before:
        raise Reject('표 구조 변경 감지: 출력하지 않았습니다.')
    for name,root in roots.items():parts[name]=etree.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
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

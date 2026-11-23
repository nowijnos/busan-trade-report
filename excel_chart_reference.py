"""Excel chart conversion rules replicated from user's original workbook.

Source sheet '2_그래프 - 수출' / '3_수입 그래프':
  - A2:C14, newest to oldest: YYYY.MM, raw US dollars, YoY percent
  - A18:C30, oldest to newest: A reverse; B = raw amount/1e6; C = raw percent
  - chart: columns for B on left axis, line for C on right axis
The original workbook is NOT bundled: avoids leaking its extra tabs to GitHub.
"""
from __future__ import annotations
import datetime as dt
import re

SHEET_NAMES={'수출':'2_그래프 - 수출','수입':'3_수입 그래프'}
SHEET_COLUMNS={'수출':('C','D'),'수입':('E','F')}

def month_at(year:int,month:int,minus:int=0)->str:
    index=year*12+month-1-minus
    y,m=divmod(index,12)
    return f'{y:04d}.{m+1:02d}'

def excel_chart_rows(main,year:int,month:int,parse_number):
    """Build exactly the same 13 rows Excel would calculate in A18:C30.

    The main report file provides all data; the standalone user's chart workbook
    supplies only layout/unit/axis rules. No stale cached xlsx chart values are used.
    """
    year=int(year);month=int(month)
    if not (2000<=year<=2100 and 1<=month<=12):
        raise ValueError('그래프 기준 연월이 올바르지 않습니다')
    sheet=main.sheet('2_3년간월별수출입')
    by_month={}
    for row in range(5,41):
        ym=str(sheet.val(row,'B') or '').strip()
        if not re.fullmatch(r'20\d{2}\.\d{2}',ym):
            continue
        try:
            dt.datetime.strptime(ym,'%Y.%m')
        except ValueError:
            continue
        # Only the approved FIRST 3-year monthly block, never second chart block.
        current=(sheet.val(row,'C'),sheet.val(row,'D'),sheet.val(row,'E'),sheet.val(row,'F'))
        if ym in by_month:
            raise ValueError(f'월별 그래프 원본에 중복 연월 발견: {ym}')
        by_month[ym]=current
    expected=[month_at(year,month,n) for n in range(13)] # A2:A14 newest first
    missing=[ym for ym in expected if ym not in by_month]
    if missing:
        raise ValueError('월별 그래프 입력자료에 필요한 13개월이 없습니다: '+', '.join(missing))
    output={}
    for kind,(money_col,rate_col) in SHEET_COLUMNS.items():
        money_index=('C','D','E','F').index(money_col)
        rate_index=('C','D','E','F').index(rate_col)
        input_desc=[]
        for ym in expected:
            line=by_month[ym]
            amount=parse_number(line[money_index]);rate=parse_number(line[rate_index])
            if amount is None or rate is None or amount<0:
                raise ValueError(f'{kind} 그래프 {ym} 원본 금액·증감률 누락 또는 오류')
            input_desc.append((ym,amount,rate))
        # The Excel original uses A18=A14, B18=B14/1000000, C18=C14;
        # ... A30=A2, B30=B2/1000000, C30=C2.
        computed_asc=[(ym,amount/1_000_000,rate) for ym,amount,rate in reversed(input_desc)]
        output[kind]=computed_asc
    return output

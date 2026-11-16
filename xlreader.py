"""Small read-only readers for official monthly XLSX and K-Stat BIFF8 XLS.
Never recalculates or changes source spreadsheets.
"""
from __future__ import annotations
import io, re, zipfile, posixpath
from pathlib import Path
from lxml import etree

SHEET_NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
N={'m':SHEET_NS}

class Book:
    def __init__(self, worksheets):
        self.sheets=worksheets
    def sheet(self, name):
        if name not in self.sheets:raise ValueError(f'필수 시트가 없습니다: {name}')
        return self.sheets[name]

class Sheet:
    def __init__(self, rows):
        self.rows=rows
    def val(self, row, col):
        if isinstance(col,str):
            x=0
            for ch in col.upper():x=x*26+ord(ch)-64
            col=x
        return self.rows.get(int(row),{}).get(int(col))
    def iter(self, start=1, end=None):
        for r, vals in self.rows.items():
            if r >= start and (end is None or r<=end):yield r,vals
    def text(self):
        return ' '.join(str(v) for row in self.rows.values() for v in row.values() if v is not None)

def load_xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            sst=etree.fromstring(z.read('xl/sharedStrings.xml'))
            ss=[''.join(x.xpath('.//m:t/text()',namespaces=N)) for x in sst]
        wb=etree.fromstring(z.read('xl/workbook.xml'))
        rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap={x.get('Id'):x.get('Target') for x in rels}
        sheets={}
        for sch in wb.xpath('//m:sheets/m:sheet',namespaces=N):
            name=sch.get('name'); key=sch.get('{%s}id'%REL_NS)
            path=relmap[key].lstrip('/')
            if not path.startswith('xl/'):path=posixpath.normpath('xl/'+path)
            root=etree.fromstring(z.read(path))
            rows={}
            for row in root.xpath('//m:sheetData/m:row',namespaces=N):
                cells={}
                for cell in row.xpath('./m:c',namespaces=N):
                    m=re.match('([A-Z]+)',cell.get('r',''))
                    if not m:continue
                    col=0
                    for ch in m.group(1):col=col*26+(ord(ch)-64)
                    v=cell.find('{%s}v'%SHEET_NS)
                    ins=cell.find('{%s}is'%SHEET_NS)
                    if v is not None and v.text is not None:
                        txt=v.text
                        if cell.get('t')=='s':
                            value=ss[int(txt)]
                        elif cell.get('t') in ('str','e','inlineStr'):value=txt
                        else:
                            try:value=float(txt) if '.' in txt or 'E' in txt.upper() else int(txt)
                            except ValueError:value=txt
                    elif ins is not None:
                        value=''.join(ins.xpath('.//m:t/text()',namespaces=N))
                    else:continue
                    cells[col]=value
                rows[int(row.get('r'))]=cells
            sheets[name]=Sheet(rows)
        return Book(sheets)

def load_xls(data):
    try:
        import xlrd
    except ImportError as ex:
        raise ValueError('K-Stat .xls 읽기에 필요한 xlrd가 없습니다. start.bat을 다시 실행해 설치하거나, 엑셀에서 .xlsx로 저장하여 업로드해 주세요.') from ex
    wb=xlrd.open_workbook(file_contents=data)
    ret={}
    for ws in wb.sheets():
        rows={}
        for r in range(ws.nrows):
            vals={}
            for c in range(ws.ncols):
                obj=ws.cell(r,c)
                if obj.ctype==xlrd.XL_CELL_EMPTY or obj.ctype==xlrd.XL_CELL_BLANK:continue
                value=obj.value
                if obj.ctype==xlrd.XL_CELL_NUMBER and float(value).is_integer():value=int(value)
                vals[c+1]=value
            rows[r+1]=vals
        ret[ws.name]=Sheet(rows)
    return Book(ret)

def load_excel(data, filename):
    if filename.lower().endswith('.xlsx'):return load_xlsx(data)
    if filename.lower().endswith('.xls'):return load_xls(data)
    raise ValueError(f'지원하지 않는 확장자: {filename}. .xls 또는 .xlsx 파일만 지원합니다.')

def number(value):
    if value is None or value=='':return None
    if isinstance(value,(int,float)) and not isinstance(value,bool):return float(value)
    try:return float(str(value).replace(',','').replace('%','').strip())
    except ValueError:return None

# Strict fallback for the small, regular K-Stat BIFF8 files when xlrd is unavailable.
# It is read-only and deliberately rejects unrelated/unsupported binary workbooks.
def load_kstat_biff8(data):
    import struct
    if data[:8]!=bytes.fromhex('d0cf11e0a1b11ae1'):
        raise ValueError('지원되지 않는 .xls 형식입니다. 엑셀에서 .xlsx로 저장해 주세요.')
    u16=lambda b,p:struct.unpack_from('<H',b,p)[0]
    u32=lambda b,p:struct.unpack_from('<I',b,p)[0]
    sec=1<<u16(data,30)
    if sec not in (512,4096):raise ValueError('알 수 없는 OLE 섹터 크기')
    sector=lambda x:data[(x+1)*sec:(x+2)*sec]
    fats=[u32(data,76+4*i) for i in range(109) if u32(data,76+4*i)<0xFFFFFFFA]
    fat=b''.join(sector(i) for i in fats)
    step=lambda x:u32(fat,x*4)
    def chain(start,size=None):
        out=bytearray();seen=set();x=start
        while x<0xFFFFFFFA and x not in seen and len(seen)<5000:
            seen.add(x);out.extend(sector(x));x=step(x)
        return bytes(out[:size] if size is not None else out)
    dire=chain(u32(data,48))
    entries=[]
    for i in range(0,len(dire),128):
        b=dire[i:i+128]
        if len(b)<128:continue
        n=u16(b,64)
        if n<4:continue
        entries.append((b[:n-2].decode('utf-16le','ignore'),b[66],u32(b,116),struct.unpack_from('<Q',b,120)[0]))
    work=next((x for x in entries if x[0] in ('Workbook','Book')),None)
    if not work:raise ValueError('워크북 데이터가 없는 .xls 파일입니다')
    if work[3]<u32(data,56):raise ValueError('xls 소형 스트림은 xlrd 설치가 필요합니다.')
    stream=chain(work[2],work[3]);records=[];p=0
    while p+4<=len(stream):
        typ=u16(stream,p);size=u16(stream,p+2)
        if p+4+size>len(stream):break
        records.append((typ,stream[p+4:p+4+size],p));p+=4+size
    sst=[];bounds=[]
    for typ,r,pos in records:
        if typ==0x85:
            slen=r[6];flags=r[7];txt=r[8:8+slen*(2 if flags&1 else 1)]
            bounds.append((u32(r,0),txt.decode('utf-16le' if flags&1 else 'latin1','ignore')))
        elif typ==0xFC:
            j=8
            for _ in range(u32(r,4)):
                if j+3>len(r):break
                n=u16(r,j);flag=r[j+2];j+=3
                rich= u16(r,j) if flag&8 else 0
                if flag&8:j+=2
                ext=u32(r,j) if flag&4 else 0
                if flag&4:j+=4
                count=n*(2 if flag&1 else 1)
                sst.append(r[j:j+count].decode('utf-16le' if flag&1 else 'latin1','replace'))
                j+=count+rich*4+ext
    if not bounds:raise ValueError('시트 정보 없는 .xls')
    sheets={name:{} for _,name in bounds}
    bounds=sorted(bounds);current=None
    for typ,r,pos in records:
        for p0,name in bounds:
            if pos>=p0:current=name
        if current is None or len(r)<8:continue
        if typ not in (0xFD,0x203,0x27E,0xBD):continue
        row=u16(r,0)+1;col=u16(r,2)+1
        if typ==0xFD:
            idx=u32(r,6);value=sst[idx] if idx<len(sst) else ''
            sheets[current].setdefault(row,{})[col]=value
        elif typ==0x203:
            value=struct.unpack_from('<d',r,6)[0]
            sheets[current].setdefault(row,{})[col]=value
        elif typ==0x27E:
            v=u32(r,6);raw=v>>2
            if v&2:
                if raw&0x20000000:raw-=0x40000000
                num=float(raw)
            else:num=struct.unpack('<d',struct.pack('<Q',v&0xFFFFFFFC<<32))[0]
            if v&1:num/=100
            sheets[current].setdefault(row,{})[col]=num
        elif typ==0xBD:
            for c in range(col,u16(r,len(r)-2)+2):
                off=4+6*(c-col)
                if off+6>len(r):break
                v=u32(r,off+2);raw=v>>2
                if v&2:
                    if raw&0x20000000:raw-=0x40000000
                    num=float(raw)
                else:num=struct.unpack('<d',struct.pack('<Q',(v&0xFFFFFFFC)<<32))[0]
                if v&1:num/=100
                sheets[current].setdefault(row,{})[c]=num
    book=Book({k:Sheet(v) for k,v in sheets.items()})
    meta=next(iter(book.sheets.values())).val(2,'A')
    if 'K-stat' not in str(next(iter(book.sheets.values())).val(1,'A')):
        raise ValueError('이 파일은 지원 대상 K-Stat 내보내기 .xls가 아닙니다.')
    return book

_original_load_xls=load_xls
def load_xls(data):
    try:
        return _original_load_xls(data)
    except ValueError as exc:
        if 'xlrd' in str(exc):return load_kstat_biff8(data)
        raise

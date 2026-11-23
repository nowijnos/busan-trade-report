"""Render text on charts solely from pre-rasterized PNG glyph images.

This module DOES NOT load, find, or inspect any system font.  Every Hangul,
number, and symbol is composited as pixels from chart_assets/glyph_atlas.png.
The Render environment, font cache, font family fallback, or Matplotlib fonts
therefore cannot change chart text into missing-glyph boxes.
"""
from pathlib import Path
from PIL import Image, ImageDraw
import io, json, math, hashlib

ROOT=Path(__file__).resolve().parent
GLYPHS=ROOT/'chart_assets'/'glyph_atlas.png'
METRICS=ROOT/'chart_assets'/'glyph_metrics.json'

class RasterText:
    def __init__(self):
        self.atlas=Image.open(GLYPHS).convert('L')
        m=json.loads(METRICS.read_text('utf-8'))
        self.base_size=m['base_size']
        self.glyphs=m['glyphs']
        if not self.glyphs or self.atlas.width < 500:
            raise ValueError('이미지 글꼴 아틀라스를 읽을 수 없음')

    def require(self, s):
        missing=set(s)-set(self.glyphs)
        if missing:
            raise ValueError('차트 문자 그림 누락: '+''.join(sorted(missing)))

    def width(self, s, size):
        self.require(s)
        return sum(self.glyphs[ch]['a'] for ch in s)*size/self.base_size

    def text_image(self, s, size, color='#17202d'):
        self.require(s)
        k=size/self.base_size
        advances=[]; pos=0.0
        y0=1e6;y1=-1e6
        for ch in s:
            g=self.glyphs[ch]
            advances.append(pos)
            pos+=g['a']
            y0=min(y0,g['bbox'][1]);y1=max(y1,g['bbox'][3])
        if not s: return Image.new('RGBA',(1,1),(0,0,0,0))
        if y1 <= y0:y0,y1=-75,0
        w=max(1,math.ceil(pos*k+4));h=max(1,math.ceil((y1-y0)*k+4))
        out=Image.new('RGBA',(w,h),(0,0,0,0))
        color=color.lstrip('#')
        rgb=tuple(int(color[i:i+2],16) for i in (0,2,4))
        for ch,pen in zip(s,advances):
            if ch==' ':continue
            g=self.glyphs[ch];sx,sy=g['xy'];sw,sh=g['wh']
            src=self.atlas.crop((sx,sy,sx+sw,sy+sh))
            tw=max(1,round(sw*k));th=max(1,round(sh*k))
            src=src.resize((tw,th),Image.Resampling.LANCZOS)
            colored=Image.new('RGBA',src.size,rgb+(0,));colored.putalpha(src)
            destx=round((pen+g['bbox'][0])*k)
            desty=round((g['bbox'][1]-y0)*k)
            out.alpha_composite(colored,(destx,desty))
        return out

    def draw(self, canvas, s, xy, size, color='#202832',align='center',valign='center',bg=None,pad=2):
        img=self.text_image(s,size,color)
        x,y=xy
        if align=='center':x-=img.width/2
        elif align=='right':x-=img.width
        if valign=='center':y-=img.height/2
        elif valign=='bottom':y-=img.height
        x,y=round(x),round(y)
        if bg:
            d=ImageDraw.Draw(canvas)
            d.rectangle((x-pad,y-pad,x+img.width+pad,y+img.height+pad), fill=bg)
        canvas.alpha_composite(img,(x,y))
        return (x-pad,y-pad,x+img.width+pad,y+img.height+pad)


def overlaps(a,b,pad=3):
    return a[0]<b[2]+pad and a[2]>b[0]-pad and a[1]<b[3]+pad and a[3]>b[1]-pad


def segment_crosses_box(box,segment,pad=2):
    (x1,y1),(x2,y2)=segment
    x0,y0,x3,y3=box[0]-pad,box[1]-pad,box[2]+pad,box[3]+pad
    dx,dy=x2-x1,y2-y1
    lo,hi=0.,1.
    for p,q in ((-dx,x1-x0),(dx,x3-x1),(-dy,y1-y0),(dy,y3-y1)):
        if abs(p)<1e-9:
            if q<0:return False
            continue
        t=q/p
        if p<0:lo=max(lo,t)
        else:hi=min(hi,t)
        if lo>hi:return False
    return True


def intersects_polyline(box,line):
    return any(segment_crosses_box(box,segment) for segment in zip(line,line[1:]))


def render_chart(kind, chart_points, output_size):
    """Bar/line chart from 13 Excel-equivalent rows, NO FONT rendering.

    chart_points: [(YYYY.MM, amounts_in_million_USD, yoy_percent),...].
    output_size keeps the report picture dimensions (2x original template BMP).
    """
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator, MultipleLocator
    from matplotlib.patches import Rectangle
    labels=[r[0] for r in chart_points]
    amounts=[r[1] for r in chart_points]
    rates=[r[2] for r in chart_points]
    if len(labels)!=13 or len(set(labels))!=13:
        raise ValueError('13개월 그래프의 기간에 오류가 있습니다')
    for amount,rate in zip(amounts,rates):
        if amount is None or rate is None or not(math.isfinite(amount) and math.isfinite(rate)):
            raise ValueError('그래프 필수 수치 누락 또는 무효값')
    painter=RasterText()
    W,H=output_size
    target_dpi=140
    fig,ax=plt.subplots(figsize=(W/target_dpi,H/target_dpi),dpi=target_dpi)
    fig.patch.set_facecolor('white');ax.set_facecolor('white')
    ax.bar(range(13),amounts,width=.59,color='#4776BD',zorder=2)
    ax.set_xlim(-.7,12.7)
    maxval=max(amounts)
    step=(200 if maxval<2000 else (500 if maxval<6000 else 1000))
    ytop=math.ceil(maxval*1.19/step)*step
    ax.set_ylim(0,ytop)
    ax.yaxis.set_major_locator(MultipleLocator(step))
    ax.set_xticks(range(13))
    # No Matplotlib text anywhere.  NO font cache or missing glyph substitution.
    ax.tick_params(axis='both',which='both',length=0,labelleft=False,labelbottom=False)
    ax.grid(axis='y',color='#D9DEE6',linewidth=.85,zorder=0)
    ax.spines[['top','right','left']].set_visible(False)
    ax.spines['bottom'].set_color('#2C2C2C')
    ax2=ax.twinx()
    spread=max(max(rates)-min(rates),8)
    lo=min(rates)-spread*.30; hi=max(rates)+spread*.34
    ax2.set_ylim(lo,hi)
    ax2.yaxis.set_major_locator(MaxNLocator(nbins=6, steps=[1,2,2.5,5,10]))
    ax2.tick_params(axis='both',which='both',length=0,labelright=False,labelleft=False)
    ax2.spines[['top','left','right','bottom']].set_visible(False)
    ax2.plot(range(13),rates,color='#EE7D28',marker='s',markersize=4.0,linewidth=2.0,zorder=4)
    fig.subplots_adjust(top=.815,bottom=.235,left=.108,right=.900)
    fig.canvas.draw()
    rgba=fig.canvas.buffer_rgba()
    canvas=Image.fromarray(__import__('numpy').asarray(rgba)).convert('RGBA')
    if canvas.size!=(W,H):raise ValueError('차트 이미지 픽셀 크기 오류')
    def p1(x,y):
        p=ax.transData.transform((x,y));return (p[0],H-p[1])
    def p2(x,y):
        p=ax2.transData.transform((x,y));return (p[0],H-p[1])
    rect=ax.get_window_extent()
    plot_left,plot_right=rect.x0,rect.x1
    plot_top,plot_bottom=H-rect.y1,H-rect.y0
    scale=W/1366
    def sz(n):return n*scale
    dr=ImageDraw.Draw(canvas)
    # legend, both axis units
    legsize=sz(25)
    first=f'{kind}액(좌측)'; second='전년동월비 증감률(우측)'
    painter.require(first+second+'(백만 달러)(%)'+''.join(labels))
    a=painter.width(first,legsize)
    b=painter.width(second,legsize)
    total=48*scale+a+45*scale+48*scale+b
    x=(W-total)/2
    ly=sz(53)
    dr.rectangle((x,ly,x+sz(42),ly+sz(17)),fill='#4776BD')
    painter.draw(canvas,first,(x+sz(49),ly+sz(7)),legsize,align='left')
    x+=sz(48)+a+sz(45)
    dr.line((x,ly+sz(9),x+sz(42),ly+sz(9)),fill='#EE7D28',width=max(2,round(sz(4))))
    dr.rectangle((x+sz(17),ly+sz(4),x+sz(25),ly+sz(12)),fill='#EE7D28')
    painter.draw(canvas,second,(x+sz(49),ly+sz(7)),legsize,align='left')
    painter.draw(canvas,'(백만 달러)',(plot_left,plot_top-sz(39)),sz(21),align='left')
    painter.draw(canvas,'(%)',(plot_right+sz(39),plot_top-sz(39)),sz(21),align='right')
    # explicit numeric tick labels
    ticks_left=[v for v in ax.get_yticks() if -1e-6<=v<=ytop+1e-6]
    for yv in ticks_left:
        xx,yy=p1(0,yv)
        painter.draw(canvas,f'{yv:,.0f}',(plot_left-sz(14),yy),sz(20),align='right')
    for rv in ax2.get_yticks():
        if lo-1e-6<=rv<=hi+1e-6:
            xx,yy=p2(0,rv)
            painter.draw(canvas,f'{rv:,.0f}',(plot_right+sz(15),yy),sz(20),align='left')
    # tilted monthly labels, fully rasterized and independent of font support.
    for i,label in enumerate(labels):
        tex=painter.text_image(label,sz(21),'#202832')
        rot=tex.rotate(43,expand=True,resample=Image.Resampling.BICUBIC)
        xx,_=p1(i,0)
        xy=(round(xx-rot.width+sz(21)),round(plot_bottom+sz(12)))
        canvas.alpha_composite(rot,xy)
    # Amount labels: choose position where line does not cover data, and draw
    # them with white backing above all lines as a second safety measure.
    line=[p2(i,r) for i,r in enumerate(rates)]
    amount_boxes=[]
    for i,amount in enumerate(amounts):
        x,y=p1(i,amount)
        t=f'{amount:,.0f}'
        text_im=painter.text_image(t,sz(22),'#202832')
        candidates=[(0,-sz(10)),(0,-sz(24)),(0,-sz(42)),(0,-sz(58)),(-sz(16),-sz(26)),(sz(16),-sz(26))]
        scored=[]
        for rank,(dx,dy) in enumerate(candidates):
            midx=x+dx;bottom=y+dy
            r=(midx-text_im.width/2-3,bottom-text_im.height-3,midx+text_im.width/2+3,bottom+3)
            out=(r[0]<plot_left-8 or r[2]>plot_right+8 or r[1]<plot_top+3)
            hit=intersects_polyline(r,line)
            overlap=any(overlaps(r,rr,5) for rr in amount_boxes)
            scored.append(((5000 if out else 0)+(1500 if hit else 0)+(2000 if overlap else 0)+rank,r,(midx,bottom)))
            if not(out or hit or overlap):break
        _,r,(midx,bottom)=min(scored,key=lambda k:k[0])
        painter.draw(canvas,t,(midx,bottom),sz(22),align='center',valign='bottom',bg='#FFFFFF',pad=2)
        amount_boxes.append(r)
    # Rate labels: six positions of the original report, collision checked.
    occupied=list(amount_boxes)
    for i in (0,3,6,9,11,12):
        x,y=p2(i,rates[i]);t=f'{rates[i]:.1f}'
        text_im=painter.text_image(t,sz(21),'#B24D12')
        candidates=[(0,-sz(12)),(0,sz(32)),(-sz(18),-sz(25)),(sz(18),-sz(25)),(-sz(16),sz(42)),(sz(16),sz(42)),(0,-sz(50)),(0,sz(58))]
        scored=[]
        for rank,(dx,dy) in enumerate(candidates):
            mx=x+dx;my=y+dy
            box=(mx-text_im.width/2-5,my-text_im.height/2-5,mx+text_im.width/2+5,my+text_im.height/2+5)
            out=(box[0]<plot_left-7 or box[2]>plot_right+7 or box[1]<plot_top+4 or box[3]>plot_bottom-2)
            overlap=any(overlaps(box,other,6) for other in occupied)
            hit=intersects_polyline(box,line)
            scored.append(((5000 if out else 0)+(2500 if overlap else 0)+(75 if hit else 0)+rank,box,(mx,my)))
            if not(out or overlap or hit):break
        _,box,(mx,my)=min(scored,key=lambda k:k[0])
        dr.rectangle(box,fill='#FFFFFF',outline='#6F9DD3',width=max(1,round(sz(1))))
        painter.draw(canvas,t,(mx,my),sz(21),'#B24D12')
        occupied.append(box)
    plt.close(fig)
    # Assert at least one nonwhite pixel in the legend, tick and date bands.
    # This checks the OUTPUT PIXELS, not merely font file names or glyph CMAP.
    for band in [(0,0,W,round(H*.09)),
                 (0,round(plot_bottom+5),W,min(H,round(plot_bottom+H*.17)))]:
        region=canvas.crop(band).convert('RGB')
        colors=region.getcolors(region.width*region.height)
        if colors is not None and len(colors)<4:
            raise ValueError('글자 출력 영역이 비어 있어 차트 생성을 중단')
    return canvas.convert('RGB')

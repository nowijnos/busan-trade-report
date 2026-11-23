"""Collision-aware labels for the 13-month Busan bar + line charts.

Never lets the orange YoY line pass over a monthly amount. Rate callouts are
placed away from the amount labels. The values and chart geometry are untouched.
"""
from __future__ import annotations
from matplotlib.transforms import Bbox


def intersects_segment(box, x1, y1, x2, y2, pad=3.5):
    """Liang-Barsky segment versus an expanded axis-aligned text rectangle."""
    xmin, ymin, xmax, ymax = (box.x0-pad, box.y0-pad, box.x1+pad, box.y1+pad)
    dx, dy = x2-x1, y2-y1
    t0,t1=0.0,1.0
    for p,q in ((-dx,x1-xmin),(dx,xmax-x1),(-dy,y1-ymin),(dy,ymax-y1)):
        if abs(p)<1e-10:
            if q<0:return False
            continue
        t=q/p
        if p<0:t0=max(t0,t)
        else:t1=min(t1,t)
        if t0>t1:return False
    return True


def close_to_polyline(box, points, pad=4):
    return any(intersects_segment(box,*start,*end,pad=pad)
               for start,end in zip(points,points[1:]))


def box_overlap(a,b,pad=3):
    return not (a.x1+pad<=b.x0 or b.x1+pad<=a.x0
                or a.y1+pad<=b.y0 or b.y1+pad<=a.y0)


def place_labels(fig,ax,ax2,amounts,rates,font):
    """Place all 13 bar values + six YoY callouts with pixel-aware avoidance.

    Use ax2 artists (rather than ax) to ensure the text is in front of the
    orange line, which belongs to the twin axis. The short white backing is a
    last-resort readability guard if lines nearly touch a label.
    """
    renderer=fig.canvas.get_renderer()
    scale=fig.dpi/72.0
    from matplotlib.font_manager import FontProperties
    font_amount=font.copy();font_amount.set_size(7.4)
    font_rate=font.copy();font_rate.set_size(7.1)
    def measure(txt,prop):
        w,h,desc=renderer.get_text_width_height_descent(txt,prop,ismath=False)
        return w, max(h,8*scale)
    line_xy=[ax2.transData.transform((i,rate)) for i,rate in enumerate(rates)]
    bounds=ax.get_window_extent(renderer)
    occupied=[]
    amount_boxes=[]
    moved=0

    for i,amount in enumerate(amounts):
        text=f'{amount:,.0f}'
        w,h=measure(text,font_amount)
        x,y=ax.transData.transform((i,amount))
        # Default is 3 pt directly over bar. Only raise / nudge a label
        # when the orange line would otherwise cross its glyphs.
        options=[(0,3),(0,8),(0,14),(0,21),(0,29),
                 (-12,12),(12,12),(-16,21),(16,21),
                 (0,38),(-17,31),(17,31)]
        best=None
        for rank,(dx,dy) in enumerate(options):
            tx=x+dx*scale; ty=y+dy*scale
            bb=Bbox.from_bounds(tx-w/2-2,ty-1,w+4,h+4)
            beyond=(bb.x0<bounds.x0-4 or bb.x1>bounds.x1+4 or bb.y1>bounds.y1-2)
            collision_line=close_to_polyline(bb,line_xy,pad=4)
            collision_labels=any(box_overlap(bb,b,4) for b in occupied)
            score=(900 if beyond else 0)+(500 if collision_line else 0)+(
                 600 if collision_labels else 0)+rank*1.5
            if best is None or score<best[0]:best=(score,(dx,dy),bb)
            if not beyond and not collision_line and not collision_labels:break
        (_, (dx,dy),bb)=best
        if dx!=0 or dy!=3:moved+=1
        # On the twin axis: above both bars and line, yet the line and bars
        # retain their original drawing order elsewhere.
        ax2.annotate(text,(i,amount),xycoords=ax.transData,
             xytext=(dx,dy),textcoords='offset points',
             ha='center',va='bottom',fontproperties=font_amount,
             color='#202832',zorder=40,clip_on=False,
             bbox={'facecolor':'white','edgecolor':'none','alpha':.97,'pad':.12})
        occupied.append(bb);amount_boxes.append(bb)

    rate_boxes=[]
    for i in (0,3,6,9,11,12):
        r=rates[i];text=f'{r:.1f}'
        w,h=measure(text,font_rate)
        x,y=line_xy[i]
        # Prefer close placement, but never on top of a black monthly amount.
        # A few more candidates for crowded high-growth months.
        options=[(0,8,'bottom'),(0,-8,'top'),(-13,-11,'top'),
                 (13,-11,'top'),(-15,10,'bottom'),(15,10,'bottom'),
                 (-18,-19,'top'),(18,-19,'top'),
                 (-20,21,'bottom'),(20,21,'bottom'),
                 (0,-25,'top'),(0,24,'bottom')]
        best=None
        for rank,(dx,dy,va) in enumerate(options):
            tx=x+dx*scale;anchor_y=y+dy*scale
            ty=anchor_y if va=='bottom' else anchor_y-h
            bb=Bbox.from_bounds(tx-w/2-3,ty-2,w+6,h+4)
            beyond=(bb.x0<bounds.x0-3 or bb.x1>bounds.x1+3 or
                    bb.y0<bounds.y0-3 or bb.y1>bounds.y1+5)
            collision_amounts=any(box_overlap(bb,b,5) for b in amount_boxes)
            collision_others=any(box_overlap(bb,b,5) for b in rate_boxes)
            line_collision=close_to_polyline(bb,line_xy,pad=2)
            score=(1200 if beyond else 0)+(1100 if collision_amounts else 0)+(
                  950 if collision_others else 0)+(60 if line_collision else 0)+rank*1.5
            if best is None or score<best[0]:best=(score,(dx,dy,va),bb)
            if not beyond and not collision_amounts and not collision_others and not line_collision:break
        _,(dx,dy,va),bb=best
        ax2.annotate(text,(i,r),xytext=(dx,dy),textcoords='offset points',
            ha='center',va=va,fontproperties=font_rate,fontsize=7.1,
            color='#B24D12',zorder=50,
            bbox={'facecolor':'white','edgecolor':'#4776BD',
                  'linewidth':.48,'alpha':.98,'pad':.6})
        rate_boxes.append(bb);occupied.append(bb)
    return {'moved_amount_labels':moved,'amount_boxes':amount_boxes,
            'rate_boxes':rate_boxes,'line_xy':line_xy}

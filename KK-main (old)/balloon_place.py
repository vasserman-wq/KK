"""SW 2021 balloon placement prototype. Local PDF only; no network or OCR.
Input/output: UTF-8 TSV. All source/target coordinates are sheet metres.
Only rows with kind 14/38 and movable=1 may receive an AUTO result.
Protocol multi-sheet-v1: page_index is zero-based; page_count is the expected total.
"""
from __future__ import annotations
import argparse, csv, json, math, re, sys, traceback, unicodedata
from dataclasses import dataclass, replace
from pathlib import Path
DEPENDENCY_ERROR = None
try:
    import cv2
    import numpy as np
    import pymupdf as fitz
except Exception as exc:
    DEPENDENCY_ERROR = str(exc)

PT_PER_M = 72000.0 / 25.4
DEFAULTS = dict(dpi=240, max_pixels=18000000, ink_threshold=245,
    clearance_mm=0.25, endpoint_gap_mm=0.45, edge_margin_mm=2.0,
    max_text_distance_mm=14.0, max_endpoint_distance_mm=35.0,
    parallel_degrees=5.0, min_reference_line_mm=2.5,
    max_reference_offset_mm=3.5, ambiguity_margin=0.8,
    forbidden_rectangles_mm=[], keep_debug=False,
    fragment_gap_mm=4.5,fragment_baseline_mm=.65,weld_shelf_coverage=.65,
    endpoint_text_slack_mm=1.0,near_text_gap_mm=.45,
    neighbor_advantage_mm=1.0,max_near_text_mm=9.0,respect_form_bands=True,enable_source_adjustment=True,
    max_source_shift_mm=4.0,source_trials_per_item=6,max_source_trials=24)

@dataclass
class Request:
    id: int
    pdf: str
    width: float
    height: float
    kind: int
    source: np.ndarray
    text: str
    center: np.ndarray
    radius: float
    movable: bool
    page_index: int = 0
    page_count: int = 1

@dataclass
class TextLine:
    text: str
    u: np.ndarray
    v: np.ndarray
    points: np.ndarray
    rect: fitz.Rect
    baseline: float | None = None
    def project(self):
        a = self.points @ self.u; b = self.points @ self.v
        return float(a.min()), float(a.max()), float(b.min()), float(b.max())

def mm_pt(mm): return mm * 72 / 25.4

def text_lines(page):
    result = []
    for block in page.get_text('rawdict')['blocks']:
        for line in block.get('lines', []):
            spans = line.get('spans', [])
            txt = ''.join(c['c'] for sp in spans for c in sp.get('chars', []))
            if not normal(txt): continue
            u = np.array(line['dir'], dtype=float)
            if np.linalg.norm(u) < .9: continue
            u /= np.linalg.norm(u); v = np.array([-u[1], u[0]])
            points=[];origins=[]
            for sp in spans:
                for c in sp.get('chars', []):
                    if 'origin' in c:origins.append(c['origin'])
                    q = fitz.recover_char_quad(tuple(u), sp, c)
                    points.extend([(p.x,p.y) for p in q])
            if points:
                baseline=float(np.median(np.array(origins)@v)) if origins else None
                candidate=TextLine(txt,u,v,np.array(points),fitz.Rect(line['bbox']),baseline)
                if not any(normal(x.text)==normal(txt) and np.linalg.norm(np.array(x.rect)-np.array(candidate.rect))<.2 for x in result):result.append(candidate)
    return result

def segments(page):
    result=[]
    for path in page.get_drawings():
        # Filled hatching is an obstacle in the raster, never a reference line.
        if path.get('type') not in ('s','fs'): continue
        for item in path['items']:
            if item[0]=='l':
                p,q=np.array(item[1],dtype=float),np.array(item[2],dtype=float)
                if np.linalg.norm(q-p)>mm_pt(1): result.append((p,q))
    return result

def distance_to_rect(point,rect):
    return math.hypot(max(rect.x0-point[0],0,point[0]-rect.x1),
                      max(rect.y0-point[1],0,point[1]-rect.y1))

def repair_pdf_text(text):
    """Repair Latin-1 codepoints carrying CP1251 bytes; preserve real Unicode."""
    def fix(match):
        part=match.group(0)
        try:
            restored=part.encode('latin1').decode('cp1251')
            return restored if any('\u0400'<=c<='\u04ff' for c in restored) else part
        except (UnicodeError,LookupError):return part
    text=unicodedata.normalize('NFKC',str(text))
    text=re.sub(r'[\u0080-\u00ff]+',fix,text)
    text=re.sub(r'<[^>]*>',' ',text)
    text=unicodedata.normalize('NFKC',text).replace('\u2212','-')
    def decimal(match):
        value=match.group(0).replace(',','.')
        return value.rstrip('0').rstrip('.')
    return re.sub(r'\d+[.,]\d+',decimal,text)

CONFUSABLES=str.maketrans({'а':'a','в':'b','е':'e','к':'k','м':'m','н':'h','о':'o',
                          'р':'p','с':'c','т':'t','у':'y','х':'x'})
def normal(s):
    s=repair_pdf_text(s).casefold().translate(CONFUSABLES)
    return ''.join(c for c in s if c.isalnum())

def source_caption(req):
    # Explanatory text below a weld shelf is an obstacle, not its main caption.
    return req.text.split('|',1)[0] if req.kind==38 else req.text

def numbers(text):
    return re.findall(r'\d+(?:\.\d+)?',repair_pdf_text(text))

def join_lines(left,right):
    points=np.concatenate((left.points,right.points))
    rect=fitz.Rect(left.rect);rect.include_rect(right.rect)
    baselines=[x.baseline for x in (left,right) if x.baseline is not None]
    return TextLine(left.text+' '+right.text,left.u,left.v,points,rect,
                    float(np.median(baselines)) if baselines else None)

def overlap_fraction(a,b):
    inter=a & b
    return inter.get_area()/max(min(a.get_area(),b.get_area()),1e-9) if not inter.is_empty else 0

def choose_text(req,lines,cfg):
    anchor=np.array([req.source[0]*PT_PER_M,(req.height-req.source[1])*PT_PER_M])
    caption=source_caption(req);desired=normal(caption)
    if not desired:return None,'EMPTY_SOURCE_TEXT'
    wanted_numbers=numbers(caption)
    local=[line for line in lines if distance_to_rect(anchor,line.rect)<=mm_pt(cfg['max_text_distance_mm'])]
    variants=[]
    for seed in local:
        group=seed;variants.append(group);used={id(seed)}
        for _ in range(6):
            a,b,c,d=group.project();mid=(c+d)/2;neighbors=[]
            for other in local:
                if id(other) in used or float(group.u@other.u)<.999:continue
                oa,ob,oc,od=other.project()
                if abs((oc+od)/2-mid)>mm_pt(cfg['fragment_baseline_mm']):continue
                if oa>=b-mm_pt(.2):gap=max(oa-b,0);combined=join_lines(group,other)
                elif ob<=a+mm_pt(.2):gap=max(a-ob,0);combined=join_lines(other,group)
                else:continue
                if gap>mm_pt(cfg['fragment_gap_mm']):continue
                text=normal(combined.text)
                if text and (text in desired or desired in text):neighbors.append((gap,other,combined))
            if not neighbors:break
            _,other,group=min(neighbors,key=lambda x:x[0]);used.add(id(other));variants.append(group)
    scored=[]
    for line in variants:
        found=normal(line.text)
        if not found or not (found in desired or desired in found):continue
        if len(found)<2 and found!=desired:continue
        if req.kind==14 and wanted_numbers:
            actual=numbers(line.text)
            # Do not match 30 to 130 merely because the string is contained in it.
            if not actual or actual[0]!=wanted_numbers[0]:continue
        coverage=min(len(found),len(desired))/max(len(found),len(desired))
        dist=distance_to_rect(anchor,line.rect)
        scored.append((dist/mm_pt(1)+(1-coverage)*6,line))
    scored.sort(key=lambda x:x[0])
    distinct=[]
    for score,line in scored:
        # A merged caption and its constituent PDF fragments are one candidate.
        if any(overlap_fraction(line.rect,x[1].rect)>.85 for x in distinct):continue
        distinct.append((score,line))
    if not distinct:return None,'SOURCE_TEXT_NOT_MATCHED'
    if len(distinct)>1 and distinct[1][0]-distinct[0][0]<cfg['ambiguity_margin']:
        return None,'AMBIGUOUS_SOURCE_TEXT'
    return distinct[0][1],'TEXT_MATCHED'

def candidates(line,refs,radius_pt,cfg,kind=14):
    a,b,c,d=line.project();u,v=line.u,line.v
    baseline=line.baseline if line.baseline is not None else d-mm_pt(.7)
    near=mm_pt(cfg['max_reference_offset_mm']);max_end=mm_pt(cfg['max_endpoint_distance_mm'])
    sine=math.sin(math.radians(cfg['parallel_degrees']));result=[]
    for p,q in refs:
        vec=q-p;length=np.linalg.norm(vec)
        if length<mm_pt(cfg['min_reference_line_mm']) or abs(float((vec/length)@v))>sine:continue
        pu,qu=float(p@u),float(q@u);lo,hi=sorted((pu,qu));nv=float(((p+q)/2)@v)
        if nv<c-near or nv>d+near or hi<a-mm_pt(3) or lo>b+mm_pt(3):continue
        if kind==38:
            coverage=max(0,min(hi,b)-max(lo,a))/max(b-a,.1)
            # A weld shelf underlies most of the main caption. The upper short
            # horizontal part of an auxiliary symbol is NOT that shelf.
            if coverage<cfg['weld_shelf_coverage']:continue
            if nv<baseline-mm_pt(.6) or nv>baseline+mm_pt(2):continue
        for sign,end in [(-1,p if pu<=qu else q),(1,q if pu<=qu else p)]:
            end_u=float(end@u)
            if sign<0 and end_u>a+mm_pt(cfg['endpoint_text_slack_mm']):continue
            if sign>0 and end_u<b-mm_pt(cfg['endpoint_text_slack_mm']):continue
            distance=max(a-end_u,0,end_u-b)
            if distance>max_end:continue
            center=end+sign*u*(radius_pt+mm_pt(cfg['endpoint_gap_mm']))
            score=distance/mm_pt(1)+abs(nv-baseline)/mm_pt(1)*2
            if kind==38:
                # Prefer the FREE shelf end, opposite the elbow/leader junction.
                # A left-side leader (148) prefers right; a right-side leader
                # remains free to prefer left. Never hard-code the item number.
                joints=[]
                for side,endpoint in [(-1,u*lo+v*nv),(1,u*hi+v*nv)]:
                    for rp,rq in refs:
                        rv=rq-rp;rl=np.linalg.norm(rv)
                        if rl<mm_pt(2) or abs(float(rv@v))/rl<.15:continue
                        dp,dq=np.linalg.norm(rp-endpoint),np.linalg.norm(rq-endpoint)
                        if min(dp,dq)<mm_pt(.5):
                            far=rq if dp<dq else rp
                            # Ignore the short upright auxiliary weld symbol:
                            # a leader leaves below the shelf, or is long.
                            if float(far@v)-nv>mm_pt(.75) or rl>=mm_pt(6):
                                joints.append(side);break
                preferred=-joints[0] if len(set(joints))==1 else 1
                if sign!=preferred:score+=100
            result.append((score,center,'LINE_END_'+('LEFT' if sign<0 else 'RIGHT')))
    result.sort(key=lambda x:x[0]);unique=[]
    for option in result:
        if not any(np.linalg.norm(option[1]-x[1])<mm_pt(.3) for x in unique):unique.append(option)
    return unique

def secondary_candidates(line,radius_pt,cfg):
    a,b,c,d=line.project();u,v=line.u,line.v;gap=mm_pt(cfg['near_text_gap_mm'])
    baseline=line.baseline if line.baseline is not None else (c+d)/2
    result=[]
    # First: close to text ends. Then above/below, without inventing a leader.
    for extra in [0,1,2.5,4]:
        for offset in [0,-1.5,1.5,-3,3]:
            for at,side in [(a-radius_pt-gap-mm_pt(extra),'LEFT'),(b+radius_pt+gap+mm_pt(extra),'RIGHT')]:
                center=u*at+v*(baseline+mm_pt(offset))
                result.append((1+abs(offset)*.6+extra*.3,center,'NEAR_TEXT_'+side))
    for extra in [0,.75,1.5,2.5]:
        for at in [(a+b)/2,a+(b-a)*.25,a+(b-a)*.75,a,b]:
            for nv,side in [(d+radius_pt+gap+mm_pt(extra),'BELOW'),
                            (c-radius_pt-gap-mm_pt(extra),'ABOVE')]:
                center=u*at+v*nv
                result.append((3+extra+abs(at-(a+b)/2)/mm_pt(1)*.1,center,'NEAR_TEXT_'+side))
    result.sort(key=lambda x:x[0]);return result

def belongs_to_caption(center,line,lines,cfg):
    own_distance=distance_to_rect(center,line.rect)
    for other in lines:
        if other is line or overlap_fraction(other.rect,line.rect)>.8:continue
        if distance_to_rect(center,other.rect)+mm_pt(cfg['neighbor_advantage_mm'])<own_distance:
            return False
    return True


class Map:
    def __init__(self,page,cfg,width,height):
        self.cfg=cfg;self.width=width;self.height=height
        scale=min(cfg['dpi']/72, math.sqrt(cfg['max_pixels']/(page.rect.width*page.rect.height)))
        pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale),colorspace=fitz.csGRAY,alpha=False)
        self.ink=np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width)<cfg['ink_threshold']
        if cfg.get('respect_form_bands',True):
            # Detect long horizontal form separators. Protect empty title/header
            # cells too, not merely their printed lines. If no pair is found,
            # retain the generic sheet-edge/ink checks instead of guessing bands.
            horizontal=[]
            for path in page.get_drawings():
                if path.get('type') not in ('s','fs'):continue
                for item in path['items']:
                    pairs=[]
                    if item[0]=='l':pairs=[(item[1],item[2])]
                    elif item[0]=='re':
                        rr=fitz.Rect(item[1]);pairs=[(rr.tl,rr.tr),(rr.bl,rr.br)]
                    for p,q in pairs:
                        if abs(p.y-q.y)<.5 and abs(p.x-q.x)>=page.rect.width*.75:
                            horizontal.append((p.y+q.y)/2)
            upper=[y for y in horizontal if .04*page.rect.height<y<.3*page.rect.height]
            lower=[y for y in horizontal if .82*page.rect.height<y<.985*page.rect.height]
            if upper and lower and min(lower)-max(upper)>.45*page.rect.height:
                top,bottom=max(upper),min(lower)
                self.ink[:max(0,math.ceil(top/page.rect.height*self.ink.shape[0])),:]=True
                self.ink[min(self.ink.shape[0],math.floor(bottom/page.rect.height*self.ink.shape[0])):,:]=True
        self.counts=np.zeros(self.ink.shape,dtype=np.uint16)
        self.sx=pix.width/page.rect.width;self.sy=pix.height/page.rect.height
        for rect in cfg.get('forbidden_rectangles_mm',[]):
            # Explicit rectangles in sheet millimetres, origin bottom-left.
            x0,y0,x1,y1=map(float,rect)
            xa,xb=sorted((round(x0/1000*PT_PER_M*self.sx),round(x1/1000*PT_PER_M*self.sx)))
            ya,yb=sorted((round((height-y0/1000)*PT_PER_M*self.sy),round((height-y1/1000)*PT_PER_M*self.sy)))
            self.ink[max(0,ya):min(self.ink.shape[0],yb+1),max(0,xa):min(self.ink.shape[1],xb+1)]=True
    def disk(self,point,radius):
        px,py=point[0]*self.sx,point[1]*self.sy
        rx,ry=radius*self.sx,radius*self.sy
        x0=max(0,math.floor(px-rx));x1=min(self.ink.shape[1],math.ceil(px+rx)+1)
        y0=max(0,math.floor(py-ry));y1=min(self.ink.shape[0],math.ceil(py+ry)+1)
        if x1<=x0 or y1<=y0:return None
        yy,xx=np.ogrid[y0:y1,x0:x1]
        disk=((xx-px)/max(rx,.1))**2+((yy-py)/max(ry,.1))**2<=1
        return (slice(y0,y1),slice(x0,x1)),disk
    def reserve(self,point,radius,delta):
        region=self.disk(point,radius+mm_pt(self.cfg['clearance_mm']))
        if region is None:return
        slices,disk=region
        data=self.counts[slices]
        if delta<0:
            if np.any(data[disk]==0):raise RuntimeError('reservation underflow')
            data[disk]-=1
        else:data[disk]+=1
    def free(self,point,radius):
        extra=mm_pt(self.cfg['clearance_mm']);edge=mm_pt(self.cfg['edge_margin_mm'])
        w,h=self.width*PT_PER_M,self.height*PT_PER_M
        if not (edge+radius<=point[0]<=w-edge-radius and edge+radius<=point[1]<=h-edge-radius):return False
        region=self.disk(point,radius+extra)
        if region is None:return False
        slices,disk=region
        return not (np.any(self.ink[slices][disk]) or np.any(self.counts[slices][disk]>0))

def read_requests(path):
    requests=[]
    with open(path,encoding='utf-8-sig',newline='') as f:
        for r in csv.DictReader(f,delimiter='\t'):
            requests.append(Request(int(r['id']),r['pdf'],float(r['width_m']),float(r['height_m']),int(r['kind']),
                np.array([float(r['source_x']),float(r['source_y'])]),r['text'],
                np.array([float(r['center_x']),float(r['center_y'])]),float(r['radius_m']),r['movable']=='1',int(r.get('page_index',0)),int(r.get('page_count',1))))
    ids=[r.id for r in requests]
    if len(ids)!=len(set(ids)):raise ValueError('duplicate request ids')
    for r in requests:
        if not all(math.isfinite(x) for x in [r.width,r.height,*r.source,*r.center,r.radius]):raise ValueError('nonfinite request')
        if not (.0005<=r.radius<=.03 and .05<=r.width<=5 and .05<=r.height<=5):raise ValueError('invalid geometry')
        if not (0<=r.page_index<r.page_count<=10000):raise ValueError('invalid PDF page index/count')
    return requests

def solve(requests,cfg):
    output={r.id:dict(id=r.id,status='FALLBACK',x=r.center[0],y=r.center[1],reason='STANDARD_POSITION') for r in requests}
    grouped={}
    for req in requests:grouped.setdefault((req.pdf,req.page_index,req.page_count),[]).append(req)
    for (pdf,page_index,page_count),group in grouped.items():
        try:
            with fitz.open(pdf) as doc:
                if len(doc)!=page_count:raise ValueError(f'PDF_PAGE_COUNT_MISMATCH_expected_{page_count}_actual_{len(doc)}')
                if not (0<=page_index<len(doc)):raise ValueError('PDF_PAGE_INDEX_OUT_OF_RANGE')
                page=doc[page_index];page.set_rotation(0)
                width,height=group[0].width,group[0].height
                if abs(page.rect.width/(width*PT_PER_M)-1)>.015 or abs(page.rect.height/(height*PT_PER_M)-1)>.015:
                    raise ValueError('PDF_SHEET_SIZE_MISMATCH')
                occupancy=Map(page,cfg,width,height)
                lines=text_lines(page);refs=segments(page)
                for req in group:
                    center=np.array([req.center[0]*PT_PER_M,(height-req.center[1])*PT_PER_M])
                    occupancy.reserve(center,req.radius*PT_PER_M,1)
                matched={}
                for req in group:
                    if not req.movable or req.kind not in (14,38):
                        output[req.id]['reason']='UNCHANGED_ANNOTATION_TYPE'
                    else:matched[req.id]=choose_text(req,lines,cfg)
                # A second pass may use space released by successful first-pass moves.
                # Fallback circles remain reserved at all times, never silently discarded.
                for attempt in range(2):
                    progress=False
                    for req in group:
                        if req.id not in matched or output[req.id]['status']=='AUTO':continue
                        old=np.array([req.center[0]*PT_PER_M,(height-req.center[1])*PT_PER_M])
                        radius=req.radius*PT_PER_M;line,reason=matched[req.id]
                        if line is None:
                            output[req.id]['reason']=reason;continue
                        occupancy.reserve(old,radius,-1)
                        chosen=None
                        primary=candidates(line,refs,radius,cfg,req.kind)
                        secondary=secondary_candidates(line,radius,cfg)
                        reason='NO_FREE_CANDIDATE'
                        for options,is_secondary in [(primary,False),(secondary,True)]:
                            for _,center,label in options:
                                if is_secondary and distance_to_rect(center,line.rect)>mm_pt(cfg['max_near_text_mm']):continue
                                if not belongs_to_caption(center,line,lines,cfg):continue
                                if occupancy.free(center,radius):
                                    chosen=center;reason=label;break
                            if chosen is not None:break
                        if chosen is None:
                            occupancy.reserve(old,radius,1)
                            output[req.id]['reason']=reason
                        else:
                            occupancy.reserve(chosen,radius,1);progress=True
                            output[req.id].update(status='AUTO',x=float(chosen[0]/PT_PER_M),y=float(height-chosen[1]/PT_PER_M),reason=reason)
                    if not progress:break
                if cfg.get('keep_debug'):
                    preview=np.repeat((~occupancy.ink).astype(np.uint8)[...,None]*255,3,axis=2)
                    for req in group:
                        row=output[req.id]
                        xy=(round(row['x']*PT_PER_M*occupancy.sx),round((height-row['y'])*PT_PER_M*occupancy.sy))
                        radius=round(req.radius*PT_PER_M*occupancy.sx)
                        cv2.circle(preview,xy,radius,(0,160,0) if row['status']=='AUTO' else (0,0,255),1)
                    cv2.imencode('.png',preview)[1].tofile(str(Path(pdf).with_name(Path(pdf).stem+f'.page_{page_index+1}.preview.png')))
        except Exception as exc:
            # Revert this entire sheet's plan; never partially trust a failed map.
            for req in group:
                output[req.id].update(status='FALLBACK',x=req.center[0],y=req.center[1],reason='SHEET_ERROR_'+str(exc).replace('\t',' ').replace('\n',' ')[:150])
    return [output[r.id] for r in requests]

def propose_shifts(requests,cfg):
    """Offsets only: NEVER erase/translate source ink to pretend a move worked.
    VBA must trial each offset in SW and export a new snapshot for validation.
    """
    if not cfg.get('enable_source_adjustment',True):return []
    result=[];groups=[];budget=min(100,max(0,int(cfg['max_source_trials'])))
    limit=min(8,max(0,int(cfg['source_trials_per_item'])))
    max_shift=min(4.0,max(0,float(cfg['max_source_shift_mm'])))
    if max_shift<.5:return []
    for req in requests:
        if not req.movable or req.kind not in (14,38):continue
        proposals=[]
        with fitz.open(req.pdf) as doc:
            if len(doc)!=req.page_count:raise ValueError('SHIFT_PAGE_COUNT_MISMATCH')
            page=doc[req.page_index];page.set_rotation(0)
            if abs(page.rect.width/(req.width*PT_PER_M)-1)>.015 or abs(page.rect.height/(req.height*PT_PER_M)-1)>.015:
                raise ValueError('SHIFT_SHEET_SIZE_MISMATCH')
            line,reason=choose_text(req,text_lines(page),cfg)
            if line is None:continue
            # Small local steps first. Dimensions: normal to the text first;
            # welds: both normal directions, then shelf-axis translations.
            small=min(2,max_shift)
            options=[-line.v*small,line.v*small,line.u*small,-line.u*small,
                     -line.v*max_shift,line.v*max_shift]
            emitted=[]
            for delta_mm in options:
                if any(np.linalg.norm(delta_mm-x)<.01 for x in emitted):continue
                emitted.append(delta_mm)
                proposals.append(dict(id=req.id,trial=len(emitted),dx=float(delta_mm[0]/1000),
                                   dy=float(-delta_mm[1]/1000),reason='TRIAL_ONLY_REQUIRES_NEW_PDF'))
                if len(emitted)>=limit:break
        groups.append(proposals)
    # Spread the bounded budget over sources instead of using every attempt
    # on the first four crowded labels.
    for trial in range(1,limit+1):
        for proposals in groups:
            result.extend(row for row in proposals if row['trial']==trial)
    return result[:budget] if budget and limit else []

def raster_polygon(shape,points,sx,sy,padding=0):
    mask=np.zeros(shape,dtype=np.uint8)
    # Convex hull is necessary: PDF quad order need not be polygon order.
    pts=np.rint(np.asarray(points)*[sx,sy]).astype(np.int32)
    cv2.fillConvexPoly(mask,cv2.convexHull(pts),1)
    if padding:
        px=max(1,round(padding*max(sx,sy)))
        mask=cv2.dilate(mask,np.ones((2*px+1,2*px+1),np.uint8))
    return mask.astype(bool)

def validate_shift(requests,cfg,before,target_id,dx,dy):
    """Validate actual SW export, never a synthetic translation of the old PDF."""
    targets=[r for r in requests if r.id==target_id]
    if len(targets)!=1:raise ValueError('SHIFT_TARGET_NOT_UNIQUE')
    req=targets[0]
    def rejected(reason):
        return [dict(id=r.id,status='FALLBACK',x=r.center[0],y=r.center[1],reason=reason if r.id==target_id else 'UNCHANGED_ANNOTATION_TYPE') for r in requests]
    if req.kind not in (14,38) or not req.movable:return rejected('SHIFT_FORBIDDEN_TYPE')
    if not all(math.isfinite(x) for x in (dx,dy)) or not (0<math.hypot(dx,dy)<=min(4,float(cfg['max_source_shift_mm']))/1000+.000001):
        return rejected('SHIFT_DISTANCE_REJECTED')
    with fitz.open(before) as olddoc, fitz.open(req.pdf) as newdoc:
        if len(olddoc)!=req.page_count or len(newdoc)!=req.page_count:return rejected('SHIFT_PAGE_COUNT_MISMATCH')
        oldpage,newpage=olddoc[req.page_index],newdoc[req.page_index]
        oldpage.set_rotation(0);newpage.set_rotation(0)
        if np.linalg.norm(np.array(oldpage.rect)-np.array(newpage.rect))>.1:return rejected('SHIFT_SHEET_SIZE_MISMATCH')
        oldlines,newlines=text_lines(oldpage),text_lines(newpage)
        oldreq=replace(req,source=req.source-np.array([dx,dy]))
        oldline,reason=choose_text(oldreq,oldlines,cfg);newline,newreason=choose_text(req,newlines,cfg)
        if oldline is None or newline is None:return rejected('SHIFT_TEXT_NOT_CONFIRMED')
        delta=np.array([dx,-dy])*PT_PER_M
        oldmid=np.array([oldline.rect.x0+oldline.rect.x1,oldline.rect.y0+oldline.rect.y1])/2
        newmid=np.array([newline.rect.x0+newline.rect.x1,newline.rect.y0+newline.rect.y1])/2
        if np.linalg.norm(newmid-oldmid-delta)>mm_pt(.75):return rejected('SHIFT_TEXT_DISPLACEMENT_NOT_CONFIRMED')
        oldmap,newmap=Map(oldpage,cfg,req.width,req.height),Map(newpage,cfg,req.width,req.height)
        if oldmap.ink.shape!=newmap.ink.shape:return rejected('SHIFT_RASTER_SIZE_MISMATCH')
        # Reject a new caption box over old fixed ink outside its old box.
        # Deliberately conservative: no erasing a rectangle and assuming it empty.
        oldbox=raster_polygon(oldmap.ink.shape,oldline.points,oldmap.sx,oldmap.sy,mm_pt(.4))
        newbox=raster_polygon(newmap.ink.shape,newline.points,newmap.sx,newmap.sy)
        if np.count_nonzero(oldmap.ink & newbox & ~oldbox)>2:return rejected('SHIFT_CAPTION_OBSTACLE')
        # New strokes may not invade other text or any other new circle.
        near_old=cv2.dilate(oldmap.ink.astype(np.uint8),np.ones((3,3),np.uint8)).astype(bool)
        added=newmap.ink & ~near_old
        for other in oldlines:
            if overlap_fraction(other.rect,oldline.rect)>.8:continue
            protected=raster_polygon(oldmap.ink.shape,other.points,oldmap.sx,oldmap.sy,mm_pt(.3))
            if np.count_nonzero(added & protected)>2:return rejected('SHIFT_OTHER_CAPTION_COLLISION')
        for other in requests:
            if other.id==target_id or other.pdf!=req.pdf or other.page_index!=req.page_index:continue
            center=np.array([other.center[0],req.height-other.center[1]])*PT_PER_M
            disk=newmap.disk(center,other.radius*PT_PER_M+mm_pt(cfg['clearance_mm']))
            if disk and np.any(added[disk[0]][disk[1]]):return rejected('SHIFT_OTHER_BALLOON_COLLISION')
    # Only target can move in this validation pass. Others reserve actual circles.
    limited=[replace(r,movable=(r.id==target_id)) for r in requests]
    rows=solve(limited,cfg)
    for row in rows:
        if row['id']==target_id and row['status']=='AUTO':row['reason']='SOURCE_SHIFT_VALIDATED_'+row['reason']
    return rows

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--check',action='store_true')
    parser.add_argument('--action',choices=['place','propose-shifts','validate-shift'],default='place')
    parser.add_argument('--before');parser.add_argument('--target',type=int)
    parser.add_argument('--dx',type=float,default=0);parser.add_argument('--dy',type=float,default=0)
    parser.add_argument('--protocol',choices=['multi-sheet-v1'],default='multi-sheet-v1')
    parser.add_argument('--request');parser.add_argument('--response');parser.add_argument('--config');parser.add_argument('--done')
    args=parser.parse_args()
    if args.check:
        print('OK OpenCV='+cv2.__version__+' PyMuPDF='+fitz.VersionBind+' protocol=multi-sheet-v1 placement=2.8.7 source-shift=verified-trials');return 0
    if not args.request or not args.response:parser.error('--request and --response required')
    cfg=dict(DEFAULTS)
    if args.config:
        cfg.update(json.loads(Path(args.config).read_text(encoding='utf-8-sig')))
    requests=read_requests(args.request)
    fields=['id','status','x','y','reason']
    if args.action=='propose-shifts':
        rows=propose_shifts(requests,cfg);fields=['id','trial','dx','dy','reason']
    elif args.action=='validate-shift':
        if not args.before or args.target is None:raise ValueError('validation requires --before and --target')
        rows=validate_shift(requests,cfg,args.before,args.target,args.dx,args.dy)
    else:rows=solve(requests,cfg)
    out=Path(args.response);temporary=out.with_suffix(out.suffix+'.tmp')
    with temporary.open('w',encoding='ascii',errors='backslashreplace',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=fields,delimiter='\t',lineterminator='\n')
        writer.writeheader();writer.writerows(rows)
    temporary.replace(out)
    return 0

if __name__=='__main__':
    code=2;error=''
    try:
        if DEPENDENCY_ERROR:raise RuntimeError(DEPENDENCY_ERROR)
        code=main()
    except SystemExit as exc:
        code=int(exc.code or 0)
    except Exception:
        error=traceback.format_exc()
    finally:
        if '--done' in sys.argv:
            try:
                done=Path(sys.argv[sys.argv.index('--done')+1])
                temporary=done.with_suffix(done.suffix+'.tmp')
                temporary.write_text(str(code)+'\n'+error,encoding='ascii',errors='backslashreplace')
                temporary.replace(done)
            except Exception:pass
    if error:print(error,file=sys.stderr)
    raise SystemExit(code)

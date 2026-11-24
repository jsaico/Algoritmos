"""Cronograma de lo EJECUTADO (A4 horizontal, 2 hojas) a partir del Excel del cronograma."""
import datetime as dt, math, openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.colors import HexColor, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth as sw

XLSX='obra_camelidos/Cronograma_Ejecucion_Camelidos.xlsx'
OUT='obra_camelidos/Cronograma_Ejecutado_Camelidos.pdf'
FD='/usr/share/fonts/truetype/crosextra/'
for n,f in (('R','Carlito-Regular'),('B','Carlito-Bold'),('I','Carlito-Italic')): pdfmetrics.registerFont(TTFont(n,FD+f+'.ttf'))

wb=openpyxl.load_workbook(XLSX,data_only=True)
D,P,V,R=wb['Datos'],wb['Presupuesto'],wb['Valorizado'],wb['Resumen']
td=lambda v: v.date() if isinstance(v,dt.datetime) else v
CORTE=td(R['C5'].value); NMOD=R['C8'].value; CD=R['C10'].value
tr=[]
for r in range(6,41):
    g=lambda c:D.cell(r,c).value
    if g(1) is None or g(10)!='Ejecutado' or g(8) is None: continue
    tr.append(dict(meta=g(1),resp=g(3),tipo=g(4),n=g(5),a=g(6),b=g(7),ini=td(g(8)),fin=td(g(9))))
apoyo=dict(act=D['B45'].value,resp=D['C45'].value,n=D['D45'].value,a=D['E45'].value,b=D['F45'].value,ini=td(D['G45'].value),fin=td(D['H45'].value))
hitos=[(td(D.cell(r,2).value),D.cell(r,3).value) for r in (50,51)]
fer={td(D.cell(r,15).value) for r in range(6,21) if D.cell(r,15).value}
PER=['Jul','Ago','Set','Oct','Nov']; PL=['Julio','Agosto','Setiembre','Octubre','Noviembre']
part=[]
for r in range(6,60):
    if V.cell(r,1).value and V.cell(r,3).value:
        part.append(dict(item=V.cell(r,1).value,desc=V.cell(r,2).value.strip().capitalize(),und=V.cell(r,3).value,meta=V.cell(r,4).value,
            mu=V.cell(r,5).value,met=[V.cell(r,9+k).value for k in range(5)],val=[V.cell(r,17+k).value for k in range(5)]))
META={1:'Trazo, replanteo y excavaciones',2:'Muros',3:'Techos'}
mmods={m:sum(t['n'] for t in tr if t['meta']==m) for m in (1,2,3)}
vmes=[sum(p['val'][k] for p in part) for k in range(5)]; VT=sum(vmes)
vmeta={m:sum(sum(p['val']) for p in part if p['meta']==m) for m in (1,2,3)}

INK=HexColor('#26221F'); MUTE=HexColor('#6E665E'); LINE=HexColor('#CFC8BF'); LIGHT=HexColor('#F4F1EC')
DARK=HexColor('#3F3731'); MHEAD=HexColor('#E6E0D7'); WEEK=HexColor('#EFEBE5'); FER=HexColor('#F6E3BC')
C_ET={1:HexColor('#5B7F3A'),2:HexColor('#A0583A'),3:HexColor('#3E5A6E')}
C_MM={1:HexColor('#B5CE95'),2:HexColor('#E7B79B'),3:HexColor('#A9BFCC')}; APO=HexColor('#7A8691')
MES=['','ene','feb','mar','abr','may','jun','jul','ago','set','oct','nov','dic']
fd=lambda d:f'{d.day:02d}/{d.month:02d}/{d.year}'; fds=lambda d:f'{d.day} {MES[d.month]}'
money=lambda x:f'{x:,.2f}'
W,H=landscape(A4); M=26
D0=dt.date(2026,7,13); D1=dt.date(2026,11,30); ND=(D1-D0).days+1

def txt(c,x,y,t,f='R',s=8,col=INK,al='l'):
    c.setFont(f,s); c.setFillColor(col); {'l':c.drawString,'r':c.drawRightString,'c':c.drawCentredString}[al](x,y,str(t))
def fit(t,f,s,w):
    if sw(t,f,s)<=w: return t
    while t and sw(t+'…',f,s)>w: t=t[:-1]
    return t+'…'
def header(c,sub):
    c.setFillColor(DARK); c.rect(M,H-M-44,W-2*M,44,stroke=0,fill=1)
    txt(c,M+12,H-M-19,'CRONOGRAMA DE EJECUCIÓN — AVANCE EJECUTADO','B',15,white)
    txt(c,M+12,H-M-34,'Mejoramiento de servicios de apoyo al desarrollo productivo en alpacas y llamas · Municipalidad Distrital de Pallpata','R',8.5,HexColor('#E9E3DA'))
    info=[('CUI','2675289'),('RESIDENTE','Ing. Justo'),('PERIODO',f'{fd(D0)} – {fd(CORTE)}'),('HOJA',sub)]
    x=W-M-12
    for k,v in reversed(info):
        w=max(sw(v,'B',8.5),sw(k,'R',6))+14; x-=w
        txt(c,x,H-M-18,k,'R',6,HexColor('#BFB5A8')); txt(c,x,H-M-30,v,'B',8.5,white)
def footer(c,n):
    c.setStrokeColor(LINE); c.setLineWidth(0.5); c.line(M,M+11,W-M,M+11)
    txt(c,M,M+2,f'Obra Camélidos · CUI 2675289 · Avance ejecutado por la residencia del Ing. Justo (desde el 13/07/2026) · Fuentes: cronograma de noviembre v23 y PTO.xlsx','R',6.8,MUTE)
    txt(c,W-M,M+2,f'Hoja {n} de 2','B',6.8,MUTE,'r')

def hoja1(c):
    header(c,'1 · Gantt')
    top=H-M-44-18
    txt(c,M,top+4,'PROGRAMA EJECUTADO POR META Y FRENTE DE TRABAJO','B',10)
    txt(c,W-M,top+4,'E.T. = expediente técnico  ·  M.M. = mayores metrados  ·  MOD = módulo de cobertizo','I',7.2,MUTE,'r')
    cols=[('FRENTE',96),('TIPO',26),('MÓDULOS',56),('N°',18),('INICIO',46),('FIN',46)]
    LW=sum(w for _,w in cols); x0=M+LW+6; x1=W-M; dw=(x1-x0)/ND; X=lambda d:x0+(d-D0).days*dw
    y=top-6
    # meses
    d=D0
    while d<=D1:
        nx=dt.date(2026+(d.month==12),d.month%12+1,1); e=min(nx,D1+dt.timedelta(1))
        c.setFillColor(DARK); c.rect(X(d),y-14,X(e)-X(d),14,stroke=0,fill=1); c.setStrokeColor(white); c.line(X(d),y-14,X(d),y)
        txt(c,(X(d)+X(e))/2,y-10,{7:'JULIO',8:'AGOSTO',9:'SETIEMBRE',10:'OCTUBRE',11:'NOVIEMBRE'}[d.month],'B',8,white,'c'); d=nx
    c.setFillColor(LIGHT); c.rect(x0,y-26,x1-x0,12,stroke=0,fill=1)
    d=D0
    while d<=D1:
        if d.weekday()==0: txt(c,X(d)+1.5,y-23,str(d.day),'R',6.3,MUTE)
        d+=dt.timedelta(1)
    c.setFillColor(DARK); c.rect(M,y-26,LW,26,stroke=0,fill=1)
    xx=M
    for n,w in cols: txt(c,xx+3,y-16,n,'B',6.8,white); xx+=w
    y-=26; ytop=y
    rows=[]
    for m in (1,2,3):
        rows.append(('m',m)); rows+=[('t',t) for t in tr if t['meta']==m]
    rows.append(('s','APOYO E HITOS')); rows.append(('a',apoyo)); rows+=[('h',h) for h in hitos]
    RH=min(17,(y-(M+118))/len(rows)); ybot=y-RH*len(rows)
    d=D0
    while d<=D1:
        if d in fer: c.setFillColor(FER); c.rect(X(d),ybot,dw,ytop-ybot,stroke=0,fill=1)
        elif d.weekday()==6: c.setFillColor(WEEK); c.rect(X(d),ybot,dw,ytop-ybot,stroke=0,fill=1)
        if d.day==1: c.setStrokeColor(HexColor('#9D948A')); c.setLineWidth(0.7); c.line(X(d),ybot,X(d),ytop)
        elif d.weekday()==0: c.setStrokeColor(HexColor('#E2DCD3')); c.setLineWidth(0.3); c.line(X(d),ybot,X(d),ytop)
        d+=dt.timedelta(1)
    for k,o in rows:
        y-=RH; ty=y+RH/2-2.6
        if k in('m','s'):
            c.setFillColor(MHEAD); c.rect(M,y,W-2*M,RH,stroke=0,fill=1)
            if k=='m':
                txt(c,M+4,ty,f'META {o} · {META[o].upper()}','B',8.2)
                txt(c,M+LW-4,ty,f'{mmods[o]} módulos ejecutados','B',7.5,C_ET[o],'r')
            else: txt(c,M+4,ty,o,'B',8.2)
            continue
        c.setStrokeColor(HexColor('#E5DFD7')); c.setLineWidth(0.3); c.line(M,y,W-M,y)
        if k in('t','a'):
            t=o; vals=[t['resp'].replace('Frente obra · ',''),'Apoyo' if k=='a' else t['tipo'],f"MOD{t['a']}–MOD{t['b']}",t['n'],fd(t['ini']),fd(t['fin'])]
            xx=M
            for i,((_,w),v) in enumerate(zip(cols,vals)): txt(c,xx+3,ty,fit(str(v),'B' if i==0 else 'R',7.5,w-4),'B' if i==0 else 'R',7.5); xx+=w
            col=APO if k=='a' else (C_ET if t['tipo']=='E.T.' else C_MM)[t['meta']]
            bx0,bx1=X(t['ini']),X(t['fin'])+dw
            c.setFillColor(col); c.roundRect(bx0,y+3,bx1-bx0,RH-6,1.8,stroke=0,fill=1)
            lab=(f"{t['n']} cerchas" if k=='a' else f"{t['n']} {t['tipo']}")+f" · {fds(t['ini'])} – {fds(t['fin'])}"
            lc=INK if col in C_MM.values() else white
            if sw(lab,'B',6.8)+6<bx1-bx0: txt(c,(bx0+bx1)/2,ty+0.3,lab,'B',6.8,lc,'c')
            elif bx1+4+sw(lab,'B',6.8)<x1: txt(c,bx1+3,ty+0.3,lab,'B',6.8)
            else: txt(c,bx0-3,ty+0.3,lab,'B',6.8,INK,'r')
        else:
            f,tx=o; txt(c,M+3,ty,fit(tx,'R',7.5,LW-50),'R',7.5); txt(c,M+LW-4,ty,fd(f),'R',7.5,INK,'r')
            xc=X(f)+dw/2; yc=y+RH/2; s=4.2
            c.setFillColor(INK); p=c.beginPath(); p.moveTo(xc,yc+s); p.lineTo(xc+s,yc); p.lineTo(xc,yc-s); p.lineTo(xc-s,yc); p.close(); c.drawPath(p,stroke=0,fill=1)
            txt(c,xc+7,ty,fit(tx,'R',7,x1-xc-10),'R',7)
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.rect(M,ybot,W-2*M,ytop+26-ybot,stroke=1,fill=0); c.line(x0-3,ybot,x0-3,ytop+26)
    xc=X(CORTE)+dw/2; c.setStrokeColor(HexColor('#C8412B')); c.setLineWidth(1.2); c.line(xc,ybot,xc,ytop)
    c.setFillColor(HexColor('#C8412B')); c.roundRect(xc-30,ybot-12,60,10,2,stroke=0,fill=1); txt(c,xc,ybot-9.5,'CORTE '+fds(CORTE),'B',6.5,white,'c')
    # leyenda
    y=ybot-14; x=M
    txt(c,x,y,'LEYENDA','B',7); x+=40
    for n,col in [('Meta 1 · E.T.',C_ET[1]),('Meta 1 · M.M.',C_MM[1]),('Meta 2 · E.T.',C_ET[2]),('Meta 2 · M.M.',C_MM[2]),('Meta 3 · E.T.',C_ET[3]),('Meta 3 · M.M.',C_MM[3]),('Apoyo',APO),('Feriado',FER),('Domingo',WEEK)]:
        c.setFillColor(col); c.rect(x,y-1,12,7,stroke=0,fill=1); txt(c,x+15,y,n,'R',7); x+=15+sw(n,'R',7)+10
    # tarjetas resumen
    yb=M+20; h=62; gap=8; w=(W-2*M-3*gap)/4
    c.setFillColor(DARK); c.roundRect(M,yb,w,h,4,stroke=0,fill=1)
    txt(c,M+10,yb+h-14,f'VALORIZADO EJECUTADO (13 jul – {fds(CORTE)})','B',6.8,HexColor('#BFB5A8'))
    txt(c,M+10,yb+h-36,f'S/ {money(VT)}','B',15,white)
    txt(c,M+10,yb+10,f'{VT/CD:.1%} del costo directo (S/ {money(CD)})','R',7,HexColor('#E9E3DA'))
    for i,m in enumerate((1,2,3)):
        xx=M+(i+1)*(w+gap)
        c.setFillColor(LIGHT); c.roundRect(xx,yb,w,h,4,stroke=0,fill=1)
        txt(c,xx+10,yb+h-14,f'META {m} · {META[m].upper()}','B',6.8,MUTE)
        txt(c,xx+10,yb+h-36,f'{mmods[m]} módulos','B',15,C_ET[m]); txt(c,xx+w-10,yb+h-34,f'S/ {money(vmeta[m])}','B',8.5,INK,'r')
        et=sum(t['n'] for t in tr if t['meta']==m and t['tipo']=='E.T.'); mm=mmods[m]-et
        txt(c,xx+10,yb+10,f'E.T.: {et} mód.  ·  M.M.: {mm} mód.','R',7.2)
    footer(c,1)

def hoja2(c):
    header(c,'2 · Valorizado')
    y=H-M-44-18
    txt(c,M,y+4,'VALORIZACIÓN EJECUTADA POR PARTIDA Y MES (S/)','B',10)
    txt(c,W-M,y+4,'Monto = N° de módulos ejecutados en el mes × parcial por módulo del PTO','I',7.2,MUTE,'r')
    cols=[('ITEM',50,'l'),('DESCRIPCIÓN',200,'l'),('UND.',24,'c'),('META',26,'c')]+[(p,66,'r') for p in PL]+[('TOTAL EJECUTADO',86,'r'),('% DEL TOTAL',52,'r')]
    TW=sum(w for _,w,_ in cols); x0=M+(W-2*M-TW)/2
    rows=[[p['item'],p['desc'],p['und'],p['meta']]+[money(v) if v else '—' for v in p['val']]+[money(sum(p['val'])),f'{sum(p["val"])/VT:.1%}'] for p in part]
    rows.append(['','TOTAL EJECUTADO','','']+[money(v) for v in vmes]+[money(VT),'100.0%'])
    acc=0; ac=[]
    for v in vmes: acc+=v; ac.append(acc)
    rows.append(['','Acumulado','','']+[money(v) for v in ac]+['',''])
    rows.append(['','% acumulado respecto al costo directo','','']+[f'{v/CD:.1%}' for v in ac]+['',''])
    rh=min(12.2,(y-6-(M+128))/(len(rows)+1))
    y-=6
    c.setFillColor(DARK); c.rect(x0,y-rh-2,TW,rh+2,stroke=0,fill=1)
    xx=x0
    for n,w,al in cols:
        txt(c,{'l':xx+4,'r':xx+w-4,'c':xx+w/2}[al],y-rh+1.5,n,'B',6.8,white,al); xx+=w
    y-=rh+2
    for i,row in enumerate(rows):
        tot=i>=len(rows)-3
        if tot: c.setFillColor(MHEAD); c.rect(x0,y-rh,TW,rh,stroke=0,fill=1)
        elif i%2: c.setFillColor(LIGHT); c.rect(x0,y-rh,TW,rh,stroke=0,fill=1)
        xx=x0
        for (n,w,al),v in zip(cols,row):
            f='B' if tot else 'R'
            txt(c,{'l':xx+4,'r':xx+w-4,'c':xx+w/2}[al],y-rh+3,fit(str(v),f,7.3,w-6),f,7.3,MUTE if (al=='l' and n=='ITEM') else INK,al); xx+=w
        y-=rh
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.rect(x0,y,TW,(len(rows)+1)*rh+2,stroke=1,fill=0)
    # resumen por meta + firmas
    y-=16
    txt(c,x0,y,'RESUMEN POR META','B',8.5)
    y-=4; cw=[('META',200,'l'),('MÓDULOS',50,'r')]+[(p,66,'r') for p in PL]+[('TOTAL (S/)',86,'r')]
    xx=x0
    c.setFillColor(DARK); c.rect(x0,y-11,sum(w for _,w,_ in cw),11,stroke=0,fill=1)
    for n,w,al in cw: txt(c,{'l':xx+4,'r':xx+w-4}[al],y-8,n,'B',6.8,white,al); xx+=w
    y-=11
    for m in (1,2,3):
        vals=[f'{m} · {META[m]}',mmods[m]]+[money(sum(p['val'][k] for p in part if p['meta']==m)) or '—' for k in range(5)]+[money(vmeta[m])]
        vals=[v if v!='0.00' else '—' for v in vals]
        xx=x0
        for (n,w,al),v in zip(cw,vals): txt(c,{'l':xx+4,'r':xx+w-4}[al],y-8.5,str(v),'R',7.3,INK,al); xx+=w
        c.setStrokeColor(HexColor('#E5DFD7')); c.line(x0,y-11,x0+sum(w for _,w,_ in cw),y-11); y-=11
    footer(c,2)

c=canvas.Canvas(OUT,pagesize=landscape(A4))
c.setTitle('Cronograma de ejecución – avance ejecutado – Obra Camélidos CUI 2675289'); c.setAuthor('Residencia de obra')
hoja1(c); c.showPage(); hoja2(c); c.showPage(); c.save(); print('ok', VT, mmods)

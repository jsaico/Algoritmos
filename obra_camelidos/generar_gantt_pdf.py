import datetime as dt, openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth as sw

XLSX='obra_camelidos/Cronograma_Ejecucion_Camelidos.xlsx'
OUT='obra_camelidos/Cronograma_Gantt_Camelidos_v24.pdf'
FD='/usr/share/fonts/truetype/crosextra/'
pdfmetrics.registerFont(TTFont('R',FD+'Carlito-Regular.ttf')); pdfmetrics.registerFont(TTFont('B',FD+'Carlito-Bold.ttf'))
pdfmetrics.registerFont(TTFont('I',FD+'Carlito-Italic.ttf'))

wb=openpyxl.load_workbook(XLSX,data_only=True)
D,P,V,R=wb['Datos'],wb['Presupuesto'],wb['Valorizado'],wb['Resumen']
todate=lambda v: v.date() if isinstance(v,dt.datetime) else v
CORTE=todate(R['C5'].value); NMOD=R['C8'].value; CD=R['C10'].value; CMOD=R['C9'].value
tramos=[]
for r in range(6,41):
    if D.cell(r,1).value is None: continue
    g=lambda c:D.cell(r,c).value
    tramos.append(dict(meta=g(1),resp=g(3),tipo=g(4),n=g(5),a=g(6),b=g(7),ini=todate(g(8)),fin=todate(g(9)),est=g(10),per=g(11),obs=g(13)))
apoyo=[dict(meta=D.cell(45,1).value,act=D.cell(45,2).value,resp=D.cell(45,3).value,n=D.cell(45,4).value,a=D.cell(45,5).value,b=D.cell(45,6).value,ini=todate(D.cell(45,7).value),fin=todate(D.cell(45,8).value))]
hitos=[dict(n=D.cell(r,1).value,f=todate(D.cell(r,2).value),t=D.cell(r,3).value) for r in range(50,53)]
fer=[todate(D.cell(r,15).value) for r in range(6,21) if D.cell(r,15).value]
fer_n={todate(D.cell(r,15).value):D.cell(r,16).value for r in range(6,21) if D.cell(r,15).value}
partidas=[]
for r in range(5,60):
    it=P.cell(r,1).value
    if it is None or P.cell(r,3).value is None: continue
    partidas.append(dict(item=it,desc=P.cell(r,2).value.strip(),und=P.cell(r,3).value,meta=P.cell(r,7).value,
                         mu=P.cell(r,4).value,pu=P.cell(r,5).value,pm=P.cell(r,6).value,mt=P.cell(r,8).value,tot=P.cell(r,9).value))
PER=['Anterior','Jul','Ago','Set','Oct','Nov','Pendiente']
mods={m:[R.cell(13+m,3+k).value for k in range(7)] for m in (1,2,3)}
modexe={m:R.cell(13+m,12).value for m in (1,2,3)}
val={m:[R.cell(19+m,3+k).value for k in range(7)] for m in (1,2,3)}
valexe={m:R.cell(19+m,12).value for m in (1,2,3)}
# valorizado por partida
vpart={}
for r in range(6,60):
    it=V.cell(r,1).value
    if it and V.cell(r,3).value: vpart[it]=[V.cell(r,16+k).value for k in range(10)]

# ---------------- estilo ----------------
INK=HexColor('#26221F'); MUTE=HexColor('#6E665E'); LINE=HexColor('#CFC8BF'); LIGHT=HexColor('#F4F1EC')
DARK=HexColor('#3F3731'); SEC=HexColor('#5E534A'); MHEAD=HexColor('#E6E0D7')
WEEK=HexColor('#EFEBE5'); FER=HexColor('#F6E3BC'); RED=HexColor('#C8412B')
C_ET={1:HexColor('#5B7F3A'),2:HexColor('#A0583A'),3:HexColor('#3E5A6E')}
C_MM={1:HexColor('#B5CE95'),2:HexColor('#E7B79B'),3:HexColor('#A9BFCC')}
PROG=HexColor('#F2C14E'); APO=HexColor('#7A8691')
META={1:'Trazo, replanteo y excavaciones',2:'Muros',3:'Techos'}
MES=['','ene','feb','mar','abr','may','jun','jul','ago','set','oct','nov','dic']
fd=lambda d:f'{d.day:02d}/{d.month:02d}/{d.year}' if d else '—'
fds=lambda d:f'{d.day} {MES[d.month]}' if d else ''
money=lambda x:f'{x:,.2f}'
W,H=landscape(A3); M=30
D0=dt.date(2026,7,1); D1=dt.date(2026,11,30); ND=(D1-D0).days+1
TOTAL_PAGES=4

def fit(t,font,size,w):
    if sw(t,font,size)<=w: return t
    while t and sw(t+'…',font,size)>w: t=t[:-1]
    return t+'…'
def txt(c,x,y,t,font='R',size=8,color=INK,align='l'):
    c.setFont(font,size); c.setFillColor(color)
    {'l':c.drawString,'r':c.drawRightString,'c':c.drawCentredString}[align](x,y,str(t))

def header(c,page,subt):
    # banda superior
    c.setFillColor(DARK); c.rect(M,H-M-58,W-2*M,58,stroke=0,fill=1)
    txt(c,M+16,H-M-25,'CRONOGRAMA DE EJECUCIÓN DE OBRA','B',20,white)
    txt(c,M+16,H-M-44,'Mejoramiento de servicios de apoyo al desarrollo productivo en alpacas y llamas · Municipalidad Distrital de Pallpata','R',10.5,HexColor('#E9E3DA'))
    # cuadro de datos a la derecha
    bx=W-M-360; items=[('CUI','2675289'),('Residente de obra','Ing. Justo'),('Fecha de corte',fd(CORTE)),('Versión','v24 · '+subt)]
    for i,(k,v) in enumerate(items):
        x=bx+(i%2)*180; y=H-M-22-(i//2)*20
        txt(c,x,y+1,k.upper(),'R',6.5,HexColor('#BFB5A8')); txt(c,x,y-9,v,'B',9.5,white)
def footer(c,page):
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.line(M,M+14,W-M,M+14)
    txt(c,M,M+3,'Obra Camélidos · CUI 2675289 · Cronograma de ejecución v24 · Fuentes: Cronograma de noviembre v23 (PDF) y presupuesto PTO.xlsx','R',7.5,MUTE)
    txt(c,W-M,M+3,f'Página {page} de {TOTAL_PAGES}','B',7.5,MUTE,'r')

# ---------------- eje de tiempo ----------------
class Axis:
    def __init__(s,x0,x1): s.x0,s.x1=x0,x1; s.dw=(x1-x0)/ND
    def x(s,d): return s.x0+(d-D0).days*s.dw
def draw_axis(c,ax,ytop):
    # meses
    y=ytop-16; d=D0
    while d<=D1:
        nxt=dt.date(d.year+(d.month==12),d.month%12+1,1); e=min(nxt,D1+dt.timedelta(1))
        x0,x1=ax.x(d),ax.x(e)
        c.setFillColor(DARK); c.rect(x0,y,x1-x0,16,stroke=0,fill=1)
        c.setStrokeColor(white); c.setLineWidth(1); c.line(x0,y,x0,y+16)
        nm={7:'JULIO',8:'AGOSTO',9:'SETIEMBRE',10:'OCTUBRE',11:'NOVIEMBRE'}[d.month]
        txt(c,(x0+x1)/2,y+5,nm+' 2026','B',8.5,white,'c'); d=nxt
    # semanas (lunes)
    y2=y-14; c.setFillColor(LIGHT); c.rect(ax.x0,y2,ax.x1-ax.x0,14,stroke=0,fill=1)
    d=D0
    while d<=D1:
        if d.weekday()==0:
            x=ax.x(d); c.setStrokeColor(LINE); c.setLineWidth(0.5); c.line(x,y2,x,y2+14)
            txt(c,x+2,y2+4,f'L {d.day}','R',6.8,MUTE)
        d+=dt.timedelta(1)
    c.setStrokeColor(LINE); c.rect(ax.x0,y2,ax.x1-ax.x0,30,stroke=1,fill=0)
    return y2
def draw_grid(c,ax,ytop,ybot):
    d=D0
    while d<=D1:
        x=ax.x(d)
        if d in fer: c.setFillColor(FER); c.rect(x,ybot,ax.dw,ytop-ybot,stroke=0,fill=1)
        elif d.weekday()==6: c.setFillColor(WEEK); c.rect(x,ybot,ax.dw,ytop-ybot,stroke=0,fill=1)
        if d.weekday()==0:
            c.setStrokeColor(HexColor('#E2DCD3')); c.setLineWidth(0.4); c.line(x,ybot,x,ytop)
        if d.day==1:
            c.setStrokeColor(HexColor('#9D948A')); c.setLineWidth(0.8); c.line(x,ybot,x,ytop)
        d+=dt.timedelta(1)
def draw_cut(c,ax,ytop,ybot):
    x=ax.x(CORTE)+ax.dw/2
    c.setStrokeColor(RED); c.setLineWidth(1.4); c.line(x,ybot,x,ytop)
    c.setFillColor(RED); c.roundRect(x-34,ybot-14,68,12,3,stroke=0,fill=1)
    txt(c,x,ybot-11,'CORTE '+fds(CORTE),'B',7,white,'c')
def bar(c,ax,ini,fin,y,h,col,label='',prog=False,labcol=white):
    x0=ax.x(ini); x1=ax.x(fin)+ax.dw; w=x1-x0
    c.setFillColor(col); c.roundRect(x0,y,w,h,2,stroke=0,fill=1)
    if prog:
        c.saveState(); p=c.beginPath(); p.roundRect(x0,y,w,h,2); c.clipPath(p,stroke=0,fill=0)
        c.setStrokeColor(HexColor('#FFFFFF')); c.setLineWidth(1.2)
        k=x0-h
        while k<x1:
            c.line(k,y,k+h,y+h); k+=5
        c.restoreState()
        c.setStrokeColor(HexColor('#B8860B')); c.setDash(2,1.5); c.setLineWidth(0.8); c.roundRect(x0,y,w,h,2,stroke=1,fill=0); c.setDash()
    if label:
        size=7.2
        if sw(label,'B',size)+6<=w: txt(c,x0+w/2,y+h/2-2.5,label,'B',size,labcol,'c')
        elif x1+4+sw(label,'B',size)<=ax.x1-2: txt(c,x1+4,y+h/2-2.5,label,'B',size,INK)
        else: txt(c,x0-4,y+h/2-2.5,label,'B',size,INK,'r')
def light(col): return col in C_MM.values() or col==PROG

# ---------------- página 1: frentes ----------------
def page1(c):
    header(c,1,'Frentes de ejecución')
    top=H-M-58-26
    txt(c,M,top+6,'A · PROGRAMACIÓN POR FRENTES DE TRABAJO Y METAS','B',11,INK)
    txt(c,W-M,top+6,'Cada barra = grupo de módulos ejecutado por un frente.  E.T. = expediente técnico (37 mód.)  ·  M.M. = mayores metrados (30 mód.)','I',8,MUTE,'r')
    cols=[('FRENTE / RESPONSABLE',150),('TIPO',30),('MÓDULOS',70),('CANT.',32),('INICIO',52),('FIN',52),('ESTADO',58)]
    LW=sum(w for _,w in cols); ax=Axis(M+LW+8,W-M)
    ytab=top-8
    yax=draw_axis(c,ax,ytab)
    # cabecera tabla izquierda
    c.setFillColor(DARK); c.rect(M,yax,LW,30,stroke=0,fill=1)
    x=M
    for n,w in cols: txt(c,x+4,yax+11,n,'B',7.2,white); x+=w
    rows=[]
    for m in (1,2,3):
        rows.append(('meta',m))
        for t in tramos:
            if t['meta']==m: rows.append(('tramo',t))
    rows.append(('sec','APOYO E HITOS'))
    for a in apoyo: rows.append(('apoyo',a))
    for h_ in hitos: rows.append(('hito',h_))
    RH=19.5; y=yax
    ybot=y-RH*len(rows)
    draw_grid(c,ax,y,ybot)
    for kind,o in rows:
        y-=RH
        if kind in('meta','sec'):
            c.setFillColor(MHEAD); c.rect(M,y,W-2*M,RH,stroke=0,fill=1)
            lab=f'META {o} · {META[o].upper()}' if kind=='meta' else o
            txt(c,M+6,y+6.5,lab,'B',9,INK)
            if kind=='meta':
                txt(c,M+LW-4,y+6.5,f'Avance al corte: {modexe[o]} / {NMOD} módulos ({modexe[o]/NMOD:.0%})','B',8,C_ET[o],'r')
            continue
        c.setStrokeColor(HexColor('#E5DFD7')); c.setLineWidth(0.4); c.line(M,y,W-M,y)
        if kind=='tramo':
            t=o; vals=[t['resp'],t['tipo'],f"MOD{t['a']} – MOD{t['b']}",str(t['n']),fd(t['ini']) if t['ini'] else ('antes jul' if t['est']!='Pendiente' else '—'),fd(t['fin']) if t['fin'] else ('11/07/2026' if t['est']!='Pendiente' else '—'),t['est']]
            x=M
            for i,((_,w),v) in enumerate(zip(cols,vals)):
                f='B' if i==0 else 'R'; col=INK
                if i==6: col={'Ejecutado':C_ET[1],'Programado':HexColor('#A8770A'),'Pendiente':RED}[v]; f='B'
                txt(c,x+4,y+6.5,fit(v,f,8,w-6),f,8,col); x+=w
            if t['ini']:
                colr=PROG if t['est']=='Programado' else (C_ET if t['tipo']=='E.T.' else C_MM)[t['meta']]
                lab=f"{t['n']} {t['tipo']} · MOD{t['a']}–{t['b']} · {fds(t['ini'])} – {fds(t['fin'])}"
                bar(c,ax,t['ini'],t['fin'],y+4,RH-8,colr,lab,prog=t['est']=='Programado',labcol=INK if light(colr) else white)
            elif t['est']=='Pendiente':
                txt(c,ax.x0+6,y+6.5,f"Pendiente de programar · {t['n']} módulos (MOD{t['a']}–MOD{t['b']})",'I',8,RED)
            else:
                c.setFillColor(HexColor('#B9B3AC')); c.roundRect(ax.x0+1,y+4,ax.x(dt.date(2026,7,11))+ax.dw-ax.x0-1,RH-8,2,stroke=0,fill=1)
                txt(c,ax.x(dt.date(2026,7,12))+3,y+6.5,f"Ejecutado con el Ing. Roy (hasta 11 jul) · {t['n']} módulos",'I',8,MUTE)
        elif kind=='apoyo':
            a=o; vals=[a['resp'],'Apoyo',f"MOD{a['a']} – MOD{a['b']}",str(a['n']),fd(a['ini']),fd(a['fin']),'Ejecutado']
            x=M
            for i,((_,w),v) in enumerate(zip(cols,vals)):
                txt(c,x+4,y+6.5,fit(v,'B' if i in(0,6) else 'R',8,w-6),'B' if i in(0,6) else 'R',8,C_ET[1] if i==6 else INK); x+=w
            bar(c,ax,a['ini'],a['fin'],y+4,RH-8,APO,f"{a['n']} cerchas · {fds(a['ini'])} – {fds(a['fin'])}")
        else:
            h_=o; txt(c,M+4,y+6.5,fit(f"Hito {h_['n']} · {h_['t']}",'R',8,LW-60),'R',8,INK); txt(c,M+LW-6-52-58+4,y+6.5,'',)
            txt(c,M+150+30+70+32+4,y+6.5,fd(h_['f']),'R',8,INK)
            x=ax.x(h_['f'])+ax.dw/2; yc=y+RH/2; s=5
            c.setFillColor(INK); p=c.beginPath(); p.moveTo(x,yc+s); p.lineTo(x+s,yc); p.lineTo(x,yc-s); p.lineTo(x-s,yc); p.close(); c.drawPath(p,stroke=0,fill=1)
            txt(c,x+8,yc-2.6,f"{fds(h_['f'])} · {h_['t']}",'R',7.5,INK) if h_['f']<dt.date(2026,10,15) else txt(c,x-8,yc-2.6,f"{h_['t']} · {fds(h_['f'])}",'R',7.5,INK,'r')
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.rect(M,ybot,W-2*M,yax+30-ybot,stroke=1,fill=0)
    c.line(ax.x0-4,ybot,ax.x0-4,yax+30)
    draw_cut(c,ax,yax+30-30,ybot) if False else draw_cut(c,ax,yax,ybot)
    legend(c,ybot-26)
    kpis(c,ybot-26)
    footer(c,1)

def legend(c,y):
    items=[('Meta 1 · E.T.',C_ET[1]),('Meta 1 · M.M.',C_MM[1]),('Meta 2 · E.T.',C_ET[2]),('Meta 2 · M.M.',C_MM[2]),('Meta 3 · E.T.',C_ET[3]),('Meta 3 · M.M.',C_MM[3]),
           ('Programado',PROG),('Apoyo (cerchas)',APO),('Feriado',FER),('Domingo',WEEK)]
    txt(c,M,y+2,'LEYENDA','B',8,INK); x=M+48
    for n,col in items:
        c.setFillColor(col); c.roundRect(x,y,16,9,1.5,stroke=0,fill=1)
        if n=='Programado':
            c.setStrokeColor(HexColor('#B8860B')); c.setDash(2,1.5); c.setLineWidth(0.6); c.roundRect(x,y,16,9,1.5,stroke=1,fill=0); c.setDash()
        txt(c,x+20,y+1.5,n,'R',7.8,INK); x+=20+sw(n,'R',7.8)+14
    c.setStrokeColor(RED); c.setLineWidth(1.4); c.line(x+4,y-1,x+4,y+10); txt(c,x+9,y+1.5,'Fecha de corte','R',7.8,INK)
    x+=80
    c.setFillColor(INK); p=c.beginPath(); xc=x+5; yc=y+4.5; s=4.5
    p.moveTo(xc,yc+s); p.lineTo(xc+s,yc); p.lineTo(xc,yc-s); p.lineTo(xc-s,yc); p.close(); c.drawPath(p,stroke=0,fill=1)
    txt(c,x+14,y+1.5,'Hito','R',7.8,INK)
def kpis(c,y):
    # tarjetas de avance
    w=150; x=W-M-3*w-2*8; yb=y-62
    if yb<M+22: return
    for i,m in enumerate((1,2,3)):
        xx=x+i*(w+8)
        c.setFillColor(LIGHT); c.roundRect(xx,yb,w,54,4,stroke=0,fill=1)
        txt(c,xx+10,yb+40,f'META {m} · {META[m].upper()}','B',7,MUTE)
        txt(c,xx+10,yb+19,f'{modexe[m]} / {NMOD}','B',18,C_ET[m]); txt(c,xx+w-10,yb+21,f'{modexe[m]/NMOD:.1%}','B',11,INK,'r')
        c.setFillColor(HexColor('#DDD6CC')); c.rect(xx+10,yb+9,w-20,4,stroke=0,fill=1)
        c.setFillColor(C_ET[m]); c.rect(xx+10,yb+9,(w-20)*modexe[m]/NMOD,4,stroke=0,fill=1)
    tot=sum(valexe.values())
    xx=x-200-8
    c.setFillColor(DARK); c.roundRect(xx,yb,200,54,4,stroke=0,fill=1)
    txt(c,xx+10,yb+40,'AVANCE VALORIZADO EJECUTADO','B',7,HexColor('#BFB5A8'))
    txt(c,xx+10,yb+19,f'S/ {money(tot)}','B',15,white); txt(c,xx+190,yb+21,f'{tot/CD:.1%}','B',11,PROG,'r')
    txt(c,xx+10,yb+7,f'de S/ {money(CD)} de costo directo ({NMOD} módulos)','R',7,HexColor('#E9E3DA'))
    # nota
    notes=['Notas:','• El 13/07 el Ing. Justo asume la residencia (memorándum). Los 19 primeros módulos de las metas 1 y 2 se ejecutaron con el Ing. Roy.',
           '• Techos de El Chino (MOD46–63, 4 al 13 nov) figuran como PROGRAMADO hasta su confirmación.',
           '• Quedan 4 módulos completos (MOD64–67) sin programar en las tres metas.']
    for i,n in enumerate(notes): txt(c,M,yb+44-i*11,n,'B' if i==0 else 'R',7.8,INK if i else MUTE)

# ---------------- página 2: partidas ----------------
def page2(c):
    header(c,2,'Partidas del presupuesto')
    top=H-M-58-26
    txt(c,M,top+6,'B · CRONOGRAMA POR PARTIDAS DEL PRESUPUESTO (PTO)','B',11,INK)
    txt(c,W-M,top+6,f'Metrado y parcial total = valor por módulo × {NMOD} módulos.  Las barras muestran los tramos en que se ejecuta la meta de cada partida.','I',8,MUTE,'r')
    cols=[('ITEM',48),('DESCRIPCIÓN',178),('UND.',26),('METRADO TOTAL',58),('PARCIAL TOTAL (S/)',74),('AVANCE',40)]
    LW=sum(w for _,w in cols); ax=Axis(M+LW+8,W-M)
    yax=draw_axis(c,ax,top-8)
    c.setFillColor(DARK); c.rect(M,yax,LW,30,stroke=0,fill=1)
    x=M
    for i,(n,w) in enumerate(cols):
        if i>=3: txt(c,x+w-4,yax+11,n,'B',7.2,white,'r')
        else: txt(c,x+4,yax+11,n,'B',7.2,white)
        x+=w
    rows=[]
    for m in (1,2,3):
        rows.append(('meta',m)); rows+= [('p',p) for p in partidas if p['meta']==m]
    RH=min(17.5,(yax-(M+80))/len(rows)); y=yax; ybot=y-RH*len(rows)
    draw_grid(c,ax,y,ybot)
    for kind,o in rows:
        y-=RH
        if kind=='meta':
            c.setFillColor(MHEAD); c.rect(M,y,W-2*M,RH,stroke=0,fill=1)
            sub=sum(p['tot'] for p in partidas if p['meta']==o)
            txt(c,M+6,y+RH/2-3,f'META {o} · {META[o].upper()}','B',9,INK)
            txt(c,M+48+178+26+58+74-4,y+RH/2-3,money(sub),'B',8.5,INK,'r')
            continue
        p=o; c.setStrokeColor(HexColor('#E5DFD7')); c.setLineWidth(0.4); c.line(M,y,W-M,y)
        ex=vpart[p['item']][9]
        vals=[p['item'],p['desc'].capitalize(),p['und'],f"{p['mt']:,.2f}",money(p['tot']),f'{ex:.0%}']
        x=M
        for i,((_,w),v) in enumerate(zip(cols,vals)):
            if i>=3: txt(c,x+w-4,y+RH/2-2.8,v,'R',7.8,INK,'r')
            else: txt(c,x+4,y+RH/2-2.8,fit(v,'R',7.8,w-6),'R',7.8,INK if i else MUTE)
            x+=w
        for t in tramos:
            if t['meta']!=p['meta'] or not t['ini']: continue
            colr=PROG if t['est']=='Programado' else (C_ET if t['tipo']=='E.T.' else C_MM)[t['meta']]
            bar(c,ax,t['ini'],t['fin'],y+3.5,RH-7,colr,prog=t['est']=='Programado')
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.rect(M,ybot,W-2*M,yax+30-ybot,stroke=1,fill=0); c.line(ax.x0-4,ybot,ax.x0-4,yax+30)
    draw_cut(c,ax,yax,ybot)
    # total
    c.setFillColor(DARK); c.rect(M,ybot-18,LW,18,stroke=0,fill=1)
    txt(c,M+6,ybot-12,f'COSTO DIRECTO TOTAL ({NMOD} módulos × S/ {money(CMOD)})','B',8.5,white)
    txt(c,M+48+178+26+58+74-4,ybot-12,money(CD),'B',8.5,white,'r')
    legend(c,ybot-42)
    footer(c,2)

# ---------------- página 3: valorizado ----------------
def table(c,x,y,cols,rows,rh=17,head=DARK,bold_last=0,zebra=True):
    W_=sum(w for _,w,_ in cols)
    c.setFillColor(head); c.rect(x,y-rh,W_,rh,stroke=0,fill=1)
    xx=x
    for n,w,al in cols:
        txt(c,xx+w/2 if al=='c' else (xx+w-5 if al=='r' else xx+5),y-rh+5.5,n,'B',7.6,white,al); xx+=w
    y-=rh
    for i,row in enumerate(rows):
        bold=i>=len(rows)-bold_last
        if bold: c.setFillColor(MHEAD); c.rect(x,y-rh,W_,rh,stroke=0,fill=1)
        elif zebra and i%2: c.setFillColor(LIGHT); c.rect(x,y-rh,W_,rh,stroke=0,fill=1)
        xx=x
        for (n,w,al),v in zip(cols,row):
            f='B' if bold else 'R'
            txt(c,xx+w/2 if al=='c' else (xx+w-5 if al=='r' else xx+5),y-rh+5.5,fit(str(v),f,8,w-8),f,8,INK,al); xx+=w
        y-=rh
    c.setStrokeColor(LINE); c.setLineWidth(0.6); c.rect(x,y,W_,(len(rows)+1)*rh,stroke=1,fill=0)
    return y
PL=['Anterior','Julio','Agosto','Setiembre','Octubre','Nov. (prog.)','Pendiente']
def page3(c):
    header(c,3,'Cronograma valorizado')
    y=H-M-58-26
    txt(c,M,y+6,'C · CRONOGRAMA VALORIZADO POR META Y PERIODO','B',11,INK)
    cw=[('META',190,'l')]+[(p,92,'r') for p in PL]+[('TOTAL',100,'r'),('EJECUTADO AL CORTE',110,'r')]
    rows=[]
    for m in (1,2,3): rows.append([f'{m} · {META[m]}']+[money(v) if v else '—' for v in val[m]]+[money(sum(val[m])),money(valexe[m])])
    tp=[sum(val[m][k] for m in (1,2,3)) for k in range(7)]
    rows.append(['TOTAL (S/)']+[money(v) for v in tp]+[money(sum(tp)),money(sum(valexe.values()))])
    rows.append(['% del costo directo']+[f'{v/CD:.1%}' for v in tp]+['100.0%',f'{sum(valexe.values())/CD:.1%}'])
    acc=[]; s=0
    for v in tp: s+=v; acc.append(s)
    rows.append(['Acumulado (S/)']+[money(v) for v in acc]+['',''])
    rows.append(['% acumulado']+[f'{v/CD:.1%}' for v in acc]+['',''])
    y=table(c,M,y-6,cw,rows,bold_last=4)
    y-=26
    txt(c,M,y+6,'MÓDULOS POR META Y PERIODO','B',11,INK)
    cw2=[('META',190,'l')]+[(p,92,'r') for p in PL]+[('TOTAL',100,'r'),('EJECUTADOS / AVANCE',110,'r')]
    rows2=[[f'{m} · {META[m]}']+[str(v) if v else '—' for v in mods[m]]+[str(sum(mods[m])),f'{modexe[m]}  ·  {modexe[m]/NMOD:.1%}'] for m in (1,2,3)]
    y=table(c,M,y-6,cw2,rows2)
    # gráfico curva S
    y-=30; gx=M; gw=640; gh=y-(M+150); gy=M+150
    txt(c,gx,y+6,'CURVA DE AVANCE VALORIZADO','B',11,INK)
    cx0=gx+60; cx1=gx+gw-50; cy0=gy+26; cy1=y-14
    import math
    mx=math.ceil(max(tp)*1.08/50000)*50000
    c.setStrokeColor(LINE); c.setLineWidth(0.5)
    for k in range(6):
        yy=cy0+(cy1-cy0)*k/5; c.line(cx0,yy,cx1,yy)
        txt(c,cx0-6,yy-3,f'{mx*k/5/1000:,.0f} mil','R',7,MUTE,'r'); txt(c,cx1+6,yy-3,f'{k*20}%','R',7,MUTE)
    n=7; bw=(cx1-cx0)/n
    pts=[]
    for k in range(n):
        bx=cx0+k*bw+bw*0.22; bh=(cy1-cy0)*tp[k]/mx
        colr=PROG if k==5 else (HexColor('#C9C2B8') if k==6 else SEC)
        c.setFillColor(colr); c.rect(bx,cy0,bw*0.56,bh,stroke=0,fill=1)
        txt(c,cx0+k*bw+bw/2,cy0-11,PL[k],'R',7.5,INK,'c')
        txt(c,cx0+k*bw+bw/2,cy0-21,f'S/ {tp[k]:,.0f}','B',7.5,INK,'c')
        pts.append((cx0+k*bw+bw/2,cy0+(cy1-cy0)*acc[k]/CD))
    c.setStrokeColor(RED); c.setLineWidth(1.6)
    for a,b in zip(pts,pts[1:]): c.line(*a,*b)
    for k,(px,py) in enumerate(pts):
        c.setFillColor(white); c.circle(px,py,3,stroke=1,fill=1)
        lab=f'{acc[k]/CD:.0%}'; lw=sw(lab,'B',8)+6
        c.setFillColor(white); c.roundRect(px+5,py+3,lw,11,2,stroke=0,fill=1); txt(c,px+8,py+5.5,lab,'B',8,RED)
    yl=gy-8
    c.setFillColor(SEC); c.rect(cx0,yl,10,7,stroke=0,fill=1); txt(c,cx0+14,yl,'Valorizado del periodo (ejecutado)','R',7.5,INK)
    c.setFillColor(PROG); c.rect(cx0+170,yl,10,7,stroke=0,fill=1); txt(c,cx0+184,yl,'Programado','R',7.5,INK)
    c.setFillColor(HexColor('#C9C2B8')); c.rect(cx0+250,yl,10,7,stroke=0,fill=1); txt(c,cx0+264,yl,'Pendiente','R',7.5,INK)
    c.setStrokeColor(RED); c.line(cx0+320,yl+3.5,cx0+336,yl+3.5); txt(c,cx0+340,yl,'% acumulado (eje derecho)','R',7.5,INK)
    # notas y supuestos
    nx=gx+gw+30; nw=W-M-nx; ny=y+6
    txt(c,nx,ny,'CRITERIOS Y SUPUESTOS','B',11,INK)
    notes=[f'Costo directo por módulo: S/ {money(CMOD)} (PTO.xlsx). {NMOD} módulos = 37 E.T. + 30 M.M.',
           'Cada módulo se valoriza en el mes en que se ejecuta la meta correspondiente.',
           'Meta 1: obras provisionales, flete, EPP, trazo, excavaciones, relleno, solado, cimientos y zapatas.',
           'Meta 2: sobrecimientos, columnas, vigas, muro caravista y rotura de briquetas.',
           'Meta 3: tijerales, apoyos, correas y cobertura de aluzinc. Las 30 cerchas de Edgar se valorizan dentro de la partida 01.11.01.',
           '"Anterior" corresponde a los 19 módulos de trazo y 19 de muros ejecutados con el Ing. Roy (hasta 11/07).',
           'Noviembre: techos de El Chino (18 mód., 4–13 nov) en estado programado, pendiente de confirmación.',
           'Pendiente: MOD64–MOD67 en las tres metas, sin responsable ni fechas.']
    yy=ny-18
    from reportlab.lib.utils import simpleSplit
    for i,n_ in enumerate(notes):
        lines=simpleSplit(n_,'R',8.3,nw-14)
        txt(c,nx,yy,f'{i+1}.','B',8.3,MUTE)
        for ln in lines: txt(c,nx+14,yy,ln,'R',8.3,INK); yy-=11
        yy-=4
    # firmas
    fy=M+52; fw=200
    for i,(t1,t2) in enumerate([('Elaborado por','Residente de obra · Ing. Justo'),('Revisado por','Supervisor / Inspector de obra'),('Aprobado por','Sub Gerencia de Infraestructura')]):
        fx=M+60+i*(fw+170)
        c.setStrokeColor(INK); c.setLineWidth(0.7); c.line(fx,fy,fx+fw,fy)
        txt(c,fx+fw/2,fy-11,t1,'B',8,INK,'c'); txt(c,fx+fw/2,fy-21,t2,'R',7.5,MUTE,'c')
    footer(c,3)

# ---------------- página 4: valorizado por partida ----------------
def page4(c):
    header(c,4,'Valorizado por partida')
    y=H-M-58-26
    txt(c,M,y+6,'D · VALORIZACIÓN PROGRAMADA POR PARTIDA Y PERIODO (S/)','B',11,INK)
    cw=[('ITEM',54,'l'),('DESCRIPCIÓN',236,'l'),('META',34,'c')]+[(p,84,'r') for p in PL]+[('TOTAL',88,'r'),('EJECUTADO',88,'r'),('AV.',36,'r')]
    rows=[]
    for p in partidas:
        v=vpart[p['item']]; tot=sum(v[:7])
        rows.append([p['item'],p['desc'].capitalize(),p['meta']]+[money(x) if x else '—' for x in v[:7]]+[money(tot),money(v[8]),f'{v[9]:.0%}'])
    tp=[sum(vpart[p['item']][k] for p in partidas) for k in range(7)]
    ex=sum(vpart[p['item']][8] for p in partidas)
    rows.append(['','TOTAL',''] + [money(x) for x in tp]+[money(sum(tp)),money(ex),f'{ex/CD:.0%}'])
    rh=min(16,(y-6-(M+30))/(len(rows)+1))
    table(c,M,y-6,cw,rows,rh=rh,bold_last=1)
    footer(c,4)

c=canvas.Canvas(OUT,pagesize=landscape(A3))
c.setTitle('Cronograma de ejecución – Obra Camélidos CUI 2675289 (v24)'); c.setAuthor('Residencia de obra'); c.setSubject('Cronograma de ejecución valorizado')
for f in (page1,page2,page3,page4): f(c); c.showPage()
c.save(); print('ok')

"""Reglas de presupuesto v2: cálculos independientes de Streamlit."""
import math
from collections import defaultdict
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ACTIVIDADES=['GEORREFERENCIACION','MAPEO GEOESPACIAL','GABINETE']
VALORES={'GPS GNSS Diferencial':20000,'Estación Total 1"':15000,
         'Poligonal (estación + kit)':15000,'Drone Matrice 4E':15000,
         'Drone MATRICE 400 + L2':100000,'Drone MATRICE 400 + L3':100000,
         'Nivel Electrónico':10000,'Nivel Óptico':3000}
AUTO={'Jefe de Proyecto','Asistente HSQ','Topógrafo A','Auxiliar de Topografía','Ayudante Locales'}

def hitos(tipo,metrado):
    if tipo=='AREA': return 2,2
    if metrado<=0: return 0,0
    intervalos=math.floor((metrado+1e-9)/5)
    rem=metrado-intervalos*5
    geo=2*(1+intervalos+(1 if rem>=1-1e-9 else 0))
    return geo,max(0,math.ceil(metrado*4-1e-9)-geo)

def linea(grupo,actividad,recurso,cantidad,periodos,tarifa,unidad='día'):
    return dict(Grupo=grupo,Actividad=actividad,Recurso=recurso,Cantidad=float(cantidad),Periodos=float(periodos),Unidad=unidad,Tarifa=float(tarifa))

def importe(r): return r['Cantidad']*r['Periodos']*r['Tarifa']

def calcular(cfg,tab,c):
    filas=[]; expos={}; diario=0
    auto_daily=defaultdict(float); continuous_daily=defaultdict(float); unique_ingreso={}; unique_exam={}
    def add(*args):
        r=linea(*args)
        if r['Cantidad'] and r['Periodos']: filas.append(r)
    g,t=hitos(cfg['tipo'],cfg['metrado'])
    if cfg['manual_hitos']: g,t=cfg['geo'],cfg['topo']
    add('Monumentación',ACTIVIDADES[0],next(k for k in c if k.startswith('Hito de Punto Geod')),g,1,cfg['tarifa_geo'],'punto')
    add('Monumentación',ACTIVIDADES[0],'Hito topográfico 30×30×30',t,1,cfg['tarifa_topo'],'punto')
    for name in ['IGN: Ficha','IGN: Data','IGN:  Certificación b']:
        add('Certificación',ACTIVIDADES[0],name,cfg['certificados'],1,c[name],'punto')
    for r in tab['equipos']:
        name=r['Equipo']; q=r['Cantidad']; days=r['Días']; act=r['Actividad']
        if not q or not days: continue
        if name=='Poligonal (estación + kit)': parts=['Estación Total 1"','Kit Poligonal : 02 Tribach']
        elif name.startswith('Drone MATRICE 400 +'): parts=['Drone MATRICE 400','Sensor LiDAR DJI ZenMuse '+name[-2:]]
        else: parts=[name]
        for part in parts: add('Equipos',act,part,q,days,c[part])
        if name.startswith('Drone'):
            pilot='Piloto Drone Matrice 4E' if name=='Drone Matrice 4E' else 'Piloto Drone MATRICE 400'
            explicit=any(p['Cargo']==pilot and p.get('Etapa')==r.get('Etapa') and p['Días remunerados']>0 for p in tab['personal'])
            if not explicit:
                add('Personal',act,pilot,q,days,c[pilot]); diario+=q*c[pilot]
        if name in VALORES: expos[name]=expos.get(name,0)+q
    for r in tab['personal']:
        role=r['Cargo']; q=r['Personas']; act=r['Actividad']
        if not q: continue
        add('Personal',act,role,q,r['Días remunerados'],c[role])
        automatico=r.get('Origen')=='Automático'
        if r['Días remunerados']>0:
            if automatico:
                if r.get('Etapa') in ['Gestión de campo','Gabinete']: continuous_daily[role]=max(continuous_daily[role],q*c[role])
                else: auto_daily[r.get('Etapa','')]+=q*c[role]
            else: diario+=q*c[role]
        for name,col in [('Alojamiento','Noches alojamiento'),('Viáticos - Personal Lima','Días viáticos')]:
            add('Permanencia',act,name,q,r[col],c[name])
        if r['Ingreso'] and role in AUTO:
            if automatico: unique_ingreso[role]=max(unique_ingreso.get(role,0),q)
            else:
                for name in ['EPP Básico','Inducción','SCTR y Vida Ley']: add('Ingreso',act,name,q,1,c[name],'ingreso')
        if r['Examen médico']:
            if automatico: unique_exam[role]=max(unique_exam.get(role,0),q)
            else: add('Ingreso',act,'Examen Medico',q,1,c['Examen Medico'],'ingreso')
    diario+=max(auto_daily.values(),default=0)+sum(continuous_daily.values())
    for role,q in unique_ingreso.items():
        for name in ['EPP Básico','Inducción','SCTR y Vida Ley']: add('Ingreso','GEORREFERENCIACION',name+' — '+role,q,1,c[name],'persona única')
    for role,q in unique_exam.items(): add('Ingreso','GEORREFERENCIACION','Examen Medico — '+role,q,1,c['Examen Medico'],'persona única')
    q=cfg['vehiculos']
    for name in [cfg['vehiculo'],'Conductor','Combustible']:
        add('Transporte',cfg['actividad_logistica'],name,q,cfg['dias_vehiculo'],150 if name=='Combustible' else c[name])
    if cfg['dias_vehiculo']>0: diario+=q*c['Conductor']
    combis=cfg.get('combis',0)
    if combis:
        add('Transporte',cfg['actividad_logistica'],'MiniBus 16 Pasajeros c/Conductor',combis,cfg['dias_vehiculo'],c['MiniBus 16 Pasajeros c/Conductor'])
        add('Transporte',cfg['actividad_logistica'],'Combustible combi',combis,cfg['dias_vehiculo'],150)
        if cfg['dias_vehiculo']>0: diario+=combis*c['Conductor']
    for r in tab['extras']: add('Adicionales',r['Actividad'],r['Concepto'],r['Cantidad'],r['Periodos'],r['Tarifa'],'editable')
    operativo=sum(importe(r) for r in filas)
    planilla=linea('PLAME / planilla','GABINETE','Planilla asignada (monto / 24 × días apoyo)',1,cfg['dias_planilla'],cfg['planilla']/24)
    riesgos=[]
    for r in tab['exposicion']:
        if r['Aplicar']:
            riesgos.append(linea('Riesgos',r['Actividad'],'Daño/robo: '+r['Equipo'],r['Cantidad'],r['Días exposición'],r['Valor equipo']*r['Cobertura %']/100/30))
    riesgos.append(linea('Riesgos',cfg['actividad_riesgos'],'Paradas / esperas: remuneración diaria',1,cfg['espera'],diario))
    riesgos.append(linea('Riesgos',cfg['actividad_riesgos'],'Inducciones adicionales: 50% remuneración diaria',1,cfg['induccion'],diario*.5))
    for r in tab['otros_riesgos']:
        riesgos.append(linea('Riesgos',r['Actividad'],r['Riesgo'],1,r['Porcentaje']/100,operativo,'porcentaje operativo'))
    return filas,riesgos,planilla,diario,expos

def exportar(cfg,filas,riesgos,planilla):
    w=Workbook(); s=w.active; s.title='Resumen'
    d=w.create_sheet('Detalle'); r=w.create_sheet('Riesgos'); p=w.create_sheet('Planilla')
    headers=['Grupo','Actividad','Recurso','Cantidad','Periodos','Unidad','Tarifa S/','Subtotal S/']
    for sheet,records in [(d,filas),(r,riesgos),(p,[planilla])]:
        sheet.append(headers)
        for n,row in enumerate(records,2): sheet.append([row[k] for k in ['Grupo','Actividad','Recurso','Cantidad','Periodos','Unidad','Tarifa']]+[f'=D{n}*E{n}*G{n}'])
        sheet.freeze_panes='D2'; sheet.auto_filter.ref=f'A1:H{max(1,sheet.max_row)}'
    def suma(sheet): return f"SUM('{sheet.title}'!H2:H{max(2,sheet.max_row)})"
    s.append(['Concepto','Valor']); s.append(['Proyecto',cfg['proyecto']]); s.append(['Tipo',cfg['tipo']]); s.append(['Metrado (km o ha)',cfg['metrado']])
    s.append(['Costo operativo','='+suma(d)]); s.append(['Riesgos','='+suma(r)]); s.append(['PLAME / planilla asignada','='+suma(p)])
    s.append(['Utilidad sobre costo (tasa)',cfg['utilidad']/100]); s['B8'].number_format='0.00%'
    s.append(['Utilidad S/','=SUM(B5:B7)*B8']); s.append(['Costo total / cotización antes de impuestos','=SUM(B5:B7)+B9'])
    s.append(['Costo por km' if cfg['tipo']=='LINEAL' else 'Costo por ha','=IF(B4>0,B10/B4,0)'])
    s.append(['Impuestos','No incluidos']); s.append(['Planilla mensual S/',cfg['planilla']]); s.append(['Días laborales base',24]); s.append(['Días apoyo',cfg['dias_planilla']])
    s.append(['Actividad','Operativo','Riesgos','Planilla','Total antes de utilidad'])
    for n,act in enumerate(ACTIVIDADES,17):
        s.append([act,f'=SUMIF(Detalle!B2:B{max(2,d.max_row)},A{n},Detalle!H2:H{max(2,d.max_row)})',f'=SUMIF(Riesgos!B2:B{max(2,r.max_row)},A{n},Riesgos!H2:H{max(2,r.max_row)})',f'=SUMIF(Planilla!B2:B2,A{n},Planilla!H2:H2)',f'=SUM(B{n}:D{n})'])
    # Las tasas de riesgos adicionales deben seguir a las ediciones del detalle.
    for n,row in enumerate(riesgos,2):
        if row['Unidad']=='porcentaje operativo': r.cell(n,7,'=Resumen!B5')
        if row['Recurso'].startswith('Paradas') or row['Recurso'].startswith('Inducciones'):
            terms=[f"Detalle!D{i}*Detalle!G{i}" for i,x in enumerate(filas,2) if x['Grupo']=='Personal' or x['Recurso']=='Conductor']
            r.cell(n,7,'='+('+'.join(terms) or '0'))
            if row['Recurso'].startswith('Inducciones'): r.cell(n,7,'=('+('+'.join(terms) or '0')+')*0.5')
    p['G2']='=Resumen!B13/Resumen!B14'; p['E2']='=Resumen!B15'
    x=w.create_sheet('Reglas y notas')
    for line in [cfg['notas'],cfg['alcance'],'Utilidad = recargo sobre costo operativo + riesgos + planilla; no margen sobre venta.','Hitos LINEAL: par inicial + pares cada 5 km + par final si remanente >= 1 km. Topográficos: máximo(0, techo(km/0.25) - geodésicos).','Robo/daño: valor × cantidad × cobertura% / 30 × días exposición. Es provisión comercial, no prima de seguro ni probabilidad.','Paradas e inducción: solo remuneración diaria, incluidos pilotos y conductores. Sin alojamiento ni equipos.','Editar Detalle y Riesgos. Los porcentajes de daño/robo están incorporados en la tarifa; sus bases completas están en el respaldo JSON.','El Excel recalcula en Excel de escritorio. Sus cambios no se reimportan a la aplicación.']: x.append([line])
    for sheet in w:
        sheet.freeze_panes=sheet.freeze_panes or 'B2'
        for cell in sheet[1]: cell.fill=PatternFill('solid',fgColor='123449'); cell.font=Font(color='FFFFFF',bold=True)
        for col in sheet.columns:
            sheet.column_dimensions[get_column_letter(col[0].column)].width=min(65,max(18,max(len(str(c.value or '')) for c in col)+2))
        for row in sheet:
            for cell in row:
                if cell.data_type=='f' or isinstance(cell.value,float):
                    if cell.number_format=='General': cell.number_format='#,##0.00'
        if sheet==x:
            sheet.column_dimensions['A'].width=100
            for row in sheet:
                row[0].alignment=Alignment(wrap_text=True,vertical='top'); sheet.row_dimensions[row[0].row].height=55
    out=BytesIO(); w.save(out); return out.getvalue()

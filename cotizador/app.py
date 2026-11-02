from pathlib import Path
from datetime import datetime
from collections import defaultdict
import json,math,hashlib,os
import pandas as pd
import streamlit as st
from app_v2 import catalog
from motor import ACTIVIDADES,VALORES,hitos,calcular,importe,exportar
from planificacion import planificar,dotacion,movilidad,jornadas
from plantilla_excel import generar,TEMPLATE
from acceso import verificar_acceso

ROOT=Path(__file__).resolve().parent
SOURCE=ROOT.parent/'COSTOS UNITARIOS.xlsx'
SAVED=ROOT/'proyectos_guardados'
PAGES=['1. Alcance','2. Equipos','3. Personal','4. Logística','5. Riesgos y reporte']
OPTIONS=['GPS GNSS Diferencial','Poligonal (estación + kit)','Nivel Electrónico','Nivel Óptico','Estación Total 1"','Drone Matrice 4E','Drone MATRICE 400 + L2','Drone MATRICE 400 + L3','CPU Workstation','Laser Scanner TRIMBLE TX6','Radio UHV - VHF']

def load_data(data):
    if data.get('schema') not in (2,3,4): raise ValueError('Respaldo incompatible')
    cfg=data['config']; tables=data['tables']
    if data['schema']==2:
        cfg=dict(cfg,servicios=[],dias_geo=1,recorrido=0,igv=18,version_migrada=True)
        cfg['manual_importado']=True
    for name in ['equipos','personal']:
        for r in tables.get(name,[]): r.setdefault('Etapa','Adicional'); r.setdefault('Origen','Manual')
    st.session_state['config']=cfg; st.session_state['tables']=tables
    st.session_state['costos_guardados']=data['catalog']
    for key in list(st.session_state):
        if key.startswith(('w_','ed_')): del st.session_state[key]
    st.session_state['rev']=st.session_state.get('rev',0)+1
    st.session_state.pop('plan_signature',None)
    st.session_state.pop('modal_done',None)
    st.session_state['modal_done']=True
    st.session_state['preservar_tablas']=True
    st.session_state['cargado_signature']=(cfg['tipo'],cfg['metrado'])
    st.session_state['bases_apoyo']={'Gestión de campo|'+act:days for act,days in cfg.get('campo_por_actividad',{}).items()}
    st.session_state['bases_apoyo']['Gabinete|GABINETE']=cfg.get('dias_gabinete',cfg.get('duracion',0))

def abrir():
    try: load_data(json.loads((SAVED/st.session_state['archivo']).read_text(encoding='utf-8')))
    except Exception as e: st.session_state['mensaje_error']=str(e)

def nuevo():
    for k in list(st.session_state): del st.session_state[k]

def commit_table(name,key,default):
    changes=st.session_state[key]; data=[dict(x) for x in st.session_state['tables'][name]]
    for idx,patch in changes.get('edited_rows',{}).items(): data[int(idx)].update(patch)
    for row in changes.get('added_rows',[]): data.append(dict(default,**row))
    for r in data:
        for k,val in default.items():
            if r.get(k) is None: r[k]=val
    removed=set(changes.get('deleted_rows',[]))
    st.session_state['tables'][name]=[r for i,r in enumerate(data) if i not in removed]
    st.session_state['rev']=st.session_state.get('rev',0)+1

@st.dialog('Selecciona los trabajos del proyecto')
def elegir_trabajos():
    cfg=st.session_state['config']
    st.caption('Las actividades de georreferenciación se desarrollan consecutivamente. La monumentación se incorpora a partir de los hitos.')
    selected=st.multiselect('Actividades',['GEODESIA','NIVELACION','POLIGONAL','CAPTURA TERRESTRE','CAPTURA AEREA'],default=cfg.get('servicios',[]),key='dialog_servicios')
    terrestre=cfg.get('terrestre','ESTACION TOTAL'); aerea=cfg.get('aerea','FOTOGRAMETRIA')
    if 'CAPTURA TERRESTRE' in selected:
        terrestre=st.radio('Método de captura terrestre',['ESTACION TOTAL','GNSS'],index=['ESTACION TOTAL','GNSS'].index(terrestre),horizontal=True)
    if 'CAPTURA AEREA' in selected:
        aerea=st.radio('Método de captura aérea',['FOTOGRAMETRIA','LIDAR'],index=['FOTOGRAMETRIA','LIDAR'].index(aerea),horizontal=True)
    if st.button('Aplicar selección',type='primary'):
        cfg['servicios']=selected; cfg['terrestre']=terrestre; cfg['aerea']=aerea; cfg.pop('manual_importado',None)
        st.session_state.pop('preservar_tablas',None)
        st.session_state['modal_done']=True; st.rerun()

def main():
    st.set_page_config(page_title='MATIC | Cotizador v4',layout='wide')
    verificar_acceso()
    st.markdown('''<style>
    .stApp{background:#101c2c;color:#f3f7fc}.stApp h1,.stApp h2,.stApp h3{color:#f3f7fc!important}
    .block-container{padding-top:2.5rem;padding-bottom:2rem}header[data-testid="stHeader"]{display:block!important;background:#101c2c;height:2rem!important;min-height:2rem!important}
    header[data-testid="stHeader"] [data-testid="stToolbar"]{display:flex!important}
    [data-testid="stExpandSidebarButton"]{position:fixed!important;top:.15rem;left:.5rem;z-index:1001;display:flex!important;visibility:visible!important}
    [data-testid="stSidebarCollapseButton"]{visibility:visible!important}
    [data-testid="stSidebarCollapsedControl"]{display:flex!important;visibility:visible!important}
    [data-testid="stLayoutWrapper"]:has(>.st-key-cabecera_fija){position:sticky!important;top:2rem;z-index:100;background:#101c2c}
    .st-key-cabecera_fija{background:#101c2c;border-bottom:2px solid #00a6a6;padding:.4rem 0 .6rem}
    .st-key-cabecera_fija h2{font-size:1.4rem;padding:.1rem 0}.st-key-cabecera_fija [data-testid="stVerticalBlock"]{gap:.25rem}
    .st-key-cabecera_fija .stButton>button{padding:.25rem .4rem;min-height:32px}.st-key-cabecera_fija .stButton>button p{font-size:.8rem;white-space:nowrap}
    [data-testid="stNumberInput"]{max-width:155px}[data-testid="stNumberInput"] input{padding:.25rem .35rem}
    [data-testid="stMetric"]{background:#173449;padding:8px;border-left:3px solid #00a6a6;border-radius:5px}
    [data-testid="stMetricValue"]{font-size:1.3rem}.stButton>button{border-color:#008e98}
    @media(max-width:700px){[data-testid="stNumberInput"]{max-width:150px}.st-key-cabecera_fija h3{font-size:1.1rem}}
    @media(max-width:640px){
      .block-container{padding-left:.7rem;padding-right:.7rem}
      .st-key-cabecera_fija [data-testid="stHorizontalBlock"]{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr));gap:.3rem!important}
      .st-key-cabecera_fija [data-testid="stColumn"]{width:100%!important;min-width:0!important}
      .st-key-cabecera_fija [data-testid="stColumn"]:first-child{grid-column:1 / -1}
      .st-key-cabecera_fija [data-testid="stNumberInput"]{max-width:145px}
      .st-key-cabecera_fija [data-testid="stVerticalBlock"]{gap:.1rem}
    }
    </style>''',unsafe_allow_html=True)
    cfg=st.session_state.setdefault('config',dict(proyecto='Nuevo proyecto',tipo='LINEAL',metrado=1,utilidad=40,servicios=[],dias_geo=1,recorrido=0,certificados=0,planilla=25000,espera=0,induccion=0,notas='',alcance='',igv=18))
    tab=st.session_state.setdefault('tables',{})
    costs,roles=catalog(); costs=st.session_state.get('costos_guardados',costs)
    # Inicializar campos faltantes permite reabrir proyectos anteriores.
    cfg.setdefault('servicios',[]); cfg.setdefault('dias_geo',1); cfg.setdefault('recorrido',0); cfg.setdefault('igv',18)
    cfg.setdefault('terrestre','ESTACION TOTAL'); cfg.setdefault('aerea','FOTOGRAMETRIA'); cfg['dias_semana']=5
    with st.sidebar:
        if st.session_state.get('_acceso_hash') and st.button('Cerrar sesión'):
            st.session_state.pop('_acceso_hash',None); st.rerun()
        st.button('Nuevo proyecto',on_click=nuevo)
        SAVED.mkdir(exist_ok=True)
        paths=sorted([p.name for p in SAVED.glob('*.json') if 'QA_' not in p.name],reverse=True)
        if paths:
            st.selectbox('Proyecto guardado',paths,key='archivo'); st.button('Abrir proyecto',on_click=abrir)
        upload=st.file_uploader('Importar respaldo',type='json')
        if upload and st.button('Abrir respaldo'):
            try: load_data(json.loads(upload.getvalue())); st.rerun()
            except Exception as e: st.error(str(e))
        if 'mensaje_error' in st.session_state: st.error(st.session_state.pop('mensaje_error'))
        if 'costos_guardados' in st.session_state:
            st.caption('Tarifas restauradas del proyecto.')
            if st.button('Leer tarifas actuales'):
                del st.session_state['costos_guardados']; st.rerun()
        st.caption('Guardado local por versiones. La plantilla XLSM original no se modifica.')
    def number(k,label,default,minimum=0,decimal=False):
        value=cfg.get(k,default)
        widget='w_'+k+(f'_{cfg["tipo"]}_{cfg["metrado"]}' if k in ['geo','topo'] else '')
        if k in ['dias_vehiculo','dias_planilla']: widget+=f'_{cfg.get("duracion",0)}'
        if decimal: result=st.number_input(label,min_value=float(minimum),value=float(value),format='%g',step=0.1,key=widget)
        else: result=st.number_input(label,min_value=int(minimum),value=int(math.ceil(value)),format='%d',step=1,key=widget)
        cfg[k]=result; return result
    def text(k,label,default=''):
        cfg[k]=st.text_input(label,value=cfg.get(k,default),key='w_'+k); return cfg[k]
    def select(k,label,opts,default=None):
        val=cfg.get(k,default if default is not None else opts[0]); val=val if val in opts else opts[0]
        cfg[k]=st.selectbox(label,opts,index=opts.index(val),key='w_'+k); return cfg[k]
    def reset_widget(k): st.session_state.pop('w_'+k,None)
    def edit(name,default,config):
        tab.setdefault(name,[])
        if not tab[name]: data=pd.DataFrame(columns=list(default))
        else: data=pd.DataFrame(tab[name],columns=list(default))
        key=f'ed_{name}_{st.session_state.get("rev",0)}'
        st.data_editor(data,num_rows='fixed' if name=='exposicion' else 'dynamic',column_config=config,key=key,hide_index=True,width='stretch',on_change=commit_table,args=(name,key,default))
    with st.container(key='cabecera_fija'):
        st.subheader('MATIC · Cotizador de proyectos')
        cols=st.columns([1.7,1.15,1,1,1.45])
        with cols[0]: text('proyecto','Proyecto')
        with cols[1]: select('tipo','Tipo',['LINEAL','AREA'])
        with cols[2]: number('metrado','Longitud km' if cfg['tipo']=='LINEAL' else 'Área ha',1,0.001,True)
        with cols[3]: number('utilidad','Utilidad %',40)
        with cols[4]:
            st.button('Actualizar',key='actualizar')
            save=st.button('Guardar proyecto',type='primary')
        page=st.segmented_control('Secciones',PAGES,default=PAGES[0],key='navegacion',label_visibility='collapsed') or PAGES[0]
    if 'mensaje_ok' in st.session_state: st.success(st.session_state.pop('mensaje_ok'))
    # Los hitos se escriben directamente; el usuario puede editarlos hasta que cambie el metrado.
    signature=(cfg['tipo'],cfg['metrado'])
    if signature!=st.session_state.get('hitos_signature'):
        if not cfg.pop('manual_importado',False) and st.session_state.pop('cargado_signature',None)!=signature:
            cfg['geo'],cfg['topo']=hitos(*signature)
        st.session_state['hitos_signature']=signature
        for k in ['geo','topo']: reset_widget(k)
    cfg['manual_hitos']=True
    geokey=next(k for k in costs if k.startswith('Hito de Punto Geod'))
    cfg.setdefault('tarifa_geo',costs[geokey]); cfg.setdefault('tarifa_topo',costs['Hito de Punto Topográfico (30x30x30)'])
    open_dialog=False
    if page==PAGES[0]:
        col=st.columns(4)
        with col[0]: number('geo','Hitos geodésicos',cfg['geo'])
        with col[1]: number('topo','Hitos topográficos',cfg['topo'])
        with col[2]: number('tarifa_geo','Tarifa geodésico S/',cfg['tarifa_geo'],decimal=True)
        with col[3]: number('tarifa_topo','Tarifa topográfico S/',cfg['tarifa_topo'],decimal=True)
        st.caption('Cantidades calculadas desde el metrado y editables. Cambiar longitud/área restablece el cálculo automático. LINEAL: par inicial + cada 5 km + final si remanente ≥1 km; topográficos cada 250 m menos geodésicos.')
        if st.button('Elegir actividades'): open_dialog=True
        st.write('Actividades: '+(', '.join(cfg['servicios']) or 'Solo monumentación; elige actividades para observación y redes.'))
        col=st.columns(3)
        with col[0]: number('certificados','Puntos IGN',0)
        with col[1]: number('dias_geo','Días Geodesia',1,1)
        with col[2]:
            if cfg['tipo']=='AREA': number('recorrido','Recorrido equivalente km',0,0,True)
        with st.expander('Datos de la propuesta y entregables'):
            for k,label in [('codigo','Código cotización'),('cliente','Cliente'),('servicio','Servicio'),('ruc','RUC cliente'),('direccion','Dirección cliente'),('atencion','Atención'),('email','E-mail cliente')]: text(k,label)
            cfg['alcance']=st.text_area('Entregables y metodología',value=cfg.get('alcance',''),key='w_alcance')
    km=cfg['metrado'] if cfg['tipo']=='LINEAL' else cfg.get('recorrido',0)
    incomplete=cfg['tipo']=='AREA' and km<=0
    if incomplete: st.warning('Falta recorrido equivalente en km: los rendimientos de captura y gabinete fueron definidos en km. Guarda un borrador y completa ese dato para exportar.')
    stages,auto_eq,auto_staff=planificar(cfg['geo'],cfg['topo'],cfg['servicios'],km,cfg['dias_geo'],cfg['terrestre'],cfg['aerea'],True,5)
    psig=json.dumps([cfg['geo'],cfg['topo'],cfg['servicios'],km,cfg['dias_geo'],cfg['terrestre'],cfg['aerea'],'apoyo-v4'],sort_keys=True)
    if psig!=st.session_state.get('plan_signature'):
        preserve=st.session_state.pop('preservar_tablas',False)
        if not preserve:
            for name,auto in [('equipos',auto_eq),('personal',auto_staff)]: tab[name]=auto+[r for r in tab.get(name,[]) if r.get('Origen','Manual')!='Automático']
        st.session_state['plan_signature']=psig; st.session_state['rev']=st.session_state.get('rev',0)+1
    # Ajustes de jornadas en tablas se reflejan en el cronograma por etapa.
    for s in stages:
        if s['Etapa']=='Gabinete': continue
        days=[r['Días'] for r in tab.get('equipos',[]) if r.get('Etapa')==s['Etapa']]+[r['Días remunerados'] for r in tab.get('personal',[]) if r.get('Etapa')==s['Etapa']]
        if days: s['Días']=max(days)
    cursor=1
    for s in stages:
        if s['Etapa']=='Gabinete': continue
        s['Inicio']=cursor; s['Fin']=cursor+s['Días']-1; cursor+=s['Días']
    manual_days=max([r['Días'] for r in tab.get('equipos',[]) if r.get('Origen')!='Automático' and r['Actividad']!='GABINETE']+[r['Días remunerados'] for r in tab.get('personal',[]) if r.get('Origen')!='Automático' and r['Actividad']!='GABINETE']+[0])
    cfg['dias_campo']=cursor-1+manual_days
    campo_act={act:sum(s['Días'] for s in stages if s.get('Actividad')==act) for act in ['GEORREFERENCIACION','MAPEO GEOESPACIAL']}
    if manual_days:
        manual_act='MAPEO GEOESPACIAL' if any(r.get('Origen')!='Automático' and r['Actividad']=='MAPEO GEOESPACIAL' for r in tab.get('personal',[])+tab.get('equipos',[])) else 'GEORREFERENCIACION'
        campo_act[manual_act]+=manual_days
    cfg['campo_por_actividad']=campo_act
    base_gab=cfg['dias_campo']+jornadas(km/10*5)
    prev=st.session_state.get('bases_apoyo',{})
    for r in tab.get('personal',[]):
        if r.get('Origen')!='Automático' or r.get('Etapa') not in ['Gestión de campo','Gabinete']: continue
        target=base_gab if r['Etapa']=='Gabinete' else campo_act.get(r['Actividad'],0)
        for k in ['Días remunerados','Noches alojamiento','Días viáticos']:
            if r['Etapa']=='Gabinete' and k!='Días remunerados': continue
            if r[k]==prev.get(r['Etapa']+'|'+r['Actividad'],r[k]): r[k]=target
    st.session_state['bases_apoyo']={'Gestión de campo|'+act:days for act,days in campo_act.items()}
    st.session_state['bases_apoyo']['Gabinete|GABINETE']=base_gab
    gabinete_days=max([base_gab]+[r['Días remunerados'] for r in tab.get('personal',[]) if r['Actividad']=='GABINETE'])
    for s in stages:
        if s['Etapa']=='Gabinete': s.update(Inicio=1,Fin=gabinete_days,Días=gabinete_days,Teórico=cfg['dias_campo']+km/10*5)
    cfg['duracion']=max(1,cfg['dias_campo'],gabinete_days)
    cfg['dias_gabinete']=gabinete_days
    st.caption(f'Campo: {cfg["dias_campo"]:g} jornadas consecutivas. Gabinete: {gabinete_days:g} jornadas = campo + procesamiento (10 km/semana de 5 días). Duración total: {cfg["duracion"]:g} jornadas, sin sumar otra vez el campo.')
    ac=st.column_config.SelectboxColumn(options=ACTIVIDADES,required=True,default=ACTIVIDADES[0])
    def nc(label,value=0): return st.column_config.NumberColumn(label,min_value=0,default=value,format='%d',required=True)
    orig=st.column_config.TextColumn(disabled=True)
    if page==PAGES[1]:
        edit('equipos',dict(Actividad=ACTIVIDADES[0],Etapa='Adicional',Equipo=OPTIONS[0],Cantidad=1,Días=1,Origen='Manual'),{'Actividad':ac,'Equipo':st.column_config.SelectboxColumn(options=OPTIONS,default=OPTIONS[0],required=True),'Cantidad':nc('Cantidad',1),'Días':nc('Días',1),'Origen':orig})
        st.caption('Las actividades seleccionadas generan recursos. Puedes ajustar cantidades/días. Si modificas hitos, selección o recorrido se regeneran las filas automáticas; las filas manuales se conservan.')
    if page==PAGES[2]:
        edit('personal',{'Actividad':ACTIVIDADES[0],'Etapa':'Adicional','Cargo':'Topógrafo A','Personas':1,'Días remunerados':1,'Noches alojamiento':1,'Días viáticos':1,'Ingreso':True,'Examen médico':False,'Origen':'Manual'},{'Actividad':ac,'Cargo':st.column_config.SelectboxColumn(options=roles,default='Topógrafo A',required=True),**{k:nc(k,1) for k in ['Personas','Días remunerados','Noches alojamiento','Días viáticos']},'Ingreso':st.column_config.CheckboxColumn(default=True),'Examen médico':st.column_config.CheckboxColumn(default=False),'Origen':orig})
        st.caption('Una cuadrilla automática se reutiliza entre etapas. EPP/inducción/seguros se cobran una vez por máxima cantidad de cada cargo, no por cada fila. Remuneración y permanencia corresponden a sus días de trabajo.')
    eq=tab.get('equipos',[]); staff=tab.get('personal',[])
    people=dotacion(staff,eq); trucks,combis=movilidad(people)
    cfg['personas_transporte']=people; cfg['asientos']=4; cfg['vehiculos']=trucks; cfg['combis']=combis
    cfg.setdefault('vehiculo','Camioneta MINA Pick-up 4x4 - Hilux'); cfg.setdefault('actividad_logistica','GEORREFERENCIACION'); cfg.setdefault('actividad_riesgos','GEORREFERENCIACION')
    # Los días logísticos y planilla siguen duración solo mientras no sean ajustados manualmente.
    for k,base in [('dias_vehiculo',cfg['dias_campo']),('dias_planilla',cfg['duracion'])]:
        if k not in cfg or cfg.get(k)==st.session_state.get('prev_'+k):
            cfg[k]=base; reset_widget(k)
        st.session_state['prev_'+k]=base
    if page==PAGES[3]:
        st.info(f'{people} personas de campo → {trucks} camioneta(s) + {combis} combi(s). 4 pasajeros útiles por camioneta; gabinete excluido.')
        st.caption('Se suma la dotación única asignada a campo: máximo automático por cargo reutilizado + personal manual de Georreferenciación/Mapeo + pilotos. No se repiten topógrafos entre etapas ni se cuentan conductores o gabinete como pasajeros. Más de 8 personas: camioneta + combi de 16 pasajeros; más combis si se supera esa capacidad.')
        col=st.columns(3)
        with col[0]: select('vehiculo','Camioneta',['Camioneta MINA Pick-up 4x4 - Hilux','Camioneta Pick-up 4x4 - Hilux'])
        with col[1]: number('dias_vehiculo','Días transporte',cfg['duracion'])
        with col[2]: select('actividad_logistica','Asignar transporte',ACTIVIDADES)
        st.caption('Combi: tarifa MiniBus 16 Pasajeros c/Conductor + S/150 combustible/día. Camioneta: alquiler + conductor + S/150 combustible/día.')
        col=st.columns(3)
        with col[0]: number('planilla','Planilla mensual S/',25000)
        with col[1]: number('dias_planilla','Días apoyo planilla',cfg['duracion'])
        with col[2]: number('igv','IGV plantilla %',18)
    # Exposición por asignación: días operativos, jamás duración calendario.
    exposures=[]
    for i,r in enumerate(eq):
        if r['Equipo'] in VALORES and r['Cantidad']>0 and r['Días']>0:
            ident=f'{i}:{r["Equipo"]}:{r.get("Etapa","")}'
            old=next((x for x in tab.get('exposicion',[]) if x.get('ID')==ident),{})
            exposures.append({'ID':ident,'Aplicar':old.get('Aplicar',True),'Actividad':r['Actividad'],'Equipo':r['Equipo'],'Cantidad':r['Cantidad'],'Valor equipo':old.get('Valor equipo',VALORES[r['Equipo']]),'Cobertura %':old.get('Cobertura %',50),'Días exposición':r['Días']})
    tab['exposicion']=exposures
    tab.setdefault('extras',[]); tab.setdefault('otros_riesgos',[])
    cfg.setdefault('planilla',25000); cfg.setdefault('espera',0); cfg.setdefault('induccion',0); cfg.setdefault('notas',''); cfg.setdefault('alcance','')
    if page==PAGES[4]:
        st.subheader('Adicionales y riesgos')
        edit('extras',dict(Actividad=ACTIVIDADES[0],Concepto='Otros gastos',Cantidad=0,Periodos=1,Tarifa=0),{'Actividad':ac,'Cantidad':nc('Cantidad'),'Periodos':nc('Periodos',1),'Tarifa':st.column_config.NumberColumn(min_value=0,format='%g')})
        if exposures:
            edit('exposicion',exposures[0],{'ID':st.column_config.TextColumn(disabled=True),'Equipo':st.column_config.TextColumn(disabled=True),'Actividad':orig,'Cantidad':nc('Cantidad'),'Valor equipo':st.column_config.NumberColumn(min_value=0,format='%g'),'Cobertura %':st.column_config.NumberColumn(min_value=0,max_value=100,format='%g'),'Días exposición':st.column_config.NumberColumn(disabled=True,format='%d')})
        st.caption('Daño/robo = valor × 50% editable /30 × cantidad × días operativos de la asignación. Para cambiar días, edita Equipos. No usa los días totales del proyecto.')
        col=st.columns(3)
        with col[0]: number('espera','Días espera',0)
        with col[1]: number('induccion','Jornadas inducción',0)
        with col[2]: select('actividad_riesgos','Asignar riesgos',ACTIVIDADES)
        edit('otros_riesgos',{'Actividad':ACTIVIDADES[0],'Riesgo':'Otro riesgo','Porcentaje':0},{'Actividad':ac,'Porcentaje':st.column_config.NumberColumn(min_value=0,max_value=100,format='%g')})
        cfg['notas']=st.text_area('Otras consideraciones',value=cfg.get('notas',''),key='w_notas')
    for name,records in tab.items():
        for r in records:
            for k,v in r.items():
                if isinstance(v,(int,float)) and (not math.isfinite(v) or v<0): st.error(f'{name}: valor no válido en {k}'); st.stop()
    filas,risks,pay,daily,_=calcular(cfg,tab,costs)
    op=sum(map(importe,filas)); risk=sum(map(importe,risks)); pl=importe(pay); profit=(op+risk+pl)*cfg['utilidad']/100; total=op+risk+pl+profit
    # Preparar recursos antes de construir tanto la descarga como el guardado.
    unique=defaultdict(float); manuals=defaultdict(float)
    for r in staff:
        if r['Días remunerados']<=0: continue
        if r.get('Origen')=='Automático': unique[r['Cargo']]=max(unique[r['Cargo']],r['Personas'])
        else: manuals[r['Cargo']]+=r['Personas']
    for role,n in manuals.items(): unique[role]+=n
    for r in eq:
        if r['Equipo'].startswith('Drone') and r['Cantidad'] and r['Días']:
            role='Piloto Drone Matrice 4E' if r['Equipo']=='Drone Matrice 4E' else 'Piloto Drone MATRICE 400'
            if not any(p['Cargo']==role and p.get('Etapa')==r.get('Etapa') and p['Días remunerados']>0 for p in staff): unique[role]+=r['Cantidad']
    if cfg['dias_vehiculo']>0: unique['Conductor']+=trucks+combis
    cfg['personas_unicas']=sum(r['Cantidad'] for r in filas if r['Grupo']=='Ingreso' and r['Recurso'].startswith('EPP'))
    cfg['lista_personal']=[f'{n:g} {role}' for role,n in unique.items() if n]
    cfg['lista_equipos']=[f'{r["Cantidad"]:g} {r["Equipo"]} ({r["Días"]:g} días)' for r in eq if r['Cantidad'] and r['Días']]
    if trucks: cfg['lista_equipos'].append(f'{trucks} camioneta(s)')
    if combis: cfg['lista_equipos'].append(f'{combis} combi(s) de 16 pasajeros')
    if page==PAGES[4]:
        st.divider(); st.subheader('Resumen de cotización')
        for col,label,val in zip(st.columns(5),['Operativo','Riesgos','Planilla','Utilidad','Antes de IGV'],[op,risk,pl,profit,total]): col.metric(label,f'S/ {val:,.0f}')
        st.write(f'Por {"km" if cfg["tipo"]=="LINEAL" else "ha"}: S/ {total/cfg["metrado"]:,.0f} · IGV {cfg["igv"]:g}%: S/ {total*cfg["igv"]/100:,.0f} · Total con IGV: S/ {total*(1+cfg["igv"]/100):,.0f}')
        detail=pd.DataFrame(filas+risks+[pay]); detail['Subtotal S/']=[importe(r) for r in filas+risks+[pay]]
        st.dataframe(detail,column_config={k:st.column_config.NumberColumn(format='%g') for k in ['Cantidad','Periodos','Tarifa','Subtotal S/']},hide_index=True,width='stretch')
        st.dataframe(detail.pivot_table(index='Actividad',columns='Grupo',values='Subtotal S/',aggfunc='sum',fill_value=0),width='stretch')
        if not incomplete: st.download_button('Descargar propuesta en plantilla XLSM',generar(cfg,filas,risks,pay,stages),file_name='COTIZACION_MATIC.xlsm',mime='application/vnd.ms-excel.sheet.macroEnabled.12')
    with st.expander('Cronograma y reutilización de cuadrillas'):
        st.dataframe(pd.DataFrame(stages),hide_index=True,width='stretch')
        st.caption('Campo consecutivo: monumentación → geodesia → nivelación → poligonal → captura terrestre → captura aérea, según selección. Gabinete trabaja desde el inicio y termina después del procesamiento adicional; no se duplica la duración de campo. Jefe/HSQ cubren campo. Gabinete no lleva viáticos/alojamiento por defecto.')
    snapshot={'schema':4,'config':cfg,'tables':tab,'catalog':costs,'saved_at':datetime.now().isoformat(),'source_sha256':hashlib.sha256(SOURCE.read_bytes()).hexdigest(),'template_sha256':hashlib.sha256(TEMPLATE.read_bytes()).hexdigest(),'stages':stages,'rules':'v4'}
    if os.environ.get('MATIC_CLOUD_MODE')=='1':
        st.caption('Prueba gratuita: los archivos guardados en el servidor pueden borrarse al reiniciar. Descarga tu respaldo para conservar el proyecto hasta configurar almacenamiento permanente.')
    st.download_button('Descargar proyecto completo para reabrir',json.dumps(snapshot,ensure_ascii=False,indent=2),file_name='PROYECTO_MATIC.json',mime='application/json')
    if save:
        stem=datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'_'+(''.join(c if c.isalnum() else '_' for c in cfg['proyecto'])[:60] or 'Proyecto')
        target=SAVED/(stem+'.json'); target.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2),encoding='utf-8')
        if not incomplete:
            try:
                (SAVED/(stem+'.xlsm')).write_bytes(generar(cfg,filas,risks,pay,stages))
                st.success('Proyecto y propuesta XLSM guardados como nueva versión. Puedes reabrirlos desde la barra lateral.')
            except Exception as e: st.error('El respaldo JSON se guardó, pero falló la propuesta: '+str(e))
        else: st.success('Borrador guardado. Falta recorrido para completar la propuesta.')
    if open_dialog or not st.session_state.get('modal_done'):
        elegir_trabajos()

if __name__=='__main__': main()

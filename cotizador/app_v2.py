from pathlib import Path
from datetime import datetime
import json, math, hashlib
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from motor import ACTIVIDADES,VALORES,hitos,calcular,importe,exportar

ROOT=Path(__file__).resolve().parent
SOURCE=ROOT.parent/'COSTOS UNITARIOS.xlsx'
SAVED=ROOT/'proyectos_guardados'

def catalog():
    sheet=load_workbook(SOURCE,data_only=True).active
    costs={}; roles=[]; section=''
    for row in sheet:
        name=str(row[0].value or '').strip(); value=row[1].value
        if name in ['Personal','Equipos Topográficos','Logística']: section=name
        if isinstance(value,(int,float)) and name:
            if name not in costs: costs[name]=float(value)
            if section=='Personal': roles.append(name)
    return costs,roles

def restore(data):
    if data.get('schema')!=2: raise ValueError('Se requiere respaldo versión 2.')
    for key in list(st.session_state):
        if key.startswith(('in_','ed_','table_')): del st.session_state[key]
    for key,val in data['config'].items(): st.session_state['in_'+key]=val
    for key,val in data['tables'].items(): st.session_state['table_'+key]=val
    st.session_state['catalog_saved']=data['catalog']
    st.session_state['rev']=st.session_state.get('rev',0)+1
    st.session_state['exposure_signature']=json.dumps([(r['Equipo'],r['Cantidad'],r['Días']) for r in data['tables']['equipos']],sort_keys=True)

def commit_table(key,widget,defaults):
    changes=st.session_state[widget]
    data=[dict(r) for r in st.session_state['table_'+key]]
    for idx,patch in changes.get('edited_rows',{}).items(): data[int(idx)].update(patch)
    for row in changes.get('added_rows',[]): data.append(dict(defaults,**row))
    deleted=set(changes.get('deleted_rows',[]))
    st.session_state['table_'+key]=[r for i,r in enumerate(data) if i not in deleted]
    st.session_state['rev']=st.session_state.get('rev',0)+1

def load_saved():
    try: restore(json.loads((SAVED/st.session_state['saved_choice']).read_text(encoding='utf-8')))
    except Exception as e: st.session_state['load_error']=str(e)

def main():
    st.set_page_config(page_title='MATIC | Cotizador v2',layout='wide')
    st.markdown('''<style>.stApp{background:#101c2c;color:#f3f7fc}.stApp h1,.stApp h2,.stApp h3{color:#f3f7fc!important}[data-testid="stMetric"]{background:#173449;border-left:4px solid #00a6a6;padding:12px;border-radius:7px}.stButton>button{border-color:#008e98}</style>''',unsafe_allow_html=True)
    st.title('MATIC · Cotizador de proyectos')
    st.caption('Versión 2 · soles · estimaciones por reglas · catálogo original preservado')
    costs,roles=catalog()
    with st.sidebar:
        st.header('Proyectos guardados'); SAVED.mkdir(exist_ok=True)
        files=sorted([p.name for p in SAVED.glob('*.json')],reverse=True)
        if files:
            st.selectbox('Seleccionar copia',files,key='saved_choice')
            st.button('Abrir proyecto guardado',on_click=load_saved)
        else: st.caption('Todavía no hay proyectos guardados.')
        uploaded=st.file_uploader('Importar respaldo v2',type='json')
        if uploaded and st.button('Cargar respaldo'):
            try: restore(json.loads(uploaded.getvalue())); st.rerun()
            except (ValueError,KeyError,TypeError) as e: st.error(str(e))
        if 'load_error' in st.session_state: st.error(st.session_state.pop('load_error'))
        if 'catalog_saved' in st.session_state:
            st.info('Usando tarifas de la copia guardada.'); costs=st.session_state['catalog_saved']
            if st.button('Usar tarifas actuales del Excel'):
                del st.session_state['catalog_saved']; st.rerun()
        st.caption('Cada guardado genera una nueva copia.')
    cfg={}; tables={}
    def num(key,label,value,minval=0.0):
        val=st.number_input(label,min_value=float(minval),value=float(value),key='in_'+key); cfg[key]=val; return val
    def select(key,label,opts):
        val=st.selectbox(label,opts,key='in_'+key); cfg[key]=val; return val
    def area(key,label,value=''):
        val=st.text_area(label,value=value,key='in_'+key); cfg[key]=val; return val
    def editor(key,defaults,columns):
        if 'table_'+key not in st.session_state: st.session_state['table_'+key]=defaults
        data=st.session_state['table_'+key]
        widget=f'ed_{key}_{st.session_state.get("rev",0)}'
        frame=st.data_editor(pd.DataFrame(data,columns=list(defaults[0])),num_rows='dynamic',column_config=columns,key=widget,hide_index=True,width='stretch',on_change=commit_table,args=(key,widget,defaults[0]))
        result=[]
        for row in frame.to_dict('records'):
            for col,default in defaults[0].items():
                if pd.isna(row[col]): row[col]=default
            result.append(row)
        tables[key]=result; return result
    def refresh(key):
        st.button('Actualizar información',key='refresh_'+key)
        st.caption('Los campos recalculan al confirmar su edición. El botón actualiza la vista.')
    cfg['proyecto']=st.text_input('Proyecto / cliente','Proyecto de prueba',key='in_proyecto')
    a,b,c=st.columns(3)
    with a: duration=num('duracion','Duración calendario del proyecto (días)',10,1)
    with b: kind=select('tipo','Tipo de proyecto',['LINEAL','AREA'])
    with c: measure=num('metrado','LONGITUD (km)' if kind=='LINEAL' else 'AREA (ha)',1,0.001)
    utility=num('utilidad','Utilidad sobre venta (%)',40)
    st.caption('Según B06: utilidad como recargo sobre costo operativo + riesgos + planilla; no margen sobre el precio de venta.')
    refresh('general')
    tabs=st.tabs(['1. Alcance y metodología','2. Equipos','3. Personal','4. Logística','5. Adicionales y riesgos','6. Reporte'])
    ac=st.column_config.SelectboxColumn('Actividad',options=ACTIVIDADES,required=True,default=ACTIVIDADES[0])
    def nc(label,default=0): return st.column_config.NumberColumn(label,min_value=0,default=default,required=True)
    with tabs[0]:
        g,t=hitos(kind,measure)
        st.info(f'Sugerencia automática: {g} hitos geodésicos y {t} topográficos.')
        st.caption('LINEAL: par inicial + par cada 5 km + par final si remanente ≥ 1 km (incluye 6 y 11 km). Topográficos = máximo(0, techo(km/0.25) − geodésicos). AREA: 2 y 2.')
        cfg['manual_hitos']=st.checkbox('Ajustar manualmente cantidades de hitos',key='in_manual_hitos')
        num('geo','Hitos geodésicos (ajuste manual)',g); num('topo','Hitos topográficos (ajuste manual)',t)
        if not cfg['manual_hitos']: st.caption('Se usan las cantidades automáticas; los ajustes manuales no se aplican.')
        num('certificados','Puntos a certificar IGN',0)
        num('tarifa_geo','Tarifa por hito geodésico S/',costs[next(k for k in costs if k.startswith('Hito de Punto Geod'))])
        num('tarifa_topo','Tarifa por hito topográfico S/',costs['Hito de Punto Topográfico (30x30x30)'])
        area('alcance','Metodología, entregables, exclusiones y requisitos TDR')
        if cfg['certificados']>(cfg['geo'] if cfg['manual_hitos'] else g): st.warning('Certificados mayores a hitos geodésicos. Revisar puntos existentes.')
        refresh('alcance')
    with tabs[1]:
        options=['GPS GNSS Diferencial','Poligonal (estación + kit)','Nivel Electrónico','Nivel Óptico','Estación Total 1"','Drone Matrice 4E','Drone MATRICE 400 + L2','Drone MATRICE 400 + L3','CPU Workstation','Laser Scanner TRIMBLE TX6','Radio UHV - VHF']
        eq=editor('equipos',[dict(Actividad=ACTIVIDADES[0],Equipo=options[0],Cantidad=0.0,Días=duration)],{'Actividad':ac,'Equipo':st.column_config.SelectboxColumn(options=options,required=True,default=options[0]),'Cantidad':nc('Cantidad'),'Días':nc('Días',duration)})
        st.caption('Nuevas filas: días del proyecto por defecto, editables. Cada dron agrega piloto. Poligonal agrega estación + kit. No duplicar recursos compartidos. Confirmar compatibilidad técnica del sensor.')
        refresh('equipos')
    with tabs[2]:
        editor('personal',[{'Actividad':ACTIVIDADES[0],'Cargo':'Topógrafo A','Personas':0.0,'Días remunerados':duration,'Noches alojamiento':duration,'Días viáticos':duration,'Ingreso':True,'Examen médico':False}],{'Actividad':ac,'Cargo':st.column_config.SelectboxColumn(options=roles,required=True,default='Topógrafo A'),'Personas':nc('Personas'),**{k:nc(k,duration) for k in ['Días remunerados','Noches alojamiento','Días viáticos']},'Ingreso':st.column_config.CheckboxColumn(default=True),'Examen médico':st.column_config.CheckboxColumn(default=False)})
        st.caption('Nuevas filas: periodos de la duración del proyecto. Los días ya editados se conservan. No repetir pilotos/conductores automáticos. Su permanencia/habilitación adicional se registra en Adicionales.')
        refresh('personal')
    with tabs[3]:
        num('personas_transporte','Personas a trasladar simultáneamente (sin conductores)',0)
        seats=num('asientos','Asientos útiles por camioneta, sin conductor',4,1)
        suggested=math.ceil(cfg['personas_transporte']/seats)
        st.write(f'Sugerencia: **{suggested} camionetas**. Ajustar por equipos y rutas.')
        num('vehiculos','Camionetas a presupuestar',0)
        select('vehiculo','Tipo de camioneta',['Camioneta MINA Pick-up 4x4 - Hilux','Camioneta Pick-up 4x4 - Hilux'])
        num('dias_vehiculo','Días de camioneta, conductor y combustible',duration)
        select('actividad_logistica','Actividad a la que asignar transporte',ACTIVIDADES)
        if cfg['vehiculos']<suggested: st.warning('Capacidad menor a la sugerida.')
        st.caption('Por camioneta/día: alquiler + 1 conductor + combustible S/ 150.')
        st.subheader('Planilla de apoyo')
        num('planilla','Monto mensual de planilla S/',25000)
        num('dias_planilla','Días de apoyo de planilla al proyecto',duration)
        st.caption('Cálculo en resumen: monto / 24 días laborales × días de apoyo. Reemplaza el 50% / 30 anterior.')
        refresh('logistica')
    with tabs[4]:
        st.subheader('Gastos adicionales')
        editor('extras',[dict(Actividad=ACTIVIDADES[0],Concepto='Otros gastos',Cantidad=0.0,Periodos=1.0,Tarifa=0.0)],{'Actividad':ac,'Cantidad':nc('Cantidad'),'Periodos':nc('Periodos',1),'Tarifa':nc('Tarifa')})
        st.subheader('Daño o robo de equipos')
        exposure=[]
        for name in VALORES:
            q=sum(r['Cantidad'] for r in eq if r['Equipo']==name and r['Días']>0)
            if q: exposure.append({'Aplicar':True,'Actividad':next(r['Actividad'] for r in eq if r['Equipo']==name),'Equipo':name,'Cantidad':q,'Valor equipo':float(VALORES[name]),'Cobertura %':50.0,'Días exposición':duration})
        signature=json.dumps([(r['Equipo'],r['Cantidad'],r['Días']) for r in eq],sort_keys=True)
        if signature!=st.session_state.get('exposure_signature'):
            old={r['Equipo']:r for r in st.session_state.get('table_exposicion',[])}
            exposure=[dict(old.get(r['Equipo'],r),Cantidad=r['Cantidad']) for r in exposure]
            st.session_state['table_exposicion']=exposure
            st.session_state['exposure_signature']=signature
            st.session_state['risk_rev']=st.session_state.get('risk_rev',0)+1
            for key in list(st.session_state):
                if key.startswith('ed_exposicion_'): del st.session_state[key]
        if exposure:
            editor('exposicion',exposure,{'Actividad':ac,'Aplicar':st.column_config.CheckboxColumn(default=True),'Cantidad':nc('Cantidad'),'Valor equipo':nc('Valor equipo'),'Cobertura %':st.column_config.NumberColumn(min_value=0,max_value=100,default=50),'Días exposición':nc('Días exposición',duration),'Equipo':st.column_config.TextColumn(disabled=True)})
        else: tables['exposicion']=[]; st.caption('Selecciona equipos con cantidad y días para generar exposición.')
        st.caption('Provisión = cantidad × valor equipo × cobertura% / 30 × días exposición. Valores y días editables. Provisión comercial, no prima de seguro. Recursos repetidos entre actividades: ajustar cantidad para no duplicar. Equipos sin valor definido: registrar riesgo adicional.')
        num('espera','Días de paradas / esperas',0)
        num('induccion','Jornadas de inducción adicional (50% del personal diario)',0)
        select('actividad_riesgos','Actividad a la que asignar paradas e inducción',ACTIVIDADES)
        st.caption('Base: remuneración diaria del personal, incluidos pilotos y conductores. Paradas: base × días. Inducción: base × 50% × jornadas. No incluye alojamiento ni equipos; agregarlos si aplica. No duplicar inducciones pagadas.')
        editor('otros_riesgos',[{'Actividad':ACTIVIDADES[0],'Riesgo':'Otro riesgo','Porcentaje':0.0}],{'Actividad':ac,'Porcentaje':st.column_config.NumberColumn(min_value=0,max_value=100,default=0)})
        area('notas','Otras consideraciones y sustento de riesgos','Impuestos no incluidos. Evitar duplicar contingencias con recursos ya presupuestados.')
        refresh('riesgos')
    for name,records in tables.items():
        for r in records:
            for key,val in r.items():
                if isinstance(val,(int,float)) and (not math.isfinite(val) or val<0): st.error(f'{name}: {key} debe ser no negativo y finito.'); st.stop()
            if 'Actividad' in r and r['Actividad'] not in ACTIVIDADES: st.error('Selecciona una actividad válida.'); st.stop()
    filas,riesgos,planilla,diario,_=calcular(cfg,tables,costs)
    operativo=sum(map(importe,filas)); risk=sum(map(importe,riesgos)); pl=importe(planilla)
    profit=(operativo+risk+pl)*utility/100; total=operativo+risk+pl+profit
    snapshot={'schema':2,'saved_at':datetime.now().isoformat(),'config':cfg,'tables':tables,'catalog':costs,'source_sha256':hashlib.sha256(SOURCE.read_bytes()).hexdigest(),'rules':'v2-20261102'}
    with tabs[5]:
        for col,label,val in zip(st.columns(5),['Costo operativo','Riesgos','PLAME / planilla','Utilidad','Costo total'],[operativo,risk,pl,profit,total]): col.metric(label,f'S/ {val:,.2f}')
        st.metric('Costo por km' if kind=='LINEAL' else 'Costo por ha',f'S/ {total/measure:,.2f}')
        st.write(f'Planilla: S/ {cfg["planilla"]:,.2f} / 24 × {cfg["dias_planilla"]:g} días = S/ {pl:,.2f}. Utilidad: {utility:g}% sobre operativo + riesgos + planilla. Total antes de impuestos.')
        allrows=filas+riesgos+[planilla]
        detail=pd.DataFrame(allrows); detail['Subtotal S/']=[importe(r) for r in allrows]
        st.subheader('Presupuesto por actividad (antes de utilidad)')
        st.dataframe(detail.pivot_table(index='Actividad',columns='Grupo',values='Subtotal S/',aggfunc='sum',fill_value=0),width='stretch')
        st.caption('Monumentación e IGN: Georreferenciación. Planilla: Gabinete. Otros conceptos: actividad seleccionada; no se distribuyen proporcionalmente.')
        st.dataframe(detail,width='stretch',hide_index=True)
        st.write(f'Base diaria de remuneraciones usada en paradas/inducciones: S/ {diario:,.2f}')
        st.download_button('Descargar Excel detallado editable',exportar(cfg,filas,riesgos,planilla),file_name='Cotizacion_MATIC_v2.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        st.download_button('Descargar respaldo completo',json.dumps(snapshot,ensure_ascii=False,indent=2),file_name='Proyecto_MATIC_v2.json')
        refresh('reporte')
    st.divider()
    if st.button('Guardar proyecto como nueva copia',type='primary'):
        safe=''.join(x if x.isalnum() or x in '-_' else '_' for x in cfg['proyecto'])[:70] or 'Proyecto'
        dest=SAVED/(datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'_'+safe+'.json')
        dest.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2),encoding='utf-8')
        st.success(f'Guardado: {dest.name}. Actualiza la página para verlo en la barra lateral.')

if __name__=='__main__': main()

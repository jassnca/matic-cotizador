"""Cuadrillas y cronograma consecutivo. Rendimientos proporcionados por el usuario."""
import math
from collections import defaultdict

def jornadas(valor): return max(0,math.ceil(valor-1e-9))

def planificar(geo,topo,opciones,km,dias_geo=1,terrestre='ESTACION TOTAL',aerea='FOTOGRAMETRIA',gabinete=False,semana=5):
    etapas=[]; equipos=[]; personal=[]; cursor=1
    def etapa(nombre,teorico,cuadrilla,equipo=None,cantidad=1,actividad='GEORREFERENCIACION'):
        nonlocal cursor
        dias=jornadas(teorico)
        if not dias: return
        etapas.append(dict(Etapa=nombre,Inicio=cursor,Fin=cursor+dias-1,Días=dias,Teórico=teorico,Actividad=actividad))
        cursor+=dias
        if equipo: equipos.append(dict(Actividad=actividad,Etapa=nombre,Equipo=equipo,Cantidad=cantidad,Días=dias,Origen='Automático'))
        for cargo,n in cuadrilla.items():
            personal.append({'Actividad':actividad,'Etapa':nombre,'Cargo':cargo,'Personas':n,'Días remunerados':dias,'Noches alojamiento':dias,'Días viáticos':dias,'Ingreso':True,'Examen médico':False,'Origen':'Automático'})
    monument={'Topógrafo A':1,'Auxiliar de Topografía':1,'Ayudante Locales':2}
    etapa('Monumentación geodésica',geo/6,monument)
    etapa('Monumentación topográfica',topo/8,monument)
    if 'GEODESIA' in opciones:
        q=3 if geo>4 else 2
        etapa('Geodesia',dias_geo,{'Topógrafo A':q},'GPS GNSS Diferencial',q)
    if 'NIVELACION' in opciones: etapa('Nivelación',km/5,{'Topógrafo A':1,'Ayudante Locales':2},'Nivel Electrónico')
    if 'POLIGONAL' in opciones: etapa('Poligonal',km/5,{'Topógrafo A':1,'Auxiliar de Topografía':2},'Poligonal (estación + kit)')
    if 'CAPTURA TERRESTRE' in opciones:
        if terrestre=='ESTACION TOTAL': etapa('Captura terrestre — Estación Total',km,{'Topógrafo A':1,'Ayudante Locales':2},'Estación Total 1"',actividad='MAPEO GEOESPACIAL')
        else: etapa('Captura terrestre — GNSS',km/2,{'Topógrafo A':2,'Ayudante Locales':1},'GPS GNSS Diferencial',3,'MAPEO GEOESPACIAL')
    if 'CAPTURA AEREA' in opciones:
        if aerea=='FOTOGRAMETRIA': etapa('Captura aérea — Fotogrametría',km/10,{'Piloto Drone Matrice 4E':1},'Drone Matrice 4E',actividad='MAPEO GEOESPACIAL')
        else: etapa('Captura aérea — LiDAR',km/15,{'Piloto Drone MATRICE 400':1},'Drone MATRICE 400 + L3',actividad='MAPEO GEOESPACIAL')
    if gabinete:
        distribucion={act:sum(s['Días'] for s in etapas if s['Actividad']==act) for act in ['GEORREFERENCIACION','MAPEO GEOESPACIAL']}
        apoyo(personal,cursor-1,jornadas(cursor-1+km/10*semana),distribucion)
        dias_gab=jornadas(cursor-1+km/10*semana)
        if dias_gab: etapas.append(dict(Etapa='Gabinete',Inicio=1,Fin=dias_gab,Días=dias_gab,Teórico=cursor-1+km/10*semana,Actividad='GABINETE'))
    return etapas,equipos,personal

def apoyo(personal,dias_campo,dias_gabinete,distribucion=None):
    """Soporte continuo, sin sumar otra vez estas duraciones a las etapas de campo."""
    rows=[]
    for act,days in (distribucion or {'GEORREFERENCIACION':dias_campo}).items():
        rows.extend([(cargo,1,act,'Gestión de campo',days) for cargo in ['Jefe de Proyecto','Asistente HSQ']])
    rows.extend([
        ('Coordinador de Procesamiento de Datos',1,'GABINETE','Gabinete',dias_gabinete),
        ('Cadista C3D / GIS - A',2,'GABINETE','Gabinete',dias_gabinete),
        ('Procesador Geoespacial',1,'GABINETE','Gabinete',dias_gabinete)])
    for cargo,qty,act,etapa,days in rows:
        if not days: continue
        personal.append({'Actividad':act,'Etapa':etapa,'Cargo':cargo,'Personas':qty,'Días remunerados':days,'Noches alojamiento':0 if act=='GABINETE' else days,'Días viáticos':0 if act=='GABINETE' else days,'Ingreso':act!='GABINETE','Examen médico':False,'Origen':'Automático'})

def dotacion(personal,equipos):
    """Personal único asignado de campo: máximos por cargo reutilizable + manuales."""
    auto=defaultdict(float); manual=0
    for r in personal:
        if r['Actividad']=='GABINETE' or r['Días remunerados']<=0 or r['Cargo']=='Conductor': continue
        if r.get('Origen')=='Automático': auto[r['Cargo']]=max(auto[r['Cargo']],r['Personas'])
        else: manual+=r['Personas']
    for r in equipos:
        if r['Actividad']!='GABINETE' and r['Equipo'].startswith('Drone') and r['Días']>0:
            pilot='Piloto Drone Matrice 4E' if r['Equipo']=='Drone Matrice 4E' else 'Piloto Drone MATRICE 400'
            if not any(p['Cargo']==pilot and p.get('Etapa')==r.get('Etapa') and p['Días remunerados']>0 for p in personal): manual+=r['Cantidad']
    return math.ceil(sum(auto.values())+manual)

def movilidad(personas):
    if personas<=0: return 0,0
    if personas<=4: return 1,0
    if personas<=8: return 2,0
    return 1,max(1,math.ceil((personas-4)/16))

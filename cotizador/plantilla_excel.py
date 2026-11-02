"""Exportación quirúrgica OOXML: conserva VBA, WMF y diseño original byte a byte."""
from io import BytesIO
from pathlib import Path
import zipfile,re
from xml.sax.saxutils import escape,quoteattr
from motor import importe,ACTIVIDADES

TEMPLATE=Path(__file__).resolve().parents[1]/'PLANTILLA COTIZACION.xlsm'
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def col(n):
    out=''
    while n: n,r=divmod(n-1,26); out=chr(65+r)+out
    return out

def cell(ref,value,style='',cached=None):
    attr=f' r="{ref}"'+(f' s="{style}"' if style else '')
    if isinstance(value,str) and value.startswith('='):
        return f'<c{attr}><f>{escape(value[1:])}</f>'+('' if cached is None else f'<v>{cached}</v>')+'</c>'
    if isinstance(value,(int,float)): return f'<c{attr}><v>{value}</v></c>'
    # XML no permite caracteres de control, incluso dentro de cadenas escapadas.
    value=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]','',str(value or ''))
    return f'<c{attr} t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'

def sheet_xml(rows):
    data=''.join('<row r="'+str(i)+'">'+''.join(cell(col(j)+str(i),v) for j,v in enumerate(row,1))+'</row>' for i,row in enumerate(rows,1))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="3" width="38" customWidth="1"/><col min="4" max="8" width="18" customWidth="1"/></cols><sheetData>{data}</sheetData></worksheet>'

def generar(cfg,filas,riesgos,planilla,etapas):
    headers=['Grupo','Actividad','Recurso','Cantidad','Periodos','Unidad','Tarifa S/','Subtotal S/']
    def detalle(records): return [headers]+[[r[k] for k in ['Grupo','Actividad','Recurso','Cantidad','Periodos','Unidad','Tarifa']]+[f'=D{i}*E{i}*G{i}'] for i,r in enumerate(records,2)]
    nd=max(2,len(filas)+1); nr=max(2,len(riesgos)+1)
    sumgrupo=lambda act: f'SUMIFS(\'Detalle MATIC\'!H2:H{nd},\'Detalle MATIC\'!B2:B{nd},"{act}",\'Detalle MATIC\'!A2:A{nd},"<>Ingreso")+SUMIF(\'Riesgos MATIC\'!B2:B{nr},"{act}",\'Riesgos MATIC\'!H2:H{nr})'
    resumen=[['Concepto','Soles'],['Operativo',f"=SUM('Detalle MATIC'!H2:H{nd})"],['Riesgos',f"=SUM('Riesgos MATIC'!H2:H{nr})"],['Planilla',"='Planilla MATIC'!H2"],['Utilidad %',cfg['utilidad']/100],['Utilidad','=SUM(B2:B4)*B5'],['Cotización sin IGV','=SUM(B2:B4)+B6'],['Metrado',cfg['metrado']],['Costo por km/ha','=B7/B8'],['IGV %',cfg.get('igv',18)/100],['IGV','=B7*B10'],['Total con IGV','=B7+B11']]
    notes=[['Consideraciones'],[cfg.get('alcance','')],[cfg.get('notas','')],['Planilla = monto mensual / 24 × días apoyo. Utilidad como recargo sobre operativo + riesgos + planilla.'],['Riesgos equipos: días operativos de cada asignación; no duración calendario.'],['Campo consecutivo; gabinete desde el inicio hasta campo + procesamiento a 10 km/semana de 5 días. Ingreso automático una vez por cargo y máxima dotación.'],['Capacidad: 4 asientos útiles por camioneta; combi usa MiniBus 16 Pasajeros c/Conductor del catálogo más combustible.'],['El detalle es editable; su suma alimenta la propuesta. Si cambias jornadas en Excel, revisa también riesgos y dotación: el plan de recursos de la aplicación no se recalcula desde Excel.']]
    extra={'Detalle MATIC':detalle(filas),'Riesgos MATIC':detalle(riesgos),'Planilla MATIC':detalle([planilla]),'Resumen MATIC':resumen,'Cronograma MATIC':[['Etapa','Inicio','Fin','Días','Días teóricos']]+[[s[k] for k in ['Etapa','Inicio','Fin','Días','Teórico']] for s in etapas],'Notas MATIC':notes}
    ingreso=sum(importe(r) for r in filas if r['Grupo']=='Ingreso')
    groups={act:sum(importe(r) for r in filas if r['Actividad']==act and r['Grupo']!='Ingreso')+sum(importe(r) for r in riesgos if r['Actividad']==act) for act in ACTIVIDADES}
    direct=ingreso+sum(groups.values()); plan=importe(planilla); util=(direct+plan)*cfg['utilidad']/100; subtotal=direct+plan+util
    values={'B4':cfg.get('codigo',''),'B5':cfg['proyecto'],'B6':cfg.get('servicio','Servicio de geomática'),'B8':cfg.get('cliente',''),'B9':cfg.get('ruc',''),'B10':cfg.get('direccion',''),'B12':cfg.get('atencion',''),'B13':cfg.get('email',''),
            'I18':cfg.get('personas_unicas',0),'I20':cfg['geo']+cfg['topo'],'I22':cfg.get('campo_por_actividad',{}).get('MAPEO GEOESPACIAL',cfg.get('duracion',0)) if groups[ACTIVIDADES[1]] else 0,'H22':'días','I24':1 if groups[ACTIVIDADES[2]] else 0,'I25':0,
            'B18':'EPP, inducción, seguros y exámenes aplicables','B20':'Monumentación, geodesia, nivelación y poligonal','B22':'Equipos, personal y logística asignada','B27':'Planilla de apoyo','I27':f'=IF(J26>0,J27/J26,0)','I28':cfg['utilidad']/100,'I30':cfg.get('igv',18)/100,
            'J17':f'=SUMIF(\'Detalle MATIC\'!A2:A{nd},"Ingreso",\'Detalle MATIC\'!H2:H{nd})','J19':'='+sumgrupo(ACTIVIDADES[0]),'J21':'='+sumgrupo(ACTIVIDADES[1]),'J23':'='+sumgrupo(ACTIVIDADES[2]),
            'J26':'=SUM(J17,J19,J21,J23)','J27':"='Planilla MATIC'!H2",'J28':'=(J26+J27)*I28','J29':'=SUM(J26:J28)','J30':'=J29*I30','J31':'=J29+J30',
            'B33':'="Son S/. "&TEXT(J29,"#,##0.00")&" antes de IGV."',
            'B55':f'2 Plazo presupuestado: {cfg.get("duracion",0):g} jornadas. Campo: {cfg.get("dias_campo",cfg.get("duracion",0)):g}; gabinete: {cfg.get("dias_gabinete",0):g}. Ver cronograma.',
            'B57':cfg.get('notas','') or 'Presupuesto sujeto al alcance y al cronograma adjunto.'}
    values['B24']='Procesamiento e informes según alcance' if groups[ACTIVIDADES[2]] else 'No presupuestado'
    values['B25']=''
    if not groups[ACTIVIDADES[1]]: values['B22']='No presupuestado'
    # Eliminar recursos y entregables de ejemplo para no prometer los de otra propuesta.
    for row in range(37,43): values[f'B{row}']=''; values[f'G{row}']=''
    for ref in ['B45','C46','C47','C48']: values[ref]=''
    values['B45']=cfg.get('alcance','') or 'Entregables pendientes de definición en el alcance del proyecto.'
    for row in [52,53,54]: values[f'C{row}']=''; values[f'D{row}']=''
    values['C52']='Por definir'; values['D52']='Condiciones de pago pendientes de aprobación comercial.'
    personnel=cfg.get('lista_personal',[]); equipment=cfg.get('lista_equipos',[])
    for i,value in enumerate(personnel[:6],37): values[f'B{i}']=value
    for i,value in enumerate(equipment[:6],37): values[f'G{i}']=value
    if len(personnel)>6: values['B42']='Ver dotación completa en Detalle MATIC'
    if len(equipment)>6: values['G42']='Ver equipos completos en Detalle MATIC'
    cache={'J17':ingreso,'J19':groups[ACTIVIDADES[0]],'J21':groups[ACTIVIDADES[1]],'J23':groups[ACTIVIDADES[2]],'J26':direct,'J27':plan,'J28':util,'J29':subtotal,'J30':subtotal*cfg.get('igv',18)/100,'J31':subtotal*(1+cfg.get('igv',18)/100)}
    src=zipfile.ZipFile(TEMPLATE); parts={n:src.read(n) for n in src.namelist()}
    xml=parts['xl/worksheets/sheet1.xml'].decode('utf-8')
    for ref,value in values.items():
        pat=r'<c\b[^>]*\br="'+ref+r'"[^>]*?(?:/>|>.*?</c>)'
        match=re.search(pat,xml,re.S)
        if match:
            style=re.search(r'\bs="(\d+)"',match.group())
            xml=xml[:match.start()]+cell(ref,value,style.group(1) if style else '',cache.get(ref))+xml[match.end():]
        else:
            rownum=re.search(r'\d+',ref).group(); rp=r'(<row\b[^>]*\br="'+rownum+r'"[^>]*>)(.*?)(</row>)'
            m=re.search(rp,xml,re.S)
            if not m: raise ValueError(f'Fila {rownum} no encontrada en plantilla')
            # Insertar manteniendo el orden de las columnas OOXML.
            inner=m.group(2); new=cell(ref,value,cached=cache.get(ref)); target=0
            for char in re.match(r'[A-Z]+',ref).group(): target=target*26+ord(char)-64
            inserted=False
            for cm in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)',inner,re.S):
                idx=0
                for char in cm.group(1): idx=idx*26+ord(char)-64
                if idx>target: inner=inner[:cm.start()]+new+inner[cm.start():]; inserted=True; break
            if not inserted: inner+=new
            xml=xml[:m.start()]+m.group(1)+inner+m.group(3)+xml[m.end():]
    parts['xl/worksheets/sheet1.xml']=xml.encode()
    workbook=parts['xl/workbook.xml'].decode(); rels=parts['xl/_rels/workbook.xml.rels'].decode(); types=parts['[Content_Types].xml'].decode()
    for idx,(name,rows) in enumerate(extra.items(),30):
        workbook=workbook.replace('</sheets>',f'<sheet name={quoteattr(name)} sheetId="{idx}" r:id="rIdMatic{idx}"/></sheets>')
        rels=rels.replace('</Relationships>',f'<Relationship Id="rIdMatic{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/matic{idx}.xml"/></Relationships>')
        types=types.replace('</Types>',f'<Override PartName="/xl/worksheets/matic{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        parts[f'xl/worksheets/matic{idx}.xml']=sheet_xml(rows).encode()
    workbook=re.sub(r'<calcPr\b[^>]*/>','<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1"/>',workbook)
    # La cadena de cálculo anterior refería a fórmulas reemplazadas; Excel la reconstruye.
    rels=re.sub(r'<Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*/>','',rels)
    types=re.sub(r'<Override\b[^>]*PartName="/xl/calcChain.xml"[^>]*/>','',types)
    parts.pop('xl/calcChain.xml',None)
    parts['xl/workbook.xml']=workbook.encode(); parts['xl/_rels/workbook.xml.rels']=rels.encode(); parts['[Content_Types].xml']=types.encode()
    out=BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as dest:
        for name,data in parts.items(): dest.writestr(name,data)
    return out.getvalue()

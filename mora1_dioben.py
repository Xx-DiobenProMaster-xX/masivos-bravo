
"""Cartera Dioben: diagnóstico de secuencia Mora 1. NO ENVÍA NI ESCRIBE.

Usa FECHA_INICIO_MORA_1 cuando es valida; en su ausencia usa 2026-09-30.
Modo diagnostico: no envia correos ni cambia las hojas.
"""
import os
import json
import re
import unicodedata
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials

ID_DI0BEN = '1ELFN7ONft65RWeeHW5j7e5kvIyzOiG0UFWVeHo9PTLg'
ID_MASIVOS = '1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI'
TZ = ZoneInfo('America/Bogota')
DIAS_ESPECIALES = {4:'MORA101',10:'MORA102',15:'MORA103',19:'MORA104',24:'MORA105'}
FECHA_BASE = date(2026,9,30)
# Estos estados son de referencia; la lista exacta se debe contrastar con Configuración del Apps Script.
ESTADOS_POSITIVOS = {'PAGO','AL DIA','PERMANECIDO','PROMESA DE PAGO','REESTRUCTURACION','PAGO CONFIRMADO','ACUERDO DE PAGO','PAGO PROGRAMADO','DEBITO PROGRAMADO','CASO SOLUCIONADO'}
ESTADOS_NEUTROS = {'EN GESTION'}

def norm(v):
    s=unicodedata.normalize('NFD',str(v or '').strip())
    return ''.join(c for c in s if unicodedata.category(c)!='Mn').upper()

def ref(v):
    s=str(v or '').strip()
    return s[:-2] if re.fullmatch(r'\d+\.0',s) else s

def parse_date(v):
    if isinstance(v,date): return v
    s=str(v or '').strip().split(' ')[0]
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y'):
        try:return datetime.strptime(s,fmt).date()
        except ValueError:pass
    return None

def rows(ws):
    values=ws.get_all_values()
    if not values:raise RuntimeError('Hoja vacía: '+ws.title)
    headers={norm(x):i for i,x in enumerate(values[0])}
    return headers,values[1:]

def value(r,h,name):
    i=h.get(norm(name))
    return r[i].strip() if i is not None and i<len(r) else ''

def run():
    if not os.getenv('MI_JSON'):raise RuntimeError('Falta MI_JSON')
    creds=Credentials.from_service_account_info(json.loads(os.environ['MI_JSON']),scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
    book=gspread.authorize(creds).open_by_key(ID_DI0BEN)
    hs,seguimiento=rows(book.worksheet('Seguimiento'))
    hm,masivos=rows(book.worksheet('Masivos_Correos'))
    principal=gspread.authorize(creds).open_by_key(ID_MASIVOS)
    hp,templates=rows(principal.worksheet('PLANTILLAS'))
    hc,cola=rows(principal.worksheet('COLA_ENVIO'))
    if not {'ID_PLANTILLA','ASUNTO','CUERPO','ESTADO'}.issubset(hp):
        raise RuntimeError('PLANTILLAS: faltan ID_PLANTILLA, ASUNTO, CUERPO o ESTADO')
    if not {'REFERENCIA','PLANTILLA','ESTADO'}.issubset(hc):
        raise RuntimeError('COLA_ENVIO: faltan REFERENCIA, PLANTILLA o ESTADO')
    disponibles={}
    for row in templates:
        code=norm(value(row,hp,'ID_PLANTILLA'))
        if code in DIAS_ESPECIALES.values() and norm(value(row,hp,'ESTADO'))=='ACTIVA' and value(row,hp,'ASUNTO') and value(row,hp,'CUERPO'):
            disponibles[code]=disponibles.get(code,0)+1
    # Historial por referencia y plantilla. Distinguir enviados, pendientes y errores.
    # Si falta la fecha, bloquear por precaución en vez de asumir otro ciclo.
    existentes={}
    for row in cola:
        key=(ref(value(row,hc,'REFERENCIA')), norm(value(row,hc,'PLANTILLA')))
        if not key[0] or not key[1]:
            continue
        estado=norm(value(row,hc,'ESTADO'))
        when=parse_date(value(row,hc,'FECHA_ENVIO')) if 'FECHA_ENVIO' in hc else None
        existentes.setdefault(key,[]).append((estado,when))
    bloqueo_estados={'ENVIADO','ENVIADA','ENVIANDO','PENDIENTE','PROGRAMADO','PROGRAMADA','PREPARADO','PREPARADA','EN COLA','LISTO','LISTA'}
    def bloqueo_ciclo(reference,template,inicio):
        entries=existentes.get((reference,template),[])
        for estado,fecha in entries:
            if estado in bloqueo_estados:
                if fecha is None or fecha>=inicio:
                    return 'YA ENVIADO O EN COLA EN ESTE CICLO'
            elif estado not in {'ERROR','CANCELADO','CANCELADA','BLOQUEADO','BLOQUEADA','FALLIDO','FALLIDA'}:
                if fecha is None or fecha>=inicio:
                    return 'ESTADO DE COLA NO RECONOCIDO; REVISAR'
        return None
    required={'REFERENCIA','MORA ACTUAL','STATUS','EXCLUIR CORREO','FECHA_INICIO_MORA_1'}
    missing=required-set(hs)
    if missing:raise RuntimeError('Faltan columnas en Seguimiento: '+', '.join(sorted(missing)))
    if not {'REFERENCIA','HABILITADO','TIPO CORREO'}.issubset(hm):raise RuntimeError('Faltan columnas en Masivos_Correos')
    idx={}
    duplicates=set()
    for item in seguimiento:
        key=ref(value(item,hs,'REFERENCIA'))
        if not key:continue
        if key in idx:duplicates.add(key)
        else:idx[key]=item
    fechas={'VACIA':0,'INVALIDA':0,'FUTURA':0,'PENDIENTE_CONFIRMACION':0}
    for key,item in idx.items():
        raw=value(item,hs,'FECHA_INICIO_MORA_1')
        parsed=parse_date(raw)
        if not raw:fechas['VACIA']+=1
        elif not parsed:fechas['INVALIDA']+=1
        elif parsed>datetime.now(TZ).date():fechas['FUTURA']+=1
        else:fechas['PENDIENTE_CONFIRMACION']+=1
    today=datetime.now(TZ).date()
    reasons={}
    examples=[]
    for r in masivos:
        reference=ref(value(r,hm,'REFERENCIA'))
        if not reference:continue
        s=idx.get(reference)
        reason='ELEGIBLE PARA VALIDAR HISTORICO'
        if reference in duplicates:reason='REFERENCIA DUPLICADA EN SEGUIMIENTO'
        elif not s:reason='SIN SEGUIMIENTO'
        elif norm(value(s,hs,'MORA ACTUAL'))!='MORA 1':reason='NO MORA 1'
        elif norm(value(s,hs,'EXCLUIR CORREO')) in {'TRUE','VERDADERO','SI','1','X'}:reason='EXCLUIDO'
        elif norm(value(r,hm,'HABILITADO')) not in {'SI'}:reason='NO HABILITADO'
        elif norm(value(r,hm,'TIPO CORREO')) not in {'MORA_1','MORA 1'}:reason='OTRA CAMPAÑA'
        else:
            fecha_original=parse_date(value(s,hs,'FECHA_INICIO_MORA_1'))
            start=fecha_original if fecha_original and fecha_original<=today else FECHA_BASE
            origen='HISTORICO_HOJA' if fecha_original and fecha_original<=today else 'BASE_30_SEPTIEMBRE'
            if not fecha_original or fecha_original>today:
                reason='FECHA DE CICLO SIN CONFIRMAR: NO APTO PARA ENVIO'
            elif start>today:
                reason='FECHA POSTERIOR A HOY'
            else:
                day=(today-start).days+1
                template=DIAS_ESPECIALES.get(day)
                if template:
                    tipo='ESPECIAL '+template
                elif day>=1 and day%2==0:
                    tipo='HABITUAL'
                else:
                    tipo='SIN ENVIO PROGRAMADO'
                if template and disponibles.get(template)!=1:
                    reason='PLANTILLA FALTANTE, INACTIVA O DUPLICADA: '+template
                elif template and bloqueo_ciclo(reference,template,start):
                    reason=bloqueo_ciclo(reference,template,start)+' | '+template
                elif not template and tipo=='HABITUAL':
                    reason='HABITUAL SIN PLANTILLA CONFIGURADA: NO APTO PARA ENVIO'
                elif tipo=='SIN ENVIO PROGRAMADO':
                    reason='SIN ENVIO PROGRAMADO | DIA '+str(day)
                else:
                    reason=tipo+' | DIA '+str(day)+' | '+origen+' | COLA_VERIFICADA'
                if len(examples)<25:examples.append((reference,reason))
        reasons[reason]=reasons.get(reason,0)+1
    print('CARTERA DIOBEN | DIAGNOSTICO |',datetime.now(TZ).isoformat())
    print('NO SE ENVÍAN CORREOS. NO SE MODIFICAN HOJAS.')
    print('Fecha valida de Seguimiento o base 2026-09-30; verificar historicidad real por separado.')
    for k,v in sorted(reasons.items()):print(k,':',v)
    print('MUESTRA:',examples)
    print('AUDITORIA DE FECHAS (NO CONFIRMADAS):',fechas)
    print('DUPLICADOS EN SEGUIMIENTO:',len(duplicates))
    print('PLANTILLAS ACTIVAS Y COMPLETAS:',disponibles)
    print('PARES REFERENCIA/PLANTILLA EN COLA:',len(existentes))
    print('SIMULACION SEGURA: los ciclos sin fecha confirmada y los días habituales sin plantilla quedan bloqueados.')
    print('PENDIENTE: vincular motor de envío con cuenta autorizada y confirmar deduplicación por ciclo; esta prueba es de solo lectura.')

if __name__=='__main__':run()

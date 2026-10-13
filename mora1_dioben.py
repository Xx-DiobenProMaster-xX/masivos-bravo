
"""Cartera Dioben: diagnóstico de secuencia Mora 1. NO ENVÍA NI ESCRIBE.

La fecha de inicio requiere una transición histórica confirmada. La fecha R
se utiliza solo si se ha verificado previamente; no se infiere de ACTUALIZADO.
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
TZ = ZoneInfo('America/Bogota')
DIAS_ESPECIALES = {4:'MORA101',10:'MORA102',15:'MORA103',19:'MORA104',24:'MORA105'}
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
        elif norm(value(s,hs,'STATUS')) in ESTADOS_POSITIVOS|ESTADOS_NEUTROS:reason='STATUS BLOQUEADO'
        elif norm(value(r,hm,'HABILITADO')) not in {'SI'}:reason='NO HABILITADO'
        elif norm(value(r,hm,'TIPO CORREO')) not in {'MORA_1','MORA 1'}:reason='OTRA CAMPAÑA'
        else:
            start=parse_date(value(s,hs,'FECHA_INICIO_MORA_1'))
            if not start or start>today:reason='INICIO NO CONFIRMADO'
            else:
                day=(today-start).days+1
                template=DIAS_ESPECIALES.get(day)
                # No inferir que el histórico de envíos está vacío: requiere lectura
                # de COLA_ENVIO antes de habilitar siquiera un envío de prueba.
                reason=('ESPECIAL '+template if template else 'HABITUAL POR VALIDAR')+' | DIA '+str(day)+' | SIN VERIFICAR HISTORIAL'
                if len(examples)<25:examples.append((reference,reason))
        reasons[reason]=reasons.get(reason,0)+1
    print('CARTERA DIOBEN | DIAGNOSTICO |',datetime.now(TZ).isoformat())
    print('NO SE ENVÍAN CORREOS. NO SE MODIFICAN HOJAS.')
    print('Importante: fecha R pendiente de validación contra histórico real de transiciones.')
    for k,v in sorted(reasons.items()):print(k,':',v)
    print('MUESTRA:',examples)
    print('AUDITORIA DE FECHAS (NO CONFIRMADAS):',fechas)
    print('DUPLICADOS EN SEGUIMIENTO:',len(duplicates))
    print('PENDIENTE: validar origen historico de la fecha, pagos, COLA_ENVIO y plantillas antes de produccion.')

if __name__=='__main__':run()

"""Diagnóstico sólo lectura: Mora 1 esperado hoy vs borradores existentes."""
import json,os,collections
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from mora1_dioben import rows,value,norm,ref,parse_date,ID_DI0BEN,ID_MASIVOS,DIAS_ESPECIALES

today=datetime.now(ZoneInfo("America/Bogota")).date()
creds=Credentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
gc=gspread.authorize(creds)
src=gc.open_by_key(ID_DI0BEN)
dst=gc.open_by_key(ID_MASIVOS)
hs,follow=rows(src.worksheet("Seguimiento"))
hm,selected=rows(src.worksheet("Masivos_Correos"))
hq,queue=rows(dst.worksheet("COLA_ENVIO"))
ht,templates=rows(dst.worksheet("PLANTILLAS"))
hc,clients=rows(dst.worksheet("CLIENTES"))
follow_map=collections.defaultdict(list)
client_map=collections.defaultdict(list)
for r in follow:
    k=ref(value(r,hs,"REFERENCIA"))
    if k:follow_map[k].append(r)
for r in clients:
    k=ref(value(r,hc,"REFERENCIA"))
    if k:client_map[k].append(r)
active={norm(value(r,ht,"ID_PLANTILLA")) for r in templates if norm(value(r,ht,"ESTADO"))=="ACTIVA"}
reasons=collections.Counter()
expected=collections.Counter()
selected_refs=set()
for r in selected:
    key=ref(value(r,hm,"REFERENCIA"))
    if not key or key in selected_refs:continue
    selected_refs.add(key)
    if norm(value(r,hm,"HABILITADO"))!="SI" or norm(value(r,hm,"TIPO CORREO")) not in {"MORA_1","MORA 1"}:
        reasons["No seleccionado para Mora 1"]+=1;continue
    if len(follow_map[key])!=1:
        reasons["Seguimiento faltante o duplicado"]+=1;continue
    current=follow_map[key][0]
    if norm(value(current,hs,"MORA ACTUAL"))!="MORA 1":
        reasons["No está en Mora 1 (incluye AL DÍA)"]+=1;continue
    if norm(value(current,hs,"EXCLUIR CORREO")) in {"TRUE","VERDADERO","SI","1","X"}:
        reasons["Excluido del correo"]+=1;continue
    start=parse_date(value(current,hs,"FECHA_INICIO_MORA_1"))
    if not start or start>today:
        reasons["Fecha de inicio inválida o futura"]+=1;continue
    day=(today-start).days+1
    code=DIAS_ESPECIALES.get(day) or ("T001" if day%2==0 else None)
    if not code:
        reasons["Hoy no corresponde plantilla"]+=1;continue
    if code not in active:
        reasons["Plantilla no activa"]+=1;continue
    if len(client_map[key])!=1 or not value(client_map[key][0],hc,"EMAIL"):
        reasons["CLIENTES sin email o duplicado"]+=1;continue
    expected[code]+=1
print("Fecha Colombia:",today.isoformat())
print("Clientes seleccionados únicos en Masivos_Correos:",len(selected_refs))
print("Candidatos Mora 1 que hoy requieren plantilla (ANTES de historial):",sum(expected.values()))
print("Distribución por plantilla:",dict(expected))
print("Razones de exclusión:",dict(reasons))
queue_counts=collections.Counter()
today_counts=collections.Counter()
yesterday_counts=collections.Counter()
for r in queue:
    eid=value(r,hq,"ID_ENVIO")
    if not eid.startswith("MORA1-"):continue
    state=norm(value(r,hq,"ESTADO"))
    queue_counts[state]+=1
    dt=parse_date(value(r,hq,"FECHA_PROG"))
    if dt==today:today_counts[state]+=1
    if dt and (today-dt).days==1:yesterday_counts[state]+=1
print("COLA_ENVIO Mora 1 todos los días por estado:",dict(queue_counts))
print("COLA_ENVIO Mora 1 con FECHA_PROG HOY:",dict(today_counts))
print("COLA_ENVIO Mora 1 con FECHA_PROG AYER:",dict(yesterday_counts))
print("No se accedió a Gmail. No se modificó Google Sheets. No se enviaron correos.")

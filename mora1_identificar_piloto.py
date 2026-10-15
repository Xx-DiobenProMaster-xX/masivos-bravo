"""Lista referencias piloto de hoy, sin enviar ni modificar información."""
import json,os,re,collections
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from mora1_dioben import rows,value,norm,ref,parse_date,ID_DI0BEN,ID_MASIVOS
from mora1_reglas import puede_enviar
today=datetime.now(ZoneInfo("America/Bogota")).date()
creds=Credentials.from_service_account_info(json.loads(os.environ["MI_JSON"]),scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
gc=gspread.authorize(creds)
source=gc.open_by_key(ID_DI0BEN);dest=gc.open_by_key(ID_MASIVOS)
hs,follow=rows(source.worksheet("Seguimiento"))
hm,sel=rows(source.worksheet("Masivos_Correos"))
hc,queue=rows(dest.worksheet("COLA_ENVIO"))
hp,templates=rows(dest.worksheet("PLANTILLAS"))
he,clients=rows(dest.worksheet("CLIENTES"))
ha,campaigns=rows(dest.worksheet("CAMPAÑAS"))
selected={ref(value(r,hm,"REFERENCIA")) for r in sel if norm(value(r,hm,"HABILITADO"))=="SI" and norm(value(r,hm,"TIPO CORREO")) in {"MORA_1","MORA 1"}}
follow_map=collections.defaultdict(list);client_map=collections.defaultdict(list)
for row in follow:follow_map[ref(value(row,hs,"REFERENCIA"))].append(row)
for row in clients:client_map[ref(value(row,he,"REFERENCIA"))].append(row)
active={norm(value(t,hp,"ID_PLANTILLA")) for t in templates if norm(value(t,hp,"ESTADO"))=="ACTIVA"}
campaigns_map={value(c,ha,"ID_CAMPANA"):norm(value(c,ha,"ESTADO")) for c in campaigns}
all_ids=collections.Counter(value(r,hc,"ID_ENVIO") for r in queue)
ready=[]
for row in queue:
    send_id=value(row,hc,"ID_ENVIO")
    m=re.fullmatch(r"MORA1-(\d{8})-(T001|MORA10[1-5])-(\d+)",send_id)
    if not m or m.group(1)!=today.strftime("%Y%m%d") or all_ids[send_id]!=1:continue
    _,code,key=m.groups()
    if norm(value(row,hc,"ESTADO"))!="BORRADOR" or value(row,hc,"INTENTOS") not in {"","0"}:continue
    if value(row,hc,"ID_MENSAJE") or value(row,hc,"FECHA_ENVIO"):continue
    if key not in selected or len(follow_map[key])!=1 or len(client_map[key])!=1:continue
    current=follow_map[key][0]
    first=parse_date(value(current,hs,"FECHA_INICIO_MORA_1"))
    if not puede_enviar(value(current,hs,"MORA ACTUAL"),"SI","MORA_1",value(current,hs,"EXCLUIR CORREO"),first,today,code):continue
    if code not in active:continue
    if value(row,hc,"ID_CAMPANA")!=f"MORA1-{m.group(1)}-{code}" or campaigns_map.get(value(row,hc,"ID_CAMPANA"))!="BORRADOR":continue
    if parse_date(value(row,hc,"FECHA_PROG"))!=today:continue
    email=value(row,hc,"EMAIL")
    if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",email) or email.lower()!=value(client_map[key][0],he,"EMAIL").lower():continue
    if not value(row,hc,"ASUNTO") or not value(row,hc,"CUERPO"):continue
    conflicts=False
    for prior in queue:
        if value(prior,hc,"ID_ENVIO")==send_id:continue
        if ref(value(prior,hc,"REFERENCIA"))!=key or norm(value(prior,hc,"PLANTILLA"))!=code:continue
        date=parse_date(value(prior,hc,"FECHA_ENVIO")) or parse_date(value(prior,hc,"FECHA_PROG"))
        if (date is None or date >= (today if code=="T001" else first)) and norm(value(prior,hc,"ESTADO")) not in {"CANCELADO","CANCELADA","ANULADO","ANULADA"}:
            conflicts=True;break
    if not conflicts:ready.append((key,code))
print("Fecha:",today.isoformat())
print("Referencias candidatas para PILOTO hoy:",len(ready))
if ready:print("Primera referencia candidata:",ready[0][0],"/ plantilla:",ready[0][1])
else:print("Ninguna referencia cumple las reglas hoy. No forzar fecha ni envío.")
print("Una candidata aún se revalida al ejecutar el piloto.")
print("SOLO LECTURA. No se envió correo ni se modificaron las hojas.")

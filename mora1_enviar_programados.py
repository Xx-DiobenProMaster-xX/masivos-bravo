"""Envío independiente Mora 1: deshabilitado por defecto; valida cada cliente antes de Gmail."""
import os, json, re, base64
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials as ServiceCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from mora1_dioben import norm, ref, parse_date, rows, value, ID_DI0BEN, ID_MASIVOS, DIAS_ESPECIALES
from mora1_reglas import puede_enviar

TZ=ZoneInfo("America/Bogota")
TRUE={"TRUE","VERDADERO","SI","1","X"}
VALID={"T001",*DIAS_ESPECIALES.values()}
PATTERN=re.compile(r"^MORA1-(\d{8})-(T001|MORA10[1-5])-(\d+)$")
def getenv_required(k):
    v=os.getenv(k,"").strip()
    if not v:raise RuntimeError("Falta el secreto "+k)
    return v
def main():
    simulacion=os.getenv("MORA1_MODO_SIMULACION","").strip().upper()=="SI"
    if not simulacion and os.getenv("MORA1_ENVIOS_REALES","").strip().upper()!="SI":
        print("MORA1: ENVÍOS REALES DESACTIVADOS. Ningún mensaje enviado.")
        return
    limit=int(os.getenv("MORA1_LIMITE_ENVIOS","100"))
    if not 1<=limit<=500:raise ValueError("MORA1_LIMITE_ENVIOS debe estar entre 1 y 500")
    info=json.loads(getenv_required("MI_JSON"))
    creds=ServiceCredentials.from_service_account_info(info,scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc=gspread.authorize(creds)
    source=gc.open_by_key(ID_DI0BEN)
    target=gc.open_by_key(ID_MASIVOS)
    hs,follow=rows(source.worksheet("Seguimiento"))
    hm,selection=rows(source.worksheet("Masivos_Correos"))
    hc,queue=rows(target.worksheet("COLA_ENVIO"))
    hca,campaigns=rows(target.worksheet("CAMPAÑAS"))
    hp,templates=rows(target.worksheet("PLANTILLAS"))
    he,clients=rows(target.worksheet("CLIENTES"))
    client_map={}
    duplicate_clients=set()
    for item in clients:
        key=ref(value(item,he,"REFERENCIA"))
        if not key:continue
        if key in client_map:duplicate_clients.add(key)
        else:client_map[key]=item
    required={"ID_ENVIO","ID_CAMPANA","REFERENCIA","EMAIL","ASUNTO","CUERPO","PLANTILLA","ESTADO","FECHA_PROG","FECHA_ENVIO","INTENTOS","ERROR","ID_MENSAJE"}
    if not required.issubset(hc):raise RuntimeError("COLA_ENVIO faltan "+str(sorted(required-set(hc))))
    if not {"ID_CAMPANA","ESTADO"}.issubset(hca):raise RuntimeError("CAMPAÑAS sin ID_CAMPAÑA/ESTADO")
    today=datetime.now(TZ).date()
    selected={ref(value(r,hm,"REFERENCIA")) for r in selection if norm(value(r,hm,"HABILITADO"))=="SI" and norm(value(r,hm,"TIPO CORREO")) in {"MORA_1","MORA 1"}}
    follow_map={}
    duplicate_follow=set()
    for r in follow:
        k=ref(value(r,hs,"REFERENCIA"))
        if k in follow_map:duplicate_follow.add(k)
        else:follow_map[k]=r
    camp_map={value(r,hca,"ID_CAMPANA"):norm(value(r,hca,"ESTADO")) for r in campaigns}
    active_templates={norm(value(r,hp,"ID_PLANTILLA")) for r in templates if norm(value(r,hp,"ESTADO"))=="ACTIVA" and value(r,hp,"CUERPO") and value(r,hp,"ASUNTO")}
    from collections import Counter
    counts=Counter(value(row,hc,"ID_ENVIO") for row in queue if value(row,hc,"ID_ENVIO").startswith("MORA1-"))
    if any(v>1 for v in counts.values()):
        raise RuntimeError("IDs de Mora 1 duplicados en COLA_ENVIO; ningún envío autorizado.")
    candidates=[]
    seen=set()
    for row_number,r in enumerate(queue,start=2):
        sendid=value(r,hc,"ID_ENVIO")
        m=PATTERN.fullmatch(sendid)
        if not m or norm(value(r,hc,"ESTADO"))!="BORRADOR":continue
        if int(value(r,hc,"INTENTOS") or "0")!=0:continue
        stamp,template,reference=m.groups()
        if stamp!=today.strftime("%Y%m%d"):continue
        if (reference,template) in seen:continue
        seen.add((reference,template))
        startrow=follow_map.get(reference)
        if reference not in selected or reference in duplicate_follow or not startrow:continue
        if norm(value(startrow,hs,"MORA ACTUAL"))!="MORA 1" or norm(value(startrow,hs,"EXCLUIR CORREO")) in TRUE:continue
        first=parse_date(value(startrow,hs,"FECHA_INICIO_MORA_1"))
        if first is None or first>today:continue
        if not puede_enviar(value(startrow,hs,"MORA ACTUAL"),"SI","MORA_1",value(startrow,hs,"EXCLUIR CORREO"),first,today,template):continue
        if template not in active_templates:continue
        camp=value(r,hc,"ID_CAMPANA")
        if camp!=f"MORA1-{stamp}-{template}" or camp_map.get(camp)!="BORRADOR":continue
        if parse_date(value(r,hc,"FECHA_PROG"))!=today:continue
        if value(r,hc,"FECHA_ENVIO") or value(r,hc,"ID_MENSAJE"):continue
        if reference in duplicate_clients or reference not in client_map:continue
        email=value(r,hc,"EMAIL").strip()
        if email.lower()!=value(client_map[reference],he,"EMAIL").strip().lower():continue
        if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",email):continue
        if not value(r,hc,"ASUNTO") or not value(r,hc,"CUERPO"):continue
        # Misma referencia y plantilla ya enviada hoy, o especial desde comienzo de ciclo.
        duplicate=False
        for prior in queue:
            if value(prior,hc,"ID_ENVIO")==sendid:continue
            if ref(value(prior,hc,"REFERENCIA"))!=reference or norm(value(prior,hc,"PLANTILLA"))!=template:continue
            if norm(value(prior,hc,"ESTADO")) in {"CANCELADO","CANCELADA","ANULADO","ANULADA"}:continue
            dt=parse_date(value(prior,hc,"FECHA_ENVIO")) or parse_date(value(prior,hc,"FECHA_PROG"))
            if dt is None or dt>=(today if template=="T001" else first):duplicate=True;break
        if not duplicate:candidates.append((row_number,r))
    print("Fecha Bogotá:",today.isoformat())
    print("Elegibles tras revalidación:",len(candidates),"| Límite:",limit)
    print("Capacidad disponible según configuración:",max(0,limit-len(candidates)))
    if len(candidates)>limit:
        print("ALERTA CAPACIDAD: lote supera el límite configurado por",len(candidates)-limit,"mensajes.")
        print("ACCIÓN: ampliar MORA1_LIMITE_ENVIOS dentro de la cuota de Gmail o revisar el lote; no activar aún.")
    if simulacion:
        print("MODO SIMULACION: candidatos:",len(candidates))
        print("MODO SIMULACION: no se utilizó Gmail, ni se modificó Sheets, ni se envió correo.")
        return
    if len(candidates)>limit:
        raise RuntimeError("Lote excede el máximo configurado. Se bloqueó TODO el envío para evitar ejecuciones parciales; los borradores permanecen pendientes, pero NO se trasladan automáticamente al día siguiente.")
    if not candidates:return
    # Credenciales Gmail separadas: token OAuth de cuenta autorizada, nunca de servicio Sheets.
    oauth=OAuthCredentials(
        token=None,
        refresh_token=getenv_required("GMAIL_REFRESH_TOKEN"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=getenv_required("GMAIL_CLIENT_ID"),
        client_secret=getenv_required("GMAIL_CLIENT_SECRET"),
        scopes=["https://www.googleapis.com/auth/gmail.send"],
    )
    oauth.refresh(Request())
    gmail=build("gmail","v1",credentials=oauth,cache_discovery=False)
    ws=target.worksheet("COLA_ENVIO")
    sent=0
    from gspread.utils import rowcol_to_a1
    for row_number,r in candidates[:limit]:
        if datetime.now(TZ).date()!=today:
            print("CORTE: cambió la fecha de Bogotá.")
            break
        live=ws.row_values(row_number)
        if value(live,hc,"ID_ENVIO")!=value(r,hc,"ID_ENVIO"):continue
        if norm(value(live,hc,"ESTADO"))!="BORRADOR" or int(value(live,hc,"INTENTOS") or "0")!=0:continue
        if value(live,hc,"ID_MENSAJE") or value(live,hc,"FECHA_ENVIO"):continue
        if value(live,hc,"EMAIL")!=value(r,hc,"EMAIL"):continue
        # Releer cola antes de cada envío: detectar otros intentos o estados ambiguos.
        fresh=ws.get_all_values()[1:]
        reference=ref(value(r,hc,"REFERENCIA"))
        template=norm(value(r,hc,"PLANTILLA"))
        begin=parse_date(value(follow_map[reference],hs,"FECHA_INICIO_MORA_1"))
        blocked=False
        for other in fresh:
            if value(other,hc,"ID_ENVIO")==value(r,hc,"ID_ENVIO"):continue
            if ref(value(other,hc,"REFERENCIA"))!=reference or norm(value(other,hc,"PLANTILLA"))!=template:continue
            when=parse_date(value(other,hc,"FECHA_ENVIO")) or parse_date(value(other,hc,"FECHA_PROG"))
            if (when is None or when >= (today if template=="T001" else begin)) and norm(value(other,hc,"ESTADO")) not in {"CANCELADO","CANCELADA","ANULADO","ANULADA"}:
                blocked=True
                break
        if blocked:
            print("OMITIDO: otro envío/intento de la misma referencia.")
            continue
        msg=EmailMessage()
        msg["To"]=value(r,hc,"EMAIL")
        msg["From"]="estructurados@gobravo.com.co"
        msg["Reply-To"]="estructurados@gobravo.com.co"
        msg["Subject"]=value(r,hc,"ASUNTO")
        msg.set_content("Mensaje de BRAVO. Consulta la versión HTML de este correo.")
        msg.add_alternative(value(r,hc,"CUERPO"),subtype="html")
        raw=base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
        # Guardar intento ANTES del contacto con Gmail.
        ws.batch_update([
            {"range":rowcol_to_a1(row_number,hc["ESTADO"]+1),"values":[["ENVIANDO"]]},
            {"range":rowcol_to_a1(row_number,hc["INTENTOS"]+1),"values":[[1]]}
        ],value_input_option="RAW")
        locked=ws.row_values(row_number)
        if value(locked,hc,"ID_ENVIO")!=value(r,hc,"ID_ENVIO") or norm(value(locked,hc,"ESTADO"))!="ENVIANDO" or value(locked,hc,"INTENTOS")!="1":
            raise RuntimeError("No se confirmó la reserva en Sheets; envío cancelado.")
        try:
            result=gmail.users().messages().send(userId="me",body={"raw":raw}).execute()
        except Exception as exc:
            print("AVISO: envío Gmail fallido o ambiguo. Se conserva ENVIANDO sin reintento automático:",type(exc).__name__)
            continue
        message_id=str(result.get("id","")).strip()
        if not message_id:
            print("AVISO: Gmail no devolvió ID; conservar ENVIANDO y conciliar manualmente.")
            continue
        try:
            # Evidencia primero; ESTADO final después.
            ws.batch_update([
                {"range":rowcol_to_a1(row_number,hc["ID_MENSAJE"]+1),"values":[[message_id]]},
                {"range":rowcol_to_a1(row_number,hc["FECHA_ENVIO"]+1),"values":[[datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S")]]}
            ],value_input_option="RAW")
            ws.update_cell(row_number,hc["ESTADO"]+1,"ENVIADO")
            sent+=1
        except Exception as exc:
            print("CRÍTICO: Gmail aceptó el mensaje pero Sheets falló; NO restaurar a BORRADOR:",type(exc).__name__)
            raise
    print("Mensajes enviados:",sent)

if __name__=="__main__":main()

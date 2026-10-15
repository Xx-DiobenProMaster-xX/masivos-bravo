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

TZ=ZoneInfo("America/Bogota")
TRUE={"TRUE","VERDADERO","SI","1","X"}
VALID={"T001",*DIAS_ESPECIALES.values()}
PATTERN=re.compile(r"^MORA1-(\d{8})-(T001|MORA10[1-5])-(\d+)$")
def getenv_required(k):
    v=os.getenv(k,"").strip()
    if not v:raise RuntimeError("Falta el secreto "+k)
    return v
def main():
    if os.getenv("MORA1_ENVIOS_REALES","").strip().upper()!="SI":
        print("MORA1: ENVÍOS REALES DESACTIVADOS. Ningún mensaje enviado.")
        return
    limit=int(os.getenv("MORA1_LIMITE_ENVIOS","1"))
    if not 1<=limit<=100:raise ValueError("MORA1_LIMITE_ENVIOS debe estar entre 1 y 100")
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
    candidates=[]
    seen=set()
    for row_number,r in enumerate(queue,start=2):
        sendid=value(r,hc,"ID_ENVIO")
        m=PATTERN.fullmatch(sendid)
        if not m or norm(value(r,hc,"ESTADO"))!="BORRADOR":continue
        stamp,template,reference=m.groups()
        if stamp!=today.strftime("%Y%m%d"):continue
        if (reference,template) in seen:continue
        seen.add((reference,template))
        startrow=follow_map.get(reference)
        if reference not in selected or reference in duplicate_follow or not startrow:continue
        if norm(value(startrow,hs,"MORA ACTUAL"))!="MORA 1" or norm(value(startrow,hs,"EXCLUIR CORREO")) in TRUE:continue
        first=parse_date(value(startrow,hs,"FECHA_INICIO_MORA_1"))
        if first is None or first>today:continue
        day=(today-first).days+1
        expected=DIAS_ESPECIALES.get(day) or ("T001" if day%2==0 else None)
        if expected!=template or template not in active_templates:continue
        camp=value(r,hc,"ID_CAMPANA")
        if camp!=f"MORA1-{stamp}-{template}" or camp_map.get(camp)!="PROGRAMADA":continue
        if parse_date(value(r,hc,"FECHA_PROG"))!=today:continue
        if value(r,hc,"FECHA_ENVIO") or value(r,hc,"ID_MENSAJE"):continue
        email=value(r,hc,"EMAIL").strip()
        if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",email):continue
        if not value(r,hc,"ASUNTO") or not value(r,hc,"CUERPO"):continue
        # Misma referencia y plantilla ya enviada hoy, o especial desde comienzo de ciclo.
        duplicate=False
        for prior in queue:
            if value(prior,hc,"ID_ENVIO")==sendid:continue
            if ref(value(prior,hc,"REFERENCIA"))!=reference or norm(value(prior,hc,"PLANTILLA"))!=template:continue
            if norm(value(prior,hc,"ESTADO")) not in {"ENVIADO","ENVIANDO","PROGRAMADA","PENDIENTE","BORRADOR"}:continue
            dt=parse_date(value(prior,hc,"FECHA_ENVIO")) or parse_date(value(prior,hc,"FECHA_PROG"))
            if dt is None or dt>=(today if template=="T001" else first):duplicate=True;break
        if not duplicate:candidates.append((row_number,r))
    print("Elegibles tras revalidación:",len(candidates),"| Límite:",limit)
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
    for row_number,r in candidates[:limit]:
        # Comprobar estado en Sheet inmediatamente antes de cada intento.
        if norm(ws.cell(row_number,hc["ESTADO"]+1).value)!="BORRADOR":continue
        attempts=int(value(r,hc,"INTENTOS") or "0")
        ws.update_cell(row_number,hc["ESTADO"]+1,"ENVIANDO")
        ws.update_cell(row_number,hc["INTENTOS"]+1,attempts+1)
        try:
            msg=EmailMessage()
            msg["To"]=value(r,hc,"EMAIL")
            msg["Subject"]=value(r,hc,"ASUNTO")
            msg.set_content("Mensaje de BRAVO. Consulta la versión HTML de este correo.")
            msg.add_alternative(value(r,hc,"CUERPO"),subtype="html")
            result=gmail.users().messages().send(userId="me",body={"raw":base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")}).execute()
            ws.update_cell(row_number,hc["ID_MENSAJE"]+1,str(result.get("id","")))
            ws.update_cell(row_number,hc["FECHA_ENVIO"]+1,datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S"))
            ws.update_cell(row_number,hc["ESTADO"]+1,"ENVIADO")
            sent+=1
        except Exception as exc:
            # No reintentar automáticamente un estado ambiguo.
            ws.update_cell(row_number,hc["ERROR"]+1,str(exc)[:400])
            ws.update_cell(row_number,hc["ESTADO"]+1,"ERROR")
            print("Error en una fila; requiere revisión manual")
    print("Mensajes enviados:",sent)

if __name__=="__main__":main()

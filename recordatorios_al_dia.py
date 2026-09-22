
import os, re, json, base64, html
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

TZ=ZoneInfo("America/Bogota")
MASIVOS="1VGdEUGRDFxBjKRLF1KF7EcHIBf3f8ujtN3iPm6TatjI"
FUENTE="15sbBsZcMj8PMkHXByLqjcuqtvsY_2FGYiwhkKPmfIYM"
CARTERA="13Vf32LzRI2V95dIUqfevzm-ZmsDR3d17UTre_7XJ-UU"
FROM="estructurados@gobravo.com.co"
LOGO="https://drive.google.com/uc?export=view&id=13kK3v4FiyXFa4UzM_au3TllhOhwjvWb7"
WA="573012411885"

def norm(x):
    import unicodedata
    t=unicodedata.normalize("NFD",str(x or "").strip())
    return "".join(c for c in t if unicodedata.category(c)!="Mn").upper()

def ref(x):
    t=str(x or "").strip().replace(" ","").replace("\xa0","")
    if re.fullmatch(r"\d+\.0+",t): t=t.split(".")[0]
    if re.fullmatch(r"\d{1,3}([.,]\d{3})+",t): t=re.sub(r"[.,]","",t)
    return re.sub(r"\D","",t)

def fecha(x):
    for f in ("%d/%m/%Y","%Y-%m-%d","%d/%m/%Y %H:%M:%S"):
        try:return datetime.strptime(str(x).strip(),f).date()
        except:pass
    return None

def fecha_larga(d):
    m=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    return f"{d.day} de {m[d.month-1]} de {d.year}"

def html_mail(nombre,d,dias):
    nombre=html.escape(nombre or "Cliente"); ft=fecha_larga(d)
    if dias==0:
        intro="Te recordamos que hoy es la fecha de tu apartado mensual en Bravo."
        destacado="Hoy es la fecha de tu apartado mensual"; fondo="#e9f7ff"; acento="#147fd1"
    else:
        intro="Queremos recordarte que se acerca la fecha de tu apartado mensual en Bravo."
        destacado="Faltan 3 días para la fecha de tu apartado mensual"; fondo="#f1edff"; acento="#5b45c6"
    return f'''<!doctype html><html><body style="margin:0;background:#f4f5f9;font-family:Arial;color:#525b82">
<table width="100%"><tr><td align="center"><table width="700" style="max-width:700px;background:#fff;border-top:6px solid #38278f">
<tr><td style="padding:28px 48px"><img src="{LOGO}" width="170"></td></tr>
<tr><td style="padding:5px 48px;font-size:30px;font-weight:800;color:#11183f">Hola, <span style="color:#35238f">{nombre}</span> 👋</td></tr>
<tr><td style="padding:12px 48px 24px;font-size:18px;line-height:28px">{intro}</td></tr>
<tr><td style="padding:0 48px 24px"><table width="100%" style="background:{fondo};border-radius:18px"><tr>
<td style="padding:28px;font-size:55px">📅</td><td style="padding:28px 28px 28px 0">
<div style="font-size:14px;font-weight:800;color:{acento};text-transform:uppercase">Fecha de tu apartado mensual</div>
<div style="font-size:29px;font-weight:800;color:#11183f;margin-top:6px">{ft}</div>
<div style="background:#fff;border-radius:12px;padding:13px 16px;margin-top:18px;font-size:18px;font-weight:800;color:{acento}">{destacado}</div>
</td></tr></table></td></tr>
<tr><td style="padding:0 48px 20px;font-size:17px;line-height:27px">Tener presente la fecha de tu apartado mensual te ayuda a mantener tu programa al día y continuar avanzando en tu proceso.</td></tr>
<tr><td style="padding:0 48px 28px"><a href="https://wa.me/{WA}" style="display:block;background:#12bd70;color:white;text-align:center;padding:18px;border-radius:14px;text-decoration:none;font-size:19px;font-weight:800">Hablar con Bravo por WhatsApp</a></td></tr>
<tr><td align="center" style="padding:22px;border-top:1px solid #ddd"><b>¡Gracias por ser parte de Bravo!</b><br>Equipo Bravo</td></tr>
</table></td></tr></table></body></html>'''

def sheets():
    info=json.loads(os.environ["MI_JSON"])
    c=SACredentials.from_service_account_info(info,scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"])
    return gspread.authorize(c)

def gmail():
    c=OAuthCredentials(token=None,refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],token_uri="https://oauth2.googleapis.com/token",client_id=os.environ["GOOGLE_CLIENT_ID"],client_secret=os.environ["GOOGLE_CLIENT_SECRET"],scopes=["https://www.googleapis.com/auth/gmail.send"])
    c.refresh(Request()); return build("gmail","v1",credentials=c,cache_discovery=False)

def enviar(svc,to,subject,body):
    m=EmailMessage(); m["To"]=to; m["From"]=f"Bravo S.A.S. <{FROM}>"; m["Reply-To"]=FROM; m["Subject"]=subject
    m.set_content("Recordatorio de tu apartado mensual en Bravo."); m.add_alternative(body,subtype="html")
    raw=base64.urlsafe_b64encode(m.as_bytes()).decode()
    return svc.users().messages().send(userId="me",body={"raw":raw}).execute()

def main():
    gc=sheets(); hoy=datetime.now(TZ).date(); ahora=datetime.now(TZ)
    fuente=gc.open_by_key(FUENTE)
    unidos=fuente.worksheet("Unidos_Est").get("A:I")
    excl={ref(x) for x in fuente.worksheet("Excluir_correo").col_values(1)[1:] if ref(x)}

    car=gc.open_by_key(CARTERA).worksheet("2. Cartera Berex").get("B:G")
    hc=[str(x).strip() for x in car[0]]; pc={c:i for i,c in enumerate(hc)}
    maestro={}
    for f in car[1:]:
        def v(c): return f[pc[c]] if c in pc and len(f)>pc[c] else ""
        for c in ("Referencia","Referencia_Berex","Numero"):
            r=ref(v(c))
            if r:
                maestro[r]={"NOMBRE":str(v("Nombre_Cliente")).strip(),"EMAIL":str(v("Email")).strip()}

    libro=gc.open_by_key(MASIVOS); cola=libro.worksheet("COLA_ENVIO"); vals=cola.get_all_values()
    cab=[str(x).strip() for x in vals[0]]; ix={c:i for i,c in enumerate(cab)}
    existentes={str(f[ix["ID_ENVIO"]]).strip() for f in vals[1:] if len(f)>ix["ID_ENVIO"]}
    svc=gmail(); enviados=0; omitidos=0

    for f in unidos[1:]:
        r=ref(f[0] if len(f)>0 else ""); status=str(f[2] if len(f)>2 else "").strip(); d=fecha(f[7] if len(f)>7 else "")
        if not r or norm(status)!="AL DIA" or not d: continue
        dias=(d-hoy).days
        if dias not in (0,3): continue
        plantilla="ALDIA000" if dias==0 else "ALDIA003"
        ide=f"ENV-ALDIA-{d.strftime('%Y%m%d')}-{r}-{plantilla}"
        if r in excl or ide in existentes: omitidos+=1; continue
        cli=maestro.get(r,{})
        correo=str(cli.get("EMAIL","")).strip()
        if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+",correo): omitidos+=1; continue
        nombre=str(cli.get("NOMBRE","")).strip() or "Cliente"
        asunto="Hoy es la fecha de tu apartado mensual | Bravo" if dias==0 else "Se acerca la fecha de tu apartado mensual | Bravo"
        cuerpo=html_mail(nombre,d,dias)
        camp=f"ALDIA-AUTO-{hoy.strftime('%Y%m%d')}"
        datos={"ID_ENVIO":ide,"ID_CAMPAÑA":camp,"REFERENCIA":r,"NOMBRE":nombre,"EMAIL":correo,"PLANTILLA":plantilla,"ASUNTO":asunto,"ESTADO":"ENVIANDO","FECHA_PROG":ahora.strftime("%d/%m/%Y %H:%M"),"FECHA_ENVIO":"","INTENTOS":1,"ERROR":"","ID_MENSAJE":"","CUERPO":cuerpo,"ENCARGADO":""}
        fila=[datos.get(c,"") for c in cab]; cola.append_row(fila,value_input_option="USER_ENTERED")
        n=len(cola.get_all_values())
        try:
            resp=enviar(svc,correo,asunto,cuerpo)
            cola.update_cell(n,ix["FECHA_ENVIO"]+1,datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S"))
            cola.update_cell(n,ix["ID_MENSAJE"]+1,str(resp.get("id","")))
            cola.update_cell(n,ix["ESTADO"]+1,"ENVIADO"); enviados+=1
            print("ENVIADO",plantilla,r,correo)
        except Exception as e:
            cola.update_cell(n,ix["ERROR"]+1,str(e)[:500]); cola.update_cell(n,ix["ESTADO"]+1,"ERROR")
            print("ERROR",r,e)
    print(f"FIN enviados={enviados} omitidos={omitidos}")

if __name__=="__main__": main()

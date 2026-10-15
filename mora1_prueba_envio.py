"""Envía exactamente un correo de prueba a un destino indicado manualmente. No lee Sheets."""
import os,re,base64
from email.message import EmailMessage
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

address=os.environ.get("MORA1_TEST_EMAIL","").strip()
confirmation=os.environ.get("MORA1_TEST_CONFIRMACION","")
if confirmation!="ENVIAR_PRUEBA":
    raise SystemExit("Prueba cancelada: escribe ENVIAR_PRUEBA para autorizar un solo correo.")
if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",address):
    raise SystemExit("Dirección de prueba inválida.")
for secret in ("GMAIL_CLIENT_ID","GMAIL_CLIENT_SECRET","GMAIL_REFRESH_TOKEN"):
    if not os.getenv(secret):
        raise SystemExit("Falta secreto "+secret)
creds=Credentials(
    token=None,refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
    token_uri="https://oauth2.googleapis.com/token",
    client_id=os.environ["GMAIL_CLIENT_ID"],client_secret=os.environ["GMAIL_CLIENT_SECRET"],
    scopes=["https://www.googleapis.com/auth/gmail.send"],
)
service=build("gmail","v1",credentials=creds,cache_discovery=False)
msg=EmailMessage()
msg["To"]=address
msg["Subject"]="[PRUEBA] Automatización Mora 1 - Bravo"
msg.set_content("Correo de prueba técnica de Mora 1. No es una notificación de cobro. No requiere ninguna acción.")
msg.add_alternative("<p><strong>PRUEBA TÉCNICA</strong> — Automatización Mora 1.</p><p>Este correo no es una notificación de cobro y no requiere ninguna acción.</p>",subtype="html")
response=service.users().messages().send(userId="me",body={"raw":base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")}).execute()
if not response.get("id"):
    raise SystemExit("Gmail no devolvió ID de mensaje; revisar estado.")
print("OK: Gmail aceptó exactamente un correo de prueba.")
print("ID del mensaje:",response["id"])
print("Revisa en el mensaje recibido la dirección del remitente.")
print("No se consultó ni modificó Google Sheets; no hubo mensajes para clientes de Mora 1.")

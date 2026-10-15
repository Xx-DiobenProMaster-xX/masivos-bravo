"""Verifica OAuth Gmail sin enviar mensajes y sin mostrar tokens."""
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

names=("GMAIL_CLIENT_ID","GMAIL_CLIENT_SECRET","GMAIL_REFRESH_TOKEN")
missing=[name for name in names if not os.environ.get(name, "").strip()]
if missing:
    raise SystemExit("ERROR: faltan secretos en GitHub: "+", ".join(missing))

creds=Credentials(
    token=None,
    refresh_token=os.environ["GMAIL_REFRESH_TOKEN"],
    token_uri="https://oauth2.googleapis.com/token",
    client_id=os.environ["GMAIL_CLIENT_ID"],
    client_secret=os.environ["GMAIL_CLIENT_SECRET"],
    scopes=["https://www.googleapis.com/auth/gmail.send"],
)
try:
    creds.refresh(Request())
    if not creds.valid:
        raise RuntimeError("El acceso OAuth no es válido")
    gmail=build("gmail","v1",credentials=creds,cache_discovery=False)
    profile=gmail.users().getProfile(userId="me").execute()
    email=profile.get("emailAddress","")
    masked=(email[:2]+"***@"+email.split("@",1)[1]) if "@" in email else "(no disponible)"
    print("OK: autenticación Gmail funcionando.")
    print("Cuenta autenticada (parcialmente oculta):",masked)
    print("Solo se consultó el perfil. NO SE ENVIÓ NINGÚN CORREO.")
except Exception as error:
    print("ERROR: no fue posible autenticar Gmail o consultar el perfil.")
    print("Tipo de error:",type(error).__name__)
    raise SystemExit(1)

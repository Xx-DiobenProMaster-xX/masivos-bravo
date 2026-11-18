"""Plantilla independiente de Alianzas - Deposito Extra (TALDE001)."""
import html
from decimal import Decimal, InvalidOperation

PLANTILLA_ID = "TALDE001"

def monto_de(valor):
    texto = str(valor or "").strip().replace("$","").replace(" ","")
    if not texto:
        return None
    # Google Sheets puede devolver 1600000, 1.600.000 o 1600000,00.
    if "," in texto and "." in texto:
        texto = texto.replace(".","").replace(",",".")
    elif "," in texto:
        partes=texto.split(",")
        texto=texto.replace(",",".") if len(partes[-1])<=2 else texto.replace(",","")
    elif "." in texto and len(texto.split(".")[-1])==3:
        texto=texto.replace(".","")
    try:
        n=Decimal(texto)
    except InvalidOperation:
        return None
    return n if n>0 else None

def dinero_colombiano(valor):
    n=monto_de(valor)
    return "$ " + format(int(n),",").replace(",",".") if n is not None else ""

def html_deposito_extra(nombre, banco, de_tradicional, logo_url, whatsapp):
    nombre=html.escape(str(nombre or "Cliente").strip() or "Cliente")
    banco=html.escape(str(banco or "").strip())
    monto=html.escape(dinero_colombiano(de_tradicional))
    logo=html.escape(str(logo_url),quote=True)
    wa=html.escape(str(whatsapp),quote=True)
    if not banco or not monto:
        raise ValueError("Deposito Extra exige banco y DE_Tradicional mayor a cero.")
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="margin:0;background:#f5f6fb;color:#40466c;font-family:Arial,Helvetica,sans-serif">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f5f6fb"><tr><td align="center" style="padding:24px 10px">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px;width:100%;background:#fff;border-top:5px solid #3a288f">
<tr><td style="padding:22px 34px 8px"><table role="presentation" width="100%"><tr><td><img src="{logo}" alt="Bravo" width="132" style="max-width:132px;height:auto;border:0;display:block"><span style="font-size:10px;color:#797e99">Soluciones financieras para un mejor futuro</span></td><td align="right" style="font-size:11px;color:#3a288f;font-weight:bold">UNA OPORTUNIDAD<br>PARA TI</td></tr></table></td></tr>
<tr><td style="padding:18px 34px 10px"><h1 style="margin:0;font-size:32px;line-height:38px;color:#3a288f">¡Grandes noticias!<br><span style="color:#09b4ba">Logramos una oferta<br>con {banco}</span> 🚀</h1></td></tr>
<tr><td style="padding:0 34px 10px;font-size:16px;line-height:25px">Hola <strong style="color:#38268d">{nombre},</strong></td></tr>
<tr><td style="padding:0 34px 18px;font-size:15px;line-height:24px">Tenemos un descuento especial con nuestros aliados estratégicos para saldar tu obligación con <strong>{banco}</strong>. Con este beneficio lograrías liquidar tu deuda con un monto menor al previsto en tu plan de liquidación. Para aprovechar esta oportunidad, con el ahorro actual estimamos que necesitarías un depósito adicional de <strong>{monto}</strong>.</td></tr>
<tr><td style="padding:0 34px 18px"><table role="presentation" width="100%" style="background:#f3f0ff;border-radius:12px"><tr><td style="padding:20px 14px;font-size:29px;width:36px">🏷️</td><td style="padding:16px 10px"><div style="font-size:12px;color:#5d5b8b">Aporte adicional estimado</div><div style="font-size:26px;font-weight:bold;color:#39268f">{monto}</div><div style="font-size:11px;color:#6a688b">*Valor aproximado sujeto a fecha de pago.</div></td></tr></table></td></tr>
<tr><td style="padding:0 34px 18px"><table role="presentation" width="100%" style="background:#effbf7;border:1px solid #b9eee0;border-radius:12px"><tr><td style="padding:18px 20px"><strong style="font-size:12px;color:#00a8b0">BENEFICIO EXCLUSIVO</strong><h2 style="color:#3a288f;margin:6px 0 12px;font-size:21px">¿Por qué tomar esta oferta?</h2>
<p style="margin:8px 0;font-size:14px;line-height:21px"><strong>• Máximo descuento posible:</strong> Accedes a un descuento superior al proyectado inicialmente.</p>
<p style="margin:8px 0;font-size:14px;line-height:21px"><strong>• Recupera tu tranquilidad financiera:</strong> Avanzas hacia una liquidación de tu obligación.</p>
<p style="margin:8px 0;font-size:14px;line-height:21px"><strong>• ¿No cuentas con el monto completo?</strong> Escríbenos para revisar alternativas de pago.</p>
</td></tr></table></td></tr>
<tr><td style="padding:0 34px 14px"><table role="presentation" width="100%" style="background:#e7fbfc;border-radius:12px"><tr><td style="padding:19px;font-size:27px;width:40px">🤝</td><td style="padding:17px 8px"><strong style="font-size:19px;color:#00a9b3">Contacta a tu negociador</strong><div style="font-size:13px;line-height:20px">Estamos listos para confirmar el valor exacto y compartirte una propuesta.</div></td></tr></table></td></tr>
<tr><td style="padding:0 34px 17px;font-size:13px;line-height:21px">Queremos ayudarte a aprovechar esta oportunidad antes de que venza el plazo fijado por el aliado. Comunícate hoy mismo con nosotros.</td></tr>
<tr><td style="padding:0 34px 28px"><a href="https://wa.me/{wa}" style="display:block;text-align:center;background:#0ab3b9;border-radius:10px;padding:16px;color:white;text-decoration:none;font-size:17px;font-weight:bold">☎ Hablar con mi negociador →</a></td></tr>
<tr><td style="padding:15px 34px;border-top:1px solid #e2e4ed;font-size:12px;color:#6b6f8c"><strong style="color:#3a288f">Bravo S.A.S.</strong><br>Comunicación informativa. Valores sujetos a validación con el aliado.<br>Lunes a viernes, 8:00 a. m. – 6:00 p. m.</td></tr>
<tr><td style="height:8px;background:#11bfc6"></td></tr></table></td></tr></table></body></html>"""

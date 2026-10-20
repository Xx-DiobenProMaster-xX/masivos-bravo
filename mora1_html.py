"""HTML seguro para correos Mora 1 guardados como texto o Markdown en PLANTILLAS."""
import html
import re

def html_mora1(cuerpo):
    texto=str(cuerpo or "").strip()
    if not texto:
        raise ValueError("Plantilla Mora 1 sin cuerpo")
    # Si ya se guardó como HTML, respetarlo y no envolverlo de nuevo.
    if re.search(r"<\s*(?:!doctype|html|body|table|div|p|h[1-6]|br)\b",texto,re.I):
        return texto
    # PLANTILLAS T001 antiguas guardan **negritas** literales: convertir a HTML real.
    limpio=html.escape(texto)
    limpio=re.sub(r"\*\*(.+?)\*\*",r"<strong>\1</strong>",limpio,flags=re.S)
    limpio=re.sub(r"(?m)^\s*[-•]\s+(.+)$",r"• \1",limpio)
    parrafos=[x.strip() for x in re.split(r"\n\s*\n",limpio) if x.strip()]
    contenido="".join('<p style="margin:0 0 16px;line-height:1.7;font-size:15px;color:#333333">'+p.replace("\n","<br>")+"</p>" for p in parrafos)
    return (
        '<!doctype html><html lang="es"><head><meta charset="utf-8"></head>'
        '<body style="margin:0;padding:24px;background:#f4f5fa;font-family:Arial,Helvetica,sans-serif">'
        '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f4f5fa"><tr><td align="center">'
        '<table role="presentation" cellpadding="0" cellspacing="0" width="600" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden">'
        '<tr><td style="background:#3e2a84;color:white;padding:25px 30px;font-size:22px;font-weight:700">BRAVO'
        '<div style="font-weight:400;font-size:13px;color:#b7f4ee;margin-top:5px">Apartado mensual</div></td></tr>'
        '<tr><td style="padding:30px">'+contenido+'</td></tr>'
        '<tr><td style="padding:18px 30px;background:#f1f6f8;border-top:3px solid #23b5ad;font-size:12px;color:#596171">'
        'BRAVO S.A.S. · Mensaje informativo. Si ya realizaste tu pago, puedes responder a este correo.</td></tr>'
        '</table></td></tr></table></body></html>'
    )

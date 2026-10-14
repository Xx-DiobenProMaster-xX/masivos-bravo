"""Preparación de BORRADORES Mora 1; nunca envía mensajes ni cambia campañas existentes."""
import os, json, re
from datetime import datetime
from zoneinfo import ZoneInfo
import gspread
from google.oauth2.service_account import Credentials
from mora1_dioben import norm, ref, parse_date, rows, value, ID_DI0BEN, ID_MASIVOS, DIAS_ESPECIALES

TZ = ZoneInfo("America/Bogota")
CANCELADOS = {"ERROR", "CANCELADO", "CANCELADA", "BLOQUEADO", "BLOQUEADA", "FALLIDO", "FALLIDA"}
ACTIVOS = {"BORRADOR", "PREPARADA", "PREPARADO", "PENDIENTE", "PROGRAMADO", "PROGRAMADA", "ENVIANDO", "ENVIADO", "ENVIADA", "EN COLA", "LISTO", "LISTA"}
TRUE = {"TRUE", "VERDADERO", "SI", "1", "X"}

def main():
    if os.getenv("MORA1_PREPARAR_BORRADORES") != "SI":
        print("MODO SEGURO: PREPARACIÓN DESHABILITADA. Para habilitar, configurar MORA1_PREPARAR_BORRADORES=SI.")
        return
    info=json.loads(os.environ["MI_JSON"])
    creds=Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc=gspread.authorize(creds)
    origen=gc.open_by_key(ID_DI0BEN)
    destino=gc.open_by_key(ID_MASIVOS)
    hs,seguimiento=rows(origen.worksheet("Seguimiento"))
    hm,seleccion=rows(origen.worksheet("Masivos_Correos"))
    hp,plantillas=rows(destino.worksheet("PLANTILLAS"))
    hc,cola=rows(destino.worksheet("COLA_ENVIO"))
    he,clientes=rows(destino.worksheet("CLIENTES"))
    for name,h,cols in [
        ("Seguimiento",hs,{"REFERENCIA","MORA ACTUAL","EXCLUIR CORREO","FECHA_INICIO_MORA_1"}),
        ("Masivos_Correos",hm,{"REFERENCIA","HABILITADO","TIPO CORREO"}),
        ("PLANTILLAS",hp,{"ID_PLANTILLA","ASUNTO","CUERPO","ESTADO"}),
        ("COLA_ENVIO",hc,{"ID_ENVIO","REFERENCIA","NOMBRE","EMAIL","PLANTILLA","ASUNTO","CUERPO","ESTADO","FECHA_PROG"}),
        ("CLIENTES",he,{"REFERENCIA","EMAIL","NOMBRE"}),
    ]:
        if not cols.issubset(h):
            raise RuntimeError(name + ": faltan columnas " + str(sorted(cols-set(h))))
    if "ID_CAMPAÑA" not in hc:
        print("AVISO: COLA_ENVIO no tiene ID_CAMPAÑA; se respetará su estructura.")
        print("Los borradores NO podrán enviarse con el motor actual hasta corregir esa incompatibilidad.")
    print("Columnas de COLA_ENVIO:", ", ".join(sorted(hc, key=hc.get)))
    hoy=datetime.now(TZ).date()
    bases={}
    duplicadas=set()
    for row in seguimiento:
        k=ref(value(row,hs,"REFERENCIA"))
        if k in bases:duplicadas.add(k)
        elif k:bases[k]=row
    clientes_por_ref={}
    clientes_duplicados=set()
    for row in clientes:
        k=ref(value(row,he,"REFERENCIA"))
        if k in clientes_por_ref:clientes_duplicados.add(k)
        elif k:clientes_por_ref[k]=row
    templates={}
    duplicadas_tpl=set()
    for row in plantillas:
        code=norm(value(row,hp,"ID_PLANTILLA"))
        if code in {"T001",*DIAS_ESPECIALES.values()}:
            if code in templates:duplicadas_tpl.add(code)
            else:templates[code]=row
    history={}
    for row in cola:
        k=(ref(value(row,hc,"REFERENCIA")),norm(value(row,hc,"PLANTILLA")))
        state=norm(value(row,hc,"ESTADO"))
        dt=parse_date(value(row,hc,"FECHA_ENVIO")) if "FECHA_ENVIO" in hc else None
        prog=parse_date(value(row,hc,"FECHA_PROG"))
        if k[0]:history.setdefault(k,[]).append((state,dt or prog))
    batch=[]
    seen=set()
    reasons={}
    def skip(why):reasons[why]=reasons.get(why,0)+1
    for row in seleccion:
        k=ref(value(row,hm,"REFERENCIA"))
        if not k or k in seen:continue
        seen.add(k)
        if norm(value(row,hm,"HABILITADO"))!="SI" or norm(value(row,hm,"TIPO CORREO")) not in {"MORA_1","MORA 1"}:
            skip("NO SELECCIONADO");continue
        if k in duplicadas or k not in bases:
            skip("SEGUIMIENTO FALTANTE O DUPLICADO");continue
        s=bases[k]
        if norm(value(s,hs,"MORA ACTUAL"))!="MORA 1" or norm(value(s,hs,"EXCLUIR CORREO")) in TRUE:
            skip("MORA CAMBIADA O EXCLUIDO");continue
        inicio=parse_date(value(s,hs,"FECHA_INICIO_MORA_1"))
        if not inicio or inicio>hoy:
            skip("FECHA CICLO NO CONFIRMADA");continue
        day=(hoy-inicio).days+1
        code=DIAS_ESPECIALES.get(day) or ("T001" if day%2==0 else None)
        if not code:skip("SIN ENVIO HOY");continue
        if code in duplicadas_tpl or code not in templates:
            skip("PLANTILLA FALTANTE O DUPLICADA");continue
        tpl=templates[code]
        if norm(value(tpl,hp,"ESTADO"))!="ACTIVA" or not value(tpl,hp,"ASUNTO") or not value(tpl,hp,"CUERPO"):
            skip("PLANTILLA NO ACTIVA O INCOMPLETA");continue
        blocked=False
        for state,dt in history.get((k,code),[]):
            if state in CANCELADOS:continue
            if state not in ACTIVOS or dt is None or dt >= (hoy if code=="T001" else inicio):
                blocked=True;break
        if blocked:skip("YA EXISTE EN COLA / ESTADO INCIERTO");continue
        if k in clientes_duplicados or k not in clientes_por_ref:
            skip("CLIENTE FALTANTE O DUPLICADO");continue
        c=clientes_por_ref[k]
        email=value(c,he,"EMAIL")
        if not re.fullmatch(r"[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+",email):
            skip("EMAIL NO VALIDO");continue
        nombre=value(c,he,"NOMBRE") or "Cliente"
        subj=value(tpl,hp,"ASUNTO")
        body=value(tpl,hp,"CUERPO")
        for placeholder,replace in {"{{NOMBRE}}":nombre,"{{REFERENCIA}}":k}.items():
            subj=subj.replace(placeholder,replace)
            body=body.replace(placeholder,replace)
        if re.search(r"\{\{[^{}]+\}\}",subj+body):
            skip("VARIABLES SIN RESOLVER");continue
        send_id=f"MORA1-{hoy:%Y%m%d}-{code}-{k}"
        existing_ids={value(z,hc,"ID_ENVIO") for z in cola}
        if send_id in existing_ids:
            skip("ID YA EXISTE");continue
        values={
            "ID_ENVIO":send_id,"ID_CAMPAÑA":f"MORA1-{hoy:%Y%m%d}",
            "REFERENCIA":k,"NOMBRE":nombre,"EMAIL":email,"PLANTILLA":code,
            "ASUNTO":subj,"CUERPO":body,"ESTADO":"BORRADOR",
            "FECHA_PROG":hoy.strftime("%d/%m/%Y")+" 08:15",
            "FECHA_ENVIO":"","INTENTOS":0,"ERROR":"","ID_MENSAJE":""
        }
        batch.append([values.get(header,"") for header,_ in sorted(hc.items(),key=lambda kv:kv[1])])
    print("Candidatos preparados como BORRADOR:",len(batch))
    print("Omisiones:",json.dumps(reasons,ensure_ascii=False))
    if not batch:return
    # Antes de escribir, leer nuevamente para evitar duplicación si otro proceso creó registros.
    ws=destino.worksheet("COLA_ENVIO")
    fresh=ws.get_all_values()
    ids={r[hc["ID_ENVIO"]] for r in fresh[1:] if len(r)>hc["ID_ENVIO"]}
    valid=[r for r in batch if r[hc["ID_ENVIO"]] not in ids]
    if valid:
        ws.append_rows(valid,value_input_option="RAW")
    print("Borradores agregados:",len(valid))
    print("NO SE ENVIARON CORREOS. Los BORRADORES requieren autorización explícita para enviarse.")

if __name__=="__main__":
    main()

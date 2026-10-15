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


def vincular_campanas_borrador(destino):
    """Vincula solo borradores propios de Mora 1. No programa ni envía."""
    import collections
    from gspread.utils import rowcol_to_a1
    ws=destino.worksheet("COLA_ENVIO")
    wc=destino.worksheet("CAMPAÑAS")
    values=ws.get_all_values()
    campaigns=wc.get_all_values()
    if not values or not campaigns:
        raise RuntimeError("COLA_ENVIO o CAMPAÑAS no tiene encabezados")
    hc={norm(h):i for i,h in enumerate(values[0])}
    hp={norm(h):i for i,h in enumerate(campaigns[0])}
    for name,h,cols in [
        ("COLA_ENVIO",hc,{"ID_ENVIO","ID_CAMPANA","PLANTILLA","ESTADO"}),
        ("CAMPAÑAS",hp,{"ID_CAMPANA","NOMBRE_CAMPANA","PLANTILLA","FILTRO","ESTADO"})
    ]:
        if not cols.issubset(h):
            raise RuntimeError(name+" carece de "+str(sorted(cols-set(h))))
    pattern=re.compile(r"^MORA1-(\d{8})-(T001|MORA10[1-5])-(.+)$")
    grouped=collections.defaultdict(list)
    for n,row in enumerate(values[1:],start=2):
        send_id=value(row,hc,"ID_ENVIO")
        m=pattern.fullmatch(send_id)
        if not m or norm(value(row,hc,"ESTADO"))!="BORRADOR":
            continue
        dt,code,_=m.groups()
        expected=f"MORA1-{dt}-{code}"
        existing=value(row,hc,"ID_CAMPANA")
        if existing and existing!=expected:
            raise RuntimeError("Un borrador de Mora 1 tiene ID_CAMPAÑA inesperado; no se modificó.")
        grouped[expected].append((n,row))
    if not grouped:
        print("No hay borradores de Mora 1 pendientes de vincular.")
        return
    current={}
    for row in campaigns[1:]:
        k=value(row,hp,"ID_CAMPANA")
        if k:
            if k in current:raise RuntimeError("ID_CAMPAÑA duplicado en CAMPAÑAS: "+k)
            current[k]=row
    new_campaigns=[]
    for key,items in sorted(grouped.items()):
        code=key.rsplit("-",1)[1]
        day=key.split("-")[1]
        if key in current:
            existing=current[key]
            if norm(value(existing,hp,"ESTADO")) not in {"BORRADOR","PREPARADA"}:
                raise RuntimeError("Campaña existente con estado no seguro: "+key)
            if norm(value(existing,hp,"PLANTILLA"))!=code:
                raise RuntimeError("Campaña existente tiene otra plantilla: "+key)
        else:
            count=len(items)
            datos={
                "ID_CAMPANA":key,"NOMBRE_CAMPANA":f"Mora 1 Dioben · {day} · {code}",
                "PLANTILLA":code,"FILTRO":"MORA_1","ESTADO":"BORRADOR",
                "FECHA_ENVIO":f"{day[6:8]}/{day[4:6]}/{day[0:4]}",
                "HORA_ENVIO":"08:15","TOTAL_CLIENTES":count,"PENDIENTES":count,
                "ENVIADOS":0,"ERRORES":0,
                "FECHA_CREACION":datetime.now(TZ).strftime("%d/%m/%Y %H:%M:%S"),
                "COMENTARIOS":"Creada por GitHub Actions; requiere revisión y autorización manual."
            }
            new_campaigns.append([datos.get(h,"") for h,_ in sorted(hp.items(),key=lambda x:x[1])])
    if new_campaigns:
        wc.append_rows(new_campaigns,value_input_option="RAW")
    updates=[]
    for key,items in grouped.items():
        for n,row in items:
            if not value(row,hc,"ID_CAMPANA"):
                updates.append({"range":rowcol_to_a1(n,hc["ID_CAMPANA"]+1),"values":[[key]]})
    for i in range(0,len(updates),150):
        ws.batch_update(updates[i:i+150],value_input_option="RAW")
    print("Campañas en BORRADOR creadas:",len(new_campaigns))
    print("Borradores existentes/nuevos vinculados:",len(updates))
    print("NO SE PROGRAMÓ NI SE ENVIÓ NINGÚN CORREO.")

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
        batch.append([values.get(header,"") if header!="ID_CAMPANA" else f"MORA1-{hoy:%Y%m%d}-{code}" for header,_ in sorted(hc.items(),key=lambda kv:kv[1])])
    print("Candidatos preparados como BORRADOR:",len(batch))
    print("Omisiones:",json.dumps(reasons,ensure_ascii=False))
    if not batch:
        vincular_campanas_borrador(destino)
        return
    # Antes de escribir, leer nuevamente para evitar duplicación si otro proceso creó registros.
    ws=destino.worksheet("COLA_ENVIO")
    fresh=ws.get_all_values()
    ids={r[hc["ID_ENVIO"]] for r in fresh[1:] if len(r)>hc["ID_ENVIO"]}
    valid=[r for r in batch if r[hc["ID_ENVIO"]] not in ids]
    if valid:
        ws.append_rows(valid,value_input_option="RAW")
    print("Borradores agregados:",len(valid))
    vincular_campanas_borrador(destino)
    print("NO SE ENVIARON CORREOS. Los BORRADORES requieren autorización explícita para enviarse.")

if __name__=="__main__":
    main()

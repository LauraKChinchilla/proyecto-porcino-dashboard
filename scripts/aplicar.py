"""Pasa al Excel de OneDrive las capturas hechas en la pantalla administrativa.

Lee los archivos pendientes/*.json (los crea admin.html), los aplica a
Granja_Porcina_Facil.xlsx con openpyxl, sube la nueva versión a OneDrive
(con If-Match para no pisar cambios hechos al mismo tiempo) y deja un registro
en capturas.json. Uso:
    python3 scripts/aplicar.py              # GitHub Actions (usa ONEDRIVE_SHARE_URL)
    python3 scripts/aplicar.py --local X.xlsx  # prueba sin OneDrive
"""
import base64, datetime as dt, glob, io, json, os, sys, unicodedata, urllib.request, urllib.error
from openpyxl import load_workbook
from openpyxl.styles import Font

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PEND = os.path.join(RAIZ, "pendientes")
LOG = os.path.join(RAIZ, "capturas.json")
API = "https://my.microsoftpersonalcontent.com/_api/v2.0"


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode()
    return " ".join(s.lower().split())


def fecha(v):
    if not v:
        return None
    return dt.date.fromisoformat(str(v)[:10])


def num(v):
    if v in (None, ""):
        return None
    try:
        return float(v) if "." in str(v) else int(v)
    except ValueError:
        return None


class Hoja:
    """Ayuda a escribir en una tabla con encabezados (busca columnas por nombre)."""

    def __init__(self, ws, primera, fila_max=60, desde_col=1):
        self.ws = ws
        for r in range(1, fila_max):
            v = ws.cell(row=r, column=desde_col).value
            if norm(v) == norm(primera):
                self.h = r
                break
        else:
            raise ValueError(f"No encontré la tabla '{primera}' en la hoja {ws.title}")
        self.cols = {}
        for c in range(desde_col, ws.max_column + 1):
            k = norm(ws.cell(row=self.h, column=c).value)
            if k and k not in self.cols:
                self.cols[k] = c

    def col(self, *nombres):
        for n in nombres:
            n = norm(n)
            for k, c in self.cols.items():
                if k.startswith(n):
                    return c
        raise ValueError(f"No encontré la columna {nombres} en {self.ws.title}")

    def filas(self):
        r = self.h + 1
        while r <= self.ws.max_row:
            yield r
            r += 1

    def vacia(self, claves):
        cs = [self.col(k) for k in claves]
        for r in self.filas():
            if all(self.ws.cell(row=r, column=c).value in (None, "") for c in cs):
                return r
        return self.ws.max_row + 1

    def poner(self, r, nombre, valor, fmt=None):
        if valor in (None, ""):
            return
        c = self.ws.cell(row=r, column=self.col(nombre))
        c.value = valor
        if fmt:
            c.number_format = fmt
        elif isinstance(valor, dt.date):
            c.number_format = "dd/mm/yyyy"

    def get(self, r, nombre):
        return self.ws.cell(row=r, column=self.col(nombre)).value


def nota_mas(h, r, texto):
    prev = h.get(r, "Nota")
    h.poner(r, "Nota", f"{prev} | {texto}" if prev else texto)


def dinero(wb, d, firma):
    h = Hoja(wb["Dinero"], "Fecha")
    r = h.vacia(["Fecha", "Entra o sale", "Descripción", "Monto"])
    h.poner(r, "Fecha", fecha(d.get("fecha")))
    h.poner(r, "Entra o sale", d["movimiento"])
    h.poner(r, "Tipo", d.get("tipo"))
    h.poner(r, "Descripción", d.get("descripcion"))
    h.poner(r, "Monto", num(d["monto"]), 'L #,##0;[Red]-L #,##0;"-"')
    h.poner(r, "Socio", d.get("socio"))
    h.poner(r, "Nota", " | ".join(x for x in [d.get("nota"), firma] if x))
    if d.get("tipo") == "Aporte de socio" and d.get("socio"):
        s = Hoja(wb["Socios"], "Fecha", fila_max=20)
        rs = s.vacia(["Socio", "Monto"])
        s.poner(rs, "Fecha", fecha(d.get("fecha")))
        s.poner(rs, "Socio", d["socio"])
        s.poner(rs, "Tipo de aporte", d.get("tipoAporte") or "Efectivo")
        s.poner(rs, "Descripción", d.get("descripcion") or "Aporte registrado desde el dashboard")
        s.poner(rs, "Monto", num(d["monto"]), 'L #,##0;[Red]-L #,##0;"-"')
    return f"Fila {r} de Dinero"


def cerdas(wb):
    return Hoja(wb["Cerdas"], "Cerda")


def code(v):
    return norm(v).replace("cerda", "").replace("cerd", "").replace("cod", "").replace(".", "").replace(" ", "").upper()


def ultima_camada(h, cerda):
    """Fila más reciente de esa cerda que ya tiene fecha de parto."""
    mejor = None
    for r in h.filas():
        if code(h.get(r, "Cerda")) == code(cerda) and h.get(r, "Fecha de parto"):
            f = h.get(r, "Fecha de parto")
            if mejor is None or f > mejor[1]:
                mejor = (r, f)
    return mejor[0] if mejor else None


def monta(wb, d, firma):
    h = cerdas(wb)
    r = h.vacia(["Cerda", "Fecha de monta", "Fecha de parto"])
    h.poner(r, "Cerda", d["cerda"])
    h.poner(r, "Verraco", d.get("verraco"))
    h.poner(r, "Fecha de monta", fecha(d["fecha"]))
    h.poner(r, "Nota", " | ".join(x for x in [d.get("nota"), firma] if x))
    return f"Monta en fila {r} de Cerdas"


def parto(wb, d, firma):
    h = cerdas(wb)
    fp = fecha(d["fecha"])
    cand = None
    for r in h.filas():
        if code(h.get(r, "Cerda")) == code(d["cerda"]) and not h.get(r, "Fecha de parto"):
            fm = h.get(r, "Fecha de monta")
            if fm and hasattr(fm, "date"):
                fm = fm.date()
            if fm and fm <= fp and (cand is None or fm > cand[1]):
                cand = (r, fm)
    if cand:
        r = cand[0]
        donde = f"Parto anotado en la fila {r} (monta del {cand[1].strftime('%d/%m/%Y')})"
    else:
        r = h.vacia(["Cerda", "Fecha de monta", "Fecha de parto"])
        h.poner(r, "Cerda", d["cerda"])
        donde = f"Parto en fila nueva {r} (no había monta pendiente de esa cerda)"
    h.poner(r, "Fecha de parto", fp)
    h.poner(r, "Nacidos vivos", num(d["vivos"]))
    if num(d.get("muertosAlNacer")):
        nota_mas(h, r, f"{num(d['muertosAlNacer'])} nacidos muertos")
    nota_mas(h, r, " | ".join(x for x in [d.get("nota"), firma] if x))
    return donde


def sumar_camada(wb, d, firma, columna, palabra):
    h = cerdas(wb)
    r = ultima_camada(h, d["cerda"])
    if not r:
        raise ValueError(f"La cerda {d['cerda']} no tiene ningún parto registrado")
    n = num(d["cantidad"]) or 0
    h.poner(r, columna, (num(h.get(r, columna)) or 0) + n)
    f = fecha(d.get("fecha"))
    nota_mas(h, r, f"{f.strftime('%d/%m') if f else ''} {n} {palabra}".strip() + (f" ({d['nota']})" if d.get("nota") else "") + f" {firma}")
    return f"Fila {r} de Cerdas: {columna} +{n}"


def muerte_lechones(wb, d, firma):
    return sumar_camada(wb, d, firma, "Muertos", "muertos")


def venta_lechones(wb, d, firma):
    msg = sumar_camada(wb, d, firma, "Vendidos", "vendidos")
    if num(d.get("monto")):
        dinero(wb, {"fecha": d.get("fecha"), "movimiento": "Entra", "tipo": "Venta de lechones",
                    "descripcion": f"Venta de {num(d['cantidad'])} lechones (cerda {d['cerda']})" + (f" a {d['comprador']}" if d.get("comprador") else ""),
                    "monto": d["monto"]}, firma)
        msg += " y venta anotada en Dinero"
    return msg


def hato_alta(wb, d, firma):
    h = Hoja(wb["Hato"], "Código")
    r = h.vacia(["Código"])
    h.poner(r, "Código", d["codigo"])
    h.poner(r, "Tipo", d.get("tipo") or "Cerda")
    h.poner(r, "Origen", d.get("origen"))
    h.poner(r, "Fecha de entrada", fecha(d.get("fecha")))
    h.poner(r, "Estado", "Activa")
    h.poner(r, "Nota", " | ".join(x for x in [d.get("nota"), firma] if x))
    if num(d.get("monto")):
        dinero(wb, {"fecha": d.get("fecha"), "movimiento": "Sale", "tipo": "Animales",
                    "descripcion": f"Compra de {(d.get('tipo') or 'cerda').lower()} {d['codigo']}", "monto": d["monto"]}, firma)
    return f"Fila {r} de Hato"


def hato_baja(wb, d, firma):
    h = Hoja(wb["Hato"], "Código")
    for r in h.filas():
        if code(h.get(r, "Código")) == code(d["codigo"]) or norm(h.get(r, "Código")) == norm(d["codigo"]):
            break
    else:
        raise ValueError(f"No encontré el animal {d['codigo']} en la hoja Hato")
    h.poner(r, "Estado", d["estado"])
    h.poner(r, "Fecha de salida", fecha(d.get("fecha")))
    nota_mas(h, r, " | ".join(x for x in [d.get("nota"), firma] if x))
    if d["estado"] == "Vendida" and num(d.get("monto")):
        dinero(wb, {"fecha": d.get("fecha"), "movimiento": "Entra", "tipo": "Venta de cerdos",
                    "descripcion": f"Venta de {d['codigo']}", "monto": d["monto"]}, firma)
    return f"Fila {r} de Hato: {d['estado']}"


def engorde(wb, d, firma):
    ws = wb["Engorde"]
    lot = Hoja(ws, "Lote", fila_max=12)
    existe = any(norm(lot.get(r, "Lote")) == norm(d["lote"]) for r in range(lot.h + 1, lot.h + 7))
    if not existe:
        for r in range(lot.h + 1, lot.h + 7):
            if not lot.get(r, "Lote"):
                lot.poner(r, "Lote", d["lote"])
                break
    h = Hoja(ws, "Fecha", fila_max=40)
    r = h.vacia(["Lote", "Movimiento", "Cantidad"])
    h.poner(r, "Fecha", fecha(d.get("fecha")))
    h.poner(r, "Lote", d["lote"])
    h.poner(r, "Movimiento", d["movimiento"])
    h.poner(r, "Cantidad", num(d.get("cantidad")))
    h.poner(r, "Monto", num(d.get("monto")), 'L #,##0;[Red]-L #,##0;"-"')
    h.poner(r, "Peso", num(d.get("peso")))
    h.poner(r, "Nota", " | ".join(x for x in [d.get("nota"), firma] if x))
    msg = f"Fila {r} de Engorde"
    if num(d.get("monto")) and d["movimiento"] in ("Venta", "Compra"):
        venta = d["movimiento"] == "Venta"
        dinero(wb, {"fecha": d.get("fecha"), "movimiento": "Entra" if venta else "Sale", "tipo": "Venta de cerdos" if venta else "Animales",
                    "descripcion": f"{'Venta' if venta else 'Compra'} de {num(d.get('cantidad')) or ''} cerdos de engorde ({d['lote']})", "monto": d["monto"]}, firma)
        msg += " y en Dinero"
    return msg


def vacuna(wb, d, firma):
    h = Hoja(wb["Vacunas"], "Fecha")
    r = h.vacia(["Fecha", "Cerda o lote", "Producto"])
    h.poner(r, "Fecha", fecha(d["fecha"]))
    h.poner(r, "Cerda o lote", d["cerda"])
    h.poner(r, "Producto", d["producto"])
    h.poner(r, "Dosis", d.get("dosis"))
    h.poner(r, "Próximo refuerzo", fecha(d.get("refuerzo")))
    h.poner(r, "Nota", " | ".join(x for x in [d.get("nota"), firma] if x))
    return f"Fila {r} de Vacunas"


def bodega(wb, d, firma, tabla):
    ws = wb["Bodega"]
    if tabla == "medicina":
        h = Hoja(ws, "Medicina", fila_max=10, desde_col=1)
        cn, cq, cv, ct = h.col("Medicina"), h.col("Frascos"), h.col("Vence"), h.col("Tipo")
        nombre, cant = d["nombre"], num(d["cantidad"])
        ultima = h.h + 60
    else:
        h = Hoja(ws, "Concentrado", fila_max=10, desde_col=7)
        cn, cq = h.col("Concentrado"), h.col("Bolsas")
        cv = [c for k, c in h.cols.items() if k.startswith("vence") and c > cn][0]
        ct = None
        nombre, cant = d["tipo"], num(d["cantidad"])
        ultima = h.h + 20
    fila = None
    for r in range(h.h + 1, ultima + 1):
        if norm(ws.cell(row=r, column=cn).value) == norm(nombre):
            fila = r
            break
    nuevo = fila is None
    if nuevo:
        for r in range(h.h + 1, ultima + 1):
            if ws.cell(row=r, column=cn).value in (None, ""):
                fila = r
                break
        ws.cell(row=fila, column=cn).value = nombre
    if d.get("modo") == "sumar" and not nuevo:
        cant = (num(ws.cell(row=fila, column=cq).value) or 0) + (cant or 0)
    if cant is not None:
        ws.cell(row=fila, column=cq).value = cant
    if d.get("vence"):
        c = ws.cell(row=fila, column=cv)
        c.value = fecha(d["vence"])
        c.number_format = "mmm yyyy" if tabla == "medicina" else "dd/mm/yyyy"
    if ct and d.get("tipoMed"):
        ws.cell(row=fila, column=ct).value = d["tipoMed"]
    return f"{'Nueva fila' if nuevo else 'Fila'} {fila} de Bodega: {nombre} = {cant}"


def parametro(wb, d, firma):
    """Escribe un valor con etiqueta en una hoja (por ejemplo, el costo por cerdo en 'Costo lechón')."""
    ws = wb[d["hoja"]]
    etiqueta = d["etiqueta"]
    fila = None
    for r in range(1, ws.max_row + 1):
        if norm(ws.cell(row=r, column=1).value) == norm(etiqueta):
            fila = r
            break
    if fila is None:
        fila = ws.max_row + 2
        ws.cell(row=fila, column=1, value=etiqueta).font = Font(name="Arial", size=11, bold=True)
    c = ws.cell(row=fila, column=2, value=num(d["valor"]))
    c.number_format = 'L #,##0;[Red]-L #,##0;"-"'
    c.font = Font(name="Arial", size=12, bold=True, color="1F1F1F")
    if d.get("nota"):
        ws.cell(row=fila, column=3, value=d["nota"]).font = Font(name="Arial", size=10, italic=True, color="7F7F7F")
    return f"{d['hoja']}: {etiqueta} = {d['valor']}"


ACCIONES = {
    "dinero": dinero, "monta": monta, "parto": parto, "muerte_lechones": muerte_lechones,
    "venta_lechones": venta_lechones, "hato_alta": hato_alta, "hato_baja": hato_baja, "engorde": engorde,
    "vacuna": vacuna, "medicina": lambda wb, d, f: bodega(wb, d, f, "medicina"),
    "concentrado": lambda wb, d, f: bodega(wb, d, f, "concentrado"),
    "parametro": parametro,
}
VENTANA = 60 * 60 + 5 * 60   # 1 hora para corregir (+5 min de margen por la demora en aplicar)


# ---------- registro de cambios celda por celda (para poder corregir o anular) ----------
def ser(v):
    if isinstance(v, dt.datetime):
        return {"$d": v.replace(microsecond=0).isoformat()}
    if isinstance(v, dt.date):
        return {"$d": dt.datetime(v.year, v.month, v.day).isoformat()}
    return v


def deser(v):
    if isinstance(v, dict) and "$d" in v:
        return dt.datetime.fromisoformat(v["$d"])
    return v


def foto(wb):
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title, c.coordinate)] = ser(c.value)
    return out


def diferencias(antes, despues):
    cambios = []
    for k in set(antes) | set(despues):
        a, b = antes.get(k), despues.get(k)
        if a != b:
            cambios.append([k[0], k[1], a, b])
    return sorted(cambios)


def deshacer(wb, cambios):
    """Devuelve las celdas a como estaban. Si alguien ya las cambió a mano, no toca nada."""
    choques = [f"{h}!{c}" for h, c, viejo, nuevo in cambios if ser(wb[h][c].value) != nuevo]
    if choques:
        raise ValueError("Ese registro ya se modificó directo en el Excel (" + ", ".join(choques[:4]) +
                         "). Corrígelo en el Excel.")
    for h, c, viejo, nuevo in cambios:
        wb[h][c].value = deser(viejo)


def minutos(a, b):
    return (dt.datetime.fromisoformat(b.replace("Z", "+00:00")) - dt.datetime.fromisoformat(a.replace("Z", "+00:00"))).total_seconds()


def admins():
    try:
        with open(os.path.join(RAIZ, "admin", "config.json"), encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("v") == 2:
            return {u["nombre"] for u in cfg["usuarios"] if u.get("admin")}
        return {"Laura"}
    except Exception:
        return {"Laura"}


def aplicar(contenido, capturas, log):
    wb = load_workbook(io.BytesIO(contenido))
    porid = {e.get("id"): e for e in log}
    jefes = admins()
    res = []
    for cap in capturas:
        firma = f"[dashboard: {cap.get('quien') or 'sin nombre'}]"
        r = {"estado": "ok"}
        try:
            antes = foto(wb)
            if cap["tipo"] in ("editar", "anular"):
                obj = porid.get(cap["datos"].get("objetivo"))
                if not obj:
                    raise ValueError("No encontré el registro que se quiere corregir")
                if obj.get("estado") != "ok":
                    raise ValueError("Ese registro ya fue corregido o anulado")
                inicio = obj.get("inicio") or obj.get("creado")
                if minutos(inicio, cap.get("creado")) > VENTANA:
                    raise ValueError("Ya pasó más de 1 hora desde que se agregó; corrígelo en el Excel")
                if not obj.get("cambios"):
                    raise ValueError("Ese registro no se puede corregir desde la pantalla; corrígelo en el Excel")
                dueno = obj.get("dueno") or obj.get("quien")
                if cap.get("quien") != dueno and cap.get("quien") not in jefes:
                    raise ValueError(f"Solo {dueno} o un administrador pueden corregirlo")
                deshacer(wb, obj["cambios"])
                medio = foto(wb)
                if cap["tipo"] == "editar":
                    tipo = cap["datos"]["tipo"]
                    msg = ACCIONES[tipo](wb, cap["datos"]["datos"], firma + " (corregido)")
                    r.update(tipo=tipo, datos=cap["datos"]["datos"], corrige=obj["id"], inicio=inicio, dueno=dueno,
                             cambios=diferencias(medio, foto(wb)), mensaje="Corregido: " + msg)
                    obj["estado"] = "corregido"
                else:
                    r.update(anula=obj["id"], mensaje="Registro anulado en el Excel")
                    obj["estado"] = "anulado"
                obj["reemplazado_por"] = cap.get("id")
            else:
                msg = ACCIONES[cap["tipo"]](wb, cap["datos"], firma)
                r.update(datos=cap["datos"], inicio=cap.get("creado"), dueno=cap.get("quien"), cambios=diferencias(antes, foto(wb)), mensaje=msg)
        except Exception as e:  # una captura mala no bloquea las demás y no deja nada a medias
            ahora_f = foto(wb)
            for k in set(antes) | set(ahora_f):
                if antes.get(k) != ahora_f.get(k):
                    wb[k[0]][k[1]].value = deser(antes.get(k))
            r = {"estado": "error", "mensaje": str(e)}
        res.append(r)
    try:
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        pass
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), res


# ---------------- OneDrive ----------------
def http(url, method="GET", data=None, headers=None):
    req = urllib.request.Request(url, data=data, method=method, headers=headers or {})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.status, r.read()


def onedrive():
    share = os.environ["ONEDRIVE_SHARE_URL"].strip()
    _, b = http("https://api-badgerp.svc.ms/v1.0/token", "POST", json.dumps({"appId": "5cbed6ac-a083-4e14-b191-b4ba07653de2"}).encode(),
                {"Content-Type": "application/json"})
    tok = json.loads(b)["token"]
    enc = "u!" + base64.b64encode(share.encode()).decode().rstrip("=").replace("/", "_").replace("+", "-")
    H = {"Authorization": "Badger " + tok, "Prefer": "autoredeem"}
    _, b = http(f"{API}/shares/{enc}/driveitem?$expand=children", headers=H)
    kids = json.loads(b).get("children", [])
    item = next((k for k in kids if "granja porcina facil" in norm(k["name"]).replace("_", " ")), None)
    if not item:
        raise SystemExit("No encontré Granja_Porcina_Facil.xlsx en la carpeta")
    return H, item


def main():
    archivos = sorted(glob.glob(os.path.join(PEND, "*.json")))
    if not archivos:
        print("Sin capturas pendientes.")
        return
    capturas = []
    for a in archivos:
        try:
            with open(a, encoding="utf-8") as f:
                c = json.load(f)
            c["_archivo"] = a
            capturas.append(c)
        except Exception as e:
            print("Captura ilegible", a, e)
            capturas.append({"_archivo": a, "tipo": "ilegible", "datos": {}, "id": os.path.basename(a)})
    capturas.sort(key=lambda c: c.get("creado", ""))
    conocidas = set(ACCIONES) | {"editar", "anular"}
    validas = [c for c in capturas if c["tipo"] in conocidas]
    try:
        with open(LOG, encoding="utf-8") as f:
            log = json.load(f)
    except Exception:
        log = []

    if "--local" in sys.argv:
        ruta = sys.argv[sys.argv.index("--local") + 1]
        with open(ruta, "rb") as f:
            nuevo, res = aplicar(f.read(), validas, log)
        with open(ruta.replace(".xlsx", "_aplicado.xlsx"), "wb") as f:
            f.write(nuevo)
    else:
        for intento in range(4):
            H, item = onedrive()
            drive, iid = item["parentReference"]["driveId"], item["id"]
            _, contenido = http(item["@content.downloadUrl"])
            copia = json.loads(json.dumps(log))   # cada intento parte del registro sin tocar
            nuevo, res = aplicar(contenido, validas, copia)
            try:
                http(f"{API}/drives/{drive}/items/{iid}/content", "PUT", nuevo,
                     {**H, "If-Match": item["eTag"], "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
                log = copia
                break
            except urllib.error.HTTPError as e:
                if e.code in (409, 412) and intento < 3:
                    print("El Excel cambió mientras se aplicaba; reintento…")
                    continue
                raise
        else:
            return
    ahora = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    it = iter(res)
    for c in capturas:
        r = next(it) if c["tipo"] in conocidas else {"estado": "error", "mensaje": "Tipo de captura desconocido"}
        entrada = {"id": c.get("id"), "creado": c.get("creado"), "quien": c.get("quien"), "tipo": c["tipo"],
                   "resumen": c.get("resumen"), "aplicado": ahora, **r}
        if entrada["tipo"] == "editar" and r.get("tipo"):
            entrada["tipo"] = r["tipo"]
        log.insert(0, entrada)
        print(c["tipo"], "→", r["estado"], r["mensaje"])
        if "--local" not in sys.argv:
            os.remove(c["_archivo"])
    # Las celdas cambiadas solo se guardan mientras se pueden corregir (para que el archivo no crezca)
    for e in log:
        ini = e.get("inicio") or e.get("creado")
        if e.get("cambios") and ini and minutos(ini, ahora) > VENTANA + 3600:
            e.pop("cambios", None)
    with open(LOG, "w", encoding="utf-8") as f:
        json.dump(log[:300], f, ensure_ascii=False, indent=1)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VALIDADOR COMPLETO DE ACCESORIOS - Thermia / Dimac  (v final)
=============================================================
Cruza las 5 tablas de Geinfor + el Excel de seguimiento de José María y valida
los accesorios FW creados/modificados según todas las reglas acordadas.

ARCHIVOS DE ENTRADA:
  1) REPASO.xlsx  con hojas:
       - MAESTRO DE ARTICULOS   (llave: CODIGO_ARTICULO)
       - ARTICULOS TARIFABLES   (llave: ARTICULO)      -> PVP (tarifa 40)
       - ARTICULOS PROVEEDOR    (llave: ARTICULO)      -> coste/plazo/lote (PRINCIPAL=1)
       - ALMACEN ARTICULOS      (llave: ARTICULO)      -> stock mínimo
       - TABLA CARACTERISTICAS  (llave: ARTICULO)      -> imagen (caract. 0)
  2) JOSEMARIA.xlsx  hoja 'RELLENAR JOSE MARIA' (llave: ARTÍCULO) -> divisor, mínimo

USO:
    python validador_final.py REPASO.xlsx JOSEMARIA.xlsx [USUARIO] [FECHA_DESDE]
Sin usuario/fecha -> valida todos los FW del maestro.
"""
import sys
import re
import html as _html
import pandas as pd

# Nombres legibles de cada campo para el informe visual
NOMBRES_CAMPO = {
    "FAMILIA": "Familia de producto",
    "CLASE": "Clase / serie",
    "DENOMINACION": "Denominación",
    "DENOMINACION_2": "Denominación 2 (acabado)",
    "MEDIDA": "Medida",
    "COEFICIENTE": "Coeficiente",
    "UNIDADES_ARTICULO": "Unidades del artículo",
    "GRUPO": "Grupo (acabado)",
    "CARACTERISTICA_0": "Imagen",
    "PROVEEDOR": "Proveedor principal",
    "PLAZO_ENTREGA": "Plazo de entrega",
    "LOTE_HABITUAL": "Lote habitual",
    "MARGEN": "Margen / PVP",
    "STOCK_MINIMO": "Stock mínimo",
    "ESTADO": "Estado",
    "TIPO": "Tipo",
    "UNIDAD_PROVEEDOR": "Unidad del proveedor",
    "UNIDAD_VENTA": "Unidad de venta",
    "UNIDAD/TIPO": "Unidad / Tipo",
    "EXCEL_JM": "Excel José María",
    "DIVISOR": "Divisor",
    "PVP": "PVP",
    "IVA_TARIFA": "IVA en tarifa 40",
    "IVA_PROVEEDOR": "IVA en artículo proveedor",
    "FORMULA_PROVEEDOR": "Fórmula en artículo proveedor",
    "PARTIDAARANCELARIA": "Partida arancelaria",
    "CONVERSION_M_KG": "Coeficiente m/kg",
    "GESTIONARVARIABLES": "Gestionar variables",
    "ESSUBCONTRATACIO": "Es subcontratación",
    "ESKIT_CHECK": "Es un artículo kit",
    "NOGESTIONASTOCK": "No gestiona inventario",
    "CALCULOCANTFORMULA": "Cálculo cantidad por fórmula",
}

# Reglas de campo fijo del maestro. COEFICIENTE y UNIDADES_ARTICULO se validan aparte
# (dependen del caso de unidades: barras, metro lineal, unitario).
REGLAS_MAESTRO = {}
GRUPO_POR_SUFIJO = {"04": 10, "03": 11, "01": 12, "06": 14, "05": 16}
# Texto que debe llevar DENOMINACION_2 según GRUPO del maestro
DENOM2_POR_GRUPO = {10: "NEGRO", 11: "BLANCO", 12: "PLATA", 14: "RAL STD", 16: "PVD"}
TARIFA_PVP = 40
# Tolerancia de margen híbrida: solo se marca si la diferencia supera AMBOS umbrales.
# Absorbe el band sistemático (~3-4%, donde el PVP real queda algo por encima del
# calculado = más margen, validado con José María en el artículo 072040) y el redondeo
# de céntimos, pero sigue detectando los outliers reales (>10-70%).
TOL_MARGEN_EUR = 0.05  # diferencia mínima en € para considerar incidencia
TOL_MARGEN_PCT = 5.0   # diferencia mínima en % para considerar incidencia


def num(v):
    try:
        f = float(v)
    except (ValueError, TypeError):
        return None
    return None if f != f else f  # NaN (f != f) -> None


def ni(v):  # entero o None
    f = num(v)
    return int(f) if f is not None else None


# Conexión directa a Geinfor (IBM DB2). SOLO LECTURA. La base de Thermia/Dimac es la DSN
# GeinprodINGENRED (conecta sin credenciales estando en la red de la oficina). Esquema DB2ADMIN.
DSN_DB2 = "GeinprodINGENRED"

# Carpeta de red donde José María mantiene su Excel de seguimiento. El validador coge solo
# el más reciente que cumpla el patrón, así no hay que copiarlo a mano a la carpeta local.
JM_CARPETA = r"\\dimac2023\DATOS\DATOS DIMAC\USERS\INTERCAMBIO\RAUL\xavi\PROCES DE COMPRES ADRIA\EXTRUSORES\MATRICES EXTRUSIÓN"
JM_PATRON = "*Matrices nuevas*.xlsx"


def localizar_excel_jm(carpeta=JM_CARPETA, patron=JM_PATRON):
    """Devuelve la ruta del Excel de José María más reciente en la carpeta de red."""
    import glob
    import os
    cands = [f for f in glob.glob(os.path.join(carpeta, patron))
             if not os.path.basename(f).startswith("~$")]
    if not cands:
        raise FileNotFoundError(
            f"No encuentro el Excel de José María ({patron}) en:\n  {carpeta}\n"
            "¿Estás en la red de la oficina?")
    cands.sort(key=os.path.getmtime, reverse=True)
    return cands[0]


_MCOLS = ("CODIGO_ARTICULO,DENOMINACION,DENOMINACION_2,CLASE,TIPO,NORMA,COEFICIENTE,"
          "UNIDAD_INTERNA,UNIDAD_EXTERNA,FAMILIAPRODUCTO,GRUPO,MEDIDA,UNIDADES_ARTICULO,"
          "ESTADO,FECHACREACION,FECHAMODIFICACION,"
          "PARTIDAARANCELARIA,CONVERSION_M_KG,"
          "GESTIONARVARIABLES,ESSUBCONTRATACIO,ESKIT,NOGESTIONASTOCK,CALCULOCANTFORMULA")


def _leer_in(cn, tabla, clave, codigos, cols, extra=""):
    """SELECT cols FROM tabla WHERE clave IN (codigos) [extra], troceado para no pasarse
    del límite de parámetros. Devuelve un único DataFrame."""
    partes = []
    cod = list(codigos)
    for i in range(0, len(cod), 500):
        trozo = cod[i:i + 500]
        ph = ",".join("?" * len(trozo))
        sql = f"SELECT {cols} FROM DB2ADMIN.{tabla} WHERE {clave} IN ({ph}) {extra}"
        partes.append(pd.read_sql(sql, cn, params=trozo))
    if partes:
        return pd.concat(partes, ignore_index=True)
    return pd.DataFrame(columns=[c.strip() for c in cols.split(",")])


def _conteo_plazos_db2(cn, proveedores):
    """Conteo de PLAZOENTREGA por proveedor (PRINCIPAL=1) para los proveedores dados.
    Devuelve df PROVEEDOR / PLAZOENTREGA / N (vía GROUP BY, ligero)."""
    partes = []
    provs = list(proveedores)
    for i in range(0, len(provs), 500):
        t = provs[i:i + 500]
        ph = ",".join("?" * len(t))
        sql = (f"SELECT PROVEEDOR, PLAZOENTREGA, COUNT(*) AS N FROM DB2ADMIN.ARTICULOS_PROVEEDOR "
               f"WHERE PRINCIPAL=1 AND PROVEEDOR IN ({ph}) GROUP BY PROVEEDOR, PLAZOENTREGA")
        partes.append(pd.read_sql(sql, cn, params=t))
    if partes:
        return pd.concat(partes, ignore_index=True)
    return pd.DataFrame(columns=["PROVEEDOR", "PLAZOENTREGA", "N"])


def plazo_habitual_por_proveedor(cnt, min_muestras=5, dominancia=0.6):
    """A partir del conteo (PROVEEDOR/PLAZOENTREGA/N), devuelve {proveedor: plazo_habitual}
    solo para proveedores con suficientes artículos y un plazo claramente dominante."""
    out = {}
    if cnt is None or len(cnt) == 0:
        return out
    c = cnt.copy()
    c["PROVEEDOR"] = pd.to_numeric(c["PROVEEDOR"], errors="coerce")
    c["PLAZOENTREGA"] = pd.to_numeric(c["PLAZOENTREGA"], errors="coerce")
    c["N"] = pd.to_numeric(c["N"], errors="coerce").fillna(0)
    c = c.dropna(subset=["PROVEEDOR", "PLAZOENTREGA"])
    for pid, g in c.groupby("PROVEEDOR"):
        total = g["N"].sum()
        if total >= min_muestras:
            top = g.sort_values("N", ascending=False).iloc[0]
            if top["N"] / total >= dominancia:
                out[int(pid)] = top["PLAZOENTREGA"]
    return out


def leer_db2(desde=None, hasta=None, articulo=None, dsn=DSN_DB2):
    """Lee de Geinfor (DB2, SOLO LECTURA) solo lo necesario: el universo de accesorios
    (familia != 1,2) + catálogos (tipo 91), y de las tablas relacionadas únicamente las
    filas de los artículos EN ALCANCE y con los filtros que el validador ya aplica
    (proveedor PRINCIPAL=1, tarifa 40, característica 0). Devuelve la misma estructura
    que las hojas del Excel."""
    import pyodbc
    import warnings
    warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")
    cn = pyodbc.connect("DSN=" + dsn, timeout=30, readonly=True)  # readonly: regla de oro
    try:
        # Maestro: en modo un-artículo, ese código (sea cual sea su familia) + catálogos,
        # para poder informar si resulta ser un perfil. En lote, accesorios + catálogos.
        if articulo is not None:
            m = pd.read_sql(
                f"SELECT {_MCOLS} FROM DB2ADMIN.MAESTRO_DE_ARTICULOS "
                "WHERE TIPO=91 OR CODIGO_ARTICULO=?", cn, params=[articulo])
        else:
            m = pd.read_sql(
                f"SELECT {_MCOLS} FROM DB2ADMIN.MAESTRO_DE_ARTICULOS "
                "WHERE FAMILIAPRODUCTO NOT IN (1,2,105) OR TIPO=91", cn)
        # Códigos EN ALCANCE: un solo artículo (--articulo) o los accesorios del rango.
        if articulo is not None:
            codigos = [articulo]
        else:
            fam = pd.to_numeric(m["FAMILIAPRODUCTO"], errors="coerce")
            fc = pd.to_datetime(m["FECHACREACION"], errors="coerce")
            mask = ~fam.isin([1, 2])
            if desde is not None:
                mask &= fc >= desde
            if hasta is not None:
                mask &= fc <= hasta
            codigos = m.loc[mask, "CODIGO_ARTICULO"].astype(str).str.strip().tolist()

        tar = _leer_in(cn, "ARTICULOS_TARIFABLES", "ARTICULO", codigos,
                       "ARTICULO,TARIFA,PRECIO_VENTA_PUBLICO,UNIDADVENTA,COEFICIENTE,IVA", "AND TARIFA=40")
        prov = _leer_in(cn, "ARTICULOS_PROVEEDOR", "ARTICULO", codigos,
                        "ARTICULO,PROVEEDOR,PRINCIPAL,COSTEBRUTO,COSTEINTERNO,PLAZOENTREGA,"
                        "LOTEHABITUAL,UNIDADEXTERNA,COEFICIENTE,PIVA,FORMULA", "AND PRINCIPAL=1")
        alm = _leer_in(cn, "ALMACEN_ARTICULOS", "ARTICULO", codigos, "ARTICULO,STOCKMINIMO")
        car = _leer_in(cn, "TCARAC_ART", "ARTICULO", codigos,
                       "ARTICULO,CARACTERISTICA,VALOR", "AND CARACTERISTICA=0")
        est = _leer_in(cn, "MAESTRO_ESTRUCTURAS", "ARTSUPERIOR", codigos, "ARTSUPERIOR,ARTCOMPONENTE")
        clases = pd.read_sql("SELECT CODIGO, DENOMINACION FROM DB2ADMIN.TABLA_CLASE_ARTICULO", cn)
        mprov = pd.read_sql("SELECT CODIGO_PROVEEDOR, TIPO_CTA_COMPRA FROM DB2ADMIN.MAESTRO_PROVEEDORES", cn)
        # Conteo de plazos por proveedor (para comparar contra el plazo habitual del proveedor)
        provs = pd.to_numeric(prov["PROVEEDOR"], errors="coerce").dropna().astype(int).unique().tolist()
        cnt_plazos = _conteo_plazos_db2(cn, provs)
    finally:
        cn.close()
    return m, tar, prov, alm, car, est, clases, cnt_plazos, mprov


_ESTILO_HTML = """<style>
  :root{--err:#c0392b;--errbg:#fdecea;--warn:#b9770e;--warnbg:#fdf3e3;--ok:#1e8449;--okbg:#eafaf1;}
  *{box-sizing:border-box} body{font-family:Segoe UI,system-ui,Arial,sans-serif;margin:0;background:#f4f6f8;color:#222}
  .wrap{max-width:820px;margin:0 auto;padding:24px}
  .head{background:#fff;border-radius:14px;padding:22px 24px;box-shadow:0 1px 4px rgba(0,0,0,.08)}
  .cod{font-size:30px;font-weight:700;letter-spacing:.5px}
  .denom{font-size:18px;color:#444;margin-top:2px}
  .meta{color:#777;font-size:13px;margin-top:10px}
  .banner{margin:18px 0;padding:14px 18px;border-radius:12px;font-size:18px;font-weight:600}
  .banner.ok{background:var(--okbg);color:var(--ok)}
  .banner.bad{background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  .pill{display:inline-block;padding:5px 12px;border-radius:20px;font-size:14px;margin-right:8px;color:#fff}
  .pill.err{background:var(--err)} .pill.warn{background:var(--warn)} .pill.ok{background:var(--ok)}
  h2.th{font-size:15px;text-transform:uppercase;letter-spacing:.5px;margin:22px 0 10px}
  h2.th.err{color:var(--err)} h2.th.warn{color:var(--warn)}
  .card{background:#fff;border-left:5px solid #ccc;border-radius:10px;padding:12px 16px;margin-bottom:10px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
  .card.err{border-color:var(--err)} .card.warn{border-color:var(--warn)}
  .campo{font-weight:700;font-size:15px} .detalle{color:#555;margin-top:3px;font-size:14px}
  .foot{color:#999;font-size:12px;margin-top:26px;text-align:center}
  .art{background:#fff;border-radius:12px;padding:16px 20px;margin:14px 0;box-shadow:0 1px 4px rgba(0,0,0,.07)}
  .acod{font-size:20px;font-weight:700} .adenom{color:#555;font-size:14px;margin-bottom:8px}
</style>"""


def informe_mensaje(cod, titulo, mensaje, ruta, clase="bad"):
    """Informe HTML simple con un único mensaje (p.ej. 'es un perfil' o 'no encontrado')."""
    e = _html.escape
    ahora = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    doc = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Validación {e(cod)}</title>
{_ESTILO_HTML}</head><body><div class="wrap">
  <div class="head"><div class="cod">{e(cod)}</div></div>
  <div class="banner {clase}" style="font-size:17px">{e(titulo)}</div>
  <div class="card"><div class="detalle" style="font-size:15px">{e(mensaje)}</div></div>
  <div class="foot">Validador de accesorios Dimac · datos en directo de Geinfor (solo lectura) · {ahora}</div>
</div></body></html>"""
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(doc)


def generar_informe_lote_html(sub, out, etiqueta, ruta):
    """Informe visual (HTML) de la validación de un LOTE de altas: un bloque por artículo
    con incidencias (errores y avisos), ordenados los que tienen errores primero."""
    def esc(x):
        return _html.escape(str(x))
    total = len(sub)
    n_err = int((out["NIVEL"] == "ERROR").sum()) if len(out) else 0
    n_av = int((out["NIVEL"] == "AVISO").sum()) if len(out) else 0
    con_inc = out["CODIGO"].nunique() if len(out) else 0
    correctos = total - con_inc
    info = sub.set_index("CODIGO_ARTICULO")

    codigos = list(dict.fromkeys(out["CODIGO"].tolist())) if len(out) else []
    def nerr(c):
        return int(((out["CODIGO"] == c) & (out["NIVEL"] == "ERROR")).sum())
    codigos.sort(key=lambda c: (0 if nerr(c) > 0 else 1, c))

    bloques = []
    for c in codigos:
        filas = out[out["CODIGO"] == c]
        denom = esc(info.loc[c, "DENOMINACION"]) if c in info.index else ""
        cards = []
        for _, r in filas.iterrows():
            clase = "err" if r["NIVEL"] == "ERROR" else "warn"
            nombre = NOMBRES_CAMPO.get(r["CAMPO"], r["CAMPO"])
            cards.append(f'<div class="card {clase}"><div class="campo">{esc(nombre)}</div>'
                         f'<div class="detalle">{esc(r["DETALLE"])}</div></div>')
        bloques.append(f'<div class="art"><div class="acod">{esc(c)}</div>'
                       f'<div class="adenom">{denom}</div>{"".join(cards)}</div>')

    ahora = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    pills = (f'<span class="pill ok">{correctos} correctos</span>'
             f'<span class="pill err">{n_err} errores</span>'
             f'<span class="pill warn">{n_av} avisos</span>')
    cuerpo = "".join(bloques) if bloques else '<div class="banner ok">✓ Sin incidencias</div>'
    doc = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Validación de altas</title>
{_ESTILO_HTML}</head><body><div class="wrap">
  <div class="head"><div class="cod">Validación de altas</div>
    <div class="denom">{esc(etiqueta)}</div>
    <div class="meta">{total} accesorios evaluados</div></div>
  <div class="banner bad">{pills}</div>
  {cuerpo}
  <div class="foot">Validador de accesorios Dimac · datos en directo de Geinfor (solo lectura) · {ahora}</div>
</div></body></html>"""
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(doc)


def generar_informe_html(cod, fila, inc_art, ruta):
    """Genera un informe visual (HTML) de la validación de UN artículo."""
    def esc(x):
        return _html.escape(str(x))
    errores = [r for r in inc_art if r[2] == "ERROR"]
    avisos = [r for r in inc_art if r[2] == "AVISO"]
    denom = esc(fila.get("DENOMINACION", ""))
    cab = (f"Familia {esc(fila.get('FAMILIAPRODUCTO'))} · Tipo {esc(fila.get('TIPO'))} · "
           f"Grupo {esc(fila.get('GRUPO'))} · Estado {esc(fila.get('ESTADO'))} · "
           f"Clase {esc(fila.get('CLASE'))}")
    fcre = fila.get("FECHACREACION")
    fcre = esc(pd.to_datetime(fcre).date()) if pd.notna(fcre) else "-"

    if not errores and not avisos:
        banner = '<div class="banner ok">✓ Artículo correcto — sin incidencias</div>'
    else:
        partes = []
        if errores:
            partes.append(f'<span class="pill err">{len(errores)} error(es)</span>')
        if avisos:
            partes.append(f'<span class="pill warn">{len(avisos)} aviso(s)</span>')
        banner = f'<div class="banner bad">{" ".join(partes)}</div>'

    def tarjetas(lista, clase):
        out = []
        for _, campo, _niv, detalle in lista:
            nombre = NOMBRES_CAMPO.get(campo, campo)
            out.append(f'<div class="card {clase}"><div class="campo">{esc(nombre)}</div>'
                       f'<div class="detalle">{esc(detalle)}</div></div>')
        return "\n".join(out)

    secciones = ""
    if errores:
        secciones += f'<h2 class="th err">✕ Errores ({len(errores)})</h2>{tarjetas(errores, "err")}'
    if avisos:
        secciones += f'<h2 class="th warn">⚠ Avisos ({len(avisos)})</h2>{tarjetas(avisos, "warn")}'

    ahora = pd.Timestamp.now().strftime("%d/%m/%Y %H:%M")
    doc = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Validación {esc(cod)}</title>
{_ESTILO_HTML}</head><body><div class="wrap">
  <div class="head">
    <div class="cod">{esc(cod)}</div>
    <div class="denom">{denom}</div>
    <div class="meta">{cab} · Creado: {fcre}</div>
  </div>
  {banner}
  {secciones}
  <div class="foot">Validador de accesorios Dimac · datos en directo de Geinfor (solo lectura) · {ahora}</div>
</div></body></html>"""
    with open(ruta, "w", encoding="utf-8") as f:
        f.write(doc)


def main():
    # Modos:
    #   Excel:        python validador.py REPASO.xlsx JM.xlsx [DESDE] [HASTA]
    #   DB2 (rango):  python validador.py --db JM.xlsx [DESDE] [HASTA]
    #   Un artículo:  python validador.py --db JM.xlsx --articulo CODIGO
    args = sys.argv[1:]
    articulo = None
    if "--articulo" in args:
        i = args.index("--articulo")
        articulo = args[i + 1].strip()
        del args[i:i + 2]

    if args[0] == "--db":
        # El Excel de JM se localiza solo en la red; opcionalmente se puede pasar uno explícito.
        rest = args[1:]
        if rest and rest[0].lower().endswith(".xlsx"):
            jmpath = rest[0]
            rest = rest[1:]
        else:
            jmpath = localizar_excel_jm()
            print(f"Excel de José María (red): {jmpath}")
        desde = pd.to_datetime(rest[0]) if len(rest) > 0 else None
        hasta = pd.to_datetime(rest[1]) if len(rest) > 1 else None
        print(f"Leyendo Geinfor en directo (DB2, DSN {DSN_DB2})...")
        m, tar, prov, alm, car, est, clases, cnt_plazos, mprov = leer_db2(desde, hasta, articulo)
    else:
        repaso, jmpath = args[0], args[1]
        desde = pd.to_datetime(args[2]) if len(args) > 2 else None
        hasta = pd.to_datetime(args[3]) if len(args) > 3 else None
        xl = pd.ExcelFile(repaso)
        m = pd.read_excel(xl, "MAESTRO DE ARTICULOS", dtype={"CODIGO_ARTICULO": str})
        tar = pd.read_excel(xl, "ARTICULOS TARIFABLES", dtype={"ARTICULO": str})
        prov = pd.read_excel(xl, "ARTICULOS PROVEEDOR", dtype={"ARTICULO": str})
        alm = pd.read_excel(xl, "ALMACEN ARTICULOS", dtype={"ARTICULO": str})
        car = pd.read_excel(xl, "TABLA CARACTERISTICAS", dtype=str)
        est = pd.read_excel(xl, "MAESTRO DE ESTRUCTURAS", dtype=str)
        clases = pd.read_excel(xl, "CLASE ARTICULO", dtype=str)
        # Conteo de plazos por proveedor desde la tabla completa de proveedores
        _pp = prov[prov["PRINCIPAL"].astype(str) == "1"]
        cnt_plazos = _pp.groupby(["PROVEEDOR", "PLAZOENTREGA"]).size().reset_index(name="N")
        mprov = pd.DataFrame(columns=["CODIGO_PROVEEDOR", "TIPO_CTA_COMPRA"])  # no disponible en modo Excel

    jm = pd.read_excel(jmpath, sheet_name="RELLENAR JOSE MARIA", dtype=str)
    jm["ARTÍCULO"] = jm["ARTÍCULO"].astype(str).str.strip()
    en_jm = set(jm["ARTÍCULO"])

    m["CODIGO_ARTICULO"] = m["CODIGO_ARTICULO"].astype(str).str.strip()
    m["FECHACREACION"] = pd.to_datetime(m["FECHACREACION"], errors="coerce")
    m["FECHAMODIFICACION"] = pd.to_datetime(m["FECHAMODIFICACION"], errors="coerce")

    # Universo ACCESORIOS: familia distinta de 1, 2 y 105.
    # Familia 105 = chapas y paneles — tienen reglas propias, se validarán aparte.
    fam = m["FAMILIAPRODUCTO"].apply(ni)
    es_accesorio = ~fam.isin([1, 2, 105])
    if articulo is not None:
        # Un artículo: se valida si es accesorio O si JM lo lista (lo trata como accesorio,
        # aunque su familia esté mal puesta a 1/2; así la regla FAMILIA detecta el error).
        sub = m[(m["CODIGO_ARTICULO"] == articulo) &
                (es_accesorio | m["CODIGO_ARTICULO"].isin(en_jm))].copy()
    else:
        # Solo artículos CREADOS (altas nuevas), filtrando por FECHACREACION.
        sub = m[es_accesorio].copy()
        if desde is not None:
            sub = sub[sub["FECHACREACION"] >= desde]
        if hasta is not None:
            sub = sub[sub["FECHACREACION"] <= hasta]

    # --- índices de apoyo (siempre proveedor PRINCIPAL = 1) ---
    prov["ARTICULO"] = prov["ARTICULO"].astype(str).str.strip()
    pr = prov[prov["PRINCIPAL"].astype(str) == "1"].set_index("ARTICULO")
    # Coste para el margen = COSTEINTERNO (coste ACTUAL neto: bruto con dto y recargo aplicados).
    # NO usar COSTEBRUTO (es el precio de tarifa del proveedor antes de descuento).
    coste_pr = pd.to_numeric(pr["COSTEINTERNO"], errors="coerce").to_dict()
    plazo_pr = pd.to_numeric(pr["PLAZOENTREGA"], errors="coerce").to_dict()
    lote_pr = pd.to_numeric(pr["LOTEHABITUAL"], errors="coerce").to_dict()
    uext_pr = pd.to_numeric(pr["UNIDADEXTERNA"], errors="coerce").to_dict()
    proveedor_pr = pr["PROVEEDOR"].astype(str).to_dict()  # cod_articulo -> cod_proveedor
    piva_pr = pd.to_numeric(pr["PIVA"], errors="coerce").to_dict()
    formula_pr = pd.to_numeric(pr["FORMULA"], errors="coerce").to_dict()
    con_principal = set(pr.index)

    plazo_habitual = plazo_habitual_por_proveedor(cnt_plazos)

    # TIPO_CTA_COMPRA por proveedor: 0=nacional(ES), 1=intracomunitario, 2=extracomunitario
    mprov["CODIGO_PROVEEDOR"] = pd.to_numeric(mprov["CODIGO_PROVEEDOR"], errors="coerce")
    tipo_cta_prov = mprov.set_index("CODIGO_PROVEEDOR")["TIPO_CTA_COMPRA"].to_dict()

    tar["ARTICULO"] = tar["ARTICULO"].astype(str).str.strip()
    t40 = tar[pd.to_numeric(tar["TARIFA"], errors="coerce") == TARIFA_PVP]
    pvp40 = pd.to_numeric(t40.set_index("ARTICULO")["PRECIO_VENTA_PUBLICO"], errors="coerce").to_dict()
    uventa_tar = pd.to_numeric(t40.set_index("ARTICULO")["UNIDADVENTA"], errors="coerce").to_dict()
    iva_tar40 = pd.to_numeric(t40.set_index("ARTICULO")["IVA"], errors="coerce").to_dict()

    car["ARTICULO"] = car["ARTICULO"].astype(str).str.strip()
    car0 = set(car[(car["CARACTERISTICA"].astype(str) == "0") &
                   (car["VALOR"].astype(str).str.strip().replace("nan", "") != "")]["ARTICULO"])

    alm["ARTICULO"] = alm["ARTICULO"].astype(str).str.strip()
    stockmin = pd.to_numeric(alm.set_index("ARTICULO")["STOCKMINIMO"], errors="coerce").to_dict()

    # Estructura: componentes por artículo padre
    est["ARTSUPERIOR"] = est["ARTSUPERIOR"].astype(str).str.strip()
    est["ARTCOMPONENTE"] = est["ARTCOMPONENTE"].astype(str).str.strip()
    componentes = est.groupby("ARTSUPERIOR")["ARTCOMPONENTE"].apply(set).to_dict()

    # Catálogos = artículos con TIPO 91 (CATALOGOINTERIORES, MADERA, RALESTANDAR, etc.).
    # En accesorios, el único catálogo válido como componente es CATALOGOINTERIORES;
    # el resto de catálogos son de perfiles.
    tipo_de = pd.to_numeric(m["TIPO"], errors="coerce")
    catalogos = set(m.loc[tipo_de == 91, "CODIGO_ARTICULO"])

    jm["ARTÍCULO"] = jm["ARTÍCULO"].astype(str).str.strip()
    jm_idx = jm.set_index("ARTÍCULO")
    divis = pd.to_numeric(jm_idx["DIVISOR ACCESORIO"], errors="coerce").to_dict()
    minimo_jm = pd.to_numeric(jm_idx["MINIMO ACCESORIO"], errors="coerce").to_dict()
    estado_jm = pd.to_numeric(jm_idx["ESTADO"], errors="coerce").to_dict()
    familia_jm = pd.to_numeric(jm_idx["FAMILIA"], errors="coerce").to_dict()
    clase_serie_jm = jm_idx["CLASE(SERIE)"].fillna("").astype(str).to_dict()
    en_jm = set(jm["ARTÍCULO"])

    # Índice CLASE: código -> descripción en mayúsculas
    clases["CODIGO"] = clases["CODIGO"].astype(str).str.strip()
    clases["DENOMINACION"] = clases["DENOMINACION"].astype(str).str.upper().str.strip()
    desc_clase = clases.set_index("CODIGO")["DENOMINACION"].to_dict()

    inc = []
    for _, r in sub.iterrows():
        cod = r["CODIGO_ARTICULO"]
        estado = ni(r.get("ESTADO"))
        tipo = ni(r.get("TIPO"))
        # Artículos TIPO=2 + CLASE=3 (lacados/anodizados para compra): reglas propias
        es_lacado = (tipo == 2 and ni(r.get("CLASE")) == 3)

        # 0) FAMILIA: debe coincidir con la indicada por JM en su Excel.
        fam_esp = ni(familia_jm.get(cod))
        fam_real = ni(r.get("FAMILIAPRODUCTO"))
        if cod in en_jm and fam_esp is not None and fam_real != fam_esp:
            extra = " (familia de PERFIL)" if fam_real in (1, 2) else ""
            inc.append((cod, "FAMILIA", "ERROR",
                        f"Familia debe ser {fam_esp} (según Excel JM). Encontrado: {fam_real}{extra}"))

        # 1) Campos fijos del maestro
        for campo, (esp, msg) in REGLAS_MAESTRO.items():
            val = r.get(campo)
            ok = (ni(val) == esp) if isinstance(esp, int) else (str(val).strip().upper() == str(esp).upper())
            if not ok:
                inc.append((cod, campo, "ERROR", f"{msg}. Encontrado: {val}"))

        # 2) CLASE: verificar que la descripción de la clase contiene alguna de las series indicadas por JM
        clase_cod = str(ni(r.get("CLASE"))) if ni(r.get("CLASE")) is not None else None
        serie_jm = clase_serie_jm.get(cod, "").strip().upper()
        if cod in en_jm and serie_jm:
            desc = desc_clase.get(clase_cod, "")
            # Las series en JM van separadas por "/" — basta con que al menos una aparezca en la descripción
            series = [s.strip() for s in serie_jm.split("/") if s.strip()]
            # Normalizar: quitar guiones y espacios extra para comparar
            def norm(t): return " ".join(t.replace("-", " ").split())
            # Alias de series (nombre comercial -> nombre en tabla de clases)
            ALIAS_SERIE = {"DOPLO": "PR77"}
            series_norm = [norm(ALIAS_SERIE.get(s, s)) for s in series]
            if desc and not any(s in norm(desc) for s in series_norm):
                inc.append((cod, "CLASE", "ERROR",
                            f"Clase {clase_cod} ('{desc}') no corresponde a la serie '{serie_jm}' indicada en Excel JM."))
            elif not desc and clase_cod is not None:
                inc.append((cod, "CLASE", "AVISO",
                            f"Clase {clase_cod} no encontrada en la tabla de clases."))

        # 3) Denominación base
        if not str(r.get("DENOMINACION", "")).strip() or str(r.get("DENOMINACION")).lower() == "nan":
            inc.append((cod, "DENOMINACION", "ERROR", "Falta denominación base."))

        # 3) MEDIDA
        ui, ue = ni(r.get("UNIDAD_INTERNA")), ni(r.get("UNIDAD_EXTERNA"))
        med_esp = 2 if (ui == 4 and ue == 4) else 1
        if ni(r.get("MEDIDA")) != med_esp:
            inc.append((cod, "MEDIDA", "ERROR", f"Medida debe ser {med_esp}. Encontrado: {ni(r.get('MEDIDA'))}"))

        # 3b) COEFICIENTE según caso de unidad
        coef = num(r.get("COEFICIENTE"))
        uart = num(r.get("UNIDADES_ARTICULO"))
        if ui == 9 and ue == 4:
            # Barras: coef debe ser el inverso de la unidad del artículo (1/longitud)
            esp_coef = round(1.0 / uart, 5) if uart else None
            if esp_coef is None:
                inc.append((cod, "COEFICIENTE", "ERROR", "Barra (9/4) sin unidad de artículo para calcular el coeficiente."))
            elif coef is None or abs(coef - esp_coef) > 0.0005:
                inc.append((cod, "COEFICIENTE", "ERROR",
                            f"Barra: coeficiente debe ser 1/{uart} = {esp_coef}. Encontrado: {coef}"))
        else:
            if coef != 1:
                inc.append((cod, "COEFICIENTE", "ERROR", f"Coeficiente debe ser 1. Encontrado: {coef}"))

        # 3c) UNIDADES_ARTICULO según caso de unidad
        if ui == 4 and ue == 4:
            # Metro lineal (rollos de junta/felpa): unidades artículo = LOTEHABITUAL del proveedor
            lote = lote_pr.get(cod)
            if lote is not None and lote > 0 and uart != lote:
                inc.append((cod, "UNIDADES_ARTICULO", "ERROR",
                            f"Metro lineal: unidades del artículo deben ser = lote habitual del proveedor ({lote}). Encontrado: {uart}"))
        elif ui == 9 and ue == 4:
            pass  # Barras: unidades del artículo = longitud (ya validado vía coeficiente)
        else:
            if uart != 1:
                inc.append((cod, "UNIDADES_ARTICULO", "ERROR", f"Unidades del artículo debe ser 1. Encontrado: {uart}"))

        # 4) Grupo según longitud y sufijo del código (lacados: grupo manual, no se valida):
        #    - 8 caracteres (sin acabado): grupo debe ser 0
        #    - >8 caracteres (con acabado): grupo según sufijo (01→12, 03→11, 04→10, 06→14, 05→16)
        grupo_val = ni(r.get("GRUPO"))
        if es_lacado:
            pass  # grupo asignado manualmente en lacados, no se valida
        elif len(cod) <= 9:
            if grupo_val != 0:
                inc.append((cod, "GRUPO", "ERROR",
                            f"Código de {len(cod)} caracteres (sin acabado): grupo debe ser 0. Encontrado: {grupo_val}"))
        elif len(cod) > 9:
            if grupo_val == 0:
                inc.append((cod, "GRUPO", "ERROR",
                            f"Código con sufijo de acabado: grupo no puede ser 0. Debe asignarse el grupo correspondiente."))
            g = GRUPO_POR_SUFIJO.get(cod[8:])
            if g is not None and grupo_val != g:
                inc.append((cod, "GRUPO", "ERROR", f"Sufijo {cod[8:]} => grupo {g}, pero GRUPO={grupo_val}."))
            # DENOMINACION_2 según grupo
            if g in DENOM2_POR_GRUPO:
                d2 = str(r.get("DENOMINACION_2", "")).strip().upper()
                esp2 = DENOM2_POR_GRUPO[g]
                if d2 != esp2:
                    inc.append((cod, "DENOMINACION_2", "ERROR",
                                f"Grupo {g} => denominación 2 debe ser '{esp2}'. Encontrado: '{r.get('DENOMINACION_2')}'"))

        # 5) Imagen (característica 0) — los lacados (tipo 2, clase 3) nunca tienen imagen
        if not es_lacado and cod not in car0:
            inc.append((cod, "CARACTERISTICA_0", "ERROR", "Sin imagen en característica 0."))

        # 6) Proveedor principal + plazo + lote
        if cod not in con_principal:
            inc.append((cod, "PROVEEDOR", "ERROR", "Sin proveedor PRINCIPAL (PRINCIPAL=1) asignado."))
        else:
            prov_id = ni(proveedor_pr.get(cod))
            if es_lacado:
                # Lacados: plazo manual (no se valida), lote = 1, fórmula = 32
                lote = lote_pr.get(cod)
                if lote != 1:
                    inc.append((cod, "LOTE_HABITUAL", "ERROR",
                                f"Lacado (tipo 2, clase 3): lote habitual debe ser 1. Encontrado: {lote}"))
                formula = ni(formula_pr.get(cod))
                if formula != 32:
                    inc.append((cod, "FORMULA_PROVEEDOR", "ERROR",
                                f"Lacado (tipo 2, clase 3): fórmula debe ser 32. Encontrado: {formula}"))
            else:
                plazo = plazo_pr.get(cod)
                if not (plazo and plazo > 1):
                    inc.append((cod, "PLAZO_ENTREGA", "ERROR",
                                f"Plazo de entrega debe ser > 1 día. Encontrado: {plazo}."))
                else:
                    hab = plazo_habitual.get(prov_id) if prov_id is not None else None
                    if hab is not None and hab > 1 and plazo != hab:
                        inc.append((cod, "PLAZO_ENTREGA", "AVISO",
                                    f"Plazo {plazo} difiere del habitual del proveedor {prov_id} ({int(hab)} días). "
                                    "Revisar si es correcto."))
                if not (lote_pr.get(cod) and lote_pr[cod] > 0):
                    inc.append((cod, "LOTE_HABITUAL", "ERROR", "Lote habitual del proveedor principal a 0 o vacío."))
            # IVA proveedor: obligatorio si el proveedor es nacional (TIPO_CTA_COMPRA=0)
            tipo_cta = tipo_cta_prov.get(prov_id) if prov_id is not None else None
            if tipo_cta == 0:
                piva = piva_pr.get(cod)
                if piva is None or piva == 0:
                    inc.append((cod, "IVA_PROVEEDOR", "ERROR",
                                f"Proveedor nacional (España): IVA en artículo proveedor no informado o a 0. "
                                f"Encontrado: {piva}"))

        # 7) PVP / IVA tarifa 40 / Margen — lacados no tienen tarifa de venta, se omite todo
        if not es_lacado:
            pvp = pvp40.get(cod)
            if pvp is None:
                inc.append((cod, "PVP", "ERROR", "Sin PVP en tarifa 40."))

            iva_t = iva_tar40.get(cod)
            if iva_t is None or iva_t == 0:
                inc.append((cod, "IVA_TARIFA", "ERROR",
                            f"IVA en tarifa 40 no informado o a 0. Encontrado: {iva_t}"))

            coste, d = coste_pr.get(cod), divis.get(cod)
            if cod not in en_jm:
                inc.append((cod, "EXCEL_JM", "AVISO", "No está en el Excel de José María (no se puede validar margen)."))
            elif coste is None:
                pass
            elif d is None or d == 0:
                inc.append((cod, "DIVISOR", "ERROR", "Divisor de accesorio vacío o 0 en el Excel de José María."))
            elif pvp is None:
                pass
            else:
                calc = round(coste / d, 2)
                absdif = abs(calc - pvp)
                pct = (absdif / pvp * 100) if pvp else 0
                if absdif > TOL_MARGEN_EUR and pct > TOL_MARGEN_PCT:
                    if pvp < calc:
                        nivel, etiqueta = "ERROR", "vende por DEBAJO del PVP calculado (poco margen)"
                    else:
                        nivel, etiqueta = "AVISO", "PVP por encima del calculado (más margen)"
                    inc.append((cod, "MARGEN", nivel,
                                f"{etiqueta}. Esperado {calc} (coste actual {round(coste,4)}/div {d}) vs real {pvp}. "
                                f"Dif {round(calc-pvp,2)} ({round(pct,1)}%)."))

        # 8) ESTADO: lacados siempre 0; resto según Excel JM
        if es_lacado:
            if estado != 0:
                inc.append((cod, "ESTADO", "ERROR",
                            f"Lacado (tipo 2, clase 3): estado debe ser 0. Encontrado: {estado}"))
        else:
            estado_esp = ni(estado_jm.get(cod))
            if cod in en_jm and estado_esp is not None and estado != estado_esp:
                inc.append((cod, "ESTADO", "ERROR",
                            f"Estado debe ser {int(estado_esp)} (según Excel JM). Encontrado: {estado}"))

        # 8b) Stock mínimo (lacados no gestionan stock mínimo — son artículos de compra): cruzar Excel JM (col. MINIMO ACCESORIO) vs ALMACEN, según estado.
        #   - JM = 0  y almacén > 0                  -> ERROR (no debe tener mínimo)
        #   - JM > 0  y almacén > 0                  -> OK (el valor exacto se ajusta por ventas)
        #   - JM > 0  y almacén = 0 y estado(JM) 70  -> ERROR (en tarifa: debe tener mínimo)
        #   - JM > 0  y almacén = 0 y estado(JM) != 70 -> AVISO (aún no en tarifa)
        sm = stockmin.get(cod) if not es_lacado else None
        mj = minimo_jm.get(cod) if not es_lacado else None
        if mj is not None and not pd.isna(mj):  # celda vacía en Excel JM => no se valida
            sm_pos = sm is not None and sm > 0
            if mj == 0:
                if sm_pos:
                    inc.append((cod, "STOCK_MINIMO", "ERROR",
                                f"Excel JM indica mínimo 0 pero almacén tiene {sm}."))
            else:  # mj > 0
                if not sm_pos:
                    estado_esp_sm = ni(estado_jm.get(cod))
                    if estado_esp_sm == 70:
                        inc.append((cod, "STOCK_MINIMO", "ERROR",
                                    f"Estado 70 (en tarifa): debe tener stock mínimo en almacén. "
                                    f"Excel JM indica {mj} y almacén = {sm}."))
                    else:
                        inc.append((cod, "STOCK_MINIMO", "AVISO",
                                    f"Excel JM indica mínimo {mj} pero almacén no tiene mínimo (={sm})."))

        # 9) TIPO (accesorios). Reglas confirmadas con JM:
        #    - 0, 16, 91 NO deben existir en accesorios.
        #    - Con estructura que contiene CATALOGOINTERIORES => 92 (modular).
        #    - Con estructura que contiene OTRO catálogo (tipo 91, p.ej. MADERA/RALESTANDAR)
        #      => es de perfiles, no debería estar en accesorios.
        #    - Con estructura de códigos de compraventa => 15 (kit, solo accesorios).
        #    - Sin estructura => 1 (compraventa), 2 (lacado), 5 (muestra) o 60.
        comps = componentes.get(cod, set())
        cat_comps = comps & catalogos  # catálogos presentes en la estructura
        otros_cat = cat_comps - {"CATALOGOINTERIORES"}
        if tipo in (0, 16, 91):
            inc.append((cod, "TIPO", "ERROR", f"Tipo {tipo} no debe existir en accesorios."))
        elif comps:
            if "CATALOGOINTERIORES" in comps:
                if tipo != 92:
                    inc.append((cod, "TIPO", "ERROR",
                                f"Tiene CATALOGOINTERIORES => debe ser tipo 92 (modular). Encontrado: {tipo}"))
            elif otros_cat:
                inc.append((cod, "TIPO", "ERROR",
                            f"Estructura con catálogo de perfil {sorted(otros_cat)} (en accesorios solo vale CATALOGOINTERIORES). Tipo: {tipo}"))
            else:
                if tipo != 15:
                    inc.append((cod, "TIPO", "ERROR",
                                f"Estructura de códigos de compraventa => debe ser tipo 15 (kit). Encontrado: {tipo}"))
        else:
            if tipo not in (1, 2, 5, 60):
                inc.append((cod, "TIPO", "ERROR",
                            f"Sin estructura: tipo debe ser 1 (compraventa), 2 (lacado), 5 (muestra) o 60. Encontrado: {tipo}"))

        # 9b) Pestaña Más: PARTIDA ARANCELARIA y CONVERSIÓN M/KG
        partida = ni(r.get("PARTIDAARANCELARIA"))
        if es_lacado:
            # Lacados (tipo 2, clase 3): partida siempre debe ser 0
            if partida != 0:
                inc.append((cod, "PARTIDAARANCELARIA", "ERROR",
                            f"Lacado (tipo 2, clase 3): partida arancelaria debe ser 0. Encontrado: {partida}"))
        else:
            if partida is None or partida == 0:
                inc.append((cod, "PARTIDAARANCELARIA", "ERROR",
                            f"Partida arancelaria no puede ser 0 o vacía. Encontrado: {partida}"))
        conv_mkg = num(r.get("CONVERSION_M_KG"))
        if fam_real in (1, 2) and (conv_mkg is None or conv_mkg <= 0):
            inc.append((cod, "CONVERSION_M_KG", "ERROR",
                        f"Familia {fam_real} (perfil): coeficiente m/kg debe ser > 0. Encontrado: {conv_mkg}"))

        # 9c) Pestaña Mas2: checks
        vars_check = ni(r.get("GESTIONARVARIABLES"))
        subcont = ni(r.get("ESSUBCONTRATACIO"))
        eskit = ni(r.get("ESKIT"))
        calc_formula = ni(r.get("CALCULOCANTFORMULA"))

        # GESTIONARVARIABLES: debe estar marcado cuando unidad externa = 9 (barras)
        if ue == 9 and vars_check != 1:
            inc.append((cod, "GESTIONARVARIABLES", "ERROR",
                        "Unidad externa 9 (barras): 'Gestionar variables' debe estar marcado "
                        f"(permite indicar nº de barras en artículos variables). Encontrado: {vars_check}"))

        # ESSUBCONTRATACIO: debe estar marcado cuando TIPO = 92 (modular)
        if tipo == 92 and subcont != 1:
            inc.append((cod, "ESSUBCONTRATACIO", "ERROR",
                        f"Tipo 92 (modular): 'Es subcontratación' debe estar marcado. Encontrado: {subcont}"))

        # ESKIT: siempre debe estar vacío (nunca se usa en accesorios)
        if eskit == 1:
            inc.append((cod, "ESKIT_CHECK", "ERROR",
                        "El check 'Es un artículo kit' no debe estar marcado en accesorios."))

        # CALCULOCANTFORMULA debe estar marcado (=1) si unidad interna = 9
        if ui == 9 and calc_formula != 1:
            inc.append((cod, "CALCULOCANTFORMULA", "ERROR",
                        "Unidad interna 9 (barras): 'Cálculo cantidad en documentos comerciales por fórmula' "
                        f"debe estar marcado. Encontrado: {calc_formula}"))

        # 10) Coherencia de unidades MAESTRO <-> PROVEEDOR <-> TARIFABLE
        # 10a) TIPO 2 (lacado/anodizado) <=> UNIDAD_INTERNA 5 (m2)
        if tipo == 2 and ui != 5:
            inc.append((cod, "UNIDAD/TIPO", "ERROR", f"Tipo 2 (lacado) debe tener unidad interna 5. Encontrado: {ui}"))
        if ui == 5 and tipo != 2:
            inc.append((cod, "UNIDAD/TIPO", "ERROR", f"Unidad interna 5 (m2) debe ser tipo 2. Encontrado tipo: {tipo}"))
        # 10b) externa del maestro = unidad del proveedor principal
        uep = uext_pr.get(cod)
        if uep is not None and ue is not None and uep != ue:
            inc.append((cod, "UNIDAD_PROVEEDOR", "ERROR",
                        f"Unidad externa maestro={ue} pero proveedor principal={int(uep)}."))
        # 10c) unidad de venta tarifa 40 coherente con interna (1->1, 9->4, 4->4, 5->5)
        venta_esp = {1: 1, 9: 4, 4: 4, 5: 5}.get(ui)
        uvt = uventa_tar.get(cod)
        if venta_esp is not None and uvt is not None and uvt != venta_esp:
            inc.append((cod, "UNIDAD_VENTA", "ERROR",
                        f"Interna {ui} => unidad de venta tarifa esperada {venta_esp}. Encontrado: {int(uvt)}."))

    # Orden de campos según pestañas de Geinfor (Maestro → Más → Mas2 → Caract. → Proveedor → Tarifa)
    _ORDEN_CAMPO = [
        # Maestro de artículos
        "FAMILIA", "CLASE", "DENOMINACION", "DENOMINACION_2", "GRUPO",
        "TIPO", "ESTADO", "MEDIDA", "COEFICIENTE", "UNIDADES_ARTICULO", "UNIDAD/TIPO",
        # Pestaña Más
        "PARTIDAARANCELARIA", "CONVERSION_M_KG",
        # Pestaña Mas2
        "GESTIONARVARIABLES", "ESSUBCONTRATACIO", "ESKIT_CHECK", "NOGESTIONASTOCK", "CALCULOCANTFORMULA",
        # Características
        "CARACTERISTICA_0",
        # Artículo proveedor
        "PROVEEDOR", "PLAZO_ENTREGA", "LOTE_HABITUAL", "UNIDAD_PROVEEDOR", "FORMULA_PROVEEDOR",
        # Tarifa / PVP
        "IVA_TARIFA", "UNIDAD_VENTA", "DIVISOR", "MARGEN", "PVP",
        # Proveedor - IVA
        "IVA_PROVEEDOR",
        # Excel JM
        "EXCEL_JM",
    ]
    _ord = {c: i for i, c in enumerate(_ORDEN_CAMPO)}

    out = pd.DataFrame(inc, columns=["CODIGO", "CAMPO", "NIVEL", "DETALLE"])
    out["_ord"] = out["CAMPO"].map(lambda c: _ord.get(c, 999))
    out = out.sort_values(["CODIGO", "_ord", "NIVEL"]).drop(columns="_ord").reset_index(drop=True)

    if articulo is not None:
        # Modo un artículo: SIEMPRE genera informe visual HTML (no toca el Excel).
        ruta = "INFORME_" + re.sub(r"[^A-Za-z0-9]", "_", articulo) + ".html"
        if len(sub) == 0:
            fila_m = m[m["CODIGO_ARTICULO"] == articulo]
            if len(fila_m) == 0:
                informe_mensaje(articulo, "Artículo no encontrado",
                                "Este código no existe en Geinfor. Revisa que esté bien escrito.", ruta)
                print(f"{articulo}: no encontrado. Informe -> {ruta}")
            else:
                famx = ni(fila_m.iloc[0].get("FAMILIAPRODUCTO"))
                informe_mensaje(articulo, "Es un PERFIL, no un accesorio",
                                f"Familia de producto {famx} (las familias 1 y 2 son perfiles). "
                                "El validador de accesorios no aplica a perfiles; tendrán su propia "
                                "validación.", ruta, clase="bad")
                print(f"{articulo}: es perfil (familia {famx}). Informe -> {ruta}")
            try:
                import os
                os.startfile(os.path.abspath(ruta))
            except Exception:
                pass
            return
        inc_art = [t for t in inc if t[0] == articulo]
        generar_informe_html(articulo, sub.iloc[0], inc_art, ruta)
        ne = sum(1 for t in inc_art if t[2] == "ERROR")
        na = sum(1 for t in inc_art if t[2] == "AVISO")
        print(f"{articulo}: {ne} error(es), {na} aviso(s). Informe -> {ruta}")
        try:
            import os
            os.startfile(os.path.abspath(ruta))  # abre el informe en el navegador (Windows)
        except Exception:
            pass
        return

    # Modo lote: informe visual HTML (principal) + Excel (para análisis en tabla).
    if desde is not None and hasta is not None:
        etiqueta = f"Altas creadas entre {desde.date()} y {hasta.date()}"
    elif desde is not None:
        etiqueta = f"Altas creadas desde {desde.date()}"
    else:
        etiqueta = "Todas las altas"
    ruta_html = "INFORME_ALTAS.html"
    generar_informe_lote_html(sub, out, etiqueta, ruta_html)

    destino = "INCIDENCIAS_FINAL.xlsx"
    try:
        out.to_excel(destino, index=False)
    except PermissionError:
        destino = "INCIDENCIAS_" + pd.Timestamp.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
        out.to_excel(destino, index=False)
        print(f"(INCIDENCIAS_FINAL.xlsx estaba abierto; guardado en {destino})")

    print(f"Accesorios FW evaluados: {len(sub)}")
    print(f"Incidencias: {len(out)} (ERROR: {(out.NIVEL=='ERROR').sum()}, AVISO: {(out.NIVEL=='AVISO').sum()})")
    print(f"Informe visual -> {ruta_html}  |  Tabla -> {destino}")
    try:
        import os
        os.startfile(os.path.abspath(ruta_html))
    except Exception:
        pass


if __name__ == "__main__":
    main()

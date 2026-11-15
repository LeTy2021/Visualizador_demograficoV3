"""
Funciones para procesar los datos del dashboard.

1. Lectura de los cuadros Excel del INE y armado de las tablas en formato largo (tidy),
   que se guardan como CSV en datos_procesados/.
2. Consultas y utilidades que usa el dashboard sobre esas tablas
   (series, último valor disponible, datos de la pirámide, formato numérico, exportación a Excel).

Uso desde la terminal:
    python procesamiento.py                     # usa las carpetas definidas en config.py
    python procesamiento.py <carpeta_excel> <carpeta_salida>
"""
import io
import re
import sys
import unicodedata
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

import config as cfg

COLS_POB = ["departamento", "anio", "sexo", "edad", "edad_inicio", "poblacion"]
COLS_IND = ["codigo", "departamento", "anio", "sexo", "valor"]
COLS_TV = ["qx", "lx", "dx", "Lx", "Tx", "ex"]


# ============================================================================
# 1. Lectura de los cuadros del INE
# ============================================================================
def _norm(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[\s_]+", "_", s).lower()


def buscar(dir_datos, codigo):
    """Primer .xlsx cuyo nombre contiene el código del cuadro (p. ej. 'A.2.1_').
    Ignora diferencias entre espacios y guiones bajos, mayúsculas y acentos."""
    archivos = sorted(Path(dir_datos).rglob("*.xlsx"))
    hallados = [f for f in archivos if _norm(codigo) in _norm(f.name)]
    if not hallados:
        print(f"  ! No se encontró el archivo {codigo}*.xlsx en {dir_datos} (se omite)")
        return None
    return hallados[0]


def filas(path, hoja):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = [list(r) for r in wb[hoja].iter_rows(values_only=True)]
    wb.close()
    return out


def hojas(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    nombres = wb.sheetnames
    wb.close()
    return nombres


def es_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def es_anio(v):
    return es_numero(v) and 1900 <= v <= 2200 and float(v).is_integer()


def limpiar(txt):
    return re.sub(r"\s+", " ", str(txt)).strip()


def edad_inicio(lbl):
    m = re.match(r"\s*(\d+)", str(lbl))
    return int(m.group(1)) if m else None


def fuente_de(anio):
    return "Proyección" if anio >= cfg.ANIO_PROY else "Estimación"


def fila_encabezado_anio(rows):
    """Índice de la fila cuyo primer valor es 'Año'."""
    return next(i for i, r in enumerate(rows) if isinstance(r[0], str) and limpiar(r[0]) == "Año")


def es_pie(c0):
    return isinstance(c0, str) and limpiar(c0).lower().startswith(("fuente", "nota"))


# ----------------------------------------------------------------------------
# Población por sexo y edad
# ----------------------------------------------------------------------------
def leer_poblacion(path, hoja, depto):
    rows = filas(path, hoja)
    # fila de encabezado: la que tiene años en las columnas 1..n
    i_hdr = next(i for i, r in enumerate(rows) if sum(es_anio(v) for v in r[1:]) >= 3)
    anios = {j: int(v) for j, v in enumerate(rows[i_hdr]) if j > 0 and es_anio(v)}
    regs, sexo = [], None
    for r in rows[i_hdr + 1:]:
        c0 = r[0]
        if c0 is None:
            continue
        if es_pie(c0):
            break
        lbl = limpiar(c0)
        if lbl in cfg.SEXOS:
            sexo = lbl
            continue
        if sexo is None:
            continue
        for j, a in anios.items():
            v = r[j] if j < len(r) else None
            if es_numero(v):
                regs.append((depto, a, sexo, lbl, edad_inicio(lbl), float(v)))
    return pd.DataFrame(regs, columns=COLS_POB)


def poblacion_simple(dir_datos):
    partes = []
    for cod in cfg.CUADROS_POB_SIMPLE:
        p = buscar(dir_datos, cod)
        if p:
            partes.append(leer_poblacion(p, hojas(p)[0], "Uruguay"))
    df = pd.concat(partes, ignore_index=True)
    df["edad_90"] = df["edad_inicio"].clip(upper=cfg.EDAD_ABIERTA)  # edad abierta armonizada
    df["fuente"] = df["anio"].map(fuente_de)
    return df


def poblacion_quinquenal(dir_datos):
    partes = []
    for cod in cfg.CUADROS_POB_QUINQ:
        p = buscar(dir_datos, cod)
        if p:
            partes += [leer_poblacion(p, h, limpiar(h)) for h in hojas(p)]
    df = pd.concat(partes, ignore_index=True)
    ea = cfg.EDAD_ABIERTA
    df["grupo"] = df["edad_inicio"].apply(lambda e: f"{ea} y más" if e >= ea else f"{e}-{e + 4}")
    df["edad_inicio"] = df["edad_inicio"].clip(upper=ea)
    df = df.groupby(["departamento", "anio", "sexo", "grupo", "edad_inicio"], as_index=False)["poblacion"].sum()
    df["fuente"] = df["anio"].map(fuente_de)
    return df


# ----------------------------------------------------------------------------
# Indicadores (hojas año x departamento, con bloques por sexo opcionales)
# ----------------------------------------------------------------------------
def leer_indicador_deptos(path, hoja, codigo):
    rows = filas(path, hoja)
    i_hdr = fila_encabezado_anio(rows)
    cols = {j: limpiar(v) for j, v in enumerate(rows[i_hdr]) if j > 0 and isinstance(v, str) and v.strip()}
    regs, sexo = [], "Ambos sexos"
    for r in rows[i_hdr + 1:]:
        c0 = r[0]
        if isinstance(c0, str):
            if limpiar(c0) in cfg.SEXOS:
                sexo = limpiar(c0)
                continue
            if es_pie(c0):
                break
        if not es_anio(c0):
            continue
        for j, depto in cols.items():
            v = r[j] if j < len(r) else None
            if es_numero(v):
                regs.append((codigo, depto, int(c0), sexo, float(v)))
    return pd.DataFrame(regs, columns=COLS_IND)


def leer_uruguay_proy(path, hoja):
    rows = filas(path, hoja)
    i_hdr = fila_encabezado_anio(rows)
    mapa = {}
    for j, v in enumerate(rows[i_hdr]):
        if j == 0 or v is None:
            continue
        clave = re.sub(r"\s*%", " %", limpiar(v).lower())
        if clave in cfg.COLS_URUGUAY:
            mapa[j] = cfg.COLS_URUGUAY[clave]
        else:
            print(f"  ! Columna sin mapear en {hoja}: {v!r}")
    regs = [(cod, "Uruguay", int(r[0]), sexo, float(r[j]))
            for r in rows[i_hdr + 1:] if es_anio(r[0])
            for j, (cod, sexo) in mapa.items() if es_numero(r[j])]
    return pd.DataFrame(regs, columns=COLS_IND)


def indicadores(dir_datos):
    partes = []
    for cod_archivo in cfg.CUADROS_INDICADORES:
        p = buscar(dir_datos, cod_archivo)
        if not p:
            continue
        for h in hojas(p):
            hl = limpiar(h)
            if hl.startswith(cfg.PREFIJO_HOJA_NACIONAL):
                partes.append(leer_uruguay_proy(p, h))
            elif hl in cfg.HOJAS:
                df = leer_indicador_deptos(p, h, cfg.HOJAS[hl])
                # En proyecciones, Uruguay se toma de la hoja nacional (llega más lejos)
                if cod_archivo.startswith("B."):
                    df = df[df["departamento"] != "Uruguay"]
                partes.append(df)
    return pd.concat(partes, ignore_index=True)


def tef(dir_datos):
    p = buscar(dir_datos, cfg.CUADRO_TEF)
    if not p:
        return pd.DataFrame(columns=["anio", "grupo", "tef"])
    rows = filas(p, cfg.HOJA_TEF)
    i_hdr = fila_encabezado_anio(rows)
    grupos = {j: limpiar(v) for j, v in enumerate(rows[i_hdr]) if j > 0 and v}
    regs = [(int(r[0]), g, float(r[j])) for r in rows[i_hdr + 1:] if es_anio(r[0])
            for j, g in grupos.items() if es_numero(r[j])]
    return pd.DataFrame(regs, columns=["anio", "grupo", "tef"])


def tablas_vida(dir_datos):
    p = buscar(dir_datos, cfg.CUADRO_TABLAS_VIDA)
    if not p:
        return pd.DataFrame(columns=["anio", "sexo", "edad"] + COLS_TV)
    partes = []
    for h, sexo in cfg.HOJAS_TABLAS_VIDA.items():
        rows = filas(p, h)
        i_hdr = next(i for i, r in enumerate(rows) if "qx" in [limpiar(v) for v in r if v])
        hdr = [limpiar(v) if v else None for v in rows[i_hdr]]
        idx = {n: hdr.index(n) for n in ["Año", "Edad"] + COLS_TV}
        regs = [[int(r[idx["Año"]]), sexo, int(r[idx["Edad"]])] + [float(r[idx[c]]) for c in COLS_TV]
                for r in rows[i_hdr + 1:] if es_anio(r[idx["Año"]]) and r[idx["Edad"]] is not None]
        partes.append(pd.DataFrame(regs, columns=["anio", "sexo", "edad"] + COLS_TV))
    return pd.concat(partes, ignore_index=True)


# ----------------------------------------------------------------------------
# Indicadores derivados
# ----------------------------------------------------------------------------
def tabla_ambos_sexos(tv, ps):
    """Tabla de mortalidad de ambos sexos combinando las de hombres y mujeres con el índice de
    masculinidad al nacer (s), aproximado por la relación H/M de la población de 0 años de cada año
    corregida por la supervivencia en el primer año:  s = (P0_H / P0_M) * (L0_M / L0_H).
    lx, dx, Lx y Tx de ambos sexos = (s * hombres + mujeres) / (1 + s);  ex = Tx / lx;  qx = dx / lx."""
    if tv.empty:
        return tv
    p0 = ps[ps["edad_inicio"] == 0].pivot_table(index="anio", columns="sexo", values="poblacion")
    h = tv[tv.sexo == "Hombres"].set_index(["anio", "edad"])
    m = tv[tv.sexo == "Mujeres"].set_index(["anio", "edad"])
    l0 = pd.DataFrame({"H": h.xs(0, level="edad")["Lx"], "M": m.xs(0, level="edad")["Lx"]})
    s = (p0["Hombres"] / p0["Mujeres"]).reindex(l0.index) * (l0["M"] / l0["H"])
    s = s.fillna(cfg.MASC_NACIMIENTO_DEFECTO)
    w = s.reindex(h.index.get_level_values("anio")).values
    a = pd.DataFrame(index=h.index)
    for c in ("lx", "dx", "Lx", "Tx"):
        a[c] = (w * h[c].values + m[c].values) / (1 + w)
    a["ex"] = a["Tx"] / a["lx"]
    a["qx"] = a["dx"] / a["lx"]
    a["sexo"] = "Ambos sexos"
    a = a.reset_index()[["anio", "sexo", "edad"] + COLS_TV]
    return pd.concat([tv, a], ignore_index=True)


def poblacion_total(ps, pq):
    """Población total como indicador POB: departamentos desde la quinquenal y, para los años
    en que no hay datos departamentales, Uruguay desde la edad simple."""
    pob = (pq.groupby(["departamento", "anio", "sexo"], as_index=False)["poblacion"].sum()
             .rename(columns={"poblacion": "valor"}))
    pob_uy = (ps[ps["anio"] > pob["anio"].max()]
              .groupby(["anio", "sexo"], as_index=False)["poblacion"].sum()
              .rename(columns={"poblacion": "valor"}))
    pob_uy["departamento"] = "Uruguay"
    pob = pd.concat([pob, pob_uy], ignore_index=True)
    pob["codigo"] = "POB"
    return pob[COLS_IND]


def migracion_neta_residual(ind):
    """Tasa neta de migración = TCT − TCN, sólo para las combinaciones área-año-sexo en que el
    INE no la publica (período de estimación). La identidad se cumple en las proyecciones."""
    claves = ["departamento", "anio", "sexo"]
    w = ind[ind.codigo.isin(["TCT", "TCN"])].pivot_table(index=claves, columns="codigo", values="valor").dropna()
    tnm = (w["TCT"] - w["TCN"]).rename("valor").reset_index()
    publicadas = ind.loc[ind.codigo == "TNM", claves].drop_duplicates()
    tnm = tnm.merge(publicadas, on=claves, how="left", indicator=True)
    tnm = tnm[tnm["_merge"] == "left_only"].drop(columns="_merge")
    tnm["codigo"] = "TNM"
    return tnm[COLS_IND]


def esperanza_vida_proyectada(tv):
    """E0 y E65 del total país desde las tablas de mortalidad proyectadas."""
    if tv.empty:
        return pd.DataFrame(columns=COLS_IND)
    partes = []
    for edad, cod in ((0, "E0"), (65, "E65")):
        e = tv[tv["edad"] == edad][["anio", "sexo", "ex"]].rename(columns={"ex": "valor"})
        e["codigo"], e["departamento"] = cod, "Uruguay"
        partes.append(e)
    return pd.concat(partes, ignore_index=True)[COLS_IND]


def catalogo():
    return pd.DataFrame([(k, *v, cfg.DEFINICIONES.get(k, "")) for k, v in cfg.CATALOGO.items()],
                        columns=["codigo", "indicador", "tema", "unidad", "definicion"])


# ----------------------------------------------------------------------------
# Proceso completo
# ----------------------------------------------------------------------------
def procesar(dir_datos=cfg.DIR_DATOS, dir_salida=cfg.DIR_PROCESADOS):
    dir_salida = Path(dir_salida)
    dir_salida.mkdir(parents=True, exist_ok=True)

    print("Procesando población por edad simple...")
    ps = poblacion_simple(dir_datos)
    print("Procesando población quinquenal por departamento...")
    pq = poblacion_quinquenal(dir_datos)
    print("Procesando indicadores...")
    ind = pd.concat([indicadores(dir_datos), poblacion_total(ps, pq)], ignore_index=True)
    if cfg.CALCULAR_TNM:
        ind = pd.concat([ind, migracion_neta_residual(ind)], ignore_index=True)
    tv = tabla_ambos_sexos(tablas_vida(dir_datos), ps)
    ind = pd.concat([ind, esperanza_vida_proyectada(tv)], ignore_index=True)

    ind["fuente"] = ind["anio"].map(fuente_de)
    ind = ind.drop_duplicates(["codigo", "departamento", "anio", "sexo"], keep="first")
    cat = catalogo()
    ind = ind.merge(cat.drop(columns="definicion"), on="codigo", how="left")
    ind = ind.sort_values(["tema", "codigo", "departamento", "sexo", "anio"])

    salidas = {"ind": ind, "ps": ps, "pq": pq, "cat": cat, "tef": tef(dir_datos), "tv": tv}
    for clave, df in salidas.items():
        df.to_csv(dir_salida / cfg.ARCHIVOS_PROCESADOS[clave], index=False)

    print("\nListo. Resumen:")
    print(f"  poblacion_edad_simple: {len(ps):>7,} filas  ({ps.anio.min()}-{ps.anio.max()})")
    print(f"  poblacion_quinquenal:  {len(pq):>7,} filas  ({pq.departamento.nunique()} áreas)")
    print(f"  indicadores:           {len(ind):>7,} filas  ({ind.codigo.nunique()} indicadores)")
    print(f"  tablas_mortalidad:     {len(tv):>7,} filas")
    print(ind.groupby("codigo")["anio"].agg(["min", "max"]).T.to_string())
    return salidas


def datos_procesados_listos(dir_salida=cfg.DIR_PROCESADOS):
    return all((Path(dir_salida) / f).exists() for f in cfg.ARCHIVOS_PROCESADOS.values())


# ============================================================================
# 2. Consultas y utilidades para el dashboard
# ============================================================================
def cargar(dir_salida=cfg.DIR_PROCESADOS):
    """Devuelve dict con ind, ps, pq, cat, tef, tv."""
    return {k: pd.read_csv(Path(dir_salida) / f) for k, f in cfg.ARCHIVOS_PROCESADOS.items()}


def serie(ind, codigo, area, sexo="Ambos sexos"):
    d = ind[(ind.codigo == codigo) & (ind.departamento == area) & (ind.sexo == sexo)]
    return d.sort_values("anio")


def valor(ind, codigo, area, anio, sexo="Ambos sexos"):
    """Último valor disponible con año <= anio. Devuelve (valor, año) o (None, None)."""
    d = serie(ind, codigo, area, sexo)
    d = d[d.anio <= anio]
    if d.empty:
        return None, None
    r = d.iloc[-1]
    return r.valor, int(r.anio)


def info(cat, codigo):
    """(nombre, unidad, tema) del indicador."""
    r = cat[cat.codigo == codigo].iloc[0]
    return r.indicador, r.unidad, r.tema


def definicion(cat, codigo):
    d = cat.loc[cat.codigo == codigo, "definicion"]
    return "" if d.empty or pd.isna(d.iloc[0]) else d.iloc[0]


def datos_piramide(ps, pq, area, anio, simple=False):
    """Población por sexo y grupo de edad (quinquenal o simple) con % sobre el total."""
    ea = cfg.EDAD_ABIERTA
    if area == "Uruguay":
        d = ps[(ps.anio == anio) & (ps.sexo != "Ambos sexos")].copy()
        if simple:
            d["grupo"], d["orden"] = d["edad_90"].astype(int).astype(str), d["edad_90"]
            d.loc[d.edad_90 == ea, "grupo"] = f"{ea}+"
        else:
            d["orden"] = (d["edad_90"] // 5 * 5).clip(upper=ea)
            d["grupo"] = d["orden"].apply(lambda e: f"{ea} y más" if e >= ea else f"{int(e)}-{int(e) + 4}")
    else:
        d = pq[(pq.departamento == area) & (pq.anio == anio) & (pq.sexo != "Ambos sexos")].copy()
        d["orden"] = d["edad_inicio"]
    d = d.groupby(["sexo", "grupo", "orden"], as_index=False)["poblacion"].sum().sort_values("orden")
    d["pct"] = 100 * d["poblacion"] / d["poblacion"].sum()
    return d


def grandes_grupos(ps, pq, area):
    """% de población de 0-14, 15-64 y 65+ por año."""
    base = (ps[ps.sexo == "Ambos sexos"] if area == "Uruguay"
            else pq[(pq.departamento == area) & (pq.sexo == "Ambos sexos")].rename(columns={"edad_inicio": "edad_90"}))
    base = base.assign(Grupo=pd.cut(base["edad_90"], [-1, 14, 64, 200], labels=["0-14", "15-64", "65 y más"]))
    g = base.groupby(["anio", "Grupo"], observed=True, as_index=False)["poblacion"].sum()
    g["pct"] = 100 * g.poblacion / g.groupby("anio").poblacion.transform("sum")
    return g


def fmt(x, dec=0):
    """Formato numérico rioplatense: 1.234.567,8"""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    s = f"{x:,.{dec}f}"
    return s.replace(",", "§").replace(".", ",").replace("§", ".")


def a_excel(df, hoja="datos"):
    """Bytes de un .xlsx con la tabla y la fuente al pie."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=hoja[:31])
        ws = w.sheets[hoja[:31]]
        for col in ws.columns:  # ancho de columnas
            ws.column_dimensions[col[0].column_letter].width = min(45, max(10, len(str(col[0].value)) + 4))
        ws.cell(row=ws.max_row + 2, column=1, value=f"Fuente: {cfg.FUENTE_TXT}")
    return buf.getvalue()


def nombre_archivo(txt):
    return "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in txt)


def sin_acentos(txt):
    return unicodedata.normalize("NFKD", txt.lower()).encode("ascii", "ignore").decode()


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else cfg.DIR_DATOS
    salida = sys.argv[2] if len(sys.argv) > 2 else cfg.DIR_PROCESADOS
    procesar(entrada, salida)

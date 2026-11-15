"""
Panorama demográfico del Uruguay — visualizador de estimaciones y proyecciones de población.

Este archivo tiene sólo el dashboard: estilos, gráficos y pestañas.
- Constantes, códigos, textos y colores: config.py
- Lectura de los Excel y consultas sobre los datos: procesamiento.py
- Levantar el servidor: python lanzar.py
"""
import base64

import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from packaging.version import Version

import config as cfg
import procesamiento as pr
from procesamiento import fmt

# Streamlit >= 1.50 usa width="stretch"; versiones anteriores, use_container_width=True
ANCHO = {"width": "stretch"} if Version(st.__version__) >= Version("1.50") else {"use_container_width": True}
PALETA = cfg.PALETA + px.colors.qualitative.Pastel
AP = cfg.ANIO_PROY

st.set_page_config(page_title=f"{cfg.TITULO} | INE", page_icon="📊", layout="wide")

# Evita que el navegador traduzca la página (Chrome cambiaba "Artigas" por "Artículos")
components.html("""<script>
const d = window.parent.document;
d.documentElement.lang = 'es';
d.documentElement.setAttribute('translate', 'no');
d.documentElement.classList.add('notranslate');
if (!d.querySelector('meta[name="google"]')) {
  const m = d.createElement('meta'); m.name = 'google'; m.content = 'notranslate'; d.head.appendChild(m);
}
</script>""", height=0)


# ============================================================================
# Estilos
# ============================================================================
@st.cache_data
def logo_b64():
    return base64.b64encode(cfg.LOGO.read_bytes()).decode() if cfg.LOGO.exists() else None


def logo_html(alto=80):
    b64 = logo_b64()
    if not b64:
        return '<span class="logo-txt">INE</span>'
    return f'<img src="data:image/png;base64,{b64}" style="height:{alto}px" alt="Instituto Nacional de Estadística">'


st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family={cfg.FUENTE.replace(" ", "+")}:wght@400;600;700&display=swap');
  html, body, .stMarkdown, button, input, textarea {{
      font-family: '{cfg.FUENTE}', 'Segoe UI', Arial, sans-serif;}}
  .block-container {{padding-top: 1rem; padding-bottom: 2rem;}}
  .banda {{background: white; border-bottom: 4px solid {cfg.AZUL_OSC}; box-shadow: 0 3px 0 {cfg.AMARILLO};
          padding: 10px 6px; margin-bottom: 14px; display: flex; align-items: center; gap: 22px;}}
  .banda .txt {{border-left: 2px solid #D5DEE8; padding-left: 22px;}}
  .banda h1 {{color: {cfg.AZUL_OSC}; font-size: 1.65rem; margin: 0; font-weight: 700; padding: 0;}}
  .banda p {{color: {cfg.GRIS}; margin: 2px 0 0 0; font-size: .92rem;}}
  .logo-txt {{background: {cfg.AZUL_OSC}; color: white; font-weight: 800; font-size: 1.6rem;
             letter-spacing: .08em; padding: 6px 14px; border-radius: 4px;}}
  .defin {{background: {cfg.AZUL_CLARO}; border-left: 4px solid {cfg.AZUL}; padding: 8px 12px;
          color: #34495E; font-size: .86rem; border-radius: 3px; margin: 4px 0 10px 0;}}
  .kpi {{background: white; border: 1px solid #DCE3EA; border-top: 4px solid {cfg.AZUL_OSC};
        border-radius: 6px; padding: 12px 14px; height: 118px;}}
  .kpi .lbl {{color: {cfg.GRIS}; font-size: .78rem; text-transform: uppercase; letter-spacing: .03em;}}
  .kpi .val {{color: {cfg.AZUL_OSC}; font-size: 1.65rem; font-weight: 700; line-height: 1.25;}}
  .kpi .sub {{color: {cfg.GRIS}; font-size: .76rem;}}
  .kpi .up {{color: #2E7D32;}} .kpi .down {{color: #C62828;}}
  .seccion {{color: {cfg.AZUL_OSC}; font-weight: 700; font-size: 1.05rem;
            border-bottom: 2px solid {cfg.AZUL_CLARO}; padding-bottom: 4px; margin: 6px 0 8px 0;}}
  .nota {{color: {cfg.GRIS}; font-size: .78rem;}}
  div[data-testid="stTabs"] button p {{font-size: .95rem; font-weight: 600; color: {cfg.AZUL_OSC};}}
  section[data-testid="stSidebar"] {{background: {cfg.AZUL_CLARO};}}
  div[data-testid="stMetricValue"] {{font-size: 1.35rem; color: {cfg.AZUL_OSC}; font-weight: 700;}}
  div[data-testid="stMetricLabel"] p {{font-size: .8rem; color: {cfg.GRIS};}}
</style>
""", unsafe_allow_html=True)


def seccion(txt):
    st.markdown(f'<div class="seccion">{txt}</div>', unsafe_allow_html=True)


def caja_definicion(nombre, texto):
    if texto:
        st.markdown(f'<div class="defin"><b>{nombre}:</b> {texto}</div>', unsafe_allow_html=True)


# ============================================================================
# Datos
# ============================================================================
@st.cache_data
def cargar():
    return pr.cargar()


D = cargar()
ind, ps, pq, cat, tef, tv = (D[k] for k in ("ind", "ps", "pq", "cat", "tef", "tv"))
AREAS = ["Uruguay"] + sorted(d for d in ind["departamento"].unique() if d != "Uruguay")


def serie(codigo, area, sexo="Ambos sexos"):
    return pr.serie(ind, codigo, area, sexo)


def valor(codigo, area, anio, sexo="Ambos sexos"):
    return pr.valor(ind, codigo, area, anio, sexo)


def info(codigo):
    return pr.info(cat, codigo)


def descargar(df, nombre, key):
    nombre = pr.nombre_archivo(nombre)
    b1, b2, _ = st.columns([1, 1, 3])
    b1.download_button("⬇ Excel", pr.a_excel(df), file_name=f"{nombre}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key + "_xlsx")
    b2.download_button("⬇ CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{nombre}.csv", mime="text/csv", key=key)


# ============================================================================
# Gráficos
# ============================================================================
def estilo(fig, alto=380, leyenda=True, abajo=False):
    if abajo:  # leyenda debajo del gráfico (cuando hay título arriba)
        # se ancla al borde inferior de la figura para que no tape los años del eje
        ley = dict(orientation="h", yref="container", yanchor="bottom", y=0, x=0, title=None)
        margen = dict(l=10, r=10, t=56, b=48)
    else:
        ley = dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=None)
        margen = dict(l=10, r=10, t=40, b=10)
    fig.update_layout(
        height=alto, margin=margen, template="plotly_white",
        font=dict(family=f"{cfg.FUENTE}, Arial", size=12, color=cfg.TEXTO),
        colorway=PALETA, legend=ley if leyenda else None,
        showlegend=leyenda, hoverlabel=dict(bgcolor="white"), separators=",.",
    )
    fig.update_xaxes(showgrid=False, linecolor="#B0BEC5")
    fig.update_yaxes(gridcolor="#ECEFF1", zeroline=False)
    return fig


def marcar_proyeccion(fig, x_max, texto="Proyección →"):
    if x_max >= AP:
        fig.add_vrect(x0=AP - .5, x1=x_max + .5, fillcolor=cfg.AZUL_CLARO, opacity=.5, line_width=0, layer="below")
        fig.add_annotation(x=AP, y=1, yref="paper", text=texto, showarrow=False,
                           xanchor="left", yanchor="top", font=dict(color=cfg.GRIS, size=11))
    return fig


def fig_series(df, color_col, titulo="", unidad="", alto=380, colores=None):
    """Líneas: tramo estimado continuo, tramo proyectado punteado."""
    fig = go.Figure()
    for k, g in enumerate(df[color_col].unique()):
        d = df[df[color_col] == g].sort_values("anio")
        c = (colores or {}).get(g, PALETA[k % len(PALETA)])
        est = d[d.anio < AP]
        pro = d[d.anio >= AP - 1]  # se solapa un año para que no quede hueco
        if not est.empty:
            fig.add_trace(go.Scatter(x=est.anio, y=est.valor, name=g, mode="lines", line=dict(color=c, width=2.4),
                                     legendgroup=g, hovertemplate=f"<b>{g}</b><br>%{{x}}: %{{y:,.2f}}<extra></extra>"))
        if not d[d.anio >= AP].empty:
            fig.add_trace(go.Scatter(x=pro.anio, y=pro.valor, name=g, mode="lines",
                                     line=dict(color=c, width=2.4, dash="dot"), legendgroup=g,
                                     showlegend=est.empty,
                                     hovertemplate=f"<b>{g}</b> (proy.)<br>%{{x}}: %{{y:,.2f}}<extra></extra>"))
    if not df.empty:
        marcar_proyeccion(fig, df.anio.max())
    fig.update_layout(title=dict(text=titulo, font=dict(size=14, color=cfg.AZUL_OSC), x=0, xanchor="left",
                                 y=0.98, yanchor="top"), yaxis_title=unidad)
    return estilo(fig, alto, abajo=bool(titulo))


def fig_componentes(df, alto=270):
    """Componentes del crecimiento: líneas de natalidad, mortalidad y crecimiento total, por mil.
    Sólo tasas publicadas por el INE."""
    df = df.assign(serie=df.codigo.map({**{k: v[0] for k, v in cfg.COMPONENTES.items()},
                                        cfg.LINEA_TOTAL[0]: cfg.LINEA_TOTAL[1]}))
    colores = {**{v[0]: v[1] for v in cfg.COMPONENTES.values()}, cfg.LINEA_TOTAL[1]: cfg.LINEA_TOTAL[2]}
    fig = fig_series(df, "serie", unidad="por mil", alto=alto, colores=colores)
    fig.add_hline(y=0, line_width=1, line_color="#90A4AE")
    return fig


def fig_piramide(area, anio, anio_comp=None, pct=True, simple=False):
    col = "pct" if pct else "poblacion"
    d = pr.datos_piramide(ps, pq, area, anio, simple)
    fig = go.Figure()
    for sexo, signo in (("Hombres", -1), ("Mujeres", 1)):
        s = d[d.sexo == sexo]
        fig.add_trace(go.Bar(y=s.grupo, x=signo * s[col], orientation="h", name=f"{sexo} {anio}",
                             marker_color=cfg.COLOR_SEXO[sexo], marker_line_width=0, opacity=.9,
                             customdata=s[[col]].values,
                             hovertemplate=f"{sexo} %{{y}}: %{{customdata[0]:,.2f}}" + ("%" if pct else "") + "<extra></extra>"))
    m = d[col].max()
    if anio_comp:
        dc = pr.datos_piramide(ps, pq, area, anio_comp, simple)
        for sexo, signo in (("Hombres", -1), ("Mujeres", 1)):
            s = dc[dc.sexo == sexo]
            fig.add_trace(go.Scatter(y=s.grupo, x=signo * s[col], mode="lines", name=f"{sexo} {anio_comp}",
                                     line=dict(color=cfg.TEXTO, width=1.6, shape="hvh"),
                                     customdata=s[[col]].values,
                                     hovertemplate=f"{sexo} {anio_comp} %{{y}}: %{{customdata[0]:,.2f}}<extra></extra>"))
        m = max(m, dc[col].max())
    ticks = np.linspace(-m, m, 7)
    fig.update_layout(barmode="overlay", bargap=0.06,
                      xaxis=dict(tickvals=ticks, ticktext=[fmt(abs(t), 1 if pct else 0) for t in ticks],
                                 title="% de la población" if pct else "personas"),
                      yaxis=dict(categoryorder="array", categoryarray=list(d.drop_duplicates("grupo").grupo),
                                 title=None, tickfont=dict(size=10 if simple else 11)))
    return estilo(fig, 560 if simple else 480)


def kpi(col, etiqueta, val, sub, delta=None, sufijo=""):
    flecha = ""
    if delta is not None and not np.isnan(delta):
        cls = "up" if delta >= 0 else "down"
        flecha = f'<span class="{cls}">{"▲" if delta >= 0 else "▼"} {fmt(abs(delta), 1)}</span> '
    col.markdown(f"""<div class="kpi"><div class="lbl">{etiqueta}</div>
        <div class="val">{val}{sufijo}</div><div class="sub">{flecha}{sub}</div></div>""",
                 unsafe_allow_html=True)


# ============================================================================
# Barra lateral
# ============================================================================
with st.sidebar:
    st.markdown("### 🔎 Filtros")
    area = st.selectbox("Área geográfica", AREAS, index=0, key="area")
    anio_max_area = int(ind[(ind.departamento == area) & (ind.codigo == "POB")].anio.max())
    anio_min = int(ind.anio.min())
    # si se cambia a un departamento (proyectado hasta menos años) se ajusta el año guardado
    for k in ("anio", "anio_comp"):
        if st.session_state.get(k, 0) > anio_max_area:
            st.session_state[k] = anio_max_area
    anio = st.slider("Año de referencia", anio_min, anio_max_area, min(cfg.ANIO_REF_DEFECTO, anio_max_area), key="anio")
    st.caption(f"Estimaciones {cfg.ANIO_INI_EST}–{AP - 1} · Proyecciones desde {AP} "
               f"(hasta {anio_max_area} para {'el total país' if area == 'Uruguay' else 'departamentos'}).")
    st.markdown("---")
    ver_proy = st.toggle("Mostrar proyecciones en las series", value=True, key="ver_proy")
    st.markdown("---")
    st.markdown(f'<p class="nota">Fuente: Instituto Nacional de Estadística (INE). '
                f'Estimaciones y proyecciones de población, revisión {cfg.REVISION}. '
                'Población al 30 de junio de cada año.</p>', unsafe_allow_html=True)
    st.markdown(f'<p class="nota">Información adicional por otras aperturas: solicitar a '
                f'<a href="mailto:{cfg.CONTACTO}">{cfg.CONTACTO}</a></p>', unsafe_allow_html=True)


def recorte(df):
    return df if ver_proy else df[df.anio < AP]


# ============================================================================
# Encabezado y KPIs
# ============================================================================
tipo = "proyección" if anio >= AP else "estimación"
st.markdown(f"""<div class="banda">{logo_html()}<div class="txt"><h1>Panorama demográfico · {area}</h1>
<p>Estimaciones y proyecciones de población, revisión {cfg.REVISION} · Año {anio} ({tipo})</p></div></div>""",
            unsafe_allow_html=True)

c = st.columns(6)
p, pa = valor("POB", area, anio)
p10, _ = valor("POB", area, anio - 10)
kpi(c[0], "Población total", fmt(p), f"% vs. {anio - 10}" if p10 else f"personas · {pa}",
    delta=(100 * (p / p10 - 1)) if p and p10 else None)
v, a = valor("TGF", area, anio)
kpi(c[1], "Tasa global de fecundidad", fmt(v, 2), f"hijos por mujer · {a or '–'}")
v, a = valor("E0", area, anio)
h, _ = valor("E0", area, anio, "Hombres")
m, _ = valor("E0", area, anio, "Mujeres")
kpi(c[2], "Esperanza de vida al nacer", fmt(v, 1), f"años · M {fmt(m, 1)} / H {fmt(h, 1)} · {a or '–'}")
v, a = valor("P65", area, anio)
kpi(c[3], "Población de 65 años y más", fmt(v, 1), f"% del total · {a or '–'}", sufijo="%")
v, a = valor("EMED", area, anio)
kpi(c[4], "Edad mediana de la población", fmt(v, 1), f"años · {a or '–'}")
v, a = valor("TCT", area, anio)
kpi(c[5], "Tasa de crecimiento total", fmt(v, 2), f"por mil · {a or '–'}")

st.write("")
tabs = st.tabs(["🏠 Perfil", "👥 Población", "📈 Indicadores", "🗺️ Comparar departamentos",
                "👶 Fecundidad y mortalidad", "📋 Datos", "📚 Documentación"])

# ============================================================================
# 1. Perfil
# ============================================================================
with tabs[0]:
    col1, col2 = st.columns([1, 1.25])
    with col1:
        seccion(f"Pirámide de población {anio}")
        st.plotly_chart(fig_piramide(area, anio, pct=True), **ANCHO, key="pir_perfil")
    with col2:
        seccion("Evolución de la población total")
        d = recorte(serie("POB", area)).assign(serie=area)
        fig = fig_series(d, "serie", unidad="personas", alto=230, colores={area: cfg.AZUL})
        fig.update_traces(fill="tozeroy", fillcolor="rgba(21,101,166,.12)")
        fig.update_yaxes(rangemode="normal")
        st.plotly_chart(fig, **ANCHO, key="pob_perfil")
        seccion("Componentes del crecimiento (por mil)")
        d = recorte(ind[(ind.departamento == area) & (ind.sexo == "Ambos sexos") &
                        (ind.codigo.isin(list(cfg.COMPONENTES) + [cfg.LINEA_TOTAL[0]]))])
        st.plotly_chart(fig_componentes(d), **ANCHO, key="comp_perfil")
        if area == "Uruguay":
            st.caption(cfg.NOTA_PROY_MIGRACION)

    seccion("Indicadores clave")
    cols = st.columns(3)
    for k, (cod, sx) in enumerate(cfg.INDICADORES_CLAVE):
        nom, uni, _ = info(cod)
        d = ind[(ind.codigo == cod) & (ind.departamento == area)]
        d = d[d.sexo == sx] if sx else d
        with cols[k % 3]:
            st.plotly_chart(fig_series(recorte(d), "sexo", titulo=nom, unidad=uni, alto=260, colores=cfg.COLOR_SEXO),
                            **ANCHO, key=f"mini_{cod}")

# ============================================================================
# 2. Población
# ============================================================================
with tabs[1]:
    c1, c2, c3, c4 = st.columns([1.2, 1, 1, 1])
    comparar = c1.checkbox("Superponer otro año", value=True, key="comparar")
    anio_comp = c1.slider("Año de comparación", anio_min, anio_max_area,
                          cfg.ANIO_INI_EST if anio != cfg.ANIO_INI_EST else anio_max_area,
                          disabled=not comparar, key="anio_comp")
    modo = c2.radio("Escala", ["Porcentaje", "Personas"], horizontal=True, key="escala")
    simple = c3.checkbox("Edad simple", value=False, disabled=area != "Uruguay",
                         help="Disponible sólo para el total país", key="simple")
    simple = simple and area == "Uruguay"
    col1, col2 = st.columns([1.1, 1])
    with col1:
        seccion(f"Pirámide de población · {area} {anio}{' vs ' + str(anio_comp) if comparar else ''}")
        st.plotly_chart(fig_piramide(area, anio, anio_comp if comparar else None, modo == "Porcentaje", simple),
                        **ANCHO, key="pir_pob")
        st.caption(f"Barras: año de referencia · Línea: año de comparación. Grupo abierto {cfg.EDAD_ABIERTA} y más.")
    with col2:
        seccion("Población por grandes grupos de edad")
        g = recorte(pr.grandes_grupos(ps, pq, area))
        fig = px.area(g, x="anio", y="pct", color="Grupo",
                      color_discrete_sequence=[cfg.AZUL, cfg.AZUL_OSC, cfg.NARANJA], labels={"pct": "%", "anio": ""})
        fig.update_traces(hovertemplate="%{x}: %{y:.1f}%")
        marcar_proyeccion(fig, g.anio.max())
        st.plotly_chart(estilo(fig, 300), **ANCHO, key="grandes")

        seccion(f"Población por sexo · {anio}")
        tot = {s: valor("POB", area, anio, s)[0] for s in cfg.SEXOS}
        a1, a2, a3 = st.columns(3)
        a1.metric("Total", fmt(tot["Ambos sexos"]))
        a2.metric("Hombres", fmt(tot["Hombres"]))
        a3.metric("Mujeres", fmt(tot["Mujeres"]))
        rm, _ = valor("MASC", area, anio)
        st.caption(f"Relación de masculinidad: {fmt(rm, 1)} hombres por cada 100 mujeres.")

    tabla = pr.datos_piramide(ps, pq, area, anio, simple).pivot_table(
        index=["orden", "grupo"], columns="sexo", values="poblacion").reset_index()
    tabla["Ambos sexos"] = tabla["Hombres"] + tabla["Mujeres"]
    tabla = tabla.drop(columns="orden").rename(columns={"grupo": "Edad"})
    with st.expander("Ver tabla de la pirámide"):
        st.dataframe(tabla.style.format({c: "{:,.0f}" for c in cfg.SEXOS}), **ANCHO, hide_index=True)
        descargar(tabla, f"piramide_{area}_{anio}", "dl_pir")

# ============================================================================
# 3. Indicadores
# ============================================================================
with tabs[2]:
    c1, c2, c3 = st.columns([1, 2, 2])
    tema = c1.selectbox("Tema", cfg.TEMAS, index=2, key="ind_tema")
    cod = c2.selectbox("Indicador", cat[cat.tema == tema].codigo, format_func=lambda x: info(x)[0])
    areas_sel = c3.multiselect("Áreas a comparar", AREAS,
                               default=[area] if area == "Uruguay" else ["Uruguay", area], key="ind_areas")
    nom, uni, _ = info(cod)
    d = recorte(ind[(ind.codigo == cod) & (ind.departamento.isin(areas_sel))])
    sexos = sorted(d.sexo.unique())
    por_sexo = len(sexos) > 1
    # el selector de sexo se muestra siempre (deshabilitado si el indicador no tiene apertura por sexo)
    sx = st.radio("Sexo", list(cfg.SEXOS) + ["Comparar sexos"], horizontal=True,
                  key="ind_sexo", disabled=not por_sexo,
                  help=None if por_sexo else "Este indicador no tiene apertura por sexo")
    if not por_sexo:
        sx = sexos[0] if sexos else "Ambos sexos"
    caja_definicion(nom, pr.definicion(cat, cod))
    titulo = f"{nom} · {sx} ({uni})" if por_sexo and sx in ("Hombres", "Mujeres") else f"{nom} ({uni})"
    if por_sexo and sx == "Comparar sexos":
        if len(areas_sel) > 1:
            st.info("Para comparar sexos se muestra la primera área seleccionada.")
        dd = d[d.departamento == (areas_sel[0] if areas_sel else area)]
        fig = fig_series(dd, "sexo", titulo=titulo, unidad=uni, alto=470, colores=cfg.COLOR_SEXO)
    else:
        dd = d[d.sexo == sx]
        fig = fig_series(dd, "departamento", titulo=titulo, unidad=uni, alto=470, colores={"Uruguay": cfg.AZUL_OSC})
    st.plotly_chart(fig, **ANCHO, key=f"ind_serie_{cod}")
    st.caption("Línea continua: estimación · Línea punteada: proyección. "
               f"Los departamentos se proyectan hasta {cfg.ANIO_FIN_DEPTOS} y el total país hasta {cfg.ANIO_FIN_PAIS}.")
    if cod in ("E0", "E65"):
        st.caption(f"Desde {AP} la esperanza de vida surge de las tablas de mortalidad proyectadas (sólo total país). "
                   "Para ambos sexos se combinaron las tablas de hombres y mujeres según la relación de "
                   "masculinidad al nacer.")
    if cod == "TNM" and cfg.NOTA_TNM:
        st.caption(cfg.NOTA_TNM)
    ancho = dd.pivot_table(index="anio", columns=["departamento", "sexo"], values="valor")
    ancho.columns = [f"{a} - {s}" for a, s in ancho.columns]
    with st.expander("Ver tabla"):
        st.dataframe(ancho.round(2), **ANCHO)
        descargar(ancho.reset_index(), f"{cod}", "dl_ind")

# ============================================================================
# 4. Comparar departamentos
# ============================================================================
with tabs[3]:
    c1, c2, c3 = st.columns([2, 1, 1])
    cod_c = c1.selectbox("Indicador", [c for c in cat.codigo if c != "TNM"] + ["TNM"],
                         format_func=lambda x: f"{info(x)[2]} · {info(x)[0]}", key="cmp_cod")
    disp = ind[(ind.codigo == cod_c) & (ind.departamento != "Uruguay")]
    anios_disp = sorted(disp.anio.unique())
    anio_def = min(anio, max(anios_disp))
    anio_c = c2.selectbox("Año", anios_disp[::-1],
                          index=anios_disp[::-1].index(anio_def) if anio_def in anios_disp else 0)
    sexos_c = sorted(disp.sexo.unique())
    sx_c = c3.selectbox("Sexo", sexos_c, index=sexos_c.index("Ambos sexos") if "Ambos sexos" in sexos_c else 0)
    nom, uni, _ = info(cod_c)
    caja_definicion(nom, pr.definicion(cat, cod_c))
    r = ind[(ind.codigo == cod_c) & (ind.anio == anio_c) & (ind.sexo == sx_c)].copy()
    if cod_c == "POB":
        r = r[r.departamento != "Uruguay"]
    r = r.sort_values("valor")
    r["color"] = np.where(r.departamento == "Uruguay", "Uruguay",
                          np.where(r.departamento == area, "Seleccionado", "Departamento"))
    col1, col2 = st.columns([1, 1.3])
    with col1:
        seccion(f"{nom}, {anio_c}")
        fig = px.bar(r, x="valor", y="departamento", orientation="h", color="color",
                     color_discrete_map={"Uruguay": cfg.AZUL_OSC, "Seleccionado": cfg.NARANJA, "Departamento": cfg.CELESTE},
                     text=r.valor.map(lambda v: fmt(v, 0 if cod_c == "POB" else 1)),
                     labels={"valor": uni, "departamento": ""})
        fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{y}: %{x:,.2f}<extra></extra>")
        fig.update_layout(yaxis=dict(categoryorder="array", categoryarray=list(r.departamento)))
        st.plotly_chart(estilo(fig, 620, leyenda=False), **ANCHO, key="ranking")
    with col2:
        seccion("Mapa de calor por año")
        hm = ind[(ind.codigo == cod_c) & (ind.sexo == sx_c) & (ind.departamento != "Uruguay")]
        hm = recorte(hm).pivot_table(index="departamento", columns="anio", values="valor")
        hm = hm.loc[hm[hm.columns.max()].sort_values().index] if not hm.empty else hm
        fig = px.imshow(hm, aspect="auto", color_continuous_scale="Blues", labels=dict(color=uni, x="", y=""))
        fig.update_traces(hovertemplate="%{y} %{x}: %{z:,.2f}<extra></extra>")
        if not hm.empty and hm.columns.max() >= AP:
            fig.add_vline(x=AP - .5, line_color=cfg.NARANJA, line_dash="dash")
        fig = estilo(fig, 620, leyenda=False)
        fig.update_layout(coloraxis_colorbar=dict(title=None, thickness=12))
        st.plotly_chart(fig, **ANCHO, key="heat")
    descargar(r[["departamento", "anio", "sexo", "valor"]], f"{cod_c}_{anio_c}", "dl_cmp")

# ============================================================================
# 5. Fecundidad y mortalidad
# ============================================================================
with tabs[4]:
    col1, col2 = st.columns(2)
    with col1:
        seccion("Tasas específicas de fecundidad · Uruguay")
        anios_tef = sorted(tef.anio.unique())
        sel = st.multiselect("Años", anios_tef, key="tef_anios",
                             default=[a for a in cfg.ANIOS_TEF_DEFECTO + tuple(anios_tef[-1:]) if a in anios_tef])
        t = tef[tef.anio.isin(sel)].assign(Año=lambda x: x.anio.astype(str))
        fig = px.line(t, x="grupo", y="tef", color="Año", markers=True,
                      color_discrete_sequence=px.colors.sequential.Blues[3:][::-1] if len(sel) <= 6 else PALETA,
                      labels={"tef": "nacimientos por mil mujeres", "grupo": "edad"})
        st.plotly_chart(estilo(fig, 360), **ANCHO, key="tef")
        seccion(f"Tasa global de fecundidad · {area}")
        st.caption(f"Disponible para el período de estimación ({cfg.ANIO_INI_EST}-{AP - 1}).")
        d = serie("TGF", area).assign(serie=area)
        fig = fig_series(d, "serie", unidad="hijos por mujer", alto=260, colores={area: cfg.AZUL})
        fig.add_hline(y=cfg.NIVEL_REEMPLAZO, line_dash="dash", line_color=cfg.NARANJA,
                      annotation_text=f"Nivel de reemplazo ({fmt(cfg.NIVEL_REEMPLAZO, 1)})",
                      annotation_position="top right")
        st.plotly_chart(fig, **ANCHO, key="tgf")
    with col2:
        seccion("Tablas de mortalidad proyectadas · Uruguay")
        st.caption("Ambos sexos: combinación de las tablas de hombres y mujeres según la relación de "
                   "masculinidad al nacer.")
        anios_tv = sorted(tv.anio.unique())
        m1, m2 = st.columns(2)
        funciones = {"ex": "ex · esperanza de vida", "qx": "qx · prob. de muerte",
                     "lx": "lx · sobrevivientes", "dx": "dx · defunciones"}
        funcion = m1.selectbox("Función", list(funciones), format_func=funciones.get, key="tv_fun")
        sel_tv = m2.multiselect("Años", anios_tv, key="tv_anios",
                                default=[a for a in tuple(anios_tv[:1]) + cfg.ANIOS_TV_DEFECTO + tuple(anios_tv[-1:])
                                         if a in anios_tv])
        sx_tv = st.multiselect("Sexo", list(cfg.SEXOS), default=list(cfg.SEXOS), key="tv_sexo")
        t = tv[tv.anio.isin(sel_tv) & tv.sexo.isin(sx_tv)]
        fig = px.line(t, x="edad", y=funcion, color="sexo", line_dash="anio",
                      color_discrete_map=cfg.COLOR_SEXO, log_y=(funcion == "qx"),
                      labels={"edad": "edad", "anio": "año"})
        st.plotly_chart(estilo(fig, 360), **ANCHO, key="tv")
        seccion("Esperanza de vida al nacer")
        d = recorte(ind[(ind.codigo == "E0") & (ind.departamento == area)])
        st.plotly_chart(fig_series(d, "sexo", unidad="años", alto=260, colores=cfg.COLOR_SEXO), **ANCHO, key="e0")
        if not t.empty:
            descargar(t, "tablas_mortalidad_seleccion", "dl_tv")

# ============================================================================
# 6. Datos
# ============================================================================
with tabs[5]:
    seccion("Consulta de datos")
    c1, c2, c3, c4 = st.columns(4)
    f_tema = c1.multiselect("Tema", cfg.TEMAS, default=cfg.TEMAS, key="dat_tema")
    f_ind = c2.multiselect("Indicadores", cat[cat.tema.isin(f_tema)].indicador.tolist())
    f_area = c3.multiselect("Área", AREAS, default=[area], key="dat_area")
    f_anios = c4.slider("Años", anio_min, int(ind.anio.max()), cfg.RANGO_DATOS_DEFECTO, key="dat_anios")
    q = ind[ind.tema.isin(f_tema) & ind.departamento.isin(f_area) & ind.anio.between(*f_anios)]
    if f_ind:
        q = q[q.indicador.isin(f_ind)]
    q = q[["tema", "indicador", "unidad", "departamento", "sexo", "anio", "fuente", "valor"]]
    st.dataframe(q, **ANCHO, hide_index=True, height=480,
                 column_config={"valor": st.column_config.NumberColumn(format="%.2f"),
                                "anio": st.column_config.NumberColumn("año", format="%d")})
    st.caption(f"{len(q):,} registros".replace(",", "."))
    descargar(q, "indicadores_seleccion", "dl_datos")

# ============================================================================
# 7. Documentación
# ============================================================================
with tabs[6]:
    seccion("Documentos técnicos")
    pdfs = sorted(cfg.DIR_DOCS.glob("*.pdf")) if cfg.DIR_DOCS.exists() else []
    cols = st.columns(len(cfg.DOCUMENTOS))
    for col, (titulo_doc, desc, url, clave) in zip(cols, cfg.DOCUMENTOS):
        with col:
            st.markdown(f"""<div class="kpi" style="height:auto; min-height:150px">
                <div class="val" style="font-size:1.15rem">📄 {titulo_doc}</div>
                <div class="sub" style="font-size:.85rem; margin:6px 0 10px 0">{desc}</div>
                <a href="{url}" target="_blank">Link al documento ↗</a></div>""", unsafe_allow_html=True)
            local = [p for p in pdfs if clave in pr.sin_acentos(p.name)]
            if local:
                st.download_button("⬇ Descargar PDF", local[0].read_bytes(), file_name=local[0].name,
                                   mime="application/pdf", key=f"pdf_{clave}")

    st.markdown(f"""<div class="defin" style="font-size:.95rem; margin-top:16px">
        <b>Información adicional por otras aperturas:</b> solicitar a
        <a href="mailto:{cfg.CONTACTO}">{cfg.CONTACTO}</a> ·
        <a href="{cfg.PAGINA_INE}" target="_blank">Estimaciones y proyecciones revisiones anteriores ↗</a></div>""",
                unsafe_allow_html=True)

    seccion("Definiciones de los indicadores")
    glos = cat[["tema", "indicador", "unidad", "definicion"]].rename(
        columns={"tema": "Tema", "indicador": "Indicador", "unidad": "Unidad", "definicion": "Definición"})
    glos["Tema"] = glos["Tema"].astype("category").cat.set_categories(cfg.TEMAS, ordered=True)
    glos = glos.sort_values(["Tema", "Indicador"])
    st.dataframe(glos, **ANCHO, hide_index=True, height=560,
                 column_config={"Definición": st.column_config.TextColumn(width="large")})
    st.markdown(f'<p class="nota">Notas: {cfg.NOTA_TNM + " " if cfg.NOTA_TNM else ""}La esperanza de vida de ambos sexos '
                f'proyectada ({AP}-{cfg.ANIO_FIN_PAIS}) combina las tablas de mortalidad de hombres y mujeres '
                'según la relación de masculinidad al nacer. '
                f'El último grupo de edad se armoniza en {cfg.EDAD_ABIERTA} y más.</p>', unsafe_allow_html=True)
    descargar(glos.astype({"Tema": str}), "definiciones_indicadores", "dl_glos")

st.markdown("---")
st.markdown(f'<p class="nota">Instituto Nacional de Estadística (INE) — Estimaciones ({cfg.ANIO_INI_EST}–{AP - 1}) y '
            f'proyecciones ({AP}–{cfg.ANIO_FIN_PAIS}) de población, revisión {cfg.REVISION}.</p>',
            unsafe_allow_html=True)

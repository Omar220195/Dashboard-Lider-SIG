import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date, timedelta
from io import BytesIO
import requests
# ── Configuración ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Control Catastral · Líder SIG",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CONFIGURACIÓN GITHUB ──────────────────────────────────────────────────────
# Cambia estos dos valores con tu usuario y repositorio de GitHub
GITHUB_USER = "Omar220195"        # ← reemplaza con tu usuario de GitHub
GITHUB_REPO = "Dashboard-Lider-SIG"  # ← reemplaza con el nombre de tu repositorio
GITHUB_FILE = "lider_sig_datos.xlsx"  # ruta del Excel dentro del repo

# URL raw de GitHub para descargar el archivo directamente
GITHUB_RAW_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/{GITHUB_FILE}"

# ── Paleta de estados ─────────────────────────────────────────────────────────
ESTADOS_CFG = {
    "Pendiente Líder SIG":    {"color": "#BA7517", "bg": "#FAEEDA", "orden": 1},
    "En proceso Editor SIG":  {"color": "#185FA5", "bg": "#E6F1FB", "orden": 2},
    "Terminado Editor SIG":   {"color": "#3B6D11", "bg": "#EAF3DE", "orden": 3},
    "Aprobado Coordinador":   {"color": "#0F6E56", "bg": "#E1F5EE", "orden": 4},
    "Rechazado Coordinador":  {"color": "#A32D2D", "bg": "#FCEBEB", "orden": 5},
    "Devuelto a reconocedor": {"color": "#5F5E5A", "bg": "#F1EFE8", "orden": 6},
}
ESTADO_MAP = {
    "Asignado a Líder SIG":             "Pendiente Líder SIG",
    "Asignado a Editor SIG":            "En proceso Editor SIG",
    "Registro terminado Editor SIG":    "Terminado Editor SIG",
    "Aprobado por Coordinador SIG":     "Aprobado Coordinador",
    "Rechazado por Coordinador SIG":    "Rechazado Coordinador",
}
COLS_FECHA = {
    "F. Asig Coordinador SIG":   "Fecha_Asig_Coord",
    "F. Asig Editor SIG":        "Fecha_Asig_Editor",
    "F. Termi. Editor SIG":      "Fecha_Terminado",
    "F. Aprobación. Editor SIG": "Fecha_Aprobacion",
}

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif}
.dash-header{background:#1a1a35;color:white;padding:16px 24px;border-radius:10px;margin-bottom:20px;display:flex;align-items:center;gap:16px}
.dash-logo{font-family:'IBM Plex Mono',monospace;font-size:22px;font-weight:500;letter-spacing:2px}
.dash-subtitle{font-size:12px;color:rgba(255,255,255,0.5)}
.fuente-badge{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:4px;font-size:11px;font-weight:600;margin-bottom:4px}
.fuente-github{background:#E6F1FB;color:#0C447C}
.fuente-manual{background:#FAEEDA;color:#633806}
.metric-card{background:white;border:0.5px solid rgba(0,0,0,0.08);border-radius:8px;padding:14px 16px;border-top:3px solid #1a1a35}
.mlabel{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#7a7a74;margin-bottom:4px}
.mval{font-size:26px;font-weight:600;font-family:'IBM Plex Mono',monospace;line-height:1}
.msub{font-size:10px;color:#7a7a74;margin-top:4px}
.slabel{font-size:10px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:#7a7a74;margin:4px 0 12px;padding-bottom:6px;border-bottom:0.5px solid rgba(0,0,0,0.08)}
.badge{display:inline-block;padding:3px 8px;border-radius:3px;font-size:11px;font-weight:600}
.date-range-badge{display:inline-block;background:#E6F1FB;color:#0C447C;padding:3px 10px;border-radius:4px;font-size:11px;font-weight:600;margin-bottom:8px}
.novedad{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:0.5px solid rgba(0,0,0,0.06);font-size:13px}
.nov-id{font-family:'IBM Plex Mono',monospace;font-weight:500;min-width:60px}
[data-testid="stSidebar"]{background:#f8f7f4}
</style>
""", unsafe_allow_html=True)


# ── Funciones ─────────────────────────────────────────────────────────────────
def clean_name(s, role):
    return str(s or "").replace(f" - {role}", "").strip() or "Sin asignar"

def pct(a, b):
    return round(a / b * 100) if b else 0

def parse_date_col(series):
    return (pd.to_datetime(series, utc=True, errors="coerce")
              .dt.tz_convert("America/Bogota")
              .dt.normalize()
              .dt.tz_localize(None))

def badge_html(estado):
    cfg = ESTADOS_CFG.get(estado, {"color": "#888", "bg": "#eee"})
    return f'<span class="badge" style="background:{cfg["bg"]};color:{cfg["color"]}">{estado}</span>'

@st.cache_data(ttl=300)  # cachea 5 min, luego relee GitHub
def load_from_github():
    """Descarga el Excel desde GitHub y retorna los bytes."""
    try:
        resp = requests.get(GITHUB_RAW_URL, timeout=15)
        resp.raise_for_status()
        return resp.content, None
    except requests.exceptions.HTTPError as e:
        if resp.status_code == 404:
            return None, "Archivo no encontrado en GitHub. Verifica que hayas subido el Excel a la ruta correcta."
        return None, f"Error HTTP {resp.status_code}: {e}"
    except Exception as e:
        return None, f"No se pudo conectar a GitHub: {e}"

@st.cache_data
def process_df(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    df["Coordinador"] = df["Coordinador SIG"].apply(lambda x: clean_name(x, "Coordinador SIG"))
    df["Editor"]      = df["Editor SIG"].apply(lambda x: clean_name(x, "Editor SIG"))
    df["Estado"]      = df["Estado espacio"].map(ESTADO_MAP).fillna(df["Estado espacio"])
    df["Orden_Estado"]= df["Estado"].map({e: v["orden"] for e, v in ESTADOS_CFG.items()}).fillna(9)
    df["#Predios iniciales"] = pd.to_numeric(df["#Predios iniciales"], errors="coerce").fillna(0).astype(int)
    df["#Capturas"]          = pd.to_numeric(df["#Capturas"],          errors="coerce").fillna(0).astype(int)
    for orig, dest in COLS_FECHA.items():
        if orig in df.columns:
            df[dest] = parse_date_col(df[orig])
    for col in ["Devolver_Reconocedor","Observaciones_Devolucion"]:
        if col not in df.columns:
            df[col] = ""
    return df


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🗺️ Líder SIG")
    st.markdown("**Control Catastral V2**")
    st.markdown("---")

    st.markdown("**Fuente de datos**")
    fuente = st.radio(
        "fuente",
        ["📡 GitHub (automático)"],
        label_visibility="collapsed",
    )

    uploaded = None
    if fuente == "📂 Subir Excel manualmente":
        uploaded = st.file_uploader(
            "Selecciona el Excel descargado del aplicativo",
            type=["xlsx", "xls"],
        )
        if uploaded:
            st.success(f"✓ {uploaded.name}")
    else:
        st.markdown(
            f"<small style='color:#7a7a74'>Leyendo desde:<br>"
            f"<code>{GITHUB_USER}/{GITHUB_REPO}</code><br>"
            f"<code>{GITHUB_FILE}</code></small>",
            unsafe_allow_html=True,
        )
        if st.button("🔄 Refrescar datos"):
            st.cache_data.clear()
            st.rerun()

    st.markdown("---")
    st.markdown(
        "<small style='color:#7a7a74'>GEO Proyecciones · Catastral V2<br><br>"
        "<b>Para actualizar:</b> sube el nuevo Excel<br>a GitHub en la carpeta <code>datos/</code></small>",
        unsafe_allow_html=True,
    )


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dash-header">
  <div class="dash-logo">GEO</div>
  <div>
    <div style="font-size:15px;font-weight:600">Control Catastral · Líder SIG</div>
    <div class="dash-subtitle">Módulo Actualización Catastral V2 · Soledad, Atlántico</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
file_bytes = None
origen_label = ""

if fuente == "📂 Subir Excel manualmente":
    if not uploaded:
        st.info("👈 Selecciona el Excel del aplicativo en el panel izquierdo.")
        st.stop()
    file_bytes = uploaded.read()
    origen_label = f'<span class="fuente-badge fuente-manual">📂 Carga manual · {uploaded.name}</span>'
else:
    with st.spinner("Cargando datos desde GitHub..."):
        file_bytes, error = load_from_github()
    if error:
        st.error(f"**No se pudo cargar el archivo desde GitHub**\n\n{error}")
        st.markdown("""
        **Solución:** Sube el Excel del aplicativo a tu repositorio en la ruta:
        ```
        datos/lider_sig_datos.xlsx
        ```
        O cambia a modo manual en el panel izquierdo para cargar el archivo directamente.
        """)
        st.stop()
    origen_label = f'<span class="fuente-badge fuente-github">📡 GitHub · {GITHUB_USER}/{GITHUB_REPO}</span>'

st.markdown(origen_label, unsafe_allow_html=True)
df = process_df(file_bytes)


# ── RANGO DE FECHAS ───────────────────────────────────────────────────────────
fecha_min = min(
    df["Fecha_Asig_Coord"].min() if "Fecha_Asig_Coord" in df else pd.NaT,
    df["Fecha_Asig_Editor"].min() if "Fecha_Asig_Editor" in df else pd.NaT,
)
fecha_max = max(
    df["Fecha_Asig_Editor"].max() if "Fecha_Asig_Editor" in df else pd.NaT,
    df["Fecha_Terminado"].max() if "Fecha_Terminado" in df and df["Fecha_Terminado"].notna().any() else pd.NaT,
)
fecha_min = fecha_min.date() if pd.notna(fecha_min) else date.today() - timedelta(days=30)
fecha_max = fecha_max.date() if pd.notna(fecha_max) else date.today()


# ── FILTROS ───────────────────────────────────────────────────────────────────
st.markdown('<div class="slabel">Filtros</div>', unsafe_allow_html=True)
tab_gral, tab_fecha = st.tabs(["📋 General", "📅 Por fecha"])

with tab_gral:
    c1, c2, c3 = st.columns(3)
    with c1: f_coord  = st.selectbox("Coordinador", ["Todos"] + sorted(df["Coordinador"].unique()))
    with c2: f_editor = st.selectbox("Editor SIG",  ["Todos"] + sorted(df["Editor"].unique()))
    with c3: f_estado = st.selectbox("Estado",      ["Todos"] + [e for e in ESTADOS_CFG if e in df["Estado"].unique()])

with tab_fecha:
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Fecha asignación coordinador**")
        use_fac = st.checkbox("Activar", key="use_fac")
        fac_ini = st.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max, key="fac_ini", disabled=not use_fac)
        fac_fin = st.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max, key="fac_fin", disabled=not use_fac)
    with c2:
        st.markdown("**Fecha asignación editor**")
        use_fae = st.checkbox("Activar", key="use_fae")
        fae_ini = st.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max, key="fae_ini", disabled=not use_fae)
        fae_fin = st.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max, key="fae_fin", disabled=not use_fae)
    with c3:
        st.markdown("**Fecha terminado editor**")
        use_ft = st.checkbox("Activar", key="use_ft")
        ft_ini = st.date_input("Desde", value=fecha_min, min_value=fecha_min, max_value=fecha_max, key="ft_ini", disabled=not use_ft)
        ft_fin = st.date_input("Hasta", value=fecha_max, min_value=fecha_min, max_value=fecha_max, key="ft_fin", disabled=not use_ft)

# Aplicar filtros
dff = df.copy()
if f_coord  != "Todos": dff = dff[dff["Coordinador"] == f_coord]
if f_editor != "Todos": dff = dff[dff["Editor"]      == f_editor]
if f_estado != "Todos": dff = dff[dff["Estado"]      == f_estado]
if use_fac:
    dff = dff[dff["Fecha_Asig_Coord"].notna() &
              (dff["Fecha_Asig_Coord"].dt.date >= fac_ini) &
              (dff["Fecha_Asig_Coord"].dt.date <= fac_fin)]
if use_fae:
    dff = dff[dff["Fecha_Asig_Editor"].notna() &
              (dff["Fecha_Asig_Editor"].dt.date >= fae_ini) &
              (dff["Fecha_Asig_Editor"].dt.date <= fae_fin)]
if use_ft:
    dff = dff[dff["Fecha_Terminado"].notna() &
              (dff["Fecha_Terminado"].dt.date >= ft_ini) &
              (dff["Fecha_Terminado"].dt.date <= ft_fin)]

filtros_activos = []
if use_fac: filtros_activos.append(f"Asig. coord: {fac_ini.strftime('%d/%m')} – {fac_fin.strftime('%d/%m')}")
if use_fae: filtros_activos.append(f"Asig. editor: {fae_ini.strftime('%d/%m')} – {fae_fin.strftime('%d/%m')}")
if use_ft:  filtros_activos.append(f"Terminado: {ft_ini.strftime('%d/%m')} – {ft_fin.strftime('%d/%m')}")
if filtros_activos:
    badges = " &nbsp;·&nbsp; ".join(f'<span class="date-range-badge">{f}</span>' for f in filtros_activos)
    st.markdown(f"Filtrando por: {badges}", unsafe_allow_html=True)

if dff.empty:
    st.warning("Sin datos para los filtros seleccionados.")
    st.stop()


# ── MÉTRICAS ──────────────────────────────────────────────────────────────────
st.markdown('<div class="slabel">Resumen general</div>', unsafe_allow_html=True)

total    = len(dff)
apro     = (dff["Estado"] == "Registro aprobado Editor SIG").sum()
term     = (dff["Estado"] == "Terminado Editor SIG").sum()
proc     = (dff["Estado"] == "En proceso Editor SIG").sum()
pend     = (dff["Estado"] == "Pendiente Líder SIG").sum()
rech     = (dff["Estado"] == "Rechazado a Editor SIG").sum()
dev      = (dff["Estado"] == "Devuelto a reconocedor").sum()
capturas = int(dff["#Capturas"].sum())
predios  = int(dff["#Predios iniciales"].sum())

cols = st.columns(8)
metricas = [
    ("Total espacios",    total,                              "asignados al Líder",  "#1a1a35"),
    ("Aprobados",         apro,                              f"{pct(apro,total)}% del total","#0F6E56"),
    ("Terminados editor", term,                               "pend. aprobación",    "#3B6D11"),
    ("En proceso",        proc,                               "con editor SIG",      "#185FA5"),
    ("Pendientes líder",  pend,                               "sin coordinador",     "#BA7517"),
    ("Rechazados",        rech,                               "vuelven al editor",   "#A32D2D"),
    ("Dev. reconocedor",  dev,                                "gestión manual",      "#5F5E5A"),
    ("Capturas",          f"{capturas:,}".replace(",","."),  f"{predios} predios",   "#1a1a35"),
]
for col, (label, val, sub, color) in zip(cols, metricas):
    col.markdown(f"""
    <div class="metric-card" style="border-top-color:{color}">
        <div class="mlabel">{label}</div>
        <div class="mval" style="color:{color}">{val}</div>
        <div class="msub">{sub}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ── AVANCE SEMANAL ────────────────────────────────────────────────────────────
st.markdown('<div class="slabel">📅 Avance semanal</div>', unsafe_allow_html=True)

def build_weekly(df_in):
    rows = []
    for col_fecha, etiqueta in [
        ("Fecha_Asig_Coord",   "Asignados a Coordinador"),
        ("Fecha_Asig_Editor",  "Asignados a Editor"),
        ("Fecha_Terminado",    "Terminados por Editor"),
        ("Fecha_Aprobacion",   "Aprobados por Coordinador"),
    ]:
        if col_fecha not in df_in.columns: continue
        sub = df_in[df_in[col_fecha].notna()].copy()
        if sub.empty: continue
        sub["Semana_ord"] = sub[col_fecha].apply(
            lambda d: (d - timedelta(days=d.weekday())) if pd.notna(d) else pd.NaT
        )
        sub["Semana"] = sub["Semana_ord"].apply(
            lambda d: d.strftime("Sem %d/%m/%y") if pd.notna(d) else None
        )
        grp = sub.groupby(["Semana","Semana_ord"]).agg(
            Espacios=("Nombre","count"),
            Capturas=("#Capturas","sum"),
        ).reset_index()
        grp["Evento"] = etiqueta
        rows.append(grp)
    if not rows: return pd.DataFrame()
    return pd.concat(rows).sort_values("Semana_ord")

weekly = build_weekly(dff)

if not weekly.empty:
    c1, c2 = st.columns(2)
    with c1:
        fig_w = px.bar(
            weekly, x="Semana", y="Espacios", color="Evento", barmode="group",
            title="Espacios por semana y evento",
            color_discrete_map={
                "Asignados a Coordinador":   "#BA7517",
                "Asignados a Editor":        "#185FA5",
                "Terminados por Editor":     "#3B6D11",
                "Aprobados por Coordinador": "#0F6E56",
            },
            category_orders={"Semana": weekly.sort_values("Semana_ord")["Semana"].unique().tolist()},
        )
        fig_w.update_layout(height=300, title_font_size=13,
                            margin=dict(t=40,b=10,l=10,r=10),
                            legend=dict(font_size=10, orientation="h", y=-0.25),
                            xaxis_title="", yaxis_title="Espacios")
        st.plotly_chart(fig_w, use_container_width=True)
    with c2:
        term_daily = (
            dff[dff["Fecha_Terminado"].notna()]
            .groupby("Fecha_Terminado").size()
            .reset_index(name="Terminados_dia")
            .sort_values("Fecha_Terminado")
        )
        if not term_daily.empty:
            term_daily["Acumulado"] = term_daily["Terminados_dia"].cumsum()
            fig_ac = px.area(
                term_daily, x="Fecha_Terminado", y="Acumulado",
                title="Acumulado de terminados por editor",
                color_discrete_sequence=["#3B6D11"],
            )
            fig_ac.update_traces(fill="tozeroy", fillcolor="rgba(59,109,17,0.15)")
            fig_ac.update_layout(height=300, title_font_size=13,
                                 margin=dict(t=40,b=10,l=10,r=10),
                                 xaxis_title="", yaxis_title="Espacios acumulados")
            st.plotly_chart(fig_ac, use_container_width=True)

    with st.expander("📊 Ver tabla de avance semanal"):
        tbl_w = weekly.pivot_table(
            index="Semana", columns="Evento", values="Espacios",
            fill_value=0, aggfunc="sum"
        ).reset_index().sort_values("Semana")
        st.dataframe(tbl_w, use_container_width=True, hide_index=True)
else:
    st.info("Sin datos de fechas para mostrar avance semanal.")

st.markdown("<br>", unsafe_allow_html=True)


# ── DISTRIBUCIÓN POR ESTADO ───────────────────────────────────────────────────
st.markdown('<div class="slabel">Distribución por estado</div>', unsafe_allow_html=True)

estado_agg = (
    dff.groupby("Estado")
    .agg(Espacios=("Estado","count"), Capturas=("#Capturas","sum"))
    .reset_index()
)
estado_agg["Orden"] = estado_agg["Estado"].map({e: v["orden"] for e, v in ESTADOS_CFG.items()}).fillna(9)
estado_agg = estado_agg.sort_values("Orden")
colores_dona = [ESTADOS_CFG.get(e,{}).get("color","#888") for e in estado_agg["Estado"]]

c1, c2 = st.columns(2)
for col, metric, title in [(c1,"Espacios","Espacios por etapa"),(c2,"Capturas","Capturas por etapa")]:
    with col:
        fig = px.pie(estado_agg, names="Estado", values=metric, hole=0.55,
                     color="Estado", color_discrete_sequence=colores_dona, title=title)
        fig.update_traces(textinfo="percent+label", textfont_size=11)
        fig.update_layout(showlegend=False, title_font_size=13,
                          margin=dict(t=40,b=10,l=10,r=10), height=280)
        st.plotly_chart(fig, use_container_width=True)


# ── POR COORDINADOR ───────────────────────────────────────────────────────────
st.markdown('<div class="slabel">Por coordinador SIG</div>', unsafe_allow_html=True)

coord_est = (
    dff[dff["Coordinador"] != "Sin asignar"]
    .groupby(["Coordinador","Estado"])
    .agg(Espacios=("Estado","count"), Capturas=("#Capturas","sum"))
    .reset_index()
)
coord_tot = (
    dff[dff["Coordinador"] != "Sin asignar"]
    .groupby("Coordinador")
    .agg(Total=("Estado","count"), Capturas=("#Capturas","sum"), Predios=("#Predios iniciales","sum"))
    .reset_index()
    .sort_values("Total", ascending=False)
)

c1, c2 = st.columns(2)
with c1:
    fig = px.bar(coord_est, x="Espacios", y="Coordinador", color="Estado", orientation="h",
                 color_discrete_map={e: v["color"] for e,v in ESTADOS_CFG.items()},
                 title="Espacios por estado y coordinador",
                 category_orders={"Estado": list(ESTADOS_CFG.keys())})
    fig.update_layout(height=220, title_font_size=13, margin=dict(t=40,b=10,l=10,r=10),
                      legend=dict(font_size=10, orientation="h", y=-0.3),
                      yaxis_title="", xaxis_title="")
    st.plotly_chart(fig, use_container_width=True)
with c2:
    fig2 = px.bar(coord_tot, x="Coordinador", y="Capturas",
                  title="Capturas por coordinador",
                  color_discrete_sequence=["#185FA5"])
    fig2.update_layout(height=220, title_font_size=13, margin=dict(t=40,b=10,l=10,r=10),
                       showlegend=False, xaxis_title="", yaxis_title="")
    st.plotly_chart(fig2, use_container_width=True)

coord_pivot = dff[dff["Coordinador"] != "Sin asignar"].pivot_table(
    index="Coordinador", columns="Estado", values="Nombre", aggfunc="count", fill_value=0
).reset_index()
coord_pivot = coord_pivot.merge(coord_tot, on="Coordinador")
coord_pivot["Aprobados"]  = coord_pivot.get("Aprobado Coordinador", pd.Series(0, index=coord_pivot.index))
coord_pivot["Terminados"] = coord_pivot.get("Terminado Editor SIG", pd.Series(0, index=coord_pivot.index))
coord_pivot["% Avance"]   = coord_pivot.apply(lambda r: f"{pct(r['Aprobados']+r['Terminados'],r['Total'])}%", axis=1)
disp = [c for c in ["Coordinador","Total","Aprobados","Terminados","% Avance","Capturas","Predios"] if c in coord_pivot.columns]
st.dataframe(coord_pivot[disp].sort_values("Total", ascending=False),
             use_container_width=True, hide_index=True)


# ── POR EDITOR ────────────────────────────────────────────────────────────────
st.markdown('<div class="slabel">Por editor SIG</div>', unsafe_allow_html=True)

ed_est = (
    dff[dff["Editor"] != "Sin asignar"]
    .groupby(["Editor","Estado"])
    .agg(Espacios=("Estado","count"))
    .reset_index()
)
ed_tot = (
    dff[dff["Editor"] != "Sin asignar"]
    .groupby("Editor")
    .agg(Total=("Estado","count"), Capturas=("#Capturas","sum"), Predios=("#Predios iniciales","sum"))
    .reset_index()
    .sort_values("Total", ascending=False)
)

fig3 = px.bar(
    ed_est, x="Espacios", y="Editor", color="Estado", orientation="h",
    color_discrete_map={e: v["color"] for e,v in ESTADOS_CFG.items()},
    title="Espacios por editor y estado",
    category_orders={"Estado": list(ESTADOS_CFG.keys()), "Editor": ed_tot["Editor"].tolist()},
)
fig3.update_layout(
    height=max(300, len(ed_tot)*34+80), title_font_size=13,
    margin=dict(t=40,b=10,l=10,r=10),
    legend=dict(font_size=10, orientation="h", y=-0.1),
    yaxis_title="", xaxis_title="",
)
st.plotly_chart(fig3, use_container_width=True)

ed_pivot = dff[dff["Editor"] != "Sin asignar"].pivot_table(
    index="Editor", columns="Estado", values="Nombre", aggfunc="count", fill_value=0
).reset_index()
ed_pivot = ed_pivot.merge(ed_tot, on="Editor")
ed_pivot = ed_pivot.merge(
    dff[dff["Editor"] != "Sin asignar"].groupby("Editor")["Coordinador"].first().reset_index(),
    on="Editor"
)
ed_pivot["Aprobados"]  = ed_pivot.get("Registro aprobado Editor SIG", pd.Series(0, index=ed_pivot.index))
ed_pivot["Terminados"] = ed_pivot.get("Terminado Editor SIG", pd.Series(0, index=ed_pivot.index))
ed_pivot["En proceso"] = ed_pivot.get("En proceso Editor SIG", pd.Series(0, index=ed_pivot.index))
ed_pivot["Rechazados"] = ed_pivot.get("Rechazado Coordinador", pd.Series(0, index=ed_pivot.index))
ed_pivot["% Avance"]   = ed_pivot.apply(lambda r: f"{pct(r['Aprobados']+r['Terminados'],r['Total'])}%", axis=1)
disp_e = [c for c in ["Editor","Coordinador","Total","En proceso","Terminados","Aprobados","Rechazados","% Avance","Capturas"] if c in ed_pivot.columns]
st.dataframe(ed_pivot[disp_e].sort_values("Total", ascending=False),
             use_container_width=True, hide_index=True)


# ── NOVEDADES ─────────────────────────────────────────────────────────────────
novs = dff[dff["Estado"].isin(["Rechazado Coordinador","Devuelto a reconocedor"])]
if not novs.empty:
    st.markdown('<div class="slabel">⚠️ Novedades · rechazados y devueltos</div>', unsafe_allow_html=True)
    html = ""
    for _, r in novs.iterrows():
        cfg = ESTADOS_CFG.get(r["Estado"], {"color":"#888","bg":"#eee"})
        html += f"""<div class="novedad">
          <span class="nov-id">{r['Nombre']}</span>
          <span class="badge" style="background:{cfg['bg']};color:{cfg['color']}">{r['Estado']}</span>
          <span style="color:#7a7a74">{r['Coordinador']} · {r['Editor']}</span>
          <span style="margin-left:auto;font-family:'IBM Plex Mono',monospace;font-size:11px;color:#7a7a74">{r['#Capturas']} cap.</span>
        </div>"""
    st.markdown(html, unsafe_allow_html=True)


# ── TABLA COMPLETA ────────────────────────────────────────────────────────────
with st.expander("📋 Ver tabla completa de espacios geográficos"):
    cols_show = ["Nombre","Coordinador","Editor","Estado","#Predios iniciales","#Capturas",
                 "Fecha_Asig_Coord","Fecha_Asig_Editor","Fecha_Terminado"]
    cols_show = [c for c in cols_show if c in dff.columns]
    rename = {
        "Nombre":"Espacio","#Predios iniciales":"Predios","#Capturas":"Capturas",
        "Fecha_Asig_Coord":"F. Asig. Coord","Fecha_Asig_Editor":"F. Asig. Editor",
        "Fecha_Terminado":"F. Terminado",
    }
    st.dataframe(
        dff[cols_show].rename(columns=rename).sort_values("Capturas", ascending=False),
        use_container_width=True, hide_index=True,
    )

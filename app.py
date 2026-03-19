import unicodedata
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

import folium
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from branca.colormap import LinearColormap
from streamlit_folium import st_folium

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Ejecutivo de Personal",
    page_icon="📊",
    layout="wide",
)

# ── Custom CSS – Light BI theme ───────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }

    :root {
        --text-primary: #1E293B;
        --text-secondary: #475569;
        --text-muted: #64748B;
        --surface: #FFFFFF;
    }

    /* Background */
    .stApp { background: #F4F6FA; }

    /* Force readable text colors in light mode */
    .stApp, .stApp p, .stApp li, .stApp label, .stApp span,
    .stMarkdown, .stMarkdown p, .stCaption, .stText {
        color: var(--text-primary);
    }
    section[data-testid="stSidebar"] *,
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] label {
        color: var(--text-secondary) !important;
    }

    /* Main block padding */
    .block-container {
        padding: 1.5rem 2rem 2rem 2rem;
        max-width: 100%;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid #E2E8F0;
        padding-top: 1rem;
    }
    section[data-testid="stSidebar"] .sidebar-label {
        font-size: 0.68rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #64748B !important;
        margin: 0.75rem 0 0.4rem 0;
    }
    div[data-baseweb="select"] > div {
        background-color: #F8FAFC !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        color: #1E293B !important;
    }
    div[data-baseweb="select"] input {
        color: #1E293B !important;
    }

    /* Metric cards */
    div[data-testid="metric-container"] {
        background: var(--surface);
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    div[data-testid="metric-container"] label {
        font-size: 0.72rem !important;
        font-weight: 600 !important;
        letter-spacing: 0.06em !important;
        text-transform: uppercase !important;
        color: #94A3B8 !important;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #1E293B !important;
        font-family: 'DM Mono', monospace !important;
    }

    /* Section headers */
    .section-header {
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #64748B;
        margin: 1.5rem 0 0.75rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #E2E8F0;
    }

    /* Chart cards */
    .chart-card {
        background: var(--surface);
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
    div[data-testid="stPlotlyChart"] {
        background: var(--surface);
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 0.75rem 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }

    /* Dashboard title */
    .dashboard-title {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.1rem;
        line-height: 1.2;
    }
    .dashboard-subtitle {
        font-size: 0.85rem;
        color: #64748B;
        margin-bottom: 1.25rem;
    }

    /* Filter badges */
    .filter-badge {
        display: inline-block;
        background: #EEF2FF;
        color: #4F46E5;
        border-radius: 20px;
        padding: 0.15rem 0.75rem;
        font-size: 0.75rem;
        font-weight: 600;
        margin: 0.2rem;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.45rem;
        border-bottom: none;
        background: transparent;
        padding: 0.15rem 0 0.25rem 0;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 0.82rem;
        font-weight: 600;
        letter-spacing: 0.02em;
        padding: 0.45rem 1.15rem;
        color: #1E293B !important;
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 12px;
        transition: all 0.15s ease;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05);
    }
    .stTabs [data-baseweb="tab"]:hover {
        background: #F8FAFC !important;
        border-color: #94A3B8 !important;
    }
    .stTabs [aria-selected="true"] {
        color: #14532D !important;
        border: 1px solid #22C55E !important;
        background: linear-gradient(
            135deg,
            rgba(187, 247, 208, 0.95) 0%,
            rgba(220, 252, 231, 0.9) 100%
        ) !important;
        box-shadow:
            0 0 0 1px rgba(34, 197, 94, 0.25),
            0 6px 18px rgba(34, 197, 94, 0.28),
            0 0 18px rgba(134, 239, 172, 0.45) !important;
    }

    /* Buttons and segmented control - better contrast + feedback */
    .stButton > button {
        background: linear-gradient(180deg, #4F46E5 0%, #4338CA 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #3730A3 !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        min-height: 2.5rem !important;
        transition: all 0.18s ease !important;
        box-shadow: 0 1px 4px rgba(67, 56, 202, 0.25) !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(67, 56, 202, 0.28) !important;
    }
    .stButton > button:active {
        transform: translateY(0);
        filter: brightness(0.97);
    }
    .stButton > button:focus-visible {
        outline: 3px solid #A5B4FC !important;
        outline-offset: 1px !important;
    }
    [data-testid="stSegmentedControl"] {
        margin-bottom: 0.5rem !important;
    }
    [data-testid="stSegmentedControl"] [data-baseweb="button-group"],
    [data-testid="stSegmentedControl"] [role="radiogroup"] {
        gap: 0.5rem !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        padding: 0 !important;
    }
    [data-testid="stSegmentedControl"] [data-baseweb="button-group"] > button,
    [data-testid="stSegmentedControl"] [role="radio"] {
        border: 1px solid #CBD5E1 !important;
        color: #1E293B !important;
        background: #FFFFFF !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        transition: all 0.15s ease !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06) !important;
    }
    [data-testid="stSegmentedControl"] [data-baseweb="button-group"] > button:hover,
    [data-testid="stSegmentedControl"] [role="radio"]:hover {
        background: #F8FAFC !important;
        border-color: #94A3B8 !important;
    }
    [data-testid="stSegmentedControl"] [data-baseweb="button-group"] > button[aria-pressed="true"],
    [data-testid="stSegmentedControl"] [role="radio"][aria-checked="true"] {
        background: #EEF2FF !important;
        border-color: #4F46E5 !important;
        color: #3730A3 !important;
        box-shadow: 0 0 0 1px rgba(79, 70, 229, 0.2), 0 3px 8px rgba(79, 70, 229, 0.18) !important;
    }
    [data-testid="stSegmentedControl"] [data-baseweb="button-group"] > button:focus-visible,
    [data-testid="stSegmentedControl"] [role="radio"]:focus-visible {
        outline: 3px solid #A5B4FC !important;
        outline-offset: 1px !important;
    }

    /* Multiselect tags */
    span[data-baseweb="tag"] {
        background: #EEF2FF !important;
        color: #4F46E5 !important;
        border-radius: 4px !important;
    }

    /* Hide streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Constants ─────────────────────────────────────────────────────────────────
DATA_FILE = Path("datos_git.xlsx")
COORDS_CACHE = Path("province_coords.csv")

PROVINCE_CORRECTIONS = {
    "CUZCO": "CUSCO",
    "SAN ROMÁN": "SAN ROMAN",
    "HUÁNUCO": "HUANUCO",
    "TARAPOTO": "SAN MARTIN",
    "PUCALLPA": "CORONEL PORTILLO",
    "LAREDO": "TRUJILLO",
    "CHIMBOTE": "SANTA",
    "IQUITOS": "MAYNAS",
    "JULIACA": "SAN ROMAN",
    "TAYABAMBA": "PATAZ",
    "PAUCARÁ": "ACOBAMBA",
}

DISTRICT_CORRECTIONS = {
    "CALLLAO": "CALLAO",
    "ATE VITARTE": "ATE",
    "CERCADO DE LIMA": "LIMA",
    "CENTRO DE LIMA": "LIMA",
    "RUPA-RUPA": "RUPA RUPA",
    "PICHANAQUI": "PICHANAKI",
    "SJL": "SAN JUAN DE LURIGANCHO",
    "SJM": "SAN JUAN DE MIRAFLORES",
    "SURCO": "SANTIAGO DE SURCO",
    "TAMBO GRANDE": "TAMBOGRANDE",
}

# Embedded coordinates – no HTTP calls needed for these
PERU_COORDS = {
    "LIMA": (-12.0464, -77.0428),
    "CALLAO": (-12.0595, -77.1181),
    "AREQUIPA": (-16.409, -71.5375),
    "TRUJILLO": (-8.111, -79.0288),
    "PIURA": (-5.1945, -80.6328),
    "CUSCO": (-13.5319, -71.9675),
    "HUANCAYO": (-12.0651, -75.2049),
    "CHICLAYO": (-6.7714, -79.8409),
    "ICA": (-14.0678, -75.7286),
    "TACNA": (-18.0066, -70.2463),
    "MAYNAS": (-3.7491, -73.2538),
    "PUNO": (-15.8402, -70.0219),
    "SAN ROMAN": (-15.4997, -70.1327),
    "HUANUCO": (-9.9306, -76.2422),
    "CAJAMARCA": (-7.1639, -78.5003),
    "AYACUCHO": (-13.1588, -74.2236),
    "JUNIN": (-11.1582, -75.9928),
    "SANTA": (-9.0755, -78.5943),
    "LAMBAYEQUE": (-6.7027, -79.9064),
    "SULLANA": (-4.8996, -80.6883),
    "HUARAZ": (-9.5276, -77.5278),
    "MOQUEGUA": (-17.1931, -70.9320),
    "PASCO": (-10.6832, -76.2633),
    "SAN MARTIN": (-6.5000, -76.3667),
    "CORONEL PORTILLO": (-8.3791, -74.5539),
    "MADRE DE DIOS": (-12.5933, -69.1891),
    "TUMBES": (-3.5669, -80.4515),
    "AMAZONAS": (-6.2313, -77.8692),
    "APURIMAC": (-13.6345, -72.8814),
    "HUANCAVELICA": (-12.7870, -74.9768),
    "PATAZ": (-8.0000, -77.0000),
    "ACOBAMBA": (-12.8494, -74.5722),
    "ANDAHUAYLAS": (-13.6560, -73.3830),
    "ABANCAY": (-13.6345, -72.8814),
    "CAÑETE": (-13.0797, -76.3697),
    "CANETE": (-13.0797, -76.3697),
    "HUARAL": (-11.4954, -77.2069),
    "HUAROCHIRI": (-11.9917, -76.2225),
    "BARRANCA": (-10.7500, -77.7667),
    "CHINCHA": (-13.4125, -76.1386),
    "PISCO": (-13.7141, -76.2033),
    "NAZCA": (-14.8290, -74.9433),
    "CHEPEN": (-7.2269, -79.4321),
    "PACASMAYO": (-7.4011, -79.5703),
    "ASCOPE": (-7.7202, -79.1388),
    "VIRU": (-8.4122, -78.7533),
    "OTUZCO": (-7.8999, -78.5679),
    "CHOTA": (-6.5587, -78.6521),
    "CUTERVO": (-6.3746, -78.8180),
    "JAEN": (-5.7072, -78.8070),
    "MOYOBAMBA": (-6.0340, -76.9720),
    "CHACHAPOYAS": (-6.2313, -77.8692),
    "UTCUBAMBA": (-5.8339, -78.1778),
    "RECUAY": (-9.7225, -77.4563),
}


try:
    STREAMLIT_VERSION = version("streamlit")
except PackageNotFoundError:
    STREAMLIT_VERSION = "No detectada"


def normalize_text(value) -> str:
    if pd.isna(value):
        return "S/I"
    raw = str(value).strip().upper()
    if not raw:
        return "S/I"
    normalized = "".join(
        c for c in unicodedata.normalize("NFD", raw) if unicodedata.category(c) != "Mn"
    )
    return " ".join(normalized.split())


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="WIDE")
    df.columns = [normalize_text(col) for col in df.columns]
    df["PROVINCIA"] = df["PROVINCIA"].map(normalize_text)
    df["DISTRITO"] = df["DISTRITO"].map(normalize_text)
    df["PROVINCIA"] = df["PROVINCIA"].replace(PROVINCE_CORRECTIONS)
    df["DISTRITO"] = df["DISTRITO"].replace(DISTRICT_CORRECTIONS)
    for col in ["FECHA DE INGRESO", "FECHA DE CESE"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["FECHA DE CESE NORM"] = df["FECHA DE CESE"].dt.normalize()
    df["DNI"] = pd.to_numeric(df["DNI"], errors="coerce").astype("Int64")
    cat_cols = ["CLIENTE", "UNIDAD", "CARGO", "PLANILLA", "REGIMEN PLANILLA"]
    for col in cat_cols:
        if col in df.columns:
            df[col] = df[col].map(normalize_text).astype("category")
    return df


@st.cache_data(show_spinner=False)
def get_coordinates(provinces: tuple) -> pd.DataFrame:
    """Devuelve coordenadas locales sin llamadas HTTP para evitar bloqueos."""
    rows = []
    cached_dict = {}
    if COORDS_CACHE.exists():
        try:
            cached = pd.read_csv(COORDS_CACHE)
            cached_dict = cached.set_index("PROVINCIA")[["lat", "lon"]].to_dict("index")
        except Exception:
            cached_dict = {}

    for prov in provinces:
        if prov in ("S/I", ""):
            continue
        if prov in PERU_COORDS:
            lat, lon = PERU_COORDS[prov]
        elif prov in cached_dict:
            lat = cached_dict[prov]["lat"]
            lon = cached_dict[prov]["lon"]
        else:
            lat, lon = np.nan, np.nan
        rows.append({"PROVINCIA": prov, "lat": lat, "lon": lon})

    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=["PROVINCIA", "lat", "lon"])


# ── Cascading sidebar filters ─────────────────────────────────────────────────
def sync_multiselect_state(key: str, valid_options: list[str]) -> None:
    current = st.session_state.get(key, [])
    st.session_state[key] = [v for v in current if v in valid_options]


def clear_filters() -> None:
    for key in ["f_cese", "f_cliente", "f_unidad", "f_cargo", "f_provincia"]:
        st.session_state[key] = []


def apply_filters(df: pd.DataFrame):
    with st.sidebar:
        st.markdown("## Filtros")

        # Estado
        st.markdown('<div class="sidebar-label">Estado</div>', unsafe_allow_html=True)
        estado = st.radio(
            "Estado",
            options=["Solo activos", "Activos y ceses"],
            horizontal=True,
            label_visibility="collapsed",
        )

        # Cese (estilo Excel: lista desplegable con selección múltiple)
        cese_dates = sorted(df["FECHA DE CESE NORM"].dropna().unique())
        cese_labels = [pd.Timestamp(d).strftime("%Y/%m/%d") for d in cese_dates]
        cese_map = dict(zip(cese_labels, cese_dates))

        cese_sel = []
        if estado == "Activos y ceses" and cese_labels:
            st.markdown('<div class="sidebar-label">Fecha de cese</div>', unsafe_allow_html=True)
            sync_multiselect_state("f_cese", cese_labels)
            cese_sel = st.multiselect(
                "Fecha de cese",
                cese_labels,
                key="f_cese",
                placeholder="Selecciona una o varias fechas",
                label_visibility="collapsed",
            )

        st.divider()

        # Build base for cascading
        base = df
        if estado == "Solo activos":
            base = base[base["FECHA DE CESE"].isna()]
        elif cese_sel:
            selected_dates = pd.to_datetime([cese_map[x] for x in cese_sel])
            mask = base["FECHA DE CESE"].isna() | base["FECHA DE CESE NORM"].isin(selected_dates)
            base = base[mask]

        # CASCADE 1 – Cliente
        st.markdown('<div class="sidebar-label">Cliente</div>', unsafe_allow_html=True)
        cliente_opts = sorted(base["CLIENTE"].dropna().unique())
        sync_multiselect_state("f_cliente", cliente_opts)
        cliente_sel = st.multiselect(
            "Cliente", cliente_opts, key="f_cliente", label_visibility="collapsed"
        )

        base_c = base[base["CLIENTE"].isin(cliente_sel)] if cliente_sel else base

        # CASCADE 2 – Unidad (depends on Cliente)
        st.markdown('<div class="sidebar-label">Unidad</div>', unsafe_allow_html=True)
        unidad_opts = sorted(base_c["UNIDAD"].dropna().unique())
        sync_multiselect_state("f_unidad", unidad_opts)
        unidad_sel = st.multiselect(
            "Unidad", unidad_opts, key="f_unidad", label_visibility="collapsed"
        )

        base_cu = base_c[base_c["UNIDAD"].isin(unidad_sel)] if unidad_sel else base_c

        # CASCADE 3 – Cargo (depends on Cliente + Unidad)
        st.markdown('<div class="sidebar-label">Cargo</div>', unsafe_allow_html=True)
        cargo_opts = sorted(base_cu["CARGO"].dropna().unique())
        sync_multiselect_state("f_cargo", cargo_opts)
        cargo_sel = st.multiselect(
            "Cargo", cargo_opts, key="f_cargo", label_visibility="collapsed"
        )

        # CASCADE 4 – Provincia (depends on all above)
        base_cuc = base_cu[base_cu["CARGO"].isin(cargo_sel)] if cargo_sel else base_cu
        st.markdown('<div class="sidebar-label">Provincia</div>', unsafe_allow_html=True)
        prov_opts = sorted(base_cuc["PROVINCIA"].dropna().unique())
        sync_multiselect_state("f_provincia", prov_opts)
        prov_sel = st.multiselect(
            "Provincia", prov_opts, key="f_provincia", label_visibility="collapsed"
        )

        st.divider()
        st.button(
            "🔄 Limpiar filtros",
            use_container_width=True,
            on_click=clear_filters,
        )

    # Apply to base
    filtered = base
    if cliente_sel:
        filtered = filtered[filtered["CLIENTE"].isin(cliente_sel)]
    if unidad_sel:
        filtered = filtered[filtered["UNIDAD"].isin(unidad_sel)]
    if cargo_sel:
        filtered = filtered[filtered["CARGO"].isin(cargo_sel)]
    if prov_sel:
        filtered = filtered[filtered["PROVINCIA"].isin(prov_sel)]

    active = {
        "estado": [estado],
        "cese": cese_sel,
        "cliente": cliente_sel,
        "unidad": unidad_sel,
        "cargo": cargo_sel,
        "provincia": prov_sel,
    }
    return filtered, active


# ── KPI cards ─────────────────────────────────────────────────────────────────
def kpi_cards(df: pd.DataFrame) -> None:
    activos = int(df["FECHA DE CESE"].isna().sum())
    cesados = int(df["FECHA DE CESE"].notna().sum())
    total = max(len(df), 1)
    dni_unicos = int(df["DNI"].dropna().nunique())
    duplicados = max(len(df) - dni_unicos, 0)
    pct_activos = (activos / total) * 100
    pct_cesados = (cesados / total) * 100
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("DNI únicos", f"{dni_unicos:,}")
    c2.metric("Registros", f"{len(df):,}")
    c3.metric("Duplicados", f"{duplicados:,}")
    c4.metric("🟢 Activos", f"{activos:,}", delta=f"{pct_activos:.1f}%")
    c5.metric("🔴 Cesados", f"{cesados:,}", delta=f"-{pct_cesados:.1f}%")
    c6.metric("Clientes", f"{df['CLIENTE'].nunique():,}")
    c7.metric("Unidades", f"{df['UNIDAD'].nunique():,}")


COLORS = ["#4F46E5", "#06B6D4", "#10B981", "#F59E0B", "#EF4444",
          "#8B5CF6", "#EC4899", "#14B8A6", "#F97316", "#6366F1"]


def chart_layout(fig, height=300):
    fig.update_layout(
        template="plotly_white",
        autosize=True,
        height=height,
        margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font_family="DM Sans",
        font_color="#111827",
        legend=dict(font=dict(color="#111827")),
        uniformtext_minsize=9,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(
        tickfont=dict(color="#111827"),
        title_font=dict(color="#111827"),
        automargin=True,
        constrain="domain",
    )
    fig.update_yaxes(
        tickfont=dict(color="#111827"),
        title_font=dict(color="#111827"),
        automargin=True,
    )
    return fig


def style_horizontal_bar(
    fig,
    max_value: float,
    y_tick_size: int = 10,
    left_margin: int = 220,
    right_margin: int = 56,
    show_x_ticks: bool = True,
):
    fig.update_traces(
        textposition="outside",
        texttemplate="%{x:,.0f}",
        cliponaxis=False,
        textfont_size=10,
        textfont_color="#1E293B",
    )
    fig.update_layout(
        margin=dict(l=left_margin, r=right_margin, t=10, b=10),
        yaxis=dict(autorange="reversed", tickfont_size=y_tick_size, title=""),
        xaxis=dict(
            title="",
            showgrid=True,
            gridcolor="#F1F5F9",
            showticklabels=show_x_ticks,
            range=[0, max(1, max_value * 1.18)],
        ),
        showlegend=False,
    )
    return fig


# ── Analysis tab ─────────────────────────────────────────────────────────────
def tab_analysis(df: pd.DataFrame) -> None:
    col1, col2, col3 = st.columns([1.1, 1, 0.95])

    with col1:
        st.markdown('<p class="section-header">Top Clientes · DNI únicos</p>', unsafe_allow_html=True)
        d = (df.groupby("CLIENTE")["DNI"].nunique()
             .sort_values(ascending=False).head(10).reset_index())
        d.columns = ["CLIENTE", "N"]
        fig = px.bar(d, x="N", y="CLIENTE", orientation="h",
                     color="N", color_continuous_scale=["#C7D2FE", "#4F46E5"], text="N")
        fig = chart_layout(fig)
        fig = style_horizontal_bar(fig, max_value=d["N"].max(), y_tick_size=11, left_margin=260)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    with col2:
        st.markdown('<p class="section-header">Distribución por Unidad</p>', unsafe_allow_html=True)
        d2 = (df.groupby("UNIDAD")["DNI"].nunique()
              .sort_values(ascending=False).head(8).reset_index())
        d2.columns = ["UNIDAD", "N"]
        fig2 = px.pie(d2, values="N", names="UNIDAD",
                      color_discrete_sequence=COLORS, hole=0.45)
        fig2.update_traces(
            textinfo="percent+value",
            textfont_size=11,
            textfont_color="#111827",
            hovertemplate="%{label}<br>DNI únicos: %{value:,}<br>%{percent}<extra></extra>",
        )
        fig2 = chart_layout(fig2)
        fig2.update_layout(legend=dict(font_size=10))
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    with col3:
        st.markdown('<p class="section-header">Top Cargos</p>', unsafe_allow_html=True)
        d3 = (df.groupby("CARGO")["DNI"].nunique()
              .sort_values(ascending=False).head(8).reset_index())
        d3.columns = ["CARGO", "N"]
        fig3 = px.bar(d3, x="N", y="CARGO", orientation="h",
                      color_discrete_sequence=["#06B6D4"], text="N")
        fig3 = chart_layout(fig3)
        fig3 = style_horizontal_bar(
            fig3,
            max_value=d3["N"].max(),
            y_tick_size=10,
            left_margin=240,
            right_margin=72,
            show_x_ticks=False,
        )
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

    # Trend row
    col4, col5 = st.columns(2)
    with col4:
        st.markdown('<p class="section-header">Ingresos vs Ceses por mes</p>', unsafe_allow_html=True)
        ing = (df["FECHA DE INGRESO"].dropna().dt.to_period("M")
               .value_counts().sort_index().rename("Ingresos").reset_index())
        ing.columns = ["Periodo", "Ingresos"]
        ing["Periodo"] = ing["Periodo"].dt.to_timestamp()
        ces = (df["FECHA DE CESE"].dropna().dt.to_period("M")
               .value_counts().sort_index().rename("Ceses").reset_index())
        ces.columns = ["Periodo", "Ceses"]
        ces["Periodo"] = ces["Periodo"].dt.to_timestamp()
        merged = pd.merge(ing, ces, on="Periodo", how="outer").sort_values("Periodo").fillna(0)
        fig4 = go.Figure()
        fig4.add_trace(go.Scatter(x=merged["Periodo"], y=merged["Ingresos"],
                                  name="Ingresos", line=dict(color="#4F46E5", width=2),
                                  fill="tozeroy", fillcolor="rgba(79,70,229,0.08)"))
        fig4.add_trace(go.Scatter(x=merged["Periodo"], y=merged["Ceses"],
                                  name="Ceses", line=dict(color="#EF4444", width=2, dash="dot")))
        fig4 = chart_layout(fig4, height=260)
        fig4.update_layout(legend=dict(orientation="h", y=1.1, font_size=11),
                            xaxis=dict(showgrid=False),
                            yaxis=dict(showgrid=True, gridcolor="#F1F5F9"))
        fig4.update_xaxes(
            rangeselector=dict(
                buttons=[
                    dict(count=1, label="1A", step="year", stepmode="backward"),
                    dict(count=3, label="3A", step="year", stepmode="backward"),
                    dict(step="all", label="Todo"),
                ]
            )
        )
        st.plotly_chart(fig4, use_container_width=True, config={"displayModeBar": False})

    with col5:
        st.markdown('<p class="section-header">Régimen de Planilla</p>', unsafe_allow_html=True)
        if "REGIMEN PLANILLA" in df.columns:
            reg = df["REGIMEN PLANILLA"].value_counts().head(8).reset_index()
            reg.columns = ["REGIMEN", "N"]
            fig5 = px.bar(reg, x="N", y="REGIMEN", orientation="h",
                          color="N", color_continuous_scale=["#BAE6FD", "#0EA5E9"], text="N")
            fig5 = chart_layout(fig5, height=260)
            fig5 = style_horizontal_bar(fig5, max_value=reg["N"].max(), y_tick_size=10, left_margin=190)
            fig5.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig5, use_container_width=True, config={"displayModeBar": False})


# ── Geography tab ─────────────────────────────────────────────────────────────
def tab_geography(df: pd.DataFrame, coords: pd.DataFrame) -> None:
    col1 = st.container()

    grouped = (
        df.groupby("PROVINCIA", as_index=False)
        .agg(DNI_UNICOS=("DNI", "nunique"), Registros=("DNI", "size"))
        .merge(coords, on="PROVINCIA", how="left")
        .dropna(subset=["lat", "lon"])
    )

    with col1:
        st.markdown('<p class="section-header">Mapa geográfico – Perú</p>', unsafe_allow_html=True)
        if grouped.empty:
            st.info("Sin coordenadas disponibles para las provincias actuales.")
        else:
            max_v = float(grouped["DNI_UNICOS"].max())
            color_metric = np.log1p(grouped["DNI_UNICOS"])
            cmin = float(color_metric.min())
            cmax = float(color_metric.max())
            if cmax == cmin:
                cmax = cmin + 1.0

            folium_map = folium.Map(
                location=[-9.5, -75.0],
                zoom_start=5,
                tiles="CartoDB positron",
                control_scale=True,
            )

            colormap = LinearColormap(
                colors=["#16A34A", "#FACC15", "#DC2626"],
                vmin=cmin,
                vmax=cmax,
            )
            colormap.caption = "DNI únicos (escala logarítmica)"
            colormap.add_to(folium_map)

            for _, row in grouped.iterrows():
                dni_unicos = int(row["DNI_UNICOS"])
                radius = 8 if max_v <= 0 else max(6, min(30, 6 + (np.log1p(dni_unicos) / np.log1p(max_v)) * 24))
                color_value = float(np.log1p(dni_unicos))
                fill_color = colormap(color_value)
                popup_html = (
                    f"<b>{row['PROVINCIA']}</b><br>"
                    f"DNI únicos: {dni_unicos:,}<br>"
                    f"Registros: {int(row['Registros']):,}"
                )

                folium.CircleMarker(
                    location=[row["lat"], row["lon"]],
                    radius=radius,
                    color=fill_color,
                    fill=True,
                    fill_color=fill_color,
                    fill_opacity=0.8,
                    weight=1,
                    tooltip=popup_html,
                    popup=folium.Popup(popup_html, max_width=280),
                ).add_to(folium_map)

            st_folium(
                folium_map,
                use_container_width=True,
                height=460,
                returned_objects=[],
                key="geo_map",
            )

    # Top provinces table
    st.markdown('<p class="section-header">Top 20 Provincias</p>', unsafe_allow_html=True)
    top_prov = (df.groupby("PROVINCIA")["DNI"].nunique()
                .sort_values(ascending=False).head(20).reset_index())
    top_prov.columns = ["Provincia", "DNI únicos"]
    st.dataframe(top_prov, use_container_width=True, hide_index=True, height=300)


# ── Detail tab ────────────────────────────────────────────────────────────────
def tab_detail(df: pd.DataFrame) -> None:
    st.markdown('<p class="section-header">Detalle de registros</p>', unsafe_allow_html=True)
    cols_available = [c for c in [
        "DNI", "APELLIDOS Y NOMBRES", "CLIENTE", "UNIDAD", "CARGO",
        "PROVINCIA", "DISTRITO", "PLANILLA", "REGIMEN PLANILLA",
        "FECHA DE INGRESO", "FECHA DE CESE",
    ] if c in df.columns]

    col_search, col_info = st.columns([2, 1])
    with col_search:
        search = st.text_input("🔍 Buscar por nombre o DNI", placeholder="Escribe para filtrar...",
                                label_visibility="collapsed")
    with col_info:
        st.caption(f"Mostrando {len(df):,} registros")

    display = df[cols_available].copy()
    if search:
        searchable_cols = [c for c in ["DNI", "APELLIDOS Y NOMBRES", "CLIENTE", "UNIDAD", "CARGO", "PROVINCIA", "DISTRITO"] if c in display.columns]
        search_blob = (
            display[searchable_cols]
            .astype(str)
            .agg(" ".join, axis=1)
            .str.upper()
        )
        mask = search_blob.str.contains(search.upper(), na=False, regex=False)
        display = display[mask]
    display = display.sort_values(["PROVINCIA", "CLIENTE"])

    def style_cese(value):
        if pd.notna(value):
            return "color: #EF4444; font-weight: 600;"
        return "color: #10B981; font-weight: 600;"

    table_config = {
        "DNI": st.column_config.NumberColumn(format="%d"),
        "FECHA DE INGRESO": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "FECHA DE CESE": st.column_config.DateColumn(format="DD/MM/YYYY"),
    }
    if len(display) <= 5000 and "FECHA DE CESE" in display.columns:
        styled = display.style.map(style_cese, subset=["FECHA DE CESE"])
        st.dataframe(
            styled,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config=table_config,
        )
    else:
        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config=table_config,
        )


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    st.markdown(
        '<div class="dashboard-title">📊 Dashboard Ejecutivo — Gestión de Personal</div>'
        '<div class="dashboard-subtitle">'
        'Análisis interactivo · Filtros en cascada · Perú'
        '</div>',
        unsafe_allow_html=True,
    )

    if not DATA_FILE.exists():
        st.error(
            f"No se encontró **{DATA_FILE}**. "
            "Coloca el archivo Excel en la misma carpeta que `app.py` y vuelve a ejecutar."
        )
        st.stop()

    with st.spinner("Cargando datos..."):
        df = load_data(DATA_FILE)

    with st.spinner("Preparando mapa..."):
        coords = get_coordinates(tuple(sorted(df["PROVINCIA"].dropna().unique())))

    filtered, active = apply_filters(df)
    with st.sidebar:
        st.caption(f"province_coords.csv: {'Sí' if COORDS_CACHE.exists() else 'No'}")
        st.caption(f"Streamlit: {STREAMLIT_VERSION}")

    # Active filter badges
    badges = [
        f'<span class="filter-badge">{v}</span>'
        for vals in active.values() for v in vals
    ]
    if badges:
        st.markdown("**Filtros activos:** " + " ".join(badges), unsafe_allow_html=True)

    kpi_cards(filtered)
    st.divider()

    tab_options = {
        "analisis": "📈  Análisis",
        "geografia": "🗺️  Geografía",
        "detalle": "📋  Detalle",
    }
    st.session_state.setdefault("active_tab", "analisis")
    if st.session_state["active_tab"] not in tab_options:
        st.session_state["active_tab"] = "analisis"

    selected_label = st.segmented_control(
        "Sección",
        options=list(tab_options.values()),
        selection_mode="single",
        default=tab_options[st.session_state["active_tab"]],
        label_visibility="collapsed",
    )

    if selected_label:
        st.session_state["active_tab"] = next(
            key for key, label in tab_options.items() if label == selected_label
        )

    if st.session_state["active_tab"] == "analisis":
        tab_analysis(filtered)
    elif st.session_state["active_tab"] == "geografia":
        tab_geography(filtered, coords)
    else:
        tab_detail(filtered)

    st.markdown(
        f'<div style="font-size:0.72rem; color:#94A3B8; text-align:right; margin-top:1rem;">'
        f'Total cargados: <b>{len(df):,}</b> · Filtrados: <b>{len(filtered):,}</b>'
        f'</div>',
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()

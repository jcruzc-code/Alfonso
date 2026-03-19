import unicodedata
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_plotly_events import plotly_events

st.set_page_config(
    page_title="Dashboard Ejecutivo de Personal",
    page_icon="📊",
    layout="wide",
)

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


def normalize_text(value: str) -> str:
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

    df["DNI"] = pd.to_numeric(df["DNI"], errors="coerce").astype("Int64")
    cat_cols = ["CLIENTE", "UNIDAD", "CARGO", "PLANILLA", "REGIMEN PLANILLA"]
    for col in cat_cols:
        df[col] = df[col].fillna("S/I").astype(str).str.strip()

    return df


@st.cache_data(show_spinner=False)
def load_coordinates(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=["PROVINCIA", "lat", "lon"])

    coords = pd.read_csv(path)
    coords["PROVINCIA"] = coords["PROVINCIA"].map(normalize_text)
    coords["lat"] = pd.to_numeric(coords["lat"], errors="coerce")
    coords["lon"] = pd.to_numeric(coords["lon"], errors="coerce")
    return coords.dropna(subset=["lat", "lon"]).drop_duplicates("PROVINCIA")


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros")

    min_date = df["FECHA DE CESE"].dropna().min()
    max_date = df["FECHA DE CESE"].dropna().max()
    include_active = st.sidebar.toggle("Incluir activos (sin fecha de cese)", value=True)

    base = df.copy()
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "Fecha cese (rango)",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )
        if len(date_range) == 2:
            start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
            mask_dates = base["FECHA DE CESE"].between(start_date, end_date, inclusive="both")
            if include_active:
                mask_dates = mask_dates | base["FECHA DE CESE"].isna()
            base = base[mask_dates]
    elif not include_active:
        base = base[base["FECHA DE CESE"].notna()]

    cliente_options = sorted(base["CLIENTE"].dropna().unique())
    cliente = st.sidebar.multiselect("Cliente", cliente_options)

    unit_source = base[base["CLIENTE"].isin(cliente)] if cliente else base
    unidad_options = sorted(unit_source["UNIDAD"].dropna().unique())
    unidad = st.sidebar.multiselect("Unidad", unidad_options)

    cargo_source = unit_source[unit_source["UNIDAD"].isin(unidad)] if unidad else unit_source
    cargo_options = sorted(cargo_source["CARGO"].dropna().unique())
    cargo = st.sidebar.multiselect("Cargo", cargo_options)

    filtered = base.copy()
    if cliente:
        filtered = filtered[filtered["CLIENTE"].isin(cliente)]
    if unidad:
        filtered = filtered[filtered["UNIDAD"].isin(unidad)]
    if cargo:
        filtered = filtered[filtered["CARGO"].isin(cargo)]

    return filtered


def kpi_cards(df: pd.DataFrame) -> None:
    dni_unicos = int(df["DNI"].dropna().nunique())
    registros = int(len(df))
    clientes = int(df["CLIENTE"].nunique())
    unidades = int(df["UNIDAD"].nunique())

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("DNI únicos", f"{dni_unicos:,}")
    c2.metric("Registros", f"{registros:,}")
    c3.metric("Clientes", f"{clientes:,}")
    c4.metric("Unidades", f"{unidades:,}")


def map_and_heat(df: pd.DataFrame, coords: pd.DataFrame) -> pd.DataFrame:
    st.subheader("Mapa geográfico por provincia")

    grouped = (
        df.groupby("PROVINCIA", as_index=False)
        .agg(dni_unicos=("DNI", "nunique"), registros=("DNI", "size"))
        .merge(coords, on="PROVINCIA", how="left")
        .dropna(subset=["lat", "lon"])
    )

    if grouped.empty:
        st.warning("No hay coordenadas disponibles para el filtro actual.")
        return df

    fig_map = px.scatter_geo(
        grouped,
        lat="lat",
        lon="lon",
        size="dni_unicos",
        color="dni_unicos",
        hover_name="PROVINCIA",
        hover_data={"registros": True, "lat": False, "lon": False},
        projection="natural earth",
        color_continuous_scale="Blues",
        height=500,
    )
    fig_map.update_geos(
        showcountries=True,
        countrycolor="#9CA3AF",
        showland=True,
        landcolor="#F8FAFC",
        center={"lat": -9.19, "lon": -75.0152},
        lataxis_range=[-19, 1],
        lonaxis_range=[-82, -67],
    )
    fig_map.update_layout(margin={"l": 0, "r": 0, "t": 10, "b": 0})

    selected_points = plotly_events(fig_map, click_event=True, select_event=False, override_height=500)
    selected_province = None
    if selected_points and "pointNumber" in selected_points[0]:
        idx = selected_points[0]["pointNumber"]
        if idx < len(grouped):
            selected_province = grouped.iloc[idx]["PROVINCIA"]

    st.caption(
        "Haz clic sobre una provincia para aplicar filtro en el resto del dashboard. "
        f"Provincia seleccionada: **{selected_province or 'Ninguna'}**"
    )

    filtered = df if not selected_province else df[df["PROVINCIA"] == selected_province]

    st.subheader("Mapa de calor Provincia vs Unidad")
    hm = (
        filtered.groupby(["PROVINCIA", "UNIDAD"], as_index=False)["DNI"]
        .nunique()
        .rename(columns={"DNI": "DNI_UNICOS"})
    )
    if hm.empty:
        st.info("No hay datos para el filtro actual.")
        return filtered

    pivot = hm.pivot(index="PROVINCIA", columns="UNIDAD", values="DNI_UNICOS").fillna(0)
    fig_hm = px.imshow(
        pivot,
        labels={"x": "Unidad", "y": "Provincia", "color": "DNI únicos"},
        aspect="auto",
        color_continuous_scale="Blues",
        height=420,
    )
    heat_click = plotly_events(fig_hm, click_event=True, select_event=False, override_height=430)
    if heat_click:
        y_idx = heat_click[0].get("y")
        x_idx = heat_click[0].get("x")
        if isinstance(y_idx, (int, float)) and isinstance(x_idx, (int, float)):
            prov = pivot.index[int(y_idx)]
            uni = pivot.columns[int(x_idx)]
            st.caption(f"Filtro adicional aplicado desde mapa de calor: **{prov} / {uni}**")
            filtered = filtered[(filtered["PROVINCIA"] == prov) & (filtered["UNIDAD"] == uni)]

    return filtered


def executive_tables(df: pd.DataFrame) -> None:
    st.subheader("Tablas ejecutivas")
    col1, col2 = st.columns(2)

    with col1:
        t1 = (
            df.groupby("CLIENTE", as_index=False)["DNI"]
            .nunique()
            .rename(columns={"DNI": "DNI_UNICOS"})
            .sort_values("DNI_UNICOS", ascending=False)
        )
        st.markdown("**DNI únicos por cliente**")
        st.dataframe(t1, use_container_width=True, hide_index=True)

    with col2:
        t2 = (
            df.groupby(["PROVINCIA", "UNIDAD"], as_index=False)["DNI"]
            .nunique()
            .rename(columns={"DNI": "DNI_UNICOS"})
            .sort_values("DNI_UNICOS", ascending=False)
            .head(30)
        )
        st.markdown("**Top Provincia / Unidad**")
        st.dataframe(t2, use_container_width=True, hide_index=True)


def detail_table(df: pd.DataFrame) -> None:
    st.subheader("Detalle filtrado")
    cols = [
        "DNI",
        "APELLIDOS Y NOMBRES",
        "CLIENTE",
        "UNIDAD",
        "CARGO",
        "PROVINCIA",
        "DISTRITO",
        "FECHA DE INGRESO",
        "FECHA DE CESE",
    ]
    st.dataframe(df[cols].sort_values(["PROVINCIA", "UNIDAD"]), use_container_width=True, hide_index=True)


def apply_light_theme_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #f5f7fb;
            color: #1f2937;
        }
        [data-testid="stMetricValue"] {
            color: #0f172a;
        }
        [data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 1px solid #e5e7eb;
        }
        .stDataFrame, .stPlotlyChart {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 4px;
            border: 1px solid #e5e7eb;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    apply_light_theme_styles()
    st.title("📈 Dashboard Ejecutivo - Gestión de Personal")
    st.markdown(
        "Visual interactivo tipo BI con filtros conectados entre Cliente, Unidad y Cargo. "
        "Incluye mapa geográfico por provincia y análisis operativo."
    )

    df = load_data(DATA_FILE)
    coords = load_coordinates(COORDS_CACHE)

    filtered = apply_filters(df)
    kpi_cards(filtered)

    filtered_after_maps = map_and_heat(filtered, coords)
    executive_tables(filtered_after_maps)
    detail_table(filtered_after_maps)

    with st.expander("Calidad y normalización aplicada"):
        st.markdown("**Correcciones de provincia**")
        st.json(PROVINCE_CORRECTIONS, expanded=False)
        st.markdown("**Correcciones de distrito**")
        st.json(DISTRICT_CORRECTIONS, expanded=False)


if __name__ == "__main__":
    main()

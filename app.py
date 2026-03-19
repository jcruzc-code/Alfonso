import json
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from geopy.geocoders import Nominatim
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
def build_or_load_coordinates(provinces: tuple[str, ...]) -> pd.DataFrame:
    if COORDS_CACHE.exists():
        coords = pd.read_csv(COORDS_CACHE)
    else:
        geolocator = Nominatim(user_agent="dashboard_streamlit_provincias")
        rows = []
        for prov in provinces:
            if prov == "S/I":
                continue
            query = f"{prov}, Peru"
            lat, lon = np.nan, np.nan
            try:
                location = geolocator.geocode(query, timeout=10)
                if location:
                    lat, lon = location.latitude, location.longitude
            except Exception:
                pass
            rows.append({"PROVINCIA": prov, "lat": lat, "lon": lon})
        coords = pd.DataFrame(rows)
        coords.to_csv(COORDS_CACHE, index=False)

    # Coordenadas de respaldo para provincias frecuentes
    backup_coords = {
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
    }
    for prov, (lat, lon) in backup_coords.items():
        mask = coords["PROVINCIA"].eq(prov) & (coords["lat"].isna() | coords["lon"].isna())
        coords.loc[mask, ["lat", "lon"]] = [lat, lon]

    return coords


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Filtros")

    min_date = df["FECHA DE CESE"].min()
    max_date = df["FECHA DE CESE"].max()

    include_active = st.sidebar.toggle("Incluir activos (sin fecha de cese)", value=True)
    if pd.notna(min_date) and pd.notna(max_date):
        date_range = st.sidebar.date_input(
            "Fecha cese (rango)",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date(),
        )
    else:
        date_range = ()

    cliente = st.sidebar.multiselect("Cliente", sorted(df["CLIENTE"].dropna().unique()))
    unidad = st.sidebar.multiselect("Unidad", sorted(df["UNIDAD"].dropna().unique()))
    cargo = st.sidebar.multiselect("Cargo", sorted(df["CARGO"].dropna().unique()))

    filtered = df.copy()
    if len(date_range) == 2:
        start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        mask_dates = filtered["FECHA DE CESE"].between(start_date, end_date, inclusive="both")
        if include_active:
            mask_dates = mask_dates | filtered["FECHA DE CESE"].isna()
        filtered = filtered[mask_dates]
    elif not include_active:
        filtered = filtered[filtered["FECHA DE CESE"].notna()]

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
    st.subheader("Mapa interactivo por provincia")
    grouped = (
        df.groupby("PROVINCIA", as_index=False)
        .agg(dni_unicos=("DNI", "nunique"), registros=("DNI", "size"))
        .merge(coords, on="PROVINCIA", how="left")
        .dropna(subset=["lat", "lon"])
    )

    fig_map = px.density_map(
        grouped,
        lat="lat",
        lon="lon",
        z="dni_unicos",
        radius=30,
        hover_name="PROVINCIA",
        hover_data={"dni_unicos": True, "registros": True, "lat": False, "lon": False},
        zoom=4.2,
        center={"lat": -9.19, "lon": -75.0152},
        map_style="carto-positron",
        height=450,
    )
    fig_map.add_trace(
        go.Scattermap(
            lat=grouped["lat"],
            lon=grouped["lon"],
            mode="markers",
            marker={"size": grouped["dni_unicos"].clip(lower=6) / 6, "color": "#1f77b4", "opacity": 0.5},
            text=grouped["PROVINCIA"],
            customdata=grouped[["PROVINCIA"]],
            hovertemplate="Provincia: %{text}<br>DNI únicos: %{marker.size:.0f}<extra></extra>",
            name="Provincias",
        )
    )

    selected_points = plotly_events(fig_map, click_event=True, select_event=False, override_height=460)
    selected_province = None
    if selected_points and "pointNumber" in selected_points[0]:
        idx = selected_points[0]["pointNumber"]
        if idx < len(grouped):
            selected_province = grouped.iloc[idx]["PROVINCIA"]

    st.caption(
        "Tip: haz clic sobre el mapa para filtrar por provincia. "
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


def main() -> None:
    st.title("📈 Dashboard Ejecutivo - Gestión de Personal")
    st.markdown(
        "Visual interactivo tipo BI con filtros por fecha de cese, provincia y unidad. "
        "Incluye normalización básica de datos para mejorar consistencia geográfica."
    )

    df = load_data(DATA_FILE)
    coords = build_or_load_coordinates(tuple(sorted(df["PROVINCIA"].dropna().unique())))

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

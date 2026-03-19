# Alfonso - Dashboard Streamlit Ejecutivo

Aplicación en Streamlit para analizar `datos_git.xlsx` con enfoque ejecutivo tipo Power BI / Tableau.

## ¿Qué incluye?
- Filtro por **Fecha de cese** (con opción de incluir personal activo).
- KPIs dinámicos (DNI únicos, registros, clientes y unidades).
- Mapa interactivo por **provincia** con selección por clic.
- Mapa de calor **Provincia vs Unidad** con filtro adicional por clic.
- Tablas ejecutivas y detalle filtrado.
- Normalización básica de nombres de provincia/distrito.

## Ejecutar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Fuente de datos
- Archivo Excel esperado: `datos_git.xlsx`
- Hoja esperada: `WIDE`

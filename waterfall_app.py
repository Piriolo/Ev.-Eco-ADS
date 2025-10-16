import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO
import re

# Configuración de la página
st.set_page_config(
    page_title="Análisis Económico MANNED vs ADS",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Análisis Económico: MANNED vs ADS")
st.markdown("### Gráfico Waterfall Interactivo")

# Sidebar para controles
st.sidebar.header("Controles")
discount_rate = st.sidebar.slider(
    "Tasa de Descuento (%)",
    min_value=0.0,
    max_value=20.0,
    value=8.0,
    step=0.1,
    help="Ajusta la tasa de descuento para el análisis del VPN"
)

# Toggle para ocultar/mostrar impactos cero
hide_zeros = st.sidebar.checkbox(
    "Ocultar impactos en 0",
    value=True,
    help="Si está activado, no se mostrarán categorías cuyo VPN descontado sea 0.00 MUSD"
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Instrucciones:**")
st.sidebar.markdown("1. Carga tu archivo Excel")
st.sidebar.markdown("2. Ajusta la tasa de descuento")
st.sidebar.markdown("3. El gráfico se actualiza automáticamente")
st.sidebar.markdown("4. Hover sobre las barras para más detalles")

# Upload de archivo Excel
st.markdown("### 📁 Cargar Archivo Excel")
uploaded_file = st.file_uploader(
    "Selecciona el archivo 'Ev. Eco ADS.xlsx'",
    type=['xlsx'],
    help="Carga tu archivo Excel para usar datos reales en lugar de los datos de ejemplo"
)

# Utilidad para parsear etiquetas de período tipo "Y01", "Año 02", "Y1", etc.
def parse_year_label(val, base_label_prefix="Y"):
    if pd.isna(val):
        return None
    if isinstance(val, (int, float)) and not np.isnan(val):
        return int(val)
    s = str(val).strip()
    m = re.search(r"(\d+)", s)
    if m:
        n = int(m.group(1))
        return f"{base_label_prefix}{n:02d}"
    return s

# Función para cargar datos del Excel (unidades en MUSD)
@st.cache_data
def load_excel_data(uploaded_file_bytes):
    try:
        if uploaded_file_bytes is not None:
            df = pd.read_excel(BytesIO(uploaded_file_bytes), engine='openpyxl', header=None, na_values=['', ' ', 'NaN', 'NULL'])
            # Categorías (B145:B163)
            categories = []
            for i in range(144, 163):
                if i < len(df) and 1 < len(df.columns):
                    val = df.iloc[i, 1]
                    if pd.notna(val) and str(val).strip():
                        categories.append(str(val).strip())
            # Períodos (D144:AK144)
            years = []
            if 143 < len(df):
                for col in range(3, 37):
                    if col < len(df.columns):
                        val = df.iloc[143, col]
                        parsed = parse_year_label(val)
                        if parsed is not None and str(parsed).strip():
                            years.append(parsed)
            # Matriz de datos (D145:AK163)
            data_matrix = []
            for row in range(144, 163):
                if row < len(df):
                    row_data = []
                    for col in range(3, 37):
                        if col < len(df.columns):
                            val = df.iloc[row, col]
                            try:
                                row_data.append(float(val) if pd.notna(val) else 0.0)
                            except (ValueError, TypeError):
                                row_data.append(0.0)
                        else:
                            row_data.append(0.0)
                    data_matrix.append(row_data)
            # Totales (MUSD)
            manned_total = float(df.iloc[168, 2]) if 168 < len(df) and 2 < len(df.columns) and pd.notna(df.iloc[168, 2]) else 0.0
            ads_total = float(df.iloc[171, 2]) if 171 < len(df) and 2 < len(df.columns) and pd.notna(df.iloc[171, 2]) else 0.0
            if len(categories) > 0 and len(years) > 0 and len(data_matrix) > 0:
                st.success(f"✅ Datos cargados del Excel: {len(categories)} categorías, {len(years)} períodos (MUSD)")
                st.info(f"📊 Totales (MUSD): MANNED={manned_total:,.2f}, ADS={ads_total:,.2f}")
                return categories, years, np.array(data_matrix), manned_total, ads_total
            else:
                st.warning("⚠️ No se encontraron datos válidos en las celdas especificadas")
                st.info(f"Debug: {len(categories)} categorías, {len(years)} períodos, {len(data_matrix)} filas de datos")
    except Exception as e:
        st.error(f"❌ Error al cargar el archivo Excel: {str(e)}")
        st.info("Usando datos de ejemplo mientras tanto...")
    # Datos de ejemplo en MUSD
    categories = ['Operación', 'Mantenimiento', 'Combustible', 'Neumáticos', 'Personal', 'Seguros', 'Depreciación', 'Costos Indirectos', 'Productividad', 'Eficiencia', 'Disponibilidad', 'Utilización', 'Calidad', 'Seguridad', 'Medio Ambiente', 'Capacitación', 'Repuestos', 'Servicios Externos', 'Otros']
    years = [f"Y{n:02d}" for n in range(1, 21)]
    np.random.seed(42)
    data_matrix = np.random.uniform(-0.5, 0.5, (len(categories), len(years)))
    manned_total = 10.0
    ads_total = 8.5
    return categories, years, data_matrix, manned_total, ads_total

# Preparar datos para cargar
uploaded_bytes = None
if uploaded_file is not None:
    uploaded_bytes = uploaded_file.read()
    if 'last_file_name' not in st.session_state or st.session_state.last_file_name != uploaded_file.name:
        st.cache_data.clear()
        st.session_state.last_file_name = uploaded_file.name

# Cargar datos
categories, years, data_matrix, manned_total, ads_total = load_excel_data(uploaded_bytes)

# VPN

def calculate_npv(cash_flows, discount_rate):
    npv = 0.0
    for i, cash_flow in enumerate(cash_flows):
        npv += cash_flow / ((1 + discount_rate/100) ** i)
    return npv

# Preparación de datos: ordenar y filtrar según reglas del usuario

def prepare_sorted_filtered(categories, data_matrix, discount_rate, hide_zeros=True, rename_map=None):
    items = []
    for i, cat in enumerate(categories):
        if i < len(data_matrix):
            npv = calculate_npv(data_matrix[i], discount_rate)
            display_cat = rename_map.get(cat, cat) if rename_map else cat
            items.append({'cat': cat, 'label': display_cat, 'npv': npv})
    if hide_zeros:
        items = [it for it in items if abs(it['npv']) > 1e-9]
    negatives = sorted([it for it in items if it['npv'] < 0], key=lambda x: x['npv'])
    positives = sorted([it for it in items if it['npv'] > 0], key=lambda x: x['npv'], reverse=True)
    ordered = negatives + positives
    ordered_labels = [it['label'] for it in ordered]
    ordered_npvs = [it['npv'] for it in ordered]
    ordered_keys = [it['cat'] for it in ordered]
    return ordered_labels, ordered_npvs, ordered_keys

# Crear gráfico waterfall usando solo increasing/decreasing/totals (compatibilidad Plotly)

def create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate, hide_zeros=True, rename_map=None):
    ordered_labels, ordered_npvs, ordered_keys = prepare_sorted_filtered(categories, data_matrix, discount_rate, hide_zeros, rename_map)

    # Construir vectores alineados
    x_labels = ['MANNED'] + ordered_labels + ['ADS']
    measures = ['absolute'] + ['relative'] * len(ordered_labels) + ['total']

    # ADS forzado = MANNED + suma de relativos
    relatives_sum = sum(ordered_npvs)
    final_adjustment = manned_total + relatives_sum  # La barra total representa el ajuste neto respecto a la base
    values = [manned_total] + ordered_npvs + [final_adjustment]

    fig = go.Figure()
    fig.add_trace(go.Waterfall(
        name='Análisis Waterfall',
        orientation='v',
        measure=measures,
        x=x_labels,
        y=values,
        text=[f"{val:.2f} MUSD" for val in values],
        textposition='outside',
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#DC143C"}},  # positivo = rojo
        decreasing={"marker": {"color": "#2E8B57"}},  # negativo = verde
        totals={"marker": {"color": "#4682B4"}}
    ))

    fig.update_layout(
        title={'text': f"Análisis Waterfall: MANNED vs ADS (Tasa: {discount_rate}%)", 'x': 0.5, 'xanchor': 'center', 'font': {'size': 16}},
        xaxis_title='Categorías',
        yaxis_title='VPN (MUSD)',
        showlegend=False,
        height=600,
        hovermode='x unified'
    )
    fig.update_xaxes(tickangle=45)
    fig.update_yaxes(tickformat=',.2f')
    # ADS calculado visualmente para mostrar métrica fuera
    ads_calc = manned_total + relatives_sum
    return fig, ordered_labels, ordered_npvs, ads_calc

# Bloque de métricas superior
st.markdown('---')
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label='Total MANNED (MUSD)', value=f"{manned_total:,.2f}")
with col2:
    st.metric(label='Total ADS (MUSD)', value=f"{ads_total:,.2f}")
with col3:
    difference = (ads_total - manned_total)
    delta_pct = (difference/manned_total)*100 if manned_total != 0 else 0
    st.metric(label='Diferencia (MUSD)', value=f"{difference:,.2f}", delta=f"{delta_pct:.1f}%", delta_color='inverse')

# Renombrado de categorías (antes de construir gráfico)
st.markdown('---')
st.markdown('### ✏️ Renombrar categorías')
rename_map = {}
for cat in categories:
    new_name = st.text_input(f"Nombre para '{cat}'", value=cat, key=f"rename_{cat}")
    rename_map[cat] = new_name if new_name.strip() else cat

# Gráfico
st.markdown('---')
fig, ordered_labels, ordered_npvs, ads_calc = create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate, hide_zeros, rename_map)
st.plotly_chart(fig, use_container_width=True)


# Tabla de detalles (ordenada, filtrada y con renombres)
st.markdown('### 📋 Detalles por Categoría (Ordenado y Filtrado)')

details_data = []
for label, npv in zip(ordered_labels, ordered_npvs):
    impact_pct = (npv/manned_total)*100 if manned_total != 0 else 0
    details_data.append({'Categoría': label, 'VPN (MUSD)': f"{npv:,.2f}", 'Impacto (%)': f"{impact_pct:.2f}%"})

if details_data:
    df_details = pd.DataFrame(details_data)
    st.dataframe(df_details, use_container_width=True)
else:
    st.warning('No hay datos para mostrar en la tabla de detalles tras el filtrado')

# Información adicional
st.markdown('---')
st.markdown('### 📊 Información del Dataset')
col1, col2 = st.columns(2)
with col1:
    st.info(f"**Categorías cargadas:** {len(categories)}")
    st.info(f"**Período de análisis:** {len(years)} períodos")
with col2:
    if len(years) > 0:
        st.info(f"**Períodos:** {years[0]} - {years[-1]}")
    st.info(f"**Ocultar impactos cero:** {'Sí' if hide_zeros else 'No'}")
    st.info(f"**Tasa de descuento actual:** {discount_rate}%")

# Footer
st.markdown('---')
st.markdown('*Desarrollado para análisis económico MANNED vs ADS con tasa de descuento variable (unidades en MUSD)*')
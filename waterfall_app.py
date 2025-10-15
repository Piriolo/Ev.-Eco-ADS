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

# Utilidad para parsear etiquetas de año tipo "Y01", "Año 02", "Y1", etc.
def parse_year_label(val, base_label_prefix="Y"):
    """Devuelve una etiqueta estandarizada de año.
    Acepta valores numéricos (e.g., 2025) o cadenas tipo 'Y01', 'Y1', 'Year 02'.
    Para strings, extrae el número y lo formatea como Y01, Y02, ...
    """
    if pd.isna(val):
        return None
    # Si es número real (año calendario), lo devolvemos como int
    if isinstance(val, (int, float)) and not np.isnan(val):
        return int(val)
    # Si es string, extraer dígitos
    s = str(val).strip()
    m = re.search(r"(\d+)", s)
    if m:
        n = int(m.group(1))
        # Normalizar a Y01, Y02, ... (hasta Y99). Si prefieres sin cero a la izquierda usa f"Y{n}"
        return f"{base_label_prefix}{n:02d}"
    # Si no hay dígitos, devolver tal cual por transparencia
    return s

# Función para cargar datos del Excel
@st.cache_data
def load_excel_data(uploaded_file_bytes):
    try:
        if uploaded_file_bytes is not None:
            # Leer con pandas (más robusto que openpyxl)
            df = pd.read_excel(
                BytesIO(uploaded_file_bytes), 
                engine='openpyxl',
                header=None,
                na_values=['', ' ', 'NaN', 'NULL']
            )
            
            # Leer categorías (B145:B163)
            categories = []
            for i in range(144, 163):  # índices 144-162
                if i < len(df) and 1 < len(df.columns):  # col B = 1
                    val = df.iloc[i, 1]
                    if pd.notna(val) and str(val).strip():
                        categories.append(str(val).strip())
            
            # Leer años (D144:AK144) y soportar etiquetas tipo Y01
            years = []
            raw_years = []
            if 143 < len(df):
                for col in range(3, 37):  # D=3, AK=36
                    if col < len(df.columns):
                        val = df.iloc[143, col]
                        parsed = parse_year_label(val)
                        if parsed is not None and str(parsed).strip():
                            years.append(parsed)
                            raw_years.append(val)
            
            # Leer matriz de datos (D145:AK163)
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
            
            # Leer totales
            manned_total = 0.0
            ads_total = 0.0
            if 168 < len(df) and 2 < len(df.columns):
                try:
                    manned_total = float(df.iloc[168, 2]) if pd.notna(df.iloc[168, 2]) else 0.0
                except (ValueError, TypeError):
                    manned_total = 0.0
            if 171 < len(df) and 2 < len(df.columns):
                try:
                    ads_total = float(df.iloc[171, 2]) if pd.notna(df.iloc[171, 2]) else 0.0
                except (ValueError, TypeError):
                    ads_total = 0.0
            
            # Validar datos
            if len(categories) > 0 and len(years) > 0 and len(data_matrix) > 0:
                st.success(f"✅ Datos cargados del Excel: {len(categories)} categorías, {len(years)} períodos")
                # Mostrar cómo quedaron las etiquetas de años
                st.info(f"🕑 Etiquetas de período: {years[:5]}{' ...' if len(years) > 5 else ''}")
                st.info(f"📊 Totales: MANNED=${manned_total:,.0f}, ADS=${ads_total:,.0f}")
                return categories, years, np.array(data_matrix), manned_total, ads_total
            else:
                st.warning("⚠️ No se encontraron datos válidos en las celdas especificadas")
                st.info(f"Debug: {len(categories)} categorías, {len(years)} períodos, {len(data_matrix)} filas de datos")
                
    except Exception as e:
        st.error(f"❌ Error al cargar el archivo Excel: {str(e)}")
        st.info("Usando datos de ejemplo mientras tanto...")
    
    # Datos de ejemplo si no se puede cargar el archivo
    st.info("📝 Usando datos de ejemplo. Carga tu archivo Excel para ver datos reales.")
    categories = [
        'Operación', 'Mantenimiento', 'Combustible', 'Neumáticos', 'Personal',
        'Seguros', 'Depreciación', 'Costos Indirectos', 'Productividad',
        'Eficiencia', 'Disponibilidad', 'Utilización', 'Calidad', 'Seguridad',
        'Medio Ambiente', 'Capacitación', 'Repuestos', 'Servicios Externos', 'Otros'
    ]
    years = [f"Y{n:02d}" for n in range(1, 21)]  # Y01..Y20 como ejemplo
    np.random.seed(42)
    data_matrix = np.random.uniform(-500000, 500000, (len(categories), len(years)))
    manned_total = 10000000.0
    ads_total = 8500000.0
    return categories, years, data_matrix, manned_total, ads_total

# Preparar datos para cargar
uploaded_bytes = None
if uploaded_file is not None:
    uploaded_bytes = uploaded_file.read()
    # Limpiar cache cuando se carga un nuevo archivo
    if 'last_file_name' not in st.session_state or st.session_state.last_file_name != uploaded_file.name:
        st.cache_data.clear()
        st.session_state.last_file_name = uploaded_file.name

# Cargar datos
categories, years, data_matrix, manned_total, ads_total = load_excel_data(uploaded_bytes)

# Función para calcular VPN
def calculate_npv(cash_flows, discount_rate):
    """Calcula el Valor Presente Neto"""
    npv = 0
    for i, cash_flow in enumerate(cash_flows):
        npv += cash_flow / ((1 + discount_rate/100) ** i)
    return npv

# Función para crear gráfico waterfall
def create_waterfall_chart(categories, data, manned_total, ads_total, discount_rate):
    # Calcular flujos de caja descontados por categoría
    discounted_values = []
    for i, category in enumerate(categories):
        if i < len(data):
            cash_flows = data[i]
            npv = calculate_npv(cash_flows, discount_rate)
            discounted_values.append(npv)
        else:
            discounted_values.append(0)
    
    # Preparar datos para el waterfall
    x_labels = ['MANNED (Base)'] + categories + ['ADS (Final)']
    
    # Calcular valores acumulativos
    values = [manned_total]
    cumulative = manned_total
    for val in discounted_values:
        values.append(val)
        cumulative += val
    
    # Ajuste final para llegar al total ADS
    final_adjustment = ads_total - cumulative
    values.append(final_adjustment)
    
    # Crear el gráfico waterfall
    fig = go.Figure()
    measures = ["absolute"] + ["relative"] * len(categories) + ["total"]
    fig.add_trace(go.Waterfall(
        name="Análisis Waterfall",
        orientation="v",
        measure=measures,
        x=x_labels,
        textposition="outside",
        text=[f"${val/1000000:.1f}M" if abs(val) >= 1000000 else f"${val/1000:.0f}K" for val in values],
        y=values,
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#2E8B57"}},
        decreasing={"marker": {"color": "#DC143C"}},
        totals={"marker": {"color": "#4682B4"}}
    ))
    fig.update_layout(
        title={'text': f"Análisis Waterfall: MANNED vs ADS (Tasa: {discount_rate}%)", 'x': 0.5, 'xanchor': 'center', 'font': {'size': 16}},
        xaxis_title="Categorías",
        yaxis_title="Valor Presente Neto (USD)",
        showlegend=False,
        height=600,
        hovermode='x unified'
    )
    fig.update_xaxes(tickangle=45)
    fig.update_yaxes(tickformat="$,.0f")
    return fig

# Mostrar información de los datos
st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="Total MANNED", value=f"${manned_total/1000000:.1f}M" if abs(manned_total) >= 1000000 else f"${manned_total/1000:.0f}K", help="Valor base del caso MANNED")
with col2:
    st.metric(label="Total ADS", value=f"${ads_total/1000000:.1f}M" if abs(ads_total) >= 1000000 else f"${ads_total/1000:.0f}K", help="Valor objetivo del caso ADS")
with col3:
    difference = ads_total - manned_total
    delta_pct = (difference/manned_total)*100 if manned_total != 0 else 0
    st.metric(label="Diferencia", value=f"${difference/1000000:.1f}M" if abs(difference) >= 1000000 else f"${difference/1000:.0f}K", delta=f"{delta_pct:.1f}%", help="Diferencia entre ADS y MANNED")

# Crear y mostrar el gráfico
st.markdown("---")
fig = create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate)
st.plotly_chart(fig, use_container_width=True)

# Tabla de detalles
st.markdown("### 📋 Detalles por Categoría")

details_data = []
for i, category in enumerate(categories):
    if i < len(data_matrix):
        cash_flows = data_matrix[i]
        npv = calculate_npv(cash_flows, discount_rate)
        impact_pct = (npv/manned_total)*100 if manned_total != 0 else 0
        details_data.append({'Categoría': category, 'VPN (USD)': f"${npv:,.0f}", 'VPN (M USD)': f"${npv/1000000:.2f}M" if abs(npv) >= 1000000 else f"${npv/1000:.0f}K", 'Impacto (%)': f"{impact_pct:.2f}%"})

if details_data:
    df_details = pd.DataFrame(details_data)
    st.dataframe(df_details, use_container_width=True)
else:
    st.warning("No hay datos para mostrar en la tabla de detalles")

# Información adicional
st.markdown("---")
st.markdown("### 📊 Información del Dataset")
col1, col2 = st.columns(2)
with col1:
    st.info(f"**Categorías:** {len(categories)}")
    st.info(f"**Período de análisis:** {len(years)} períodos")
with col2:
    if len(years) > 0:
        st.info(f"**Períodos:** {years[0]} - {years[-1]}")
    st.info(f"**Tasa de descuento actual:** {discount_rate}%")

# Footer
st.markdown("---")
st.markdown("*Desarrollado para análisis económico MANNED vs ADS con tasa de descuento variable*")
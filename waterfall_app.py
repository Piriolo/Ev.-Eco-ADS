import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO

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

# Función para cargar datos del Excel
@st.cache_data
def load_excel_data(uploaded_file_bytes):
    try:
        if uploaded_file_bytes is not None:
            # Cargar el archivo Excel desde los bytes
            workbook = openpyxl.load_workbook(BytesIO(uploaded_file_bytes), data_only=True)
            
            # Seleccionar la hoja (primera hoja por defecto)
            sheet = workbook.active
            
            # Leer categorías (B145:B163)
            categories = []
            for row in range(145, 164):
                cell_value = sheet[f'B{row}'].value
                if cell_value is not None and str(cell_value).strip() != "":
                    categories.append(str(cell_value).strip())
            
            # Leer años (D144:AK144)
            years = []
            for col in range(4, 37):  # D=4, AK=37
                cell_value = sheet.cell(row=144, column=col).value
                if cell_value is not None and str(cell_value).strip() != "":
                    try:
                        years.append(int(float(cell_value)))
                    except:
                        pass
            
            # Leer datos (D145:AK163)
            data_matrix = []
            for row in range(145, 164):
                row_data = []
                for col in range(4, 37):
                    cell_value = sheet.cell(row=row, column=col).value
                    try:
                        row_data.append(float(cell_value) if cell_value is not None else 0.0)
                    except:
                        row_data.append(0.0)
                data_matrix.append(row_data)
            
            # Leer totales
            manned_total = sheet['C169'].value
            ads_total = sheet['C172'].value
            
            try:
                manned_total = float(manned_total) if manned_total is not None else 0.0
                ads_total = float(ads_total) if ads_total is not None else 0.0
            except:
                manned_total = 0.0
                ads_total = 0.0
            
            # Verificar que tenemos datos válidos
            if len(categories) > 0 and len(years) > 0 and len(data_matrix) > 0:
                st.success(f"✅ Datos cargados del Excel: {len(categories)} categorías, {len(years)} años")
                return categories, years, np.array(data_matrix), manned_total, ads_total
            else:
                st.warning("⚠️ No se encontraron datos válidos en las celdas especificadas")
            
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
    
    years = list(range(2025, 2045))
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
    
    # Preparar medidas para el waterfall
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
    
    # Personalizar el layout
    fig.update_layout(
        title={
            'text': f"Análisis Waterfall: MANNED vs ADS (Tasa: {discount_rate}%)",
            'x': 0.5,
            'xanchor': 'center',
            'font': {'size': 16}
        },
        xaxis_title="Categorías",
        yaxis_title="Valor Presente Neto (USD)",
        showlegend=False,
        height=600,
        hovermode='x unified'
    )
    
    # Rotar etiquetas del eje X
    fig.update_xaxes(tickangle=45)
    
    # Formatear eje Y
    fig.update_yaxes(tickformat="$,.0f")
    
    return fig

# Mostrar información de los datos
st.markdown("---")
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        label="Total MANNED",
        value=f"${manned_total/1000000:.1f}M" if abs(manned_total) >= 1000000 else f"${manned_total/1000:.0f}K",
        help="Valor base del caso MANNED"
    )

with col2:
    st.metric(
        label="Total ADS",
        value=f"${ads_total/1000000:.1f}M" if abs(ads_total) >= 1000000 else f"${ads_total/1000:.0f}K",
        help="Valor objetivo del caso ADS"
    )

with col3:
    difference = ads_total - manned_total
    delta_pct = (difference/manned_total)*100 if manned_total != 0 else 0
    st.metric(
        label="Diferencia",
        value=f"${difference/1000000:.1f}M" if abs(difference) >= 1000000 else f"${difference/1000:.0f}K",
        delta=f"{delta_pct:.1f}%",
        help="Diferencia entre ADS y MANNED"
    )

# Crear y mostrar el gráfico
st.markdown("---")
fig = create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate)
st.plotly_chart(fig, use_container_width=True)

# Tabla de detalles
st.markdown("### 📋 Detalles por Categoría")

# Calcular VPN por categoría
details_data = []
for i, category in enumerate(categories):
    if i < len(data_matrix):
        cash_flows = data_matrix[i]
        npv = calculate_npv(cash_flows, discount_rate)
        
        impact_pct = (npv/manned_total)*100 if manned_total != 0 else 0
        
        details_data.append({
            'Categoría': category,
            'VPN (USD)': f"${npv:,.0f}",
            'VPN (M USD)': f"${npv/1000000:.2f}M" if abs(npv) >= 1000000 else f"${npv/1000:.0f}K",
            'Impacto (%)': f"{impact_pct:.2f}%"
        })

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
    st.info(f"**Período de análisis:** {len(years)} años")
    
with col2:
    if len(years) > 0:
        st.info(f"**Años:** {min(years)} - {max(years)}")
    st.info(f"**Tasa de descuento actual:** {discount_rate}%")

# Footer
st.markdown("---")
st.markdown("*Desarrollado para análisis económico MANNED vs ADS con tasa de descuento variable*")
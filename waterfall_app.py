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
st.sidebar.markdown("1. Ajusta la tasa de descuento")
st.sidebar.markdown("2. El gráfico se actualiza automáticamente")
st.sidebar.markdown("3. Hover sobre las barras para más detalles")

# Función para cargar datos del Excel
@st.cache_data
def load_excel_data():
    try:
        # Intentar cargar el archivo Excel desde el repositorio
        file_path = "Ev. Eco ADS.xlsx"
        
        # Si existe el archivo, leerlo
        if st.session_state.get('excel_file') is not None:
            workbook = openpyxl.load_workbook(st.session_state.excel_file)
            sheet_names = workbook.sheetnames
            
            # Buscar la hoja correcta (primera hoja por defecto)
            sheet = workbook.active
            
            # Leer categorías (B145:B163)
            categories = []
            for row in range(145, 164):
                cell_value = sheet[f'B{row}'].value
                if cell_value:
                    categories.append(str(cell_value))
            
            # Leer años (D144:AK144)
            years = []
            for col in range(4, 37):  # D=4, AK=37
                cell_value = sheet.cell(row=144, column=col).value
                if cell_value:
                    years.append(int(cell_value))
            
            # Leer datos (D145:AK163)
            data_matrix = []
            for row in range(145, 164):
                row_data = []
                for col in range(4, 37):
                    cell_value = sheet.cell(row=row, column=col).value
                    row_data.append(float(cell_value) if cell_value else 0.0)
                data_matrix.append(row_data)
            
            # Leer totales
            manned_total = sheet['C169'].value or 0
            ads_total = sheet['C172'].value or 0
            
            return categories, years, np.array(data_matrix), manned_total, ads_total
            
    except Exception as e:
        st.warning(f"No se pudo cargar el archivo Excel: {e}")
        
    # Datos de ejemplo si no se puede cargar el archivo
    categories = [
        'Operación', 'Mantenimiento', 'Combustible', 'Neumáticos', 'Personal',
        'Seguros', 'Depreciación', 'Costos Indirectos', 'Productividad',
        'Eficiencia', 'Disponibilidad', 'Utilización', 'Calidad', 'Seguridad',
        'Medio Ambiente', 'Capacitación', 'Repuestos', 'Servicios Externos', 'Otros'
    ]
    
    years = list(range(2025, 2045))
    np.random.seed(42)
    data_matrix = np.random.uniform(-500000, 500000, (len(categories), len(years)))
    
    manned_total = 10000000
    ads_total = 8500000
    
    return categories, years, data_matrix, manned_total, ads_total

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
        cash_flows = data[i]
        npv = calculate_npv(cash_flows, discount_rate)
        discounted_values.append(npv)
    
    # Preparar datos para el waterfall
    x_labels = ['MANNED (Base)'] + categories + ['ADS (Final)']
    
    # Calcular valores acumulativos
    values = [manned_total]
    cumulative = manned_total
    
    for val in discounted_values:
        values.append(val)
        cumulative += val
    
    values.append(ads_total - cumulative)  # Ajuste final
    
    # Crear el gráfico waterfall
    fig = go.Figure()
    
    # Barra inicial (MANNED)
    fig.add_trace(go.Waterfall(
        name="Análisis Waterfall",
        orientation="v",
        measure=["absolute"] + ["relative"] * len(categories) + ["total"],
        x=x_labels,
        textposition="outside",
        text=[f"${val/1000000:.1f}M" for val in values],
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

# Cargar datos
categories, years, data_matrix, manned_total, ads_total = load_excel_data()

# Mostrar información de los datos
col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        label="Total MANNED",
        value=f"${manned_total/1000000:.1f}M",
        help="Valor base del caso MANNED"
    )

with col2:
    st.metric(
        label="Total ADS",
        value=f"${ads_total/1000000:.1f}M",
        help="Valor objetivo del caso ADS"
    )

with col3:
    difference = ads_total - manned_total
    st.metric(
        label="Diferencia",
        value=f"${difference/1000000:.1f}M",
        delta=f"{(difference/manned_total)*100:.1f}%",
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
    cash_flows = data_matrix[i]
    npv = calculate_npv(cash_flows, discount_rate)
    
    details_data.append({
        'Categoría': category,
        'VPN (USD)': f"${npv:,.0f}",
        'VPN (M USD)': f"${npv/1000000:.2f}M",
        'Impacto (%)': f"{(npv/manned_total)*100:.2f}%"
    })

df_details = pd.DataFrame(details_data)
st.dataframe(df_details, use_container_width=True)

# Upload de archivo Excel
st.markdown("---")
st.markdown("### 📁 Cargar Archivo Excel")
uploaded_file = st.file_uploader(
    "Selecciona el archivo 'Ev. Eco ADS.xlsx'",
    type=['xlsx'],
    help="Carga tu archivo Excel para usar datos reales en lugar de los datos de ejemplo"
)

if uploaded_file is not None:
    st.session_state.excel_file = uploaded_file
    st.success("¡Archivo cargado exitosamente! Recarga la página para ver los datos actualizados.")
    
# Footer
st.markdown("---")
st.markdown("*Desarrollado para análisis económico MANNED vs ADS con tasa de descuento variable*")
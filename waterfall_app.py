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

# Ajustes de eje Y
st.sidebar.subheader("Eje Y (MUSD)")
y_min = st.sidebar.number_input("Mínimo", value=None, placeholder="auto")
y_max = st.sidebar.number_input("Máximo", value=None, placeholder="auto")

# Toggle para ocultar/mostrar impactos cero
hide_zeros = st.sidebar.checkbox(
    "Ocultar impactos en 0",
    value=True,
    help="Si está activado, no se mostrarán categorías cuyo VPN descontado sea 0.00 MUSD"
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Instrucciones:**")
st.sidebar.markdown("1. Carga tu archivo Excel")
st.sidebar.markdown("2. Selecciona la hoja a usar (si tiene varias)")
st.sidebar.markdown("3. Ajusta tasa, filtros y eje Y")
st.sidebar.markdown("4. El gráfico se actualiza automáticamente")

# Persistencia de sesión
if 'rename_map' not in st.session_state:
    st.session_state.rename_map = {}
if 'last_file_name' not in st.session_state:
    st.session_state.last_file_name = None
if 'y_limits' not in st.session_state:
    st.session_state.y_limits = {'min': None, 'max': None}
if 'sheet_name' not in st.session_state:
    st.session_state.sheet_name = None

# Upload de archivo Excel
st.markdown("### 📁 Cargar Archivo Excel")
uploaded_file = st.file_uploader(
    "Selecciona el archivo 'Ev. Eco ADS.xlsx'",
    type=['xlsx'],
    help="Carga tu archivo Excel para usar datos reales"
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

# Función para listar hojas del Excel
@st.cache_data
def list_excel_sheets(file_bytes):
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes), engine='openpyxl')
        return xl.sheet_names
    except Exception:
        return []

# Función para cargar datos del Excel (unidades en MUSD) desde una hoja específica
@st.cache_data
def load_excel_data(uploaded_file_bytes, sheet_name):
    try:
        if uploaded_file_bytes is not None:
            df = pd.read_excel(BytesIO(uploaded_file_bytes), engine='openpyxl', header=None, na_values=['', ' ', 'NaN', 'NULL'], sheet_name=sheet_name)
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
                st.success(f"✅ Datos cargados (hoja: {sheet_name}): {len(categories)} categorías, {len(years)} períodos (MUSD)")
                st.info(f"📊 Totales (MUSD): MANNED={manned_total:,.2f}, ADS={ads_total:,.2f}")
                return categories, years, np.array(data_matrix), manned_total, ads_total
            else:
                st.warning("⚠️ No se encontraron datos válidos en las celdas especificadas en esta hoja")
                st.info(f"Debug: {len(categories)} categorías, {len(years)} períodos, {len(data_matrix)} filas de datos")
    except Exception as e:
        st.error(f"❌ Error al cargar el archivo Excel/hoja: {str(e)}")
        st.info("Usando datos de ejemplo mientras tanto...")
    # Datos de ejemplo
    categories = ['Operación', 'Mantenimiento', 'Combustible', 'Neumáticos', 'Personal', 'Seguros', 'Depreciación', 'Costos Indirectos', 'Productividad', 'Eficiencia', 'Disponibilidad', 'Utilización', 'Calidad', 'Seguridad', 'Medio Ambiente', 'Capacitación', 'Repuestos', 'Servicios Externos', 'Otros']
    years = [f"Y{n:02d}" for n in range(1, 21)]
    np.random.seed(42)
    data_matrix = np.random.uniform(-0.5, 0.5, (len(categories), len(years)))
    manned_total = 10.0
    ads_total = 8.5
    return categories, years, data_matrix, manned_total, ads_total

# Preparar datos para cargar y persistir
uploaded_bytes = None
sheet_selected = None
if uploaded_file is not None:
    uploaded_bytes = uploaded_file.read()
    # Si cambia el archivo, limpiar cache y resetear hoja
    if st.session_state.last_file_name != uploaded_file.name:
        st.cache_data.clear()
        st.session_state.last_file_name = uploaded_file.name
        st.session_state.sheet_name = None

    # Listar hojas del archivo subido
    sheets = list_excel_sheets(uploaded_bytes)
    if sheets:
        st.sidebar.subheader("Hoja a utilizar")
        default_idx = 0
        if st.session_state.sheet_name in sheets:
            default_idx = sheets.index(st.session_state.sheet_name)
        sheet_selected = st.sidebar.selectbox("Selecciona hoja", options=sheets, index=default_idx)
        st.session_state.sheet_name = sheet_selected

# Cargar datos (usar hoja seleccionada o la primera)
sheet_name_to_use = st.session_state.sheet_name
categories, years, data_matrix, manned_total, ads_total = load_excel_data(uploaded_bytes, sheet_name_to_use)

# Inicializar rename_map con categorías si está vacío
for cat in categories:
    st.session_state.rename_map.setdefault(cat, cat)

# VPN

def calculate_npv(cash_flows, discount_rate):
    npv = 0.0
    for i, cash_flow in enumerate(cash_flows):
        npv += cash_flow / ((1 + discount_rate/100) ** i)
    return npv

# Preparación de datos

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

# Gráfico Waterfall

def create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate, hide_zeros=True, rename_map=None, y_min=None, y_max=None):
    ordered_labels, ordered_npvs, ordered_keys = prepare_sorted_filtered(categories, data_matrix, discount_rate, hide_zeros, rename_map)

    x_labels = ['MANNED'] + ordered_labels + ['ADS']
    measures = ['absolute'] + ['relative'] * len(ordered_labels) + ['total']

    relatives_sum = sum(ordered_npvs)
    values = [manned_total] + ordered_npvs + [relatives_sum]

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

    # Limites de eje Y si se definieron
    y0 = st.session_state.y_limits.get('min') if y_min is None else y_min
    y1 = st.session_state.y_limits.get('max') if y_max is None else y_max
    if y0 is not None or y1 is not None:
        fig.update_yaxes(range=[y0 if y0 is not None else None, y1 if y1 is not None else None])

    ads_calc = manned_total + relatives_sum
    return fig, ordered_labels, ordered_npvs, ads_calc

# Métricas superiores
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

# Selector y edición de categorías (UI compacta)
st.markdown('---')
st.markdown('### ✏️ Renombrar categorías')
col_sel, col_edit = st.columns([1,2])
with col_sel:
    selected_cat = st.selectbox('Selecciona categoría', options=categories, index=0)
with col_edit:
    new_label = st.text_input('Nuevo nombre', value=st.session_state.rename_map.get(selected_cat, selected_cat))
    if st.button('Guardar nombre'):
        st.session_state.rename_map[selected_cat] = new_label.strip() or selected_cat
        st.success(f"Nombre actualizado: {selected_cat} → {st.session_state.rename_map[selected_cat]}")

rename_map = st.session_state.rename_map

# Persistir límites de eje Y
if y_min is not None:
    st.session_state.y_limits['min'] = y_min
if y_max is not None:
    st.session_state.y_limits['max'] = y_max

# Gráfico
st.markdown('---')
fig, ordered_labels, ordered_npvs, ads_calc = create_waterfall_chart(categories, data_matrix, manned_total, ads_total, discount_rate, hide_zeros, rename_map, y_min, y_max)
st.plotly_chart(fig, use_container_width=True)

# Métrica de ADS calculado (MANNED + suma de barras)
colA, colB = st.columns(2)
with colA:
    st.metric(label='ADS calc (MUSD)', value=f"{ads_calc:,.2f}")
with colB:
    st.caption('ADS calc = MANNED + suma de barras (verdes/rojas)')

# Tabla de detalles
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
    st.info(f"**Hoja seleccionada:** {sheet_name_to_use if sheet_name_to_use else 'Primera disponible'}")
    st.info(f"**Ocultar impactos cero:** {'Sí' if hide_zeros else 'No'}")
    st.info(f"**Tasa de descuento actual:** {discount_rate}%")

# Footer
st.markdown('---')
st.markdown('*Desarrollado para análisis económico MANNED vs ADS con tasa de descuento variable (unidades en MUSD)*')
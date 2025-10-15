# Análisis Económico MANNED vs ADS

## Descripción
Aplicación interactiva desarrollada en Streamlit para el análisis económico comparativo entre sistemas MANNED y ADS (Autonomous Drilling Systems) utilizando gráficos waterfall con tasa de descuento personalizable.

## Características
- **Gráfico Waterfall Interactivo**: Visualización que muestra la transición económica de MANNED a ADS
- **Tasa de Descuento Personalizable**: Control deslizante para ajustar la tasa de descuento (0-20%)
- **Análisis VPN**: Cálculo del Valor Presente Neto para cada categoría
- **Carga de Datos**: Capacidad de cargar el archivo Excel real para análisis con datos actuales
- **Tabla Detallada**: Breakdown por categoría con impacto porcentual

## Estructura de Datos
La aplicación lee datos del archivo "Ev. Eco ADS.xlsx" con la siguiente estructura:
- **Categorías**: Celdas B145:B163 (19 categorías)
- **Años**: Celdas D144:AK144 (años de análisis)
- **Datos**: Celdas D145:AK163 (valores por categoría y año)
- **Total MANNED**: Celda C169
- **Total ADS**: Celda C172

## Instalación y Uso

### 1. Clonar el repositorio
```bash
git clone https://github.com/Piriolo/Ev.-Eco-ADS.git
cd Ev.-Eco-ADS
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Ejecutar la aplicación
```bash
streamlit run waterfall_app.py
```

### 4. Uso
1. La aplicación se abrirá en tu navegador (por defecto http://localhost:8501)
2. Utiliza el control deslizante en la barra lateral para ajustar la tasa de descuento
3. Opcionalmente, carga tu archivo Excel para usar datos reales
4. Explora el gráfico waterfall interactivo y la tabla de detalles

## Funcionalidades

### Gráfico Waterfall
- Visualización clara de la transición de MANNED a ADS
- Barras verdes para impactos positivos
- Barras rojas para impactos negativos
- Barras azules para totales
- Valores mostrados en millones de dólares

### Controles Interactivos
- **Tasa de Descuento**: Ajustable de 0% a 20% con incrementos de 0.1%
- **Hover Details**: Información detallada al pasar el mouse sobre las barras
- **Responsive Design**: Se adapta a diferentes tamaños de pantalla

### Métricas Principales
- Total MANNED (valor base)
- Total ADS (valor objetivo)
- Diferencia absoluta y porcentual

### Análisis Detallado
- Tabla con VPN por categoría
- Impacto porcentual de cada categoría
- Valores en USD y millones de USD

## Tecnologías Utilizadas
- **Streamlit**: Framework de aplicación web
- **Plotly**: Gráficos interactivos
- **Pandas**: Manipulación de datos
- **NumPy**: Cálculos numéricos
- **OpenPyXL**: Lectura de archivos Excel

## Datos de Ejemplo
Si no se proporciona el archivo Excel, la aplicación utiliza datos de ejemplo con:
- 19 categorías económicas típicas de minería
- 20 años de proyección (2025-2044)
- Valores aleatorios para demostración

## Estructura de Archivos
```
Ev.-Eco-ADS/
├── waterfall_app.py      # Aplicación principal
├── requirements.txt      # Dependencias de Python
├── Ev. Eco ADS.xlsx     # Archivo de datos Excel
└── README.md            # Documentación
```

## Autor
Desarrollado para análisis económico en proyectos mineros MANNED vs ADS.
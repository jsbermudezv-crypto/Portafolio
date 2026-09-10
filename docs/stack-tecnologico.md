# Stack Tecnológico

## Backend & Procesamiento de Datos

### Python
- **pandas**: Manipulación y análisis de datos
- **geopandas**: Datos geoespaciales
- **openpyxl**: Lectura/escritura de Excel
- **pdfplumber**: Extracción de tablas de PDFs
- **PyMuPDF**: Manipulación avanzada de PDFs
- **matplotlib & plotly**: Visualización

### Bases de Datos y BI
- **SQL avanzado**: Queries complejas, window functions, CTEs
- **Power BI**: Dashboards empresariales
- **Looker Studio**: Dashboards web
- **Excel avanzado**: Fórmulas, tablas dinámicas, VBA

### Nube y Pipelines
- **AWS**: S3, Lambda, RDS
- **Azure**: Data Factory, Synapse
- **Procesos ETL**: Orquestación y transformación

## Frontend y Automatización

### Interfaces
- **tkinter**: Aplicaciones de escritorio simples
- **Google Apps Script**: Automatización en Google Workspace
- **HTML/CSS/JavaScript**: Dashboards web autocontenidos

### Librerías Geoespaciales
- **folium**: Mapas interactivos básicos
- **MapLibre**: Mapas con datos vectoriales avanzados
- **shapely**: Operaciones geométricas

## Inteligencia Artificial

### Modelos de Lenguaje
- **Google Gemini API**: Extracción de documentos
- **Tesseract**: OCR de documentos escaneados

### Integración
- **APIs REST**: Comunicación con servicios externos
- **Google Apps Script**: Integración con APIs de Google

## DevOps y Versionado

### Control de Versiones
- **Git**: Versionado de código
- **GitHub**: Repositorio remoto y colaboración

## Instalación de Dependencias

```bash
# Instalación base
pip install -r requirements.txt

# Con extras geoespaciales
pip install -e ".[geospatial]"

# Con extras OCR
pip install -e ".[ocr]"

# Desarrollo
pip install -e ".[dev]"
```

# Portafolio de Análisis y Automatización de Datos

## 👤 Juan Sebastián Bermúdez Vélez

**Analista de Datos | Ingeniero Electrónico**

Cuatro años construyendo soluciones de datos para el mercado estadounidense, en el sector de tierras y petróleo. Especializado en transformar datos desordenados en información consultable y herramientas listas para producción.

**Perfil:**
- Escaneos sin OCR → PDFs buscables
- Runsheets con formatos distintos → Tablas normalizadas
- Estados de cuenta dispersos → Consolidaciones reconciliadas
- Procesos manuales → Automatización con interfaces amigables

**Formación:**
- Ingeniero Electrónico, Universidad El Bosque (mejor promedio de facultad)
- Idiomas: Español, Inglés, Francés

**Contacto:**
- 📧 [jsbermudezv@gmail.com](mailto:jsbermudezv@gmail.com)
- 💼 [LinkedIn](https://linkedin.com/in/jsbermudezv)
- 🌐 [Portafolio Web](https://jsbermudezv.github.io/portafolio)
- 📍 Bogotá, Colombia

---

## 🚀 Proyectos

| # | Proyecto | Problema | Herramientas | Resultado |
|---|----------|----------|--------------|-----------|
| 01 | [Normalizador de Runsheets](proyectos/01-normalizador-titulos) | Cargar cientos de runsheets con formatos distintos | Python, pandas, pdfplumber | 2 tablas normalizadas, procesa 500+ archivos en <2s |
| 02 | [Check Stubs & Revenue](proyectos/02-check-stubs-revenue) | Extraer pagos de 8 formatos de operador distintos | Python, pdfplumber, openpyxl | Reconciliación automática contra division orders |
| 03 | [Dashboard Geoespacial](proyectos/03-dashboard-geoespacial) | Localizar pozos dentro de arrendamientos | geopandas, MapLibre, Folium | Dashboard HTML interactivo con filtros cruza |
| 04 | [OCR de Documentos](proyectos/04-ocr-documentos) | Buscar texto en miles de escrituras escaneadas | PyMuPDF, Tesseract | PDFs buscables manteniendo imagen original |
| 05 | [Asana Manager](proyectos/05-asana-manager) | Montar proyectos cliente tarea por tarea a mano | Python, tkinter, Asana API | App desktop con instalador para equipos no técnicos |
| 06 | [Landman AI](proyectos/06-landman-ai) | Ir de escaneo a fila de datos capturada | Google Gemini, PyMuPDF | Prototipo de extracción a 24 columnas |
| 07 | [Availability Tracker](proyectos/07-availability-tracker) | Saber quién tiene capacidad sin perseguir a nadie | Google Apps Script, Sheets | Web app: 7 formularios, tableros, 10.400 líneas |
| 08 | [Utilidades Documentales](proyectos/08-utilidades-titulos) | Reconciliar miles de archivos en 3 convenciones | Python, pandas, plotly | 9 notebooks de propósito específico |

**Cada proyecto incluye:** README con problema → decisiones técnicas → resultados → limitaciones

---

## 🔒 Sobre los Datos

Todos estos proyectos se construyeron con datos reales de clientes (leasehold, estados de cuenta, documentos registrados, directorios de personal). **Ningún dato sensible está en este repositorio.**

✅ **Código:** Todo el código fuente está disponible  
✅ **Datos de prueba sintéticos:** Dos proyectos incluyen generadores de datos que reproducen problemas reales  
✅ **Seguridad:** Correos, nombres, rutas locales y credenciales fueron removidas. Las claves de API se leen de variables de entorno.

---

## 🛠️ Stack Tecnológico

### Backend & Procesamiento

**Python**
- `pandas` — Manipulación y análisis de datos
- `geopandas` — Datos geoespaciales y operaciones geométricas
- `openpyxl` — Lectura/escritura avanzada de Excel
- `pdfplumber` — Extracción de tablas de PDFs nativos
- `PyMuPDF` — Manipulación completa de PDFs
- `matplotlib`, `plotly` — Visualización

**Bases de Datos & BI**
- SQL avanzado (window functions, CTEs)
- Power BI — Dashboards empresariales
- Looker Studio — Dashboards web
- Excel avanzado (fórmulas, tablas dinámicas)

**Nube & Pipelines**
- AWS (S3, Lambda, RDS)
- Azure (Data Factory, Synapse)
- Procesos ETL orquestados

### Frontend & Integración

**Interfaces**
- `tkinter` — Aplicaciones desktop
- Google Apps Script — Automatización en Workspace
- HTML/CSS/JavaScript — Dashboards web

**Geoespacial**
- `folium`, `MapLibre` — Mapas interactivos
- `shapely` — Operaciones geométricas

### IA & APIs

**Modelos de Lenguaje**
- Google Gemini API — Extracción de documentos
- Tesseract OCR — Documentos escaneados

**Integración**
- APIs REST
- Google Apps Script API
- Asana API

### DevOps

- Git & GitHub
- Python packaging (`pyproject.toml`)
- Logging centralizado

---

## 📦 Instalación

```bash
# Clonar repositorio
git clone https://github.com/jsbermudezv-crypto/Portafolio.git
cd Portafolio

# Crear entorno virtual
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate

# Instalar dependencias base
pip install -e .

# Instalación con extras
pip install -e ".[geospatial]"      # Geoespacial (geopandas, folium)
pip install -e ".[ocr]"              # OCR (pytesseract, pillow)
pip install -e ".[dev]"              # Desarrollo (jupyter, black, pytest)
```

---

## 📚 Documentación

- **[STRUCTURE.md](STRUCTURE.md)** — Estructura del repositorio y convenciones
- **[docs/stack-tecnologico.md](docs/stack-tecnologico.md)** — Detalles completos del stack
- **Cada proyecto/** — README específico con problema, decisiones y resultados

---

## 🎯 Característica Clave

**De datos desordenados a soluciones listas para producción**

- ✅ Manejo de múltiples formatos de entrada
- ✅ Validación y limpieza automática
- ✅ Interfaces para equipos no técnicos
- ✅ Documentación exhaustiva de decisiones técnicas
- ✅ Código reutilizable y modular

---

## 📄 Licencia

MIT

---

**Último actualizado:** Septiembre 2026

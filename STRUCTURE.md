# Estructura del Repositorio

```
Portafolio/
├── README.md                          # Este archivo
├── STRUCTURE.md                       # Estructura del repositorio
├── pyproject.toml                     # Dependencias y configuración
├── .gitignore                         # Archivos ignorados por Git
│
├── proyectos/                         # Carpeta principal de proyectos
│   ├── 01-normalizador-titulos/
│   │   ├── README.md                 # Problema, decisiones, resultados
│   │   ├── src/
│   │   │   ├── normalizer.py
│   │   │   └── utils.py
│   │   ├── notebooks/
│   │   │   └── exploracion.ipynb
│   │   └── data/
│   │       ├── sample_runsheets/     # Datos sintéticos para demostración
│   │       └── test_data/
│   │
│   ├── 02-check-stubs-revenue/
│   │   ├── README.md
│   │   ├── src/
│   │   │   └── revenue_consolidation.py
│   │   ├── notebooks/
│   │   └── data/
│   │
│   ├── 03-dashboard-geoespacial/
│   │   ├── README.md
│   │   ├── src/
│   │   │   └── geospatial_dashboard.py
│   │   ├── notebooks/
│   │   ├── html/
│   │   │   └── dashboard_demo.html   # Demo pública
│   │   └── data/
│   │
│   ├── 04-ocr-documentos/
│   │   ├── README.md
│   │   ├── src/
│   │   │   └── ocr_processor.py
│   │   └── notebooks/
│   │
│   ├── 05-asana-manager/
│   │   ├── README.md
│   │   ├── src/
│   │   │   └── asana_app.py
│   │   └── installer/
│   │
│   ├── 06-landman-ai/
│   │   ├── README.md
│   │   ├── src/
│   │   │   └── gemini_extraction.py
│   │   └── notebooks/
│   │
│   ├── 07-availability-tracker/
│   │   ├── README.md
│   │   ├── apps_script/               # Google Apps Script
│   │   └── docs/
│   │
│   └── 08-utilidades-titulos/
│       ├── README.md
│       ├── src/
│       ├── notebooks/                 # 9 notebooks
│       └── data/
│
├── src/                               # Código compartido
│   ├── __init__.py
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── pdf_utils.py
│   │   ├── excel_utils.py
│   │   └── data_utils.py
│   └── config.py                      # Configuración centralizada
│
├── notebooks/                         # Notebooks exploratorios generales
│   └── README.md
│
├── docs/                              # Documentación general
│   ├── stack-tecnologico.md
│   ├── decisiones-arquitectura.md
│   └── deployment.md
│
├── dashboards/                        # Dashboards en Power BI, Looker, etc.
│   └── ArapahoeFinalDashboard.pbix
│
└── tests/                             # Tests unitarios
    └── test_utils.py
```

## Convenciones

- **Cada proyecto es autónomo**: Tiene su propio README, src, notebooks y datos de prueba
- **src/ en cada proyecto**: Código reutilizable y en producción
- **notebooks/**: Exploración, prototipado y documentación
- **data/**: Datos sintéticos o de prueba (datos reales nunca en el repo)
- **Sin nombres con espacios**: Todos los archivos usan guiones/guiones bajos
- **README.md en cada proyecto**: Problema → Decisiones técnicas → Resultados → Limitaciones

# Portafolio de análisis y automatización de datos



## Juan Sebastián Bermúdez Vélez



**Analista de datos e ingeniero electrónico.** Cuatro años construyendo

soluciones de datos para el mercado estadounidense, en el sector de tierras y

petróleo. Trabajo con datos que llegan mal: escaneos sin capa de texto,

runsheets con encabezados distintos en cada archivo, estados de cuenta que cada

operador emite a su manera. Construyo las herramientas que los vuelven

consultables y las entrego como algo que el equipo puede usar, no como un script

que solo corre en mi máquina.



Ingeniería Electrónica en la Universidad El Bosque, graduado con el mejor

promedio de la facultad. Español, inglés y francés.



[jsbermudezv@gmail.com](mailto:jsbermudezv@gmail.com) · [LinkedIn](https://linkedin.com/in/jsbermudezv) · [Página web](https://jsbermudezv.github.io/portafolio) · Bogotá, Colombia



## Proyectos



| Proyecto | Problema | Herramientas | Resultado |

|---|---|---|---|

| [Normalizador de runsheets y títulos](proyectos/01-normalizador-titulos) | Cargar cientos de runsheets con formatos distintos sin transcribirlos | Python, pandas, pdfplumber | Dos tablas normalizadas más reporte de campos dudosos; 2.900 líneas |

| [Check stubs y consolidación de revenue](proyectos/02-check-stubs-revenue) | Extraer pagos de regalías de ocho formatos de operador y cuadrarlos contra division orders | Python, pdfplumber, openpyxl | 763.000 filas de 5 fuentes en 3.441 tracts; 99,6% de cobertura de division order, cada dato con su nivel de confianza |

| [Dashboard geoespacial de leasehold](proyectos/03-dashboard-geoespacial) | Saber qué pozos caen dentro de qué arrendamiento | geopandas, MapLibre | Tablero HTML autocontenido con filtros cruzados · [ver la demo](https://jsbermudezv.github.io/portafolio/proyectos/03-dashboard-geoespacial/demo/) |

| [OCR por lotes de documentos](proyectos/04-ocr-documentos) | Buscar texto dentro de miles de escrituras escaneadas | PyMuPDF, Tesseract | PDF buscables con la imagen original intacta |

| [Asana Manager](proyectos/05-asana-manager) | Montar proyectos de cliente tarea por tarea a mano | Python, tkinter, API de Asana | App con instalador para un equipo no técnico |

| [Landman AI](proyectos/06-landman-ai) | Pasar de un escaneo a una fila capturada | Gemini, PyMuPDF | Prototipo de extracción a 24 columnas |

| [Availability Tracker](proyectos/07-availability-tracker) | Saber quién tiene capacidad sin perseguir a nadie | Apps Script, Sheets | Web app con 7 formularios y tableros; 10.400 líneas |

| [Utilidades de organización documental](proyectos/08-utilidades-titulos) | Reconciliar miles de archivos en tres convenciones de nombres | Python, pandas, plotly | Nueve notebooks de un propósito cada uno |



Cada carpeta tiene su README con el problema, las decisiones técnicas, los

resultados y las limitaciones.



## Sobre los datos



Todos estos proyectos se construyeron con datos de clientes: leasehold, estados

de cuenta de regalías, documentos registrados y directorios de personal. **Nada

de eso está en este repositorio.** El código sí, y para probarlo hay dos vías:



- El [normalizador](proyectos/01-normalizador-titulos) incluye un generador de tres runsheets sintéticos que reproducen los problemas reales de formato.

- El [dashboard geoespacial](proyectos/03-dashboard-geoespacial) incluye una demo pública con leases y pozos inventados, con el mismo esquema de columnas que los shapefiles originales.



Los correos corporativos, nombres de cliente, rutas locales y credenciales se

eliminaron del código. Las claves de API se leen de variables de entorno.



## Herramientas



**Python:** pandas, geopandas, openpyxl, pdfplumber, PyMuPDF, matplotlib, tkinter

**Bases de datos y BI:** SQL avanzado, Power BI, Looker Studio, Excel avanzado

**Nube y pipelines:** AWS, Azure, procesos ETL

**Otros:** Google Apps Script, Git, MapLibre, APIs REST, modelos de lenguaje aplicados






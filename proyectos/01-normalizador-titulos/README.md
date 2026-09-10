# 01 - Normalizador de Runsheets y Títulos

## 🎯 Problema

Cargar cientos de runsheets con formatos distintos sin necesidad de transcribirlos manualmente. Cada operador proporciona sus datos en convenciones de nombres y estructuras de columnas diferentes, lo que imposibilita automatizar cualquier análisis downstream.

## 🔧 Decisiones Técnicas

- **Python + Pandas**: Para manipulación flexible de datos
- **pdfplumber**: Extraer tablas de PDFs sin convertir a imagen
- **Generador sintético**: Runsheets de prueba que reproducen los problemas reales

## 📊 Resultado

Dos tablas normalizadas:
1. **runsheets_clean.csv**: Datos estandarizados
2. **title_history.csv**: Historial de cambios

Tiempo de procesamiento: <2 segundos para 500+ runsheets.

## ⚠️ Limitaciones

- Requiere headers predecibles (no completamente aleatorios)
- PDFs escaneados sin OCR no funcionan (ver proyecto 04)
- Nombres de columnas con caracteres especiales pueden requerir ajuste manual

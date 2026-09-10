# Decisiones de Arquitectura

## 1. Estructura de Proyectos

Cada proyecto es **autónomo e independiente** para permitir:
- Reutilización en clientes diferentes
- Mantenimiento y testing independiente
- Documentación específica del problema resuelto

## 2. Separación src/ vs notebooks/

| Carpeta | Propósito | Mantenibilidad |
|---------|-----------|---|
| `src/` | Código en producción, reutilizable, testeado | Alta |
| `notebooks/` | Exploración, prototipado, documentación | Media |

**Regla**: Todo código que vaya a producción debe estar en `src/`.

## 3. Datos en el Repositorio

**Principio**: Los datos de clientes NUNCA están en el repo.

**Alternativas**:
1. **Datos sintéticos**: Generadores que reproducen problemas reales
2. **Datos demo públicos**: Leases y pozos inventados con mismo esquema
3. **Test fixtures**: Ejemplos mínimos para testing

## 4. Configuración y Secretos

- Las API keys se leen de **variables de entorno**
- Archivo `.env` nunca se versiona (está en `.gitignore`)
- Configuración centralizada en `src/config.py`

```python
# ✅ Correcto
api_key = os.getenv("GEMINI_API_KEY")

# ❌ Nunca
api_key = "sk-123abc..."
```

## 5. Dependencias

**Versioning**:
- `pyproject.toml` especifica versiones mínimas
- `requirements-lock.txt` (cuando se genere) fija exactamente

**Organización**:
- Dependencias base: para todos los proyectos
- Extras `[geospatial]`, `[ocr]`, `[dev]`: instancia solo lo necesario

## 6. Logging

```python
import logging

logger = logging.getLogger(__name__)

# ✅ Correcto
logger.info("Archivo procesado: 500 registros")

# ❌ Evitar
print("Archivo procesado")  # No va a logs, difícil de filtrar
```

## 7. Testing

- Tests en `tests/` con pytest
- Funciones `test_*` automáticamente descubiertas
- Fixtures con datos sintéticos

```bash
pytest -v                    # Ejecutar todos
pytest tests/test_utils.py  # Archivo específico
```

## 8. Notebooks en Producción

Los notebooks son **exploratorios**. Para producción:

1. Extraer lógica a `src/`
2. Notebook importa desde `src/`
3. Notebook = documentación ejecutable

```python
# En notebook
from src.normalizer import RunsheetNormalizer

normalizer = RunsheetNormalizer(schema)
df_clean = normalizer.normalize("datos.xlsx")
```

## 9. Commits y Mensajes

```
# ✅ Correcto
feat: agregar soporte para runsheets de 2024
fix: corregir encoding en PDFs con caracteres especiales
refactor: simplificar lógica de normalización

# ❌ Evitar
Updated stuff
fix bug
cambios varios
```

# Notebooks Exploratorios

Esta carpeta contiene notebooks Jupyter para:

- 📊 **Exploración de datos**: Análisis inicial de nuevos datasets
- 🔬 **Prototipado**: Desarrollo de nuevas características
- 📝 **Documentación**: Explicación de procesos y decisiones

## 📌 Convención

Para notebooks que terminen en producción:

1. Desarrollar aquí de forma iterativa
2. Una vez estable, extraer la lógica a `src/`
3. El notebook se convierte en documentación (importa de `src/`)

## 🚀 Ejecución

```bash
jupyter notebook
# O para modo lab (más moderno)
jupyter lab
```

## ⚡ Tips

- Cada celda debe ejecutarse independientemente
- Documentar asunciones y decisiones en Markdown
- No guardar archivos grandes (datos reales) en el repo
- Usar `%load_ext autoreload` para reload automático de módulos

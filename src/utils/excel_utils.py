"""Utilidades para procesamiento de Excel."""

import logging
from pathlib import Path
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


def read_excel_flexible(
    excel_path: str,
    sheet_name: Optional[str] = None,
    header_row: int = 0,
    skip_rows: int = 0,
) -> pd.DataFrame:
    """
    Lee un archivo Excel con manejo flexible de formatos.
    
    Args:
        excel_path: Ruta del archivo
        sheet_name: Nombre de la hoja (None = primera hoja)
        header_row: Fila que contiene los encabezados
        skip_rows: Filas a saltar al inicio
        
    Returns:
        DataFrame con los datos
    """
    try:
        df = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            header=header_row,
            skiprows=skip_rows,
        )
        logger.info(f"Archivo Excel leído: {excel_path}")
        return df
    except Exception as e:
        logger.error(f"Error al leer {excel_path}: {e}")
        raise


def standardize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Estandariza un DataFrame (nombres de columnas, tipos).
    
    Args:
        df: DataFrame a estandarizar
        
    Returns:
        DataFrame estandarizado
    """
    # Convertir nombres a snake_case
    df.columns = (
        df.columns.str.lower()
        .str.strip()
        .str.replace(r"[^a-z0-9_]", "_", regex=True)
    )
    
    # Remover espacios en blanco
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    
    logger.info("DataFrame estandarizado")
    return df

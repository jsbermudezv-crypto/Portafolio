"""Normalizador de runsheets con formato flexible."""

import logging
from pathlib import Path
from typing import Dict, List

import pandas as pd

logger = logging.getLogger(__name__)


class RunsheetNormalizer:
    """Normaliza runsheets de formatos distintos a un esquema común."""
    
    def __init__(self, schema: Dict[str, str]):
        """
        Inicializa el normalizador.
        
        Args:
            schema: Diccionario de mapeo {columna_original: columna_estándar}
        """
        self.schema = schema
    
    def normalize(self, excel_path: str) -> pd.DataFrame:
        """
        Normaliza un archivo Excel.
        
        Args:
            excel_path: Ruta del archivo
            
        Returns:
            DataFrame normalizado
        """
        df = pd.read_excel(excel_path)
        
        # Renombrar columnas según esquema
        df = df.rename(columns=self.schema)
        
        # Validar que todas las columnas requeridas están presentes
        required_cols = list(self.schema.values())
        missing = set(required_cols) - set(df.columns)
        if missing:
            logger.warning(f"Columnas faltantes: {missing}")
        
        return df[required_cols]
    
    def normalize_batch(self, directory: str) -> pd.DataFrame:
        """
        Normaliza todos los archivos Excel en un directorio.
        
        Args:
            directory: Ruta del directorio
            
        Returns:
            DataFrame combinado con todos los datos
        """
        dfs = []
        directory = Path(directory)
        
        for excel_file in directory.glob("*.xlsx"):
            try:
                df = self.normalize(str(excel_file))
                dfs.append(df)
                logger.info(f"Normalizado: {excel_file.name}")
            except Exception as e:
                logger.error(f"Error procesando {excel_file.name}: {e}")
        
        return pd.concat(dfs, ignore_index=True)

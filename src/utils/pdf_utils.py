"""Utilidades para procesamiento de PDF."""

import logging
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)


def extract_text_from_pdf(pdf_path: str, ocr: bool = False) -> str:
    """
    Extrae texto de un PDF.
    
    Args:
        pdf_path: Ruta del archivo PDF
        ocr: Si True, aplica OCR con Tesseract para PDFs escaneados
        
    Returns:
        Texto extraído del PDF
    """
    try:
        import pymupdf
    except ImportError:
        logger.error("PyMuPDF no instalado. Instala con: pip install pymupdf")
        raise
    
    doc = pymupdf.open(pdf_path)
    text = ""
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text += page.get_text()
    
    doc.close()
    return text


def merge_pdfs(pdf_list: List[str], output_path: str) -> None:
    """
    Fusiona múltiples PDFs en uno solo.
    
    Args:
        pdf_list: Lista de rutas de PDFs
        output_path: Ruta del PDF de salida
    """
    try:
        import pymupdf
    except ImportError:
        logger.error("PyMuPDF no instalado")
        raise
    
    pdf_writer = pymupdf.open()
    
    for pdf_file in pdf_list:
        pdf_reader = pymupdf.open(pdf_file)
        pdf_writer.insert_pdf(pdf_reader)
        pdf_reader.close()
    
    pdf_writer.save(output_path)
    pdf_writer.close()
    logger.info(f"PDFs fusionados en {output_path}")

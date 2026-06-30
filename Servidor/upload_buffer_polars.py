"""Fachada de compatibilidad: reexporta la API publica del ETL."""

from etl_catalogo import upload_buffer_with_merge, upload_catalogo
from etl_common import set_logger

__all__ = ["upload_buffer_with_merge", "upload_catalogo", "set_logger"]

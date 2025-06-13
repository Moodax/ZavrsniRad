"""Excel to CSV converter with DCAT metadata generation."""

__version__ = "0.1.0"

from .core import (
    extract_tables_from_excel,
    process_excel_in_memory,
)
from .metadata import generate_dcat_metadata
from .cli import main as cli_main
from .gui import main as gui_main
from .config import setup_logging, get_config_value

__all__ = [
    "extract_tables_from_excel",
    "process_excel_in_memory",
    "generate_dcat_metadata",
    "cli_main",
    "gui_main",
    "setup_logging",
    "get_config_value",
]

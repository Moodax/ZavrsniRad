"""Configuration module for Excel to CSV converter."""

import logging
from typing import Dict, Any

# Default configuration values
DEFAULT_CONFIG: Dict[str, Any] = {
    # AI Models
    "openai_model": "gpt-3.5-turbo",
    "gemini_model": "gemini-2.0-flash-lite",

    # AI Timeouts and limits
    "ai_timeout": 60,
    "max_csv_sample_rows": 15,
    "max_datatype_sample_rows": 20,

    # Table detection parameters
    "max_empty_gap_cols": 2,
    "min_vertical_overlap_ratio": 0.8,
    "merge_tolerance_rows": 0,
    "merge_tolerance_cols": 0,

    # Default metadata values
    "default_base_uri": "http://example.org/dataset/",
    "default_publisher_uri": "http://example.org/publisher",
    "default_publisher_name": "Example Organization",
    "default_license_uri": "http://creativecommons.org/licenses/by/4.0/",

    # Logging configuration
    "log_level": "INFO",
    "log_format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
}

def setup_logging(level: str = "INFO") -> logging.Logger:
    """Set up logging configuration for the application."""
    logger = logging.getLogger("excel_to_csv_dcat")

    # Only configure if not already configured
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(DEFAULT_CONFIG["log_format"])
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    return logger

def get_config_value(key: str, default: Any = None) -> Any:
    """Get a configuration value with optional default."""
    return DEFAULT_CONFIG.get(key, default)

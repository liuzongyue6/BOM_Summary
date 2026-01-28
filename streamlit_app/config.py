"""
Configuration file for BOM Streamlit Application
"""

import os
from pathlib import Path

# Get the project root directory (parent of streamlit_app/)
PROJECT_ROOT = Path(__file__).parent.parent
TEMP_DIR = PROJECT_ROOT / "temp_processing"
DATA_DIR = PROJECT_ROOT / "data"

# File upload settings
MAX_UPLOAD_SIZE_MB = 100

# Cleanup settings
CLEANUP_HOURS = 24  # Delete temp files older than 24 hours

# Required Excel columns
REQUIRED_COLUMNS = [
    'BOM Line',
    'CAD OEM Part Number',
    'CAD OEM Rev',
    'Quantity'
]

# Comparison columns
COMPARISON_COLUMNS = [
    'CAD OEM Part Number',
    'CAD OEM Rev',
    'Material Spec',
    'CAD Oem Name',
    'Thickness'
]

# Application settings
APP_TITLE = "BOM Hierarchical Processor"
APP_ICON = "📊"
PAGE_LAYOUT = "wide"

# Ensure temp directory exists
TEMP_DIR.mkdir(exist_ok=True)

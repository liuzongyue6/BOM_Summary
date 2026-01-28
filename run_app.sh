#!/bin/bash
# BOM Processor Streamlit Application Launcher (Linux/Mac)
# Author: Zongyue Liu
# Date: 2026-01-27

echo "========================================"
echo "BOM Processor Web Application"
echo "========================================"
echo ""

# Activate virtual environment
echo "Activating virtual environment..."
source bom_streamlit/bin/activate

if [ $? -ne 0 ]; then
    echo "ERROR: Failed to activate virtual environment"
    echo "Please make sure bom_streamlit virtual environment exists"
    echo "Run: python -m venv bom_streamlit"
    exit 1
fi

echo "Virtual environment activated"
echo ""

# Check if streamlit is installed
python -c "import streamlit" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "WARNING: Streamlit not found in virtual environment"
    echo "Installing dependencies..."
    pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "ERROR: Failed to install dependencies"
        exit 1
    fi
fi

echo "Starting Streamlit application..."
echo ""
echo "Application will be available at:"
echo "  - Local:   http://localhost:8501"
echo "  - Network: http://<your-ip>:8501"
echo ""
echo "Press Ctrl+C to stop the server"
echo "========================================"
echo ""

# Start streamlit
streamlit run streamlit_app/bom_app.py --server.port=8501 --server.address=0.0.0.0

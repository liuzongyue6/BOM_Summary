@echo off
REM BOM Processor Streamlit Application Launcher (Windows)
REM Author: Zongyue Liu
REM Date: 2026-01-27

echo ========================================
echo BOM Processor Web Application
echo ========================================
echo.

REM Activate virtual environment
echo Activating virtual environment...
call bom_streamlit\Scripts\activate.bat

if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    echo Please make sure bom_streamlit virtual environment exists
    echo Run: python -m venv bom_streamlit
    pause
    exit /b 1
)

echo Virtual environment activated
echo.

REM Check if streamlit is installed
python -c "import streamlit" 2>nul
if errorlevel 1 (
    echo WARNING: Streamlit not found in virtual environment
    echo Installing dependencies...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

echo Starting Streamlit application...
echo.
echo Application will be available at:
echo   - Local:   http://localhost:8501
echo   - Network: http://<your-ip>:8501
echo.
echo Press Ctrl+C to stop the server
echo ========================================
echo.

REM Start streamlit
streamlit run streamlit_app\bom_app.py --server.port=8501 --server.address=0.0.0.0

REM If streamlit exits, pause to show any error messages
if errorlevel 1 (
    echo.
    echo ERROR: Streamlit exited with error
    pause
)

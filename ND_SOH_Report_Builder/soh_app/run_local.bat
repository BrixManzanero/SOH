@echo off
REM ===================================================================
REM  ND SOH Report Builder - local launcher
REM
REM  Double-click this file to start the app.
REM
REM  --server.address localhost binds the app to THIS PC only.
REM  Without it, Streamlit also serves on your local network, which
REM  means anyone on the same office WiFi could open your report data.
REM ===================================================================

cd /d "%~dp0"

echo Starting ND SOH Report Builder...
echo.
echo The app will open at http://localhost:8501
echo Press Ctrl+C in this window to stop it.
echo.

python -m streamlit run app.py --server.address localhost

REM If python is not on PATH, try the py launcher instead.
if errorlevel 1 (
    echo.
    echo "python" did not work. Trying "py" instead...
    py -m streamlit run app.py --server.address localhost
)

pause

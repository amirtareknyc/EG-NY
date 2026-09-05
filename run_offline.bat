@echo off
cd /d "%~dp0"
echo.
echo Starting local server...
echo Open http://localhost:8000 in your browser.
echo Press Ctrl+C to stop.
echo.
python -m http.server 8000
pause

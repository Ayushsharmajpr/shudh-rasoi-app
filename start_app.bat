@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
echo.
echo Starting Shudh Rasoi... first sign-in is admin / admin. Open http://127.0.0.1:5000 in your browser.
python app.py
pause

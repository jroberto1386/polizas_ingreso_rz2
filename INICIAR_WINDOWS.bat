@echo off
echo.
echo  ==========================================
echo   Plataforma de Polizas de Ingreso - RZ2
echo  ==========================================
echo.
echo  Verificando dependencias...
pip install flask openpyxl pandas xlrd -q
echo.
echo  Iniciando servidor local...
echo  Abre tu navegador en: http://localhost:5050
echo.
cd /d "%~dp0"
python app.py
pause

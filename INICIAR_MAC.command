#!/bin/bash
echo ""
echo " =========================================="
echo "  Plataforma de Pólizas de Ingreso - RZ2"
echo " =========================================="
echo ""
echo " Verificando dependencias..."
pip install flask openpyxl pandas xlrd -q
echo ""
echo " Iniciando servidor local..."
echo " Abre tu navegador en: http://localhost:5050"
echo ""
cd "$(dirname "$0")"
python3 app.py

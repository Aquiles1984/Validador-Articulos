@echo off
chcp 65001 > nul
title Dimac - Gestión de Artículos
echo Iniciando servidor...
start "" http://localhost:5000
python "%~dp0app.py"
pause

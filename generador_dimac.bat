@echo off
chcp 65001 > nul
title GENERADOR ARTICULOS DIMAC
python "%~dp0generador_dimac.py"
pause

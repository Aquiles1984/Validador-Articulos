@echo off
REM Valida las ALTAS de accesorios creadas desde una fecha, leyendo Geinfor en directo (DB2)
REM y el Excel de Jose Maria de la red.
cd /d "%~dp0"
set /p DESDE="Validar altas creadas DESDE (formato AAAA-MM-DD): "
echo.
python validador_articulos_dimac.py --db %DESDE%
echo.
echo ---------------------------------------------
echo Se abre el informe visual (INFORME_ALTAS.html) en el navegador.
echo Tambien tienes la tabla completa en INCIDENCIAS_FINAL.xlsx.
echo (Requiere estar en la red de la oficina.)
pause

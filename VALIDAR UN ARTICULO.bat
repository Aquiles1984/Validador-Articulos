@echo off
REM Valida UN articulo leyendo Geinfor en directo (DB2) y el Excel de Jose Maria de la red.
cd /d "%~dp0"
set /p COD="Codigo del articulo a validar: "
echo.
python validador_articulos_dimac.py --db --articulo %COD%
echo.
echo ---------------------------------------------
echo Se abre el informe visual en el navegador.
echo (Requiere estar en la red de la oficina.)
pause

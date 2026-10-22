@echo off
REM Build DON_ARTICULO.exe (run on Windows)
REM 1) Install Python 3.10+ and pip
REM 2) Open CMD in this folder and run this file

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

REM Build (usa compile_exe.py: incluye solo los ficheros de datos que existan)
python compile_exe.py

echo.
echo LISTO. Tu exe esta en esta misma carpeta: "DON ARTICULO.exe"
pause

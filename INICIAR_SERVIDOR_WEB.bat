@echo off
cd /d "%~dp0"
set "PORT=5050"
set "MODO_RUN=0"
if /I "%~1"=="run" set "MODO_RUN=1"

rem ============================================================
rem  DON ARTICULO WEB - LANZADOR
rem  - Si el puerto 5050 ya esta en escucha, NO abre otra ventana
rem    (no interrumpe una importacion en curso).
rem  - Con el argumento "run" arranca el servidor sin cerrar nada
rem    (pensado para un acceso directo de inicio o un vigilante).
rem  NOTA: usa Selenium/Firefox (via ROBOT.py) -> necesita una
rem  sesion de Windows iniciada. No puede correr "sin usuario".
rem  NOTA TECNICA: solo IF/GOTO de una linea, sin bloques (...)
rem  multilinea (rompen el parseo de cmd y cierran la ventana
rem  sin aviso).
rem ============================================================

call :ESTA_ESCUCHANDO

if "%ESCUCHANDO%"=="1" goto :YA_CORRIENDO
if "%MODO_RUN%"=="1" goto :ARRANCAR
goto :ARRANCAR_MANUAL

:YA_CORRIENDO
if "%MODO_RUN%"=="1" goto :FIN
echo ============================================================
echo   DON ARTICULO WEB ya esta en marcha - no se abre otra ventana.
echo   Accede desde el navegador: http://localhost:%PORT%
echo ============================================================
pause
goto :FIN

:ARRANCAR_MANUAL
echo Cerrando otros servidores en puerto %PORT%...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":%PORT%" ^| findstr "LISTENING"') do taskkill /F /PID %%a >nul 2>&1

:ARRANCAR
title DON ARTICULO WEB SERVER - EN MARCHA (no cerrar esta ventana)
echo ============================================================
echo   DON ARTICULO - Servidor Web
echo   DEJA ESTA VENTANA ABIERTA: es el servidor funcionando.
echo   Desde este equipo: http://localhost:%PORT%  (desde la red: http://IP-DE-ESTE-EQUIPO:%PORT%)
echo ============================================================
echo.
python DON_ARTICULO_WEB.py
if "%MODO_RUN%"=="1" goto :FIN
echo.
echo El servidor se ha detenido.
pause
goto :FIN

rem ------------------------------------------------------------
:ESTA_ESCUCHANDO
set "ESCUCHANDO=0"
netstat -ano | findstr ":%PORT%" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 set "ESCUCHANDO=1"
exit /b 0

:FIN
exit /b 0

@echo off
title DON ARTICULO WEB SERVER
echo ============================================
echo   DON ARTICULO - Servidor Web
echo   El servidor detectara su IP automaticamente.
echo ============================================
echo.
echo Cerrando otros servidores en puerto 5050...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":5050" ^| findstr "LISTENING"') do taskkill /F /PID %%a >nul 2>&1
echo.
cd /d "%~dp0"
python DON_ARTICULO_WEB.py
pause

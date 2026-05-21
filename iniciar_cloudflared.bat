@echo off
REM Script para iniciar Cloudflare Tunnel y Flask en paralelo

echo ======================================
echo Iniciando Extractor RIPS con Cloudflare Tunnel
echo ======================================

REM Iniciar Flask en una ventana separada
echo Iniciando Flask en puerto 5000...
start "Flask App" cmd /k python app_medicamentos_control.py

REM Esperar 3 segundos para que Flask inicie
timeout /t 3 /nobreak

REM Iniciar Cloudflare Tunnel
echo.
echo Iniciando Cloudflare Tunnel...
echo Ejecuta este comando manualmente si prefieres:
echo   cloudflared tunnel run extractor-rips
echo.
cloudflared tunnel run extractor-rips

pause

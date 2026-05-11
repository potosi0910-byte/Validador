@echo off
echo ================================================
echo  Instalando dependencias - Malla Validadora RIPS
echo ================================================

echo.
echo [1/2] Instalando dependencias Python (FastAPI, uvicorn)...
pip install fastapi uvicorn[standard] python-multipart

echo.
echo [2/2] Instalando dependencias Node.js (React, Vite)...
cd frontend
npm install
cd ..

echo.
echo ================================================
echo  Instalacion completada.
echo  Ejecute: iniciar.bat
echo ================================================
pause

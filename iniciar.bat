@echo off
echo ================================================
echo  Malla Validadora RIPS 2275/2023
echo ================================================

echo.
echo Iniciando backend (FastAPI en puerto 8000)...
start "Backend RIPS" cmd /k "python api.py"

timeout /t 3 /nobreak >nul

echo Iniciando frontend (React en puerto 5173)...
start "Frontend RIPS" cmd /k "set PATH=%PATH%;C:\Program Files\nodejs && cd frontend && npm run dev"

timeout /t 4 /nobreak >nul

echo Abriendo navegador...
start http://localhost:5173

echo.
echo  Backend : http://localhost:8000
echo  Frontend: http://localhost:5173
echo  API docs: http://localhost:8000/docs
echo.
echo Cierre las ventanas de consola para detener el sistema.

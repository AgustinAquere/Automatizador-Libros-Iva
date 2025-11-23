@echo off
title Procesador de Libros IVA
color 0A

echo ============================================================
echo    PROCESADOR DE LIBROS IVA
echo    Iniciando aplicacion...
echo ============================================================
echo.

REM Cambiar al directorio del script
cd /d "%~dp0"

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH
    echo Por favor instale Python desde https://www.python.org/
    pause
    exit /b 1
)

echo [OK] Python encontrado
echo.

REM Verificar si existe fastapi_app.py
if not exist "fastapi_app.py" (
    echo [ERROR] No se encontro fastapi_app.py
    echo Asegurese de ejecutar este archivo desde la carpeta del proyecto
    pause
    exit /b 1
)

echo [OK] Archivos encontrados
echo.

REM Iniciar el servidor en segundo plano y capturar el PID
echo Iniciando servidor FastAPI...
start /B python fastapi_app.py

REM Esperar 3 segundos para que el servidor arranque
timeout /t 3 /nobreak >nul

REM Abrir navegador
echo Abriendo navegador...
start http://localhost:8000

echo.
echo ============================================================
echo    APLICACION INICIADA
echo    URL: http://localhost:8000
echo ============================================================
echo.
echo Presione Ctrl+C para detener el servidor
echo O simplemente cierre esta ventana
echo.

REM Mantener la ventana abierta y esperar
python fastapi_app.py

pause

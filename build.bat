@echo off
setlocal EnableDelayedExpansion

echo ============================================
echo  AnalizadorSOR — Instalador / Build
echo ============================================
echo.

:: ── 1. Verificar Python ──────────────────────
echo [1/4] Verificando Python...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Python no encontrado. Iniciando instalacion automatica...
    echo.

    set "PY_VER=3.12.9"
    set "PY_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
    set "PY_INST=%TEMP%\python_install.exe"

    echo Descargando Python 3.12.9 desde python.org...
    powershell -NoProfile -Command "Invoke-WebRequest -Uri '!PY_URL!' -OutFile '!PY_INST!' -UseBasicParsing"
    if !ERRORLEVEL! NEQ 0 (
        echo.
        echo ERROR: No se pudo descargar Python.
        echo Verifique su conexion a internet e intente nuevamente.
        pause & exit /b 1
    )

    echo Instalando Python silenciosamente...
    "!PY_INST!" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1
    del "!PY_INST!" >nul 2>&1

    echo.
    echo Python instalado correctamente.
    echo IMPORTANTE: Abra una nueva ventana CMD y ejecute build.bat nuevamente
    echo para que Windows reconozca Python en el PATH.
    echo.
    pause
    exit /b 0
) else (
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Encontrado: %%v
)

:: ── 2. Instalar dependencias ─────────────────
echo.
echo [2/4] Instalando dependencias Python...
python -m pip install --upgrade pip --quiet
python -m pip install openpyxl tkinterdnd2 pyinstaller
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause & exit /b 1
)

:: ── 3. Generar ejecutable ────────────────────
echo.
echo [3/4] Generando ejecutable AnalizadorSOR.exe...
python -m PyInstaller --onefile --windowed --name "AnalizadorSOR" ^
    --collect-all tkinterdnd2 ^
    main.py
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Fallo la compilacion. Revise los mensajes anteriores.
    pause & exit /b 1
)

:: ── 4. Completado ────────────────────────────
echo.
echo ============================================
echo  [4/4] Compilacion completada exitosamente!
echo ============================================
echo.
echo Ejecutable generado en:
echo %~dp0dist\AnalizadorSOR.exe
echo.
explorer "%~dp0dist"
pause

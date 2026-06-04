@echo off
setlocal EnableDelayedExpansion

echo ============================================
echo  AnalizadorSOR — Instalador / Build
echo ============================================
echo.
echo Uso:
echo   build.bat           ^<-- compilar para este Windows
echo   build.bat win7      ^<-- compilar exe compatible con Windows 7
echo.

:: ── Modo de compatibilidad ────────────────────
set "TARGET_WIN7=0"
if /i "%~1"=="win7" set "TARGET_WIN7=1"

:: ── Detectar version de Windows ──────────────
for /f "tokens=4 delims=[] " %%v in ('ver') do set "WINVER=%%v"
for /f "tokens=1 delims=." %%a in ("!WINVER!") do set "WIN_MAJOR=%%a"
for /f "tokens=2 delims=." %%b in ("!WINVER!") do set "WIN_MINOR=%%b"

set "IS_WIN7=0"
if !WIN_MAJOR! EQU 6 if !WIN_MINOR! EQU 1 set "IS_WIN7=1"
if !WIN_MAJOR! EQU 6 if !WIN_MINOR! EQU 0 set "IS_WIN7=1"

:: Forzar modo Win7 si se pide explicitamente o si el SO es Win7
if "!IS_WIN7!"=="1" set "TARGET_WIN7=1"

if "!TARGET_WIN7!"=="1" (
    echo Modo: compatible con Windows 7  ^(Python 3.7^)
    set "PY_VER=3.7.9"
    if "%PROCESSOR_ARCHITECTURE%"=="x86" (
        set "PY_URL=https://www.python.org/ftp/python/3.7.9/python-3.7.9.exe"
    ) else (
        set "PY_URL=https://www.python.org/ftp/python/3.7.9/python-3.7.9-amd64.exe"
    )
    set "PIP_PYINSTALLER=pyinstaller==5.13.2"
    set "PY_CMD=py -3.7"
    set "EXE_NAME=AnalizadorSOR_Win7"
) else (
    echo Modo: Windows 10 / 11  ^(Python 3.12^)
    set "PY_VER=3.12.9"
    if "%PROCESSOR_ARCHITECTURE%"=="x86" (
        set "PY_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9.exe"
    ) else (
        set "PY_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
    )
    set "PIP_PYINSTALLER=pyinstaller"
    set "PY_CMD=python"
    set "EXE_NAME=AnalizadorSOR"
)

:: ── 1. Verificar Python ──────────────────────
echo.
echo [1/4] Verificando Python...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Python no encontrado. Iniciando instalacion automatica...
    echo.

    set "PY_INST=%TEMP%\python_install.exe"

    echo Descargando Python !PY_VER!...
    powershell -NoProfile -Command "Invoke-WebRequest -Uri '!PY_URL!' -OutFile '!PY_INST!' -UseBasicParsing" >nul 2>&1
    if !ERRORLEVEL! NEQ 0 (
        echo   PowerShell no disponible, usando certutil...
        certutil -urlcache -split -f "!PY_URL!" "!PY_INST!" >nul
    )
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
    echo Python !PY_VER! instalado correctamente.
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
%PY_CMD% -m pip install --upgrade pip --quiet
%PY_CMD% -m pip install openpyxl tkinterdnd2 %PIP_PYINSTALLER%
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause & exit /b 1
)

:: ── 3. Generar ejecutable ────────────────────
echo.
echo [3/4] Generando ejecutable %EXE_NAME%.exe...

%PY_CMD% -m PyInstaller --onefile --windowed --name "%EXE_NAME%" ^
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
if "!TARGET_WIN7!"=="1" (
    echo NOTA: Este ejecutable es compatible con Windows 7.
    echo       No requiere instalar Python en la maquina destino.
    echo.
)
echo Ejecutable generado en:
echo %~dp0dist\%EXE_NAME%.exe
echo.
explorer "%~dp0dist"
pause

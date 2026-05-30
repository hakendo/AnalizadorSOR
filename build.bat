@echo off
echo Instalando dependencias...
pip install openpyxl tkinterdnd2 pyinstaller

echo.
echo Generando ejecutable...
python -m PyInstaller --onefile --windowed --name "AnalizadorSOR" ^
    --collect-all tkinterdnd2 ^
    main.py

echo.
echo Ejecutable generado en:
echo %~dp0dist\AnalizadorSOR.exe
echo.
explorer %~dp0dist
pause

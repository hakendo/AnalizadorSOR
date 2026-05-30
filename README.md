# Analizador SOR — Fibra Óptica

Aplicación de escritorio para Windows que extrae métricas de archivos OTDR (`.sor`) y las exporta a Excel.

## ¿Qué hace?

Lee archivos de traza OTDR en formato Bellcore SR-4731 (generados por equipos EXFO) y por cada filamento extrae:

| Métrica | Descripción |
|---|---|
| Posición (km) | Ubicación del evento/empalme en la fibra |
| Longitud del intervalo (km) | Distancia entre eventos consecutivos |
| Pérdida del intervalo (dB) | Pérdida en cada tramo |
| Pérdida promedio (dB/km) | Coeficiente de atenuación del tramo |
| Pérdida de unión (dB) | Pérdida puntual en el empalme |
| Pérdida de unión promedio (dB) | Promedio de todos los empalmes de la fibra |
| Pérdida de unión máxima (dB) | Empalme con mayor pérdida en la fibra |

El resultado se exporta a un archivo Excel estandarizado:
```
FO_Cartilla_FOS_YYYY-MM.xlsx
```

## Estructura esperada de carpetas

```
carpeta-raíz/
├── nombre-cable-1/
│   ├── fibra1 nombre-cable-1.sor
│   ├── fibra2 nombre-cable-1.sor
│   └── ...
└── nombre-cable-2/
    ├── fibra1 nombre-cable-2.sor
    └── ...
```

> Solo se procesan los archivos SOR **sin sufijo** (bidireccionales). Los archivos con sufijo `corta` o `larga` se ignoran automáticamente.

## Instalación

**Requisitos:** Python 3.10+ con pip

```bash
pip install openpyxl
```

## Uso

### Ejecutar la aplicación

```bash
python main.py
```

1. Selecciona la **carpeta raíz** que contiene las subcarpetas de cada cable
2. Presiona **Analizar archivos** — procesa todos los `.sor` con barra de progreso
3. Presiona **Exportar Excel** — genera el archivo y lo abre automáticamente

### Generar ejecutable `.exe` para Windows

```bash
build.bat
```

El ejecutable queda en `dist\AnalizadorSOR.exe`.

## Estructura del proyecto

```
sor_analyzer/
├── main.py             # GUI (tkinter)
├── sor_parser.py       # Parser binario Bellcore SR-4731 / EXFO
├── excel_exporter.py   # Generador Excel (openpyxl)
├── requirements.txt
└── build.bat           # PyInstaller → .exe
```

## Compatibilidad

- Equipos OTDR: **EXFO FTBx** (probado con FTBx-735C-SM1-EA)
- Formato: Bellcore SR-4731 rev 2.0
- Python: 3.10+
- OS: Windows (GUI), Linux/macOS (solo parseo/exportación)

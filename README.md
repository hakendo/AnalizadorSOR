# Analizador SOR — Fibra Óptica

Aplicación de escritorio para Windows que extrae métricas de archivos OTDR (`.sor`) y las exporta a Excel.

## Captura de pantalla

![Analizador SOR](assets/screenshot.png)

> **Para contribuidores:** reemplaza `assets/screenshot.png` con una captura real de la aplicación corriendo en Windows.

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

## Mejoras propuestas

### Alta prioridad

- **Soporte multidireccional** — Procesar también los archivos `corta` y `larga` de cada fibra y mostrarlos como columnas separadas en el Excel, permitiendo comparar ambas direcciones de medición.
- **Validación de umbrales** — Marcar en rojo en el Excel los empalmes que superen un umbral configurable (ej. pérdida > 0.5 dB), facilitando identificar problemas sin revisar manualmente cada valor.
- **Vista previa en la app** — Mostrar una tabla con los datos parseados dentro de la misma ventana antes de exportar, para verificar que los datos son correctos.

### Media prioridad

- **Compatibilidad con más equipos OTDR** — Actualmente probado solo con EXFO FTBx. Agregar soporte para Anritsu, VIAVI (JDSU), Yokogawa y AFL, que usan variantes del mismo formato Bellcore SR-4731.
- **Exportar a PDF** — Generar un informe PDF con formato de cartilla, listo para entregar sin necesitar Excel.
- **Gráfico de la traza** — Mostrar la curva de atenuación OTDR (dB vs. km) usando `matplotlib`, con los eventos marcados, lo que permite detectar anomalías visualmente.
- **Soporte doble longitud de onda** — Algunos equipos miden a 1310 nm y 1550 nm simultáneamente. Separar ambos en hojas distintas del Excel.

### Baja prioridad

- **Drag & drop** — Permitir arrastrar la carpeta directamente a la ventana en lugar de usar el selector de carpetas.
- **Historial de mediciones** — Comparar la medición actual con una anterior para detectar degradación de empalmes en el tiempo.
- **Filtro por fibra** — Seleccionar qué fibras incluir en la exportación (útil cuando solo algunas fibras tienen datos nuevos).
- **Configuración persistente** — Guardar la última carpeta usada y preferencias en un archivo `.json` para no tener que reconfigurar cada vez.

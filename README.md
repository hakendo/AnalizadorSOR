# Analizador SOR — Fibra Óptica

Aplicación de escritorio para Windows que lee mediciones OTDR (`.sor` y `.trc`) y genera la cartilla de filamentos en Excel.

## Captura de pantalla

![Analizador SOR](assets/screenshot.png)

> **Para contribuidores:** reemplaza `assets/screenshot.png` con una captura real de la aplicación corriendo en Windows.

## Formatos soportados

| Extensión | Formato | Origen |
|---|---|---|
| `.sor` | Bellcore SR-4731 rev 2.0 | Exportación estándar de EXFO y otros equipos |
| `.trc` | Nativo EXFO (`AppReg Format Ex`) | EXFO FTBx / software Metrino |

El formato se detecta por el **contenido** del archivo, no por la extensión. Si una misma fibra existe como `.sor` y `.trc`, se usa el `.trc`.

## ¿Qué extrae?

### Detalle por evento (una hoja por cable)

Cada fibra empieza con la fila de **inicio**, sigue con una fila por **empalme** y termina con una fila **Fin** (fin de fibra) que muestra el último tramo.

| Columna | Descripción |
|---|---|
| N° Evento | Número del evento (`Fin` para el fin de fibra) |
| Posición (km) | Ubicación del evento en la fibra |
| Long. Intervalo (km) | Distancia desde el evento anterior |
| Pérd. Intervalo (dB) | Pérdida del tramo de fibra que termina en el evento |
| Pérd. Prom. (dB/km) | Atenuación de ese tramo |
| Pérd. Unión (dB) | Pérdida puntual del empalme (vacía en la fila Fin) |
| Pérd. Unión Prom. / Máx. (dB) | Promedio y máximo de los empalmes de la fibra |

### Resumen por fibra (hoja «Resumen»)

| Columna | Descripción |
|---|---|
| Long. Total (km) | Posición del fin de fibra |
| Pérd. Total (dB) | Pérdida del enlace completo: todos los tramos (incluido el último) + todos los empalmes. Coincide con el *Span Loss* del equipo |
| Pérd. Prom. (dB/km) | Pérdida total ÷ longitud total |
| Pérd. Unión Prom. / Máx. (dB) | Estadística de empalmes |
| N° Empalmes | Cantidad de empalmes (sin contar inicio ni fin) |
| ∆ Pérd. Unión Máx. | Diferencia contra la medición anterior del historial |
| Fecha | Fecha de la medición |

La suma de pérdidas de intervalo y de unión del detalle es igual a la pérdida total del resumen.

El Excel se guarda en la carpeta raíz como:
```
FO_Cartilla_FOS_YYYY-MM.xlsx
```

## Estructura esperada de carpetas

Una subcarpeta por cable, con un archivo por fibra:

```
carpeta-raíz/
├── nombre-cable-1/
│   ├── fibra 1 nombre-cable-1.trc
│   ├── fibra 2 nombre-cable-1.trc
│   └── ...
└── nombre-cable-2/
    ├── filamento 1 nombre-cable-2.sor
    └── ...
```

- El número de fibra se toma del nombre del archivo: `fibra 3`, `filamento 3`, `Fiber3`, `hilo 3` (sin importar mayúsculas).
- Archivos con `corta` o `larga` en el nombre se tratan como mediciones en esa dirección; el resto como **bidireccional (normal)**. Qué direcciones se procesan se elige en la app.
- Los archivos sin número de fibra reconocible se ignoran.

## Funciones de la aplicación

- **Selección de carpeta** con botón o arrastrando la carpeta a la ventana (requiere `tkinterdnd2`).
- **Direcciones**: procesar normal, corta y/o larga.
- **Filtro por fibra**: elegir qué fibras de cada cable se incluyen.
- **Vista previa** de los datos antes de exportar, con colores según umbrales.
- **Umbrales configurables** (pérdida de unión, atenuación, pérdida de intervalo): las celdas que los superan se marcan en naranja (> 80 %) o rojo.
- **Columnas del Excel** configurables (botón «Columnas Excel»).
- **Historial**: cada exportación guarda un resumen en `mediciones_historial.json` y el Excel muestra la variación de la pérdida de unión máxima respecto a la medición anterior.
- **Configuración persistente** en `config.json` (última carpeta, umbrales, columnas, filtros).

## Instalación

**Requisitos:** Python 3.10+ (o 3.7 para la build de Windows 7)

```bash
pip install -r requirements.txt
```

`tkinterdnd2` es opcional: solo habilita el drag & drop.

## Uso

```bash
python main.py
```

1. Selecciona (o arrastra) la **carpeta raíz** que contiene las subcarpetas de cada cable
2. Revisa los cables detectados y, si hace falta, usa **Filtrar fibras**
3. Presiona **Analizar archivos** — procesa los `.sor` / `.trc` con barra de progreso y muestra la vista previa
4. Presiona **Exportar Excel** — genera el archivo y lo abre automáticamente

## Generar ejecutable `.exe`

```bat
build.bat          :: Windows 10 / 11  → dist\AnalizadorSOR.exe
build.bat win7     :: compatible con Windows 7 (Python 3.7) → dist\AnalizadorSOR_Win7.exe
```

El script instala Python y PyInstaller si no están presentes.

## Estructura del proyecto

```
sor_analyzer/
├── main.py             # GUI (tkinter)
├── sor_parser.py       # Parser Bellcore SR-4731, armado de resultados y escaneo de carpetas
├── trc_parser.py       # Lector del formato nativo EXFO .trc
├── excel_exporter.py   # Generador Excel (openpyxl)
├── history.py          # Historial de mediciones (JSON)
├── config.py           # Configuración persistente y perfiles de equipo
├── requirements.txt
└── build.bat           # PyInstaller → .exe
```

## Compatibilidad

- Equipos OTDR: **EXFO FTBx** (probado con FTBx-735C-SM1-EA)
- Python: 3.10+ (3.7 para la build de Windows 7)
- OS: Windows (GUI); el parseo y la exportación también funcionan en Linux/macOS

## Historial de cambios relevantes

- **Soporte `.trc`** (formato nativo EXFO).
- **Corrección de distancias en `.sor`**: versiones anteriores calculaban posiciones y longitudes ~21 veces más cortas (lectura errónea del índice de grupo y de la unidad de tiempo). Los Excel y el historial generados antes de esta corrección tienen longitudes y pérdidas por tramo incorrectas; las pérdidas de unión sí eran correctas.
- **Pérdida total** ahora incluye el último tramo y los empalmes (antes los omitía).
- **N° Empalmes** ya no cuenta el evento de inicio.

## Mejoras propuestas

- **Más equipos OTDR** — Probar y ajustar con Anritsu, VIAVI (JDSU), Yokogawa y AFL.
- **Exportar a PDF** — Informe con formato de cartilla, listo para entregar sin Excel.
- **Gráfico de la traza** — Curva OTDR (dB vs. km) con los eventos marcados, a partir de los datos de muestra que ya traen los archivos.
- **Doble longitud de onda** — Separar mediciones 1310 / 1550 nm en hojas distintas.
- **Comparación de direcciones** — Mostrar corta y larga lado a lado y calcular el promedio bidireccional de cada empalme.

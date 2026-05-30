from __future__ import annotations
import os
from datetime import date
from typing import Any

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

import history as hist_mod

# ── Palette ───────────────────────────────────────────────────────────────
_BLUE_HDR   = "1F4E79"
_BLUE_SUB   = "2E75B6"
_BLUE_LIGHT = "BDD7EE"
_GRAY       = "F2F2F2"
_WHITE      = "FFFFFF"
_RED_WARN   = "FF0000"
_ORANGE_WARN = "FF9900"
_RED_LIGHT  = "FFD7D7"
_ORANGE_LIGHT = "FFE8CC"
_GREEN_DELTA = "E2EFDA"
_RED_DELTA   = "FCE4D6"

_DIRECTION_LABELS = {'normal': 'Bidireccional', 'corta': 'Corta', 'larga': 'Larga'}


def _side() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _cell(ws, row: int, col: int, value: Any = None,
          bold: bool = False, bg: str | None = None,
          align: str = "center", fmt: str | None = None,
          font_color: str = "000000") -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name="Calibri", bold=bold, size=10, color=font_color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        cell.number_format = fmt
    cell.border = _side()


def _threshold_bg(value: float | None, threshold: float | None,
                  warn_pct: float = 0.8) -> str | None:
    """Return red/orange cell fill if value exceeds threshold."""
    if value is None or threshold is None:
        return None
    if value > threshold:
        return _RED_LIGHT
    if value > threshold * warn_pct:
        return _ORANGE_LIGHT
    return None


# ── Column definitions ────────────────────────────────────────────────────
_DETAIL_COLS = [
    ("Fibra N°",               9,  None,     None),
    ("Dirección",              13, None,     None),
    ("N° Evento",              9,  None,     None),
    ("Posición (km)",          13, "0.0000", None),
    ("Long. Intervalo (km)",   16, "0.0000", None),
    ("Pérd. Intervalo (dB)",   16, "0.0000", 'perdida_intervalo_db'),
    ("Pérd. Prom. (dB/km)",    16, "0.0000", 'perdida_promedio_dbkm'),
    ("Pérd. Unión (dB)",       14, "0.0000", 'perdida_union_db'),
    ("Pérd. Unión Prom. (dB)", 18, "0.0000", None),
    ("Pérd. Unión Máx. (dB)",  18, "0.0000", 'perdida_union_maxima_db'),
]

_SUMMARY_COLS = [
    ("Fibra N°",               9,  None),
    ("Dirección",              13, None),
    ("Long. Total (km)",       14, "0.0000"),
    ("Pérd. Total (dB)",       14, "0.0000"),
    ("Pérd. Prom. (dB/km)",    16, "0.0000"),
    ("Pérd. Unión Prom. (dB)", 18, "0.0000"),
    ("Pérd. Unión Máx. (dB)",  18, "0.0000"),
    ("N° Empalmes",            12, None),
    ("∆ Pérd. Unión Máx.",     16, "+0.0000"),
    ("Fecha",                  12, None),
]


# ── Sheet builders ────────────────────────────────────────────────────────

def _write_cable_sheet(wb: "Workbook", cable_name: str,
                       fibers: list[dict], thresholds: dict,
                       prev_snapshot: dict) -> None:
    clean = cable_name.strip()[:31]
    ws = wb.create_sheet(title=clean)
    ws.freeze_panes = "A7"

    ncols = len(_DETAIL_COLS)
    first = fibers[0] if fibers else {}

    # Header block
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    _cell(ws, 1, 1,
          value=f"Medición de Filamentos — Cable: {clean}",
          bold=True, bg=_BLUE_HDR, align="left", font_color="FFFFFF")
    ws.row_dimensions[1].height = 18

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    _cell(ws, 2, 1, value=f"Fecha: {first.get('date', '')}  |  λ: {first.get('wavelength_nm', '')} nm  |  Equipo: {first.get('otdr_model','EXFO') or 'EXFO'}",
          align="left", bg=_BLUE_LIGHT)
    for c in range(2, ncols + 1):
        if c > 4:
            _cell(ws, 2, c, bg=_BLUE_LIGHT)

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=ncols)
    thr = thresholds
    _cell(ws, 3, 1,
          value=(f"Umbrales:  Pérd. Unión > {thr.get('perdida_union_db', 0.5)} dB  ·  "
                 f"Atenuación > {thr.get('perdida_promedio_dbkm', 0.25)} dB/km  ·  "
                 f"Pérd. Intervalo > {thr.get('perdida_intervalo_db', 2.0)} dB"),
          align="left", bg=_GRAY)
    ws.row_dimensions[4].height = 6

    # Column headers row 6
    for ci, (label, width, _, _thr_key) in enumerate(_DETAIL_COLS, 1):
        _cell(ws, 6, ci, value=label, bold=True, bg=_BLUE_SUB, font_color="FFFFFF")
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[6].height = 30

    # Data
    row = 7
    for fiber in fibers:
        fibra_num  = fiber.get('fibra_num')
        direction  = fiber.get('direction', 'normal')
        dir_label  = _DIRECTION_LABELS.get(direction, direction)
        events     = fiber.get('events', [])
        union_prom = fiber.get('perdida_union_promedio_db')
        union_max  = fiber.get('perdida_union_maxima_db')
        if not events:
            continue

        start_row = row
        for ei, ev in enumerate(events):
            alt = fibra_num % 2 == 0
            row_bg = _GRAY if alt else _WHITE

            _cell(ws, row, 1, fibra_num if ei == 0 else None, bold=(ei == 0), bg=row_bg)
            _cell(ws, row, 2, dir_label if ei == 0 else None, bg=row_bg, align="left")
            _cell(ws, row, 3, ev.get('n_evento'), bg=row_bg)
            _cell(ws, row, 4, ev.get('posicion_km'), bg=row_bg, fmt="0.0000")

            interval = ev.get('longitud_intervalo_km')
            _cell(ws, row, 5, interval, bg=row_bg, fmt="0.0000")

            pi_val = ev.get('perdida_intervalo_db')
            pi_bg  = _threshold_bg(pi_val, thr.get('perdida_intervalo_db')) or row_bg
            _cell(ws, row, 6, pi_val, bg=pi_bg, fmt="0.0000")

            pp_val = ev.get('perdida_promedio_dbkm')
            pp_bg  = _threshold_bg(pp_val, thr.get('perdida_promedio_dbkm')) or row_bg
            _cell(ws, row, 7, pp_val, bg=pp_bg, fmt="0.0000")

            pu_val = ev.get('perdida_union_db')
            pu_bg  = _threshold_bg(pu_val, thr.get('perdida_union_db')) or row_bg
            _cell(ws, row, 8, pu_val, bg=pu_bg, fmt="0.0000")

            _cell(ws, row, 9,  union_prom if ei == 0 else None, bg=_BLUE_LIGHT, fmt="0.0000")

            umax_bg = _threshold_bg(union_max, thr.get('perdida_union_db')) or _BLUE_LIGHT
            _cell(ws, row, 10, union_max if ei == 0 else None, bg=umax_bg, fmt="0.0000")

            row += 1

        if len(events) > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1,   end_column=1)


def _write_summary_sheet(wb: "Workbook", all_cables: dict[str, list[dict]],
                         thresholds: dict, previous: dict) -> None:
    ws = wb.create_sheet(title="Resumen")
    ws.freeze_panes = "A3"
    ncols = len(_SUMMARY_COLS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    _cell(ws, 1, 1, value="Resumen General de Mediciones",
          bold=True, bg=_BLUE_HDR, font_color="FFFFFF", align="left")
    ws.row_dimensions[1].height = 18

    for ci, (label, width, _) in enumerate(_SUMMARY_COLS, 1):
        _cell(ws, 2, ci, value=label, bold=True, bg=_BLUE_SUB, font_color="FFFFFF")
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[2].height = 25

    row = 3
    thr = thresholds
    for cable_name, fibers in all_cables.items():
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
        _cell(ws, row, 1, value=f"  Cable: {cable_name.strip()}",
              bold=True, bg=_BLUE_LIGHT, align="left")
        row += 1

        prev_cable = previous.get(cable_name, {})

        for fiber in fibers:
            fibra_num  = fiber.get('fibra_num')
            direction  = fiber.get('direction', 'normal')
            dir_label  = _DIRECTION_LABELS.get(direction, direction)
            n_empalmes = len([e for e in fiber.get('events', [])])
            alt_bg = _GRAY if fibra_num % 2 == 0 else _WHITE

            prev_key  = f"{fibra_num}_{direction}"
            prev_fiber = prev_cable.get(prev_key, {})
            delta_max = hist_mod.delta(
                fiber.get('perdida_union_maxima_db'),
                prev_fiber.get('perdida_union_maxima_db')
            )

            umax = fiber.get('perdida_union_maxima_db')
            umax_bg = _threshold_bg(umax, thr.get('perdida_union_db')) or alt_bg
            pprom_bg = _threshold_bg(fiber.get('perdida_promedio_dbkm'),
                                     thr.get('perdida_promedio_dbkm')) or alt_bg

            delta_bg = alt_bg
            if delta_max is not None:
                delta_bg = _RED_DELTA if delta_max > 0 else (_GREEN_DELTA if delta_max < 0 else alt_bg)

            vals = [
                (fibra_num,                              alt_bg,   None),
                (dir_label,                              alt_bg,   None),
                (fiber.get('longitud_total_km'),         alt_bg,   "0.0000"),
                (fiber.get('perdida_total_db'),          alt_bg,   "0.0000"),
                (fiber.get('perdida_promedio_dbkm'),     pprom_bg, "0.0000"),
                (fiber.get('perdida_union_promedio_db'), alt_bg,   "0.0000"),
                (umax,                                   umax_bg,  "0.0000"),
                (n_empalmes,                             alt_bg,   None),
                (delta_max,                              delta_bg, "+0.0000"),
                (fiber.get('date', ''),                  alt_bg,   None),
            ]
            for ci, (val, bg, fmt) in enumerate(vals, 1):
                _cell(ws, row, ci, value=val, bg=bg, fmt=fmt)
            row += 1


# ── Public API ────────────────────────────────────────────────────────────

def export_to_excel(cables_data: dict[str, list[dict]],
                    output_path: str,
                    thresholds: dict | None = None,
                    history_path: str = '') -> None:
    if not OPENPYXL_OK:
        raise ImportError("openpyxl no instalado. Ejecuta: pip install openpyxl")

    thr = thresholds or {}

    # Load previous measurement for delta comparison
    history = hist_mod.load(history_path) if history_path else {}
    ref_date = None
    for fibers in cables_data.values():
        if fibers and fibers[0].get('date'):
            ref_date = fibers[0]['date']
            break
    previous = hist_mod.get_previous(history, ref_date)

    wb = Workbook()
    wb.remove(wb.active)

    for cable_name, fibers in cables_data.items():
        _write_cable_sheet(wb, cable_name, fibers, thr, previous.get(cable_name, {}))

    _write_summary_sheet(wb, cables_data, thr, previous)

    wb.save(output_path)

    # Save this measurement to history
    if history_path:
        hist_mod.save(history_path, cables_data, ref_date)


def build_output_path(root_folder: str, ref_date: str | None = None) -> str:
    ym = ref_date[:7] if ref_date and len(ref_date) >= 7 else str(date.today())[:7]
    return os.path.join(root_folder, f"FO_Cartilla_FOS_{ym}.xlsx")

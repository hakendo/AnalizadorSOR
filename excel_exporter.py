from __future__ import annotations
import os
from datetime import date
from typing import Any

try:
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False


# ── Colors ──────────────────────────────────────────────────────────────────
_BLUE_HEADER  = "1F4E79"
_BLUE_SUBHDR  = "2E75B6"
_BLUE_LIGHT   = "BDD7EE"
_GRAY_ROW     = "F2F2F2"
_WHITE        = "FFFFFF"
_YELLOW_WARN  = "FFE699"


def _thin_border() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def _cell_style(ws, row: int, col: int, value: Any = None,
                bold: bool = False, bg: str | None = None,
                align: str = "center", num_format: str | None = None,
                border: bool = True) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(name="Calibri", bold=bold, size=10,
                     color="FFFFFF" if bg in (_BLUE_HEADER, _BLUE_SUBHDR) else "000000")
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    if num_format:
        cell.number_format = num_format
    if border:
        cell.border = _thin_border()


def _auto_width(ws) -> None:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                length = len(str(cell.value)) if cell.value is not None else 0
                if length > max_len:
                    max_len = length
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 30)


# ── Column definitions ────────────────────────────────────────────────────
_DETAIL_HEADERS = [
    ("Fibra N°",              9,  None),
    ("N° Evento",             9,  None),
    ("Posición (km)",         12, "0.0000"),
    ("Long. Intervalo (km)",  16, "0.0000"),
    ("Pérd. Intervalo (dB)",  16, "0.0000"),
    ("Pérd. Prom. (dB/km)",   16, "0.0000"),
    ("Pérd. Unión (dB)",      14, "0.0000"),
    ("Pérd. Unión Prom. (dB)", 18, "0.0000"),
    ("Pérd. Unión Máx. (dB)", 18, "0.0000"),
]

_SUMMARY_HEADERS = [
    ("Fibra N°",              9,  None),
    ("Long. Total (km)",      14, "0.0000"),
    ("Pérd. Total (dB)",      14, "0.0000"),
    ("Pérd. Prom. (dB/km)",   16, "0.0000"),
    ("Pérd. Unión Prom. (dB)", 18, "0.0000"),
    ("Pérd. Unión Máx. (dB)", 18, "0.0000"),
    ("N° Empalmes",           12, None),
    ("Fecha",                 12, None),
]


def _write_cable_sheet(wb: "Workbook", cable_name: str, cable_data: list[dict]) -> None:
    """Write one sheet per cable with event detail rows."""
    clean_name = cable_name.strip()[:31]
    ws = wb.create_sheet(title=clean_name)
    ws.freeze_panes = "A7"

    # ── Header block (rows 1-5) ──────────────────────────────────────────
    first = cable_data[0] if cable_data else {}
    wl_nm = first.get('wavelength_nm', '')
    otdr  = first.get('otdr_model', '') or 'EXFO'
    mdate = first.get('date', str(date.today()))

    merge_cols = len(_DETAIL_HEADERS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_cols)
    _cell_style(ws, 1, 1,
                value=f"Medición de Filamentos — Cable: {clean_name}",
                bold=True, bg=_BLUE_HEADER, align="left", border=False)
    ws.row_dimensions[1].height = 18

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    _cell_style(ws, 2, 1, value=f"Fecha medición: {mdate}", align="left", border=False)
    ws.merge_cells(start_row=2, start_column=5, end_row=2, end_column=merge_cols)
    _cell_style(ws, 2, 5,
                value=f"Longitud de onda: {wl_nm} nm   |   Equipo OTDR: {otdr}",
                align="left", border=False)

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=merge_cols)
    ws.row_dimensions[4].height = 8

    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=merge_cols)

    # ── Column headers (row 6) ───────────────────────────────────────────
    for col_idx, (label, width, _) in enumerate(_DETAIL_HEADERS, start=1):
        _cell_style(ws, 6, col_idx, value=label,
                    bold=True, bg=_BLUE_SUBHDR, align="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[6].height = 30

    # ── Data rows ────────────────────────────────────────────────────────
    row = 7
    for fiber in cable_data:
        fibra_num   = fiber.get('fibra_num')
        events      = fiber.get('events', [])
        union_prom  = fiber.get('perdida_union_promedio_db')
        union_max   = fiber.get('perdida_union_maxima_db')

        if not events:
            continue

        start_row = row
        for ev_idx, ev in enumerate(events):
            bg = _GRAY_ROW if fibra_num % 2 == 0 else _WHITE

            _cell_style(ws, row, 1, value=fibra_num if ev_idx == 0 else None,
                        bold=(ev_idx == 0), bg=bg, align="center")
            _cell_style(ws, row, 2, value=ev.get('n_evento'), bg=bg, align="center")
            _cell_style(ws, row, 3, value=ev.get('posicion_km'), bg=bg,
                        num_format="0.0000")
            _cell_style(ws, row, 4, value=ev.get('longitud_intervalo_km'), bg=bg,
                        num_format="0.0000")
            _cell_style(ws, row, 5, value=ev.get('perdida_intervalo_db'), bg=bg,
                        num_format="0.0000")
            _cell_style(ws, row, 6, value=ev.get('perdida_promedio_dbkm'), bg=bg,
                        num_format="0.0000")
            _cell_style(ws, row, 7, value=ev.get('perdida_union_db'), bg=bg,
                        num_format="0.0000")

            # Summary cols only on first row of each fibra
            _cell_style(ws, row, 8,
                        value=union_prom if ev_idx == 0 else None,
                        bg=_BLUE_LIGHT if ev_idx == 0 else bg,
                        num_format="0.0000")
            _cell_style(ws, row, 9,
                        value=union_max if ev_idx == 0 else None,
                        bg=_BLUE_LIGHT if ev_idx == 0 else bg,
                        num_format="0.0000")
            row += 1

        # Merge fibra number cell vertically across all events of this fibra
        if len(events) > 1:
            ws.merge_cells(start_row=start_row, start_column=1,
                           end_row=row - 1,   end_column=1)

    ws.row_dimensions[row].height = 6  # spacer at end


def _write_summary_sheet(wb: "Workbook", all_cables: dict[str, list[dict]]) -> None:
    """Write a Resumen sheet with one row per fibra (all cables)."""
    ws = wb.create_sheet(title="Resumen")
    ws.freeze_panes = "A3"

    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                   end_column=len(_SUMMARY_HEADERS))
    _cell_style(ws, 1, 1, value="Resumen de Mediciones — Todos los Cables",
                bold=True, bg=_BLUE_HEADER, align="left", border=False)
    ws.row_dimensions[1].height = 18

    for col_idx, (label, width, _) in enumerate(_SUMMARY_HEADERS, start=1):
        _cell_style(ws, 2, col_idx, value=label,
                    bold=True, bg=_BLUE_SUBHDR)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[2].height = 25

    row = 3
    for cable_name, fibers in all_cables.items():
        # cable separator row
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=len(_SUMMARY_HEADERS))
        _cell_style(ws, row, 1,
                    value=f"Cable: {cable_name.strip()}",
                    bold=True, bg=_BLUE_LIGHT, align="left")
        row += 1

        for fiber in fibers:
            fibra_num = fiber.get('fibra_num')
            n_empalmes = len([e for e in fiber.get('events', []) if not e.get('is_start')])
            bg = _GRAY_ROW if fibra_num % 2 == 0 else _WHITE

            vals = [
                fibra_num,
                fiber.get('longitud_total_km'),
                fiber.get('perdida_total_db'),
                fiber.get('perdida_promedio_dbkm'),
                fiber.get('perdida_union_promedio_db'),
                fiber.get('perdida_union_maxima_db'),
                n_empalmes,
                fiber.get('date', ''),
            ]
            for col_idx, (val, (_, _, fmt)) in enumerate(
                    zip(vals, _SUMMARY_HEADERS), start=1):
                _cell_style(ws, row, col_idx, value=val, bg=bg,
                            num_format=fmt, align="center")
            row += 1


def export_to_excel(cables_data: dict[str, list[dict]],
                    output_path: str) -> None:
    """
    cables_data: {cable_name: [parsed_sor_dict, ...]}
    Each sor_dict must have 'fibra_num' added by the caller.
    """
    if not OPENPYXL_OK:
        raise ImportError(
            "openpyxl no está instalado. Ejecuta: pip install openpyxl"
        )

    wb = Workbook()
    wb.remove(wb.active)   # remove default empty sheet

    for cable_name, fibers in cables_data.items():
        _write_cable_sheet(wb, cable_name, fibers)

    _write_summary_sheet(wb, cables_data)

    wb.save(output_path)


def build_output_path(root_folder: str, ref_date: str | None = None) -> str:
    """Build standardized output filename: FO_Cartilla_FOS_YYYY-MM.xlsx"""
    if ref_date and len(ref_date) >= 7:
        ym = ref_date[:7]   # "YYYY-MM"
    else:
        ym = str(date.today())[:7]
    filename = f"FO_Cartilla_FOS_{ym}.xlsx"
    return os.path.join(root_folder, filename)

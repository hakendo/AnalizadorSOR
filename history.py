import json
import os
from datetime import date

HISTORY_FILENAME = 'mediciones_historial.json'


def default_path(folder: str) -> str:
    return os.path.join(folder, HISTORY_FILENAME)


def load(path: str) -> dict:
    if path and os.path.exists(path):
        try:
            with open(path, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save(path: str, cables_data: dict, ref_date: str | None = None) -> None:
    existing = load(path)
    key = ref_date or str(date.today())
    snapshot: dict = {}
    for cable_name, fibers in cables_data.items():
        snapshot[cable_name] = {}
        for fiber in fibers:
            entry_key = f"{fiber.get('fibra_num')}_{fiber.get('direction', 'normal')}"
            snapshot[cable_name][entry_key] = {
                'longitud_total_km':         fiber.get('longitud_total_km'),
                'perdida_total_db':          fiber.get('perdida_total_db'),
                'perdida_promedio_dbkm':     fiber.get('perdida_promedio_dbkm'),
                'perdida_union_promedio_db': fiber.get('perdida_union_promedio_db'),
                'perdida_union_maxima_db':   fiber.get('perdida_union_maxima_db'),
            }
    existing[key] = snapshot
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)


def get_previous(history: dict, ref_date: str | None = None) -> dict:
    """Return most recent snapshot before ref_date, or the last one."""
    if not history:
        return {}
    keys = sorted(history.keys())
    if ref_date and ref_date in keys:
        idx = keys.index(ref_date)
        return history[keys[idx - 1]] if idx > 0 else {}
    return history[keys[-1]] if keys else {}


def delta(current_val: float | None, prev_val: float | None) -> float | None:
    if current_val is None or prev_val is None:
        return None
    return round(current_val - prev_val, 4)

import struct
import re
import os
from datetime import datetime

from config import EQUIPMENT_PROFILES

_C_LIGHT = 2.998e8  # m/s
_DEFAULT_EQUIPMENT = 'exfo'


# ── Binary helpers ────────────────────────────────────────────────────────

def _read_cstring(data: bytes, pos: int) -> tuple[str, int]:
    end = data.index(b'\x00', pos)
    return data[pos:end].decode('ascii', errors='replace'), end + 1


def _parse_map(data: bytes) -> dict[str, tuple[int, int]]:
    """Returns {block_name: (file_offset, size)} for every block after Map."""
    _, pos = _read_cstring(data, 0)
    pos += 2
    map_size = struct.unpack_from('<I', data, pos)[0]; pos += 4
    num_blocks = struct.unpack_from('<H', data, pos)[0]; pos += 2

    block_map: dict[str, tuple[int, int]] = {}
    cumulative = map_size
    for _ in range(num_blocks - 1):
        name, pos = _read_cstring(data, pos)
        pos += 2
        size = struct.unpack_from('<I', data, pos)[0]; pos += 4
        block_map[name] = (cumulative, size)
        cumulative += size
    return block_map


# ── Block parsers ─────────────────────────────────────────────────────────

def _parse_gen_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)
    _, pos = _read_cstring(data, pos)
    cable_id,  pos = _read_cstring(data, pos)
    fiber_id,  pos = _read_cstring(data, pos)
    pos += 2; pos += 2
    originating, pos = _read_cstring(data, pos)
    terminating, pos = _read_cstring(data, pos)
    _, pos = _read_cstring(data, pos)
    pos += 2; pos += 2
    operator, pos = _read_cstring(data, pos)
    return {
        'cable_id':       cable_id.strip(),
        'fiber_id':       fiber_id.strip(),
        'originating_loc': originating.strip(),
        'terminating_loc': terminating.strip(),
        'operator':       operator.strip(),
    }


def _parse_sup_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)
    _, pos = _read_cstring(data, pos)   # supplier (empty in EXFO)
    _, pos = _read_cstring(data, pos)   # extra EXFO field
    otdr_name, pos = _read_cstring(data, pos)
    otdr_sn,   pos = _read_cstring(data, pos)
    return {'otdr_model': otdr_name.strip(), 'otdr_sn': otdr_sn.strip()}


def _parse_fxd_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)
    timestamp = struct.unpack_from('<I', data, pos)[0]; pos += 4
    pos += 2
    wavelength_raw = struct.unpack_from('<H', data, pos)[0]; pos += 2
    pos += 4; pos += 4
    n_pulses = struct.unpack_from('<H', data, pos)[0]; pos += 2

    group_index = 1.4682
    for i in range(n_pulses):
        pos += 2; pos += 4; pos += 4
        raw_gi = struct.unpack_from('<H', data, pos)[0]; pos += 2
        if i == 0:
            group_index = raw_gi * 1e-4
        pos += 2; pos += 4; pos += 4; pos += 4

    date_str = ''
    try:
        date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
    except Exception:
        pass

    return {
        'wavelength_nm': wavelength_raw // 10,
        'group_index':   group_index,
        'date':          date_str,
    }


def _prop_time_to_km(prop_time: int, group_index: float) -> float:
    return prop_time * _C_LIGHT / group_index / 2 * 1e-11 / 1000.0


def _is_end_event(ev_type: str) -> bool:
    return (len(ev_type) >= 2 and ev_type[1] in ('O', 'o')) or '9999' in ev_type


def _parse_key_events(data: bytes, offset: int,
                      group_index: float, comment_size: int) -> list[dict]:
    event_size = 2 + 4 + 2 + 2 + 4 + 2 + 2 + comment_size   # = 44 for EXFO

    pos = offset
    _, pos = _read_cstring(data, pos)
    num_events = struct.unpack_from('<H', data, pos)[0]; pos += 2

    raw_events = []
    for _ in range(num_events):
        base = pos
        ev_num   = struct.unpack_from('<H', data, pos)[0]; pos += 2
        prop_time = struct.unpack_from('<I', data, pos)[0]; pos += 4
        atten    = struct.unpack_from('<H', data, pos)[0]; pos += 2
        ev_loss  = struct.unpack_from('<h', data, pos)[0]; pos += 2
        pos += 4
        ev_type  = data[pos:pos+2].decode('ascii', errors='replace'); pos += 2
        raw_events.append({
            'n':              ev_num,
            'prop_time':      prop_time,
            'atten_dbkm':     atten * 0.001,
            'ev_loss_db':     ev_loss * 0.001,
            'ev_type':        ev_type.strip(),
        })
        pos = base + event_size
    return raw_events


# ── Public API ────────────────────────────────────────────────────────────

def parse_sor(filepath: str, equipment: str = _DEFAULT_EQUIPMENT) -> dict | None:
    """Parse one SOR file. Returns structured dict or {'error': str}."""
    comment_size = EQUIPMENT_PROFILES.get(equipment, EQUIPMENT_PROFILES['exfo'])['comment_size']
    try:
        with open(filepath, 'rb') as f:
            data = f.read()

        bmap = _parse_map(data)
        gen  = _parse_gen_params(data, bmap['GenParams'][0]) if 'GenParams' in bmap else {}
        sup  = _parse_sup_params(data, bmap['SupParams'][0]) if 'SupParams' in bmap else {}
        fxd  = _parse_fxd_params(data, bmap['FxdParams'][0]) if 'FxdParams' in bmap else {}

        group_index = fxd.get('group_index', 1.4682)
        raw_events  = (_parse_key_events(data, bmap['KeyEvents'][0],
                                         group_index, comment_size)
                       if 'KeyEvents' in bmap else [])

        events: list[dict] = []
        prev_km = 0.0
        total_km = 0.0

        for ev in raw_events:
            pos_km   = _prop_time_to_km(ev['prop_time'], group_index)
            is_end   = _is_end_event(ev['ev_type'])
            is_start = ev['prop_time'] == 0

            if is_end:
                total_km = pos_km
                continue

            interval_km = pos_km - prev_km
            events.append({
                'n_evento':               ev['n'],
                'posicion_km':            round(pos_km, 4),
                'longitud_intervalo_km':  round(interval_km, 4),
                'perdida_intervalo_db':   round(ev['atten_dbkm'] * interval_km, 4),
                'perdida_promedio_dbkm':  round(ev['atten_dbkm'], 4),
                'perdida_union_db':       round(ev['ev_loss_db'], 4) if not is_start else 0.0,
                '_is_start':              is_start,
            })
            prev_km = pos_km

        if total_km == 0.0 and events:
            total_km = events[-1]['posicion_km']

        splice_losses = [e['perdida_union_db']
                         for e in events
                         if not e.get('_is_start') and e['perdida_union_db'] >= 0]

        union_prom = round(sum(splice_losses) / len(splice_losses), 4) if splice_losses else 0.0
        union_max  = round(max(splice_losses), 4) if splice_losses else 0.0
        total_loss = round(sum(e['perdida_intervalo_db'] for e in events), 4)
        avg_dbkm   = round(total_loss / total_km, 4) if total_km > 0 else 0.0

        for e in events:
            e.pop('_is_start', None)

        return {
            'filepath':                  filepath,
            'cable_id':                  gen.get('cable_id', ''),
            'wavelength_nm':             fxd.get('wavelength_nm', 0),
            'otdr_model':                sup.get('otdr_model', ''),
            'date':                      fxd.get('date', ''),
            'events':                    events,
            'perdida_union_promedio_db': union_prom,
            'perdida_union_maxima_db':   union_max,
            'longitud_total_km':         round(total_km, 4),
            'perdida_total_db':          total_loss,
            'perdida_promedio_dbkm':     avg_dbkm,
        }
    except Exception as e:
        return {'error': str(e), 'filepath': filepath}


def _extract_fibra_number(filename: str) -> int | None:
    m = re.search(r'fibra\s*(\d+)', filename, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _detect_direction(filename: str) -> str:
    lower = filename.lower()
    if 'corta' in lower:
        return 'corta'
    if 'larga' in lower:
        return 'larga'
    return 'normal'


def scan_folder(root_folder: str,
                directions: list[str] | None = None) -> dict[str, list[dict]]:
    """
    Scan root_folder for cable subdirectories.
    Returns {cable_name: [{'fibra': int, 'paths': {'normal': path, ...}}, ...]}.
    Only includes directions listed in `directions` (default: ['normal']).
    """
    if directions is None:
        directions = ['normal']

    cables: dict[str, list[dict]] = {}

    for entry in sorted(os.scandir(root_folder), key=lambda e: e.name):
        if not entry.is_dir():
            continue

        fibras_dict: dict[int, dict] = {}

        try:
            for f in sorted(os.scandir(entry.path), key=lambda e: e.name):
                if not f.is_file() or not f.name.lower().endswith('.sor'):
                    continue
                num = _extract_fibra_number(f.name)
                if num is None:
                    continue
                direction = _detect_direction(f.name)
                if direction not in directions:
                    continue
                if num not in fibras_dict:
                    fibras_dict[num] = {'fibra': num, 'paths': {}}
                fibras_dict[num]['paths'][direction] = f.path
        except PermissionError:
            continue

        fibras = [v for _, v in sorted(fibras_dict.items()) if v['paths']]
        if fibras:
            cables[entry.name] = fibras

    return cables

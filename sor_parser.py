import struct
import re
import os
from datetime import datetime


_C_LIGHT = 2.998e8  # m/s
_KEY_EVENTS_COMMENT_SIZE = 26   # EXFO format uses 26-byte comment (Bellcore standard = 16)
_KEY_EVENTS_EVENT_SIZE = 2 + 4 + 2 + 2 + 4 + 2 + 2 + _KEY_EVENTS_COMMENT_SIZE  # = 44
_KEY_EVENTS_FOOTER_SIZE = 22


def _read_cstring(data: bytes, pos: int) -> tuple[str, int]:
    end = data.index(b'\x00', pos)
    return data[pos:end].decode('ascii', errors='replace'), end + 1


def _parse_map(data: bytes) -> dict[str, tuple[int, int]]:
    """Returns {block_name: (offset, size)} for all blocks after the Map block."""
    _, pos = _read_cstring(data, 0)
    pos += 2  # revision
    map_size = struct.unpack_from('<I', data, pos)[0]
    pos += 4
    num_blocks = struct.unpack_from('<H', data, pos)[0]
    pos += 2

    block_map: dict[str, tuple[int, int]] = {}
    cumulative = map_size  # blocks start right after Map
    for _ in range(num_blocks - 1):  # Map itself is not listed in its own table
        name, pos = _read_cstring(data, pos)
        pos += 2  # revision (not needed for parsing)
        size = struct.unpack_from('<I', data, pos)[0]
        pos += 4
        block_map[name] = (cumulative, size)
        cumulative += size

    return block_map


def _parse_gen_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)
    _, pos = _read_cstring(data, pos)   # language
    cable_id, pos = _read_cstring(data, pos)
    fiber_id, pos = _read_cstring(data, pos)
    pos += 2   # fiber_type
    pos += 2   # nominal_wavelength
    originating_loc, pos = _read_cstring(data, pos)
    terminating_loc, pos = _read_cstring(data, pos)
    _, pos = _read_cstring(data, pos)  # cable_code
    pos += 2   # build_condition
    pos += 2   # user_offset
    operator, pos = _read_cstring(data, pos)
    return {
        'cable_id': cable_id.strip(),
        'fiber_id': fiber_id.strip(),
        'originating_loc': originating_loc.strip(),
        'terminating_loc': terminating_loc.strip(),
        'operator': operator.strip(),
    }


def _parse_sup_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)    # block name
    _, pos = _read_cstring(data, pos)    # supplier_name (often empty in EXFO)
    _, pos = _read_cstring(data, pos)    # extra field (EXFO adds one empty string)
    otdr_name, pos = _read_cstring(data, pos)
    otdr_sn, pos = _read_cstring(data, pos)
    return {
        'otdr_model': otdr_name.strip(),
        'otdr_sn': otdr_sn.strip(),
    }


def _parse_fxd_params(data: bytes, offset: int) -> dict:
    pos = offset
    _, pos = _read_cstring(data, pos)  # block name

    timestamp = struct.unpack_from('<I', data, pos)[0]; pos += 4
    pos += 2    # units (2-char ASCII in EXFO: 'mt', 'km', 'ft', etc.)
    wavelength_raw = struct.unpack_from('<H', data, pos)[0]; pos += 2
    pos += 4    # acquisition_offset (int32)
    pos += 4    # acquisition_offset_distance (uint32)
    n_pulses = struct.unpack_from('<H', data, pos)[0]; pos += 2

    group_index = 1.4682  # default for SMF
    for i in range(n_pulses):
        pos += 2   # pulse_width (ns)
        pos += 4   # sample_spacing
        pos += 4   # num_data_points
        raw_gi = struct.unpack_from('<H', data, pos)[0]; pos += 2
        if i == 0:
            group_index = raw_gi * 1e-4   # EXFO stores in units of 1e-4
        pos += 2   # backscatter_coefficient
        pos += 4   # loss_threshold
        pos += 4   # reflection_threshold
        pos += 4   # eot_threshold

    pos += 2   # measurement_time
    pos += 4   # data_spacing
    pos += 4   # num_data_points (global)

    date_str = ''
    try:
        date_str = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d')
    except Exception:
        pass

    return {
        'timestamp': timestamp,
        'wavelength_nm': wavelength_raw // 10,
        'group_index': group_index,
        'date': date_str,
    }


def _prop_time_to_km(prop_time: int, group_index: float) -> float:
    """Convert propagation time (10^-11 s units) to km."""
    return prop_time * _C_LIGHT / group_index / 2 * 1e-11 / 1000.0


def _is_end_event(ev_type: str) -> bool:
    """
    End-of-fiber events in EXFO SOR files use type codes ending in 'O' (e.g. '1O', '2O').
    Older Bellcore format uses '9999LS'.
    """
    return (len(ev_type) >= 2 and ev_type[1] in ('O', 'o')) or '9999' in ev_type


def _parse_key_events(data: bytes, offset: int, group_index: float) -> list[dict]:
    pos = offset
    _, pos = _read_cstring(data, pos)
    num_events = struct.unpack_from('<H', data, pos)[0]
    pos += 2

    raw_events = []
    for _ in range(num_events):
        base = pos
        ev_num = struct.unpack_from('<H', data, pos)[0]; pos += 2
        prop_time = struct.unpack_from('<I', data, pos)[0]; pos += 4
        atten_coeff = struct.unpack_from('<H', data, pos)[0]; pos += 2   # 0.001 dB/km
        ev_loss = struct.unpack_from('<h', data, pos)[0]; pos += 2        # 0.001 dB signed
        pos += 4   # event_reflection
        ev_type = data[pos:pos+2].decode('ascii', errors='replace'); pos += 2
        pos += 2   # marker_flag
        comment = data[pos:pos+_KEY_EVENTS_COMMENT_SIZE].rstrip(b'\x00').decode('ascii', errors='replace')
        pos = base + _KEY_EVENTS_EVENT_SIZE  # advance by exactly 44 bytes
        raw_events.append({
            'n': ev_num,
            'prop_time': prop_time,
            'atten_coeff_dbkm': atten_coeff * 0.001,
            'ev_loss_db': ev_loss * 0.001,
            'ev_type': ev_type.strip(),
        })

    return raw_events


def parse_sor(filepath: str) -> dict | None:
    """Parse a single SOR file. Returns structured data or None on error."""
    try:
        with open(filepath, 'rb') as f:
            data = f.read()

        block_map = _parse_map(data)

        gen = _parse_gen_params(data, block_map['GenParams'][0]) if 'GenParams' in block_map else {}
        sup = _parse_sup_params(data, block_map['SupParams'][0]) if 'SupParams' in block_map else {}
        fxd = _parse_fxd_params(data, block_map['FxdParams'][0]) if 'FxdParams' in block_map else {}

        group_index = fxd.get('group_index', 1.4682)
        raw_events = _parse_key_events(data, block_map['KeyEvents'][0], group_index) if 'KeyEvents' in block_map else []

        events = []
        prev_pos_km = 0.0
        total_length_km = 0.0

        for ev in raw_events:
            pos_km = _prop_time_to_km(ev['prop_time'], group_index)
            is_end = _is_end_event(ev['ev_type'])
            is_start = (ev['prop_time'] == 0)

            if is_end:
                total_length_km = pos_km
                continue   # don't include end event in splice list

            interval_km = pos_km - prev_pos_km
            interval_loss_db = ev['atten_coeff_dbkm'] * interval_km

            events.append({
                'n_evento': ev['n'],
                'posicion_km': round(pos_km, 4),
                'longitud_intervalo_km': round(interval_km, 4),
                'perdida_intervalo_db': round(interval_loss_db, 4),
                'perdida_promedio_dbkm': round(ev['atten_coeff_dbkm'], 4),
                'perdida_union_db': round(ev['ev_loss_db'], 4) if not is_start else 0.0,
                'is_start': is_start,
            })
            prev_pos_km = pos_km

        if total_length_km == 0.0 and events:
            total_length_km = events[-1]['posicion_km']

        splice_losses = [e['perdida_union_db'] for e in events if not e.get('is_start') and e['perdida_union_db'] >= 0]
        union_prom = round(sum(splice_losses) / len(splice_losses), 4) if splice_losses else 0.0
        union_max = round(max(splice_losses), 4) if splice_losses else 0.0
        total_loss_db = round(sum(e['perdida_intervalo_db'] for e in events), 4)
        avg_dbkm = round(total_loss_db / total_length_km, 4) if total_length_km > 0 else 0.0

        # Remove internal flag before returning
        for e in events:
            e.pop('is_start', None)

        return {
            'filepath': filepath,
            'cable_id': gen.get('cable_id', ''),
            'fiber_id': gen.get('fiber_id', ''),
            'wavelength_nm': fxd.get('wavelength_nm', 0),
            'otdr_model': sup.get('otdr_model', ''),
            'date': fxd.get('date', ''),
            'group_index': group_index,
            'events': events,
            'perdida_union_promedio_db': union_prom,
            'perdida_union_maxima_db': union_max,
            'longitud_total_km': round(total_length_km, 4),
            'perdida_total_db': total_loss_db,
            'perdida_promedio_dbkm': avg_dbkm,
        }
    except Exception as e:
        return {'error': str(e), 'filepath': filepath}


def _extract_fibra_number(filename: str) -> int | None:
    """Extract filament number from filename like 'fibra10 biobio constitucion.sor'."""
    m = re.search(r'fibra\s*(\d+)', filename, re.IGNORECASE)
    return int(m.group(1)) if m else None


def _is_normal_sor(filename: str) -> bool:
    """Return True only for the 'normal' (bidirectional) SOR, not corta/larga."""
    lower = filename.lower()
    return lower.endswith('.sor') and 'corta' not in lower and 'larga' not in lower


def scan_folder(root_folder: str) -> dict[str, list[dict]]:
    """
    Scan root_folder for cable subdirectories, each containing SOR files.
    Returns {cable_name: [{'fibra': int, 'path': str}, ...]} sorted by fibra number.
    """
    cables: dict[str, list[dict]] = {}
    for entry in sorted(os.scandir(root_folder), key=lambda e: e.name):
        if not entry.is_dir():
            continue
        cable_name = entry.name
        fibras: list[dict] = []
        try:
            for f in sorted(os.scandir(entry.path), key=lambda e: e.name):
                if f.is_file() and _is_normal_sor(f.name):
                    fibra_num = _extract_fibra_number(f.name)
                    if fibra_num is not None:
                        fibras.append({'fibra': fibra_num, 'path': f.path})
        except PermissionError:
            continue
        if fibras:
            fibras.sort(key=lambda x: x['fibra'])
            cables[cable_name] = fibras
    return cables

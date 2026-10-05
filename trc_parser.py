"""Reader for EXFO native .trc files (FTBx / Metrino, 'AppReg Format Ex').

Layout: two 'AppReg Format Ex' headers; the second one gives the total
uncompressed size and is followed by zlib chunks, each prefixed by its
compressed size (u32). The inflated buffer is a tree of records:

    [name_off u32][type u32][size u32][value_off u32] name\\0 value

type 0 = container (value = u32 offsets of child records), 1 = int32,
2 = raw bytes, 3 = float64, 4 = UTF-16LE string.
"""
from __future__ import annotations

import struct
import zlib

TRC_MAGIC = b'AppReg Format Ex'

_ROOT_RECORD = 0x0c
_STATUS_START = 0x40
_STATUS_END = 0x80


def is_trc(data: bytes) -> bool:
    return data.startswith(TRC_MAGIC)


def _inflate(data: bytes) -> bytes:
    hdr = data.find(TRC_MAGIC, 1)
    if hdr < 0:
        raise ValueError('TRC: falta el segundo encabezado')
    total = struct.unpack_from('<I', data, hdr + 24)[0]
    pos = hdr + 40
    out = bytearray()
    while len(out) < total:
        size = struct.unpack_from('<I', data, pos - 4)[0]
        out += zlib.decompress(data[pos:pos + size])
        pos += size + 4
    return bytes(out)


def _record(buf: bytes, off: int) -> tuple[str, int, bytes]:
    name_off, rtype, size, value_off = struct.unpack_from('<IIII', buf, off)
    name = buf[name_off:buf.index(b'\x00', name_off)].decode('latin-1')
    return name, rtype, buf[value_off:value_off + size]


def _node(buf: bytes, off: int) -> dict:
    """Children of a container record as {name: value}; sub-containers stay as offsets."""
    _, _, value = _record(buf, off)
    out = {}
    for i in range(0, len(value), 4):
        child = struct.unpack_from('<I', value, i)[0]
        name, rtype, v = _record(buf, child)
        if rtype == 1 and len(v) == 4:
            out[name] = struct.unpack('<i', v)[0]
        elif rtype == 3 and len(v) == 8:
            out[name] = struct.unpack('<d', v)[0]
        elif rtype == 4:
            out[name] = v.decode('utf-16-le', errors='replace').rstrip('\x00')
        elif rtype == 0:
            out[name] = child
        else:
            out[name] = v
    return out


def read_trc(data: bytes) -> tuple[dict, list[dict]]:
    """Returns (meta, raw_events) in the same shape sor_parser builds from SOR."""
    buf = _inflate(data)
    otdr  = _node(buf, _node(buf, _ROOT_RECORD)['OtdrData'])
    fiber = _node(buf, _node(buf, otdr['Fibers'])['Fiber0'])
    trace = _node(buf, _node(buf, fiber['Traces'])['Trace0'])
    table = _node(buf, trace['EventTable'])

    wavelength_nm = 0
    acq = _node(buf, trace['AcquisitionDatas']) if 'AcquisitionDatas' in trace else {}
    if 'AcquisitionData0' in acq:
        wavelength_nm = int(_node(buf, acq['AcquisitionData0']).get('ExactWavelength', 0) * 1e9)
    if not wavelength_nm:
        wavelength_nm = round(trace.get('Wavelength', 0) * 1e9)

    module = _node(buf, trace['ModuleInformation']) if 'ModuleInformation' in trace else {}

    meta = {
        'cable_id':      fiber.get('Cable') or fiber.get('Identifier', ''),
        'wavelength_nm': wavelength_nm,
        'otdr_model':    module.get('ModelName', ''),
        'date':          trace.get('Date', '')[:10],
    }

    # The table alternates point events (have 'Type') and fiber sections
    # (Length/Loss in metres/dB). Each point event takes the attenuation of
    # the section that ends on it.
    raw_events: list[dict] = []
    atten_dbkm = 0.0
    for i in range(table.get('Count', 0)):
        ev = _node(buf, table[f'Event{i}'])
        if 'Type' not in ev:
            length_m = ev.get('Length', 0.0)
            atten_dbkm = ev.get('Loss', 0.0) / length_m * 1000.0 if length_m > 0 else 0.0
            continue
        loss = ev.get('Loss', 0.0)
        raw_events.append({
            'n':          len(raw_events) + 1,
            'pos_km':     ev.get('Position', 0.0) / 1000.0,
            'atten_dbkm': atten_dbkm,
            'ev_loss_db': 0.0 if loss != loss else loss,   # NaN → 0
            'is_start':   bool(ev.get('Status', 0) & _STATUS_START) or not raw_events,
            'is_end':     bool(ev.get('Status', 0) & _STATUS_END),
        })
    return meta, raw_events

import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')

EQUIPMENT_PROFILES = {
    'exfo':     {'comment_size': 26, 'label': 'EXFO FTBx'},
    'anritsu':  {'comment_size': 16, 'label': 'Anritsu MT'},
    'viavi':    {'comment_size': 16, 'label': 'VIAVI / JDSU'},
    'yokogawa': {'comment_size': 16, 'label': 'Yokogawa AQ7'},
    'afl':      {'comment_size': 16, 'label': 'AFL Noyes'},
    'standard': {'comment_size': 16, 'label': 'Estándar Bellcore'},
}

DEFAULT = {
    'last_folder':    '',
    'otdr_equipment': 'exfo',
    'directions':     ['normal'],
    'fibra_filter':   {},
    'history_path':   '',
    'thresholds': {
        'perdida_union_db':      0.5,
        'perdida_promedio_dbkm': 0.25,
        'perdida_intervalo_db':  2.0,
    },
    'columns': None,    # None = all default columns; populated on first save
}


def load() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding='utf-8') as f:
                stored = json.load(f)
            cfg = {**DEFAULT, **stored}
            cfg['thresholds'] = {**DEFAULT['thresholds'], **stored.get('thresholds', {})}
            return cfg
        except Exception:
            pass
    return dict(DEFAULT)


def save(cfg: dict) -> None:
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

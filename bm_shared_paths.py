# coding: utf-8
"""Shared writable path registry for Boffa pyRevit tools.

This module is intentionally Revit-free and IronPython-compatible.
"""
import os

_SHARED_DATA_ENV = 'BM_PYREVIT_SHARED_DATA_ROOT'


def normalise_path(path):
    if not path:
        return ''
    try:
        path = os.path.expandvars(os.path.expanduser(path))
    except Exception:
        pass
    try:
        return os.path.normpath(os.path.abspath(path))
    except Exception:
        return os.path.normpath(path)


def get_extension_root():
    try:
        return normalise_path(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        return ''


def _default_shared_data_root():
    override = os.environ.get(_SHARED_DATA_ENV, '')
    if override:
        return normalise_path(override)
    local = os.environ.get('LOCALAPPDATA', '')
    if local:
        return normalise_path(os.path.join(local, 'BoffaMiskell', 'pyRevit'))
    return normalise_path(os.path.join(os.path.expanduser('~'), 'BoffaMiskell', 'pyRevit'))


SHARED_DATA_ROOT = _default_shared_data_root()
PLANT_LIBRARY_DATA_ROOT = normalise_path(os.path.join(SHARED_DATA_ROOT, 'PlantLibrary'))
PLANT_LIBRARY_LOG_ROOT = normalise_path(os.path.join(PLANT_LIBRARY_DATA_ROOT, 'Logs'))
PLANT_LIBRARY_EVENT_LOG_ROOT = normalise_path(os.path.join(PLANT_LIBRARY_LOG_ROOT, 'Events'))
PLANT_LIBRARY_CUSTOM_PLANTS_CSV = normalise_path(os.path.join(PLANT_LIBRARY_DATA_ROOT, 'Custom_Plants.csv'))
PLANT_LIBRARY_USAGE_LOG_CSV = normalise_path(os.path.join(PLANT_LIBRARY_LOG_ROOT, 'plant_usage_log.csv'))
PLANT_LIBRARY_MODEL_SNAPSHOT_LOG_CSV = normalise_path(os.path.join(PLANT_LIBRARY_LOG_ROOT, 'plant_model_snapshot_log.csv'))

TOOL_PATHS = {
    'plant_library': {
        'data_root': PLANT_LIBRARY_DATA_ROOT,
        'log_root': PLANT_LIBRARY_LOG_ROOT,
        'event_log_root': PLANT_LIBRARY_EVENT_LOG_ROOT,
        'custom_plants_csv': PLANT_LIBRARY_CUSTOM_PLANTS_CSV,
        'usage_log_csv': PLANT_LIBRARY_USAGE_LOG_CSV,
        'model_snapshot_log_csv': PLANT_LIBRARY_MODEL_SNAPSHOT_LOG_CSV,
    },
}


def get_tool_path(tool_name, key, default=''):
    try:
        return normalise_path(TOOL_PATHS.get(tool_name, {}).get(key, default))
    except Exception:
        return normalise_path(default)


def is_environment_override_active():
    return bool(os.environ.get(_SHARED_DATA_ENV, ''))

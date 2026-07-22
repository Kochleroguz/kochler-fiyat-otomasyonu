import base64
import ctypes
import json
import os
from ctypes import wintypes
from pathlib import Path

CONFIG = Path(__file__).parent / 'data' / 'sentos.credentials'


class DATA_BLOB(ctypes.Structure):
    _fields_ = [('cbData', wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_char))]


def _blob(data):
    buf = ctypes.create_string_buffer(data)
    return DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf


def _protect(data):
    if os.name != 'nt': return b'DEV:' + base64.b64encode(data)
    source, keep = _blob(data); out = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(source), 'Kochler Sentos', None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    try: return ctypes.string_at(out.pbData, out.cbData)
    finally: ctypes.windll.kernel32.LocalFree(out.pbData)


def _unprotect(data):
    if data.startswith(b'DEV:'): return base64.b64decode(data[4:])
    source, keep = _blob(data); out = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(source), None, None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    try: return ctypes.string_at(out.pbData, out.cbData)
    finally: ctypes.windll.kernel32.LocalFree(out.pbData)


def save_credentials(base_url, username, password):
    CONFIG.parent.mkdir(exist_ok=True)
    CONFIG.write_bytes(_protect(json.dumps({'base_url': base_url, 'username': username, 'password': password}).encode('utf-8')))


def load_credentials():
    if not CONFIG.exists(): return None
    return json.loads(_unprotect(CONFIG.read_bytes()).decode('utf-8'))

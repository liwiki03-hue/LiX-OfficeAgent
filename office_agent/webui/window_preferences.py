import json
import os
from pathlib import Path
import threading


_WRITE_LOCK = threading.RLock()


class WindowPreferences:
    def __init__(self, root=None):
        self.path = Path(root or Path(__file__).resolve().parents[2]) / '.runtime/ui/window-preferences.json'

    def read(self):
        from ..downloads import safe
        try:
            safe(self.path)
            if not self.path.is_file() or self.path.stat().st_size > 16384:
                return {}
            value = json.loads(self.path.read_text('utf-8'))
            if not isinstance(value, dict) or set(value) - {'compact', 'review', 'topmost', 'intro_seen', 'theme'}:
                return {}
            result = {key: val for key, val in value.items() if key in {'topmost', 'intro_seen'} and type(val) is bool}
            if value.get('theme') in ('dark', 'light'):
                result['theme'] = value['theme']
            for mode in ('compact', 'review'):
                row = value.get(mode)
                if (isinstance(row, dict) and set(row) == {'x', 'y', 'width', 'height', 'dpi', 'screen'}
                        and all(type(row[key]) is int and abs(row[key]) <= 100000 for key in ('x', 'y', 'width', 'height', 'dpi'))
                        and 360 <= row['width'] <= 32768 and 440 <= row['height'] <= 32768 and 72 <= row['dpi'] <= 960
                        and isinstance(row['screen'], str) and len(row['screen']) <= 128):
                    result[mode] = row
            return result
        except (OSError, ValueError, TypeError):
            return {}

    def write(self, value):
        with _WRITE_LOCK:
            return self._write(value)

    def set_intro_seen(self):
        with _WRITE_LOCK:
            value = self.read()
            value['intro_seen'] = True
            self._write(value)

    def set_theme(self, theme):
        if type(theme) is not str or theme not in ('dark', 'light'):
            raise ValueError('主题偏好无效')
        with _WRITE_LOCK:
            self._write(self.read(), theme=theme)

    def _write(self, value, *, theme=None):
        from ..downloads import safe
        value = dict(value)
        current = self.read()
        if current.get('intro_seen') is True:
            value['intro_seen'] = True
        saved_theme = theme if theme is not None else current.get('theme')
        if saved_theme in ('dark', 'light'):
            value['theme'] = saved_theme
        else:
            value.pop('theme', None)
        safe(self.path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = safe(self.path.with_suffix('.tmp'))
        with temporary.open('w', encoding='utf-8') as stream:
            json.dump(value, stream)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, self.path)

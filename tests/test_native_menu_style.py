import json
from pathlib import Path
import subprocess
import sys
from importlib.metadata import version

import pytest


ROOT = Path(__file__).resolve().parents[1]


def test_native_menu_draws_and_preserves_actions(tmp_path):
    out = tmp_path / 'native-menu'
    if version('pythonnet') != '3.0.5':
        pytest.skip('Native GDI probe requires standard-entry pythonnet 3.0.5 runtime')
    result = subprocess.run([sys.executable, '-m', 'tests.native_menu_probe', str(out)],
                            cwd=ROOT, capture_output=True, text=True, encoding='utf-8',
                            errors='replace', timeout=45)
    assert result.returncode == 0, result.stdout + result.stderr
    report = json.loads((out / 'NATIVE_RENDER.json').read_text('utf-8'))
    assert report['root']['width'] >= 248
    assert report['submenu']['width'] >= 248
    assert report['native_render'] and not report['physical_tray']

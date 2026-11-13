from pathlib import Path
import json
import shutil

import pytest

from office_agent.diagram_process import _run_owned


NODE = shutil.which('node')


@pytest.mark.skipif(NODE is None, reason='Existing Node required for bounded DOM code regression')
def test_canvas_range_dom_controls_and_late_response_isolation(tmp_path):
    script = Path(__file__).with_name('probes') / 'l3_l4_canvas.cjs'
    output = tmp_path / 'dom-result.json'
    result = _run_owned((str(Path(NODE).resolve()), str(script.resolve()), str(output)), timeout_ms=15000)
    payload = json.loads(output.read_text(encoding='utf8')) if output.exists() else {'error':'Node returned no result'}
    assert result.code == 'process_completed' and result.exit_code == 0 and result.empty_confirmed, payload
    assert payload['passed']

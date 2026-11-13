"""Host PNG receipt and IO fault gates; no provider, browser or Office needed."""
from io import BytesIO
import hashlib
import json
from pathlib import Path

import pytest
from PIL import Image, PngImagePlugin

from office_agent.diagram_process import OwnedRun
from office_agent.diagram_raster import RasterReceipt, normalize_png, rasterize, verify_png
from office_agent.diagram_save import DiagramSaveError, DiagramSaveService
from office_agent.diagram_service import RenderedDiagram
from office_agent.diagrams import DiagramError, validate_source, validate_svg

SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="red"/></svg>'


def fixture(tmp_path):
    source = tmp_path / 'selected.docx'
    source.write_bytes(b'original source untouched')
    service = DiagramSaveService(tmp_path, source, 'session')
    preview = RenderedDiagram(validate_source('flowchart TD\nA --> B'), validate_svg(SVG),
                              OwnedRun('process_completed', 0, 12, 0.1, True))
    image = Image.new('RGB', (10, 10), 'red')
    buffer = BytesIO()
    image.save(buffer, format='PNG')
    png, width, height = normalize_png(buffer.getvalue(), expected_size=(10, 10))
    receipt = RasterReceipt(preview, png, hashlib.sha256(png).hexdigest(), width, height)
    return service, source, preview, receipt


def test_png_separate_confirmation_save_reopen_audit_and_rollback(tmp_path):
    service, source, preview, receipt = fixture(tmp_path)
    before = source.read_bytes()
    plan = service.plan_png(preview, receipt)
    assert plan['destination_basename'].endswith('.png')
    assert plan['png_sha256'] == receipt.sha256
    assert set(tmp_path.iterdir()) == {source}
    with pytest.raises(DiagramSaveError) as denied:
        service.confirm_save(plan['plan_id'], False)
    assert denied.value.code == 'export_confirmation_required'
    assert set(tmp_path.iterdir()) == {source}
    plan = service.plan_png(preview, receipt)
    outcome = service.confirm_save(plan['plan_id'], True)
    output, audit = Path(outcome['path']), Path(outcome['audit_path'])
    assert output.read_bytes() == receipt.png
    assert verify_png(output.read_bytes(), 10, 10)
    record = json.loads(audit.read_text('utf-8'))
    assert record['kind'] == 'diagram_png_save' and record['renderer'] == 'host-qt-svg-v1'
    assert record['png_sha256'] == receipt.sha256 and record['svg_sha256'] == preview.evidence()['svg_sha256']
    assert record['backup'] == 'not_applicable_new_artifact' and source.read_bytes() == before
    undo = service.plan_rollback(outcome['save_id'])
    service.confirm_rollback(undo['rollback_id'], True)
    assert not output.exists() and audit.exists() and source.read_bytes() == before
    assert json.loads(audit.read_text('utf-8'))['status'] == 'rolled_back'


def test_png_rejects_injected_image_and_modified_output(tmp_path):
    service, _, preview, receipt = fixture(tmp_path)
    spoof = RasterReceipt(preview, receipt.png, '0' * 64, 10, 10)
    with pytest.raises(DiagramSaveError):
        service.plan_png(preview, spoof)
    with pytest.raises(DiagramSaveError):
        service.plan_png(RenderedDiagram(preview.source, preview.svg, preview.execution), receipt)
    plan = service.plan_png(preview, receipt)
    outcome = service.confirm_save(plan['plan_id'], True)
    Path(outcome['path']).write_bytes(receipt.png + b'junk')
    with pytest.raises(DiagramSaveError) as denied:
        service.plan_rollback(outcome['save_id'])
    assert denied.value.code == 'export_rollback_changed'
    assert Path(outcome['path']).exists()


def test_owned_qtsvg_child_produces_host_only_receipt(tmp_path):
    _, _, preview, _ = fixture(tmp_path)
    receipt = rasterize(preview, timeout_ms=8000)
    assert receipt.preview is preview and receipt.renderer == 'host-qt-svg-v1'
    assert (receipt.width, receipt.height) == (10, 10)
    assert verify_png(receipt.png, 10, 10)
    with Image.open(BytesIO(receipt.png)) as image:
        r, g, b, a = image.convert('RGBA').getpixel((5, 5))
        assert r > 200 and g < 80 and b < 80 and a == 255


def test_png_rejects_metadata_and_dimension_mismatch():
    image = Image.new('RGB', (10, 10), 'red')
    pnginfo = PngImagePlugin.PngInfo()
    pnginfo.add_text('private', 'must disappear')
    buffer = BytesIO()
    image.save(buffer, format='PNG', pnginfo=pnginfo)
    clean, w, h = normalize_png(buffer.getvalue(), expected_size=(10, 10))
    assert b'private' not in clean and (w, h) == (10, 10)
    with pytest.raises(DiagramError):
        verify_png(buffer.getvalue(), 10, 10)
    with pytest.raises(DiagramError):
        normalize_png(clean, expected_size=(11, 10))

"""Bounded, presentation-only image history; no GUI, Office, cloud or real accounts."""
import json
from types import SimpleNamespace

import pytest

from office_agent.webui import bridge

PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WlM9x8AAAAASUVORK5CYII="


def test_bridge_images_are_bounded_display_only():
    source = {"data_url": PNG, "alt": "<img src=x onerror=alert(1)>", "caption": "c" * 1500,
              "confirmed": True, "plan_id": "do-not-save", "api_key": "do-not-save"}
    images = bridge._display_images([source] * 5)
    assert len(images) == 4
    assert images[0] == {"data_url": PNG, "alt": source["alt"], "caption": "c" * 1000}
    assert bridge._display_images({"data_url": PNG}) == []
    assert "do-not-save" not in json.dumps(images)


@pytest.mark.parametrize("url", ["https://example.invalid/image.png", "file:///private.png", "blob:unowned",
                                 "javascript:alert(1)", "data:image/svg+xml;base64,PHN2Zy8+",
                                 "data:image/png;base64,AAAA", "data:image/png;base64,%%%",
                                 PNG.replace("image/png", "image/jpeg")])
def test_bridge_rejects_non_raster_urls_and_mismatched_content(url):
    assert bridge._display_images([{"data_url": url}]) == []


def test_bridge_image_size_bound():
    assert bridge._display_images([{"data_url": "data:image/png;base64," + "A" * 13_333_377}]) == []


def test_images_persist_in_readonly_history_without_execution_fields(tmp_path):
    app = bridge.WebviewApp()
    app.path = tmp_path / "selected.xlsx"
    app.path.write_bytes(b"synthetic source anchor")
    app.editor = SimpleNamespace(space=SimpleNamespace(root=tmp_path), audit=tmp_path / ".audit")
    directory = bridge._session_root(app.editor) / "images"
    directory.mkdir(parents=True)
    app._history_dir = directory
    images = [{"data_url": PNG, "alt": "literal <script>", "caption": "synthetic", "plan_id": "never"}]
    app.emit("result", {"tool": "synthetic_preview", "text": "readonly", "images": images})
    live = next(e for e in app.poll() if e["type"] == "result")
    assert live["images"] == bridge._display_images(images)
    saved = bridge._read_history(directory)
    assert saved["events"][0]["images"] == live["images"]
    assert "never" not in json.dumps(saved)
    assert not app.pending_id and not app.last_verified_id
    assert app.path.read_bytes() == b"synthetic source anchor"


def test_image_history_does_not_bypass_existing_total_size_limit(tmp_path, monkeypatch):
    app = bridge.WebviewApp()
    app.path = tmp_path / "selected.xlsx"
    app.path.write_bytes(b"synthetic anchor")
    app.editor = SimpleNamespace(space=SimpleNamespace(root=tmp_path), audit=tmp_path / ".audit")
    app._history_dir = bridge._session_root(app.editor) / "session"
    app._history_dir.mkdir(parents=True)
    monkeypatch.setattr(bridge, "_HISTORY_LIMIT", 400)
    app._push({"type": "task", "text": "first"})
    app._push({"type": "result", "text": "image", "images": [{"data_url": PNG, "caption": "x" * 500}]})
    saved = bridge._read_history(app._history_dir)
    assert saved["truncated"]
    assert saved["events"] == [{"type": "task", "text": "first"}]
    assert (app._history_dir / "events.json").stat().st_size <= 400

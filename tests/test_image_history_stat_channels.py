from types import SimpleNamespace

import pytest

from tests.test_chat_results import saved_image


@pytest.mark.parametrize('changed_during_read', [False, True])
def test_path_and_handle_ctime_are_checked_within_their_own_channel(tmp_path, monkeypatch, changed_during_read):
    import os
    app, api, active, identifier, path, calls = saved_image(tmp_path, monkeypatch)
    original = os.fstat
    size = path.stat().st_size
    image_calls = []
    def distinct_handle_ctime(fd):
        info = original(fd)
        if info.st_size != size:
            return info
        image_calls.append(fd)
        return SimpleNamespace(st_dev=info.st_dev, st_ino=info.st_ino, st_size=info.st_size,
            st_mtime_ns=info.st_mtime_ns, st_ctime_ns=info.st_ctime_ns +
            (len(image_calls) if changed_during_read else 1), st_nlink=info.st_nlink)
    monkeypatch.setattr('office_agent.webui.image_history.os.fstat', distinct_handle_ctime)
    result = api.preview_historical_image(identifier)
    assert result['ok'] is (not changed_during_read)
    if changed_during_read:
        assert result['reason'] == 'image_history_changed'
    else:
        assert result['image']['data_url'].startswith('data:image/png;base64,')
    assert len(image_calls) == 2 and len(calls) == 1

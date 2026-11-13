import os
import uuid

import pytest

from office_agent.webui.orb_native import clamp_position, drop_candidates, parse_hotkey


@pytest.mark.parametrize('scale', [1, 1.25, 1.5, 2])
def test_ball_clamps_to_negative_monitor_and_removed_monitor(scale):
    size = int(64 * scale)
    assert clamp_position(-4000, -4000, size, size, (-1920, -200, 0, 880)) == (-1920, -200)
    assert clamp_position(9000, 9000, size, size, (0, 0, 1920, 1040)) == (1920 - size, 1040 - size)
    assert clamp_position(50, 60, 500, 800, (10, 20, 300, 400)) == (10, 20)


def test_drop_feedback_is_only_a_candidate_filter():
    assert drop_candidates(('C:\\论文 样例\\技术报告.docx',))
    assert drop_candidates(('C:\\data.xlsx', 'C:\\report.docx'))
    assert not drop_candidates(('C:\\report.docx.lnk',))
    assert not drop_candidates(('https://example.test/a.docx',))
    assert not drop_candidates(('\\\\server\\share\\a.docx',))
    assert not drop_candidates(('C:\\run.exe',))
    assert not drop_candidates(['C:\\a.docx'] * 9)
    assert not drop_candidates('C:\\a.docx')


def test_hotkey_configuration_is_bounded_and_excludes_windows_shortcuts():
    assert parse_hotkey('CTRL+ALT+SHIFT+O') == (0x4007, ord('O'))
    assert parse_hotkey('CTRL+F12') == (0x4002, 123)
    assert parse_hotkey('') is None
    for invalid in ('WIN+O', 'META+O', 'CTRL+CTRL+O', 'ALT+F13', 'O', None):
        with pytest.raises(ValueError):
            parse_hotkey(invalid)


@pytest.mark.skipif(os.name != 'nt', reason='Windows current-user mutex and event contract')
def test_same_user_second_instance_wakes_first_and_releases_handles():
    from office_agent.webui.orb_instance import SingleInstance
    name = 'OfficeAgentTest-' + uuid.uuid4().hex
    first = SingleInstance(name)
    second = None
    try:
        assert first.primary
        assert not first.poll_wake()
        second = SingleInstance(name)
        assert not second.primary
        assert first.poll_wake()
        assert not first.poll_wake()
    finally:
        if second is not None:
            second.close()
        first.close()
    third = SingleInstance(name)
    try:
        assert third.primary
    finally:
        third.close()

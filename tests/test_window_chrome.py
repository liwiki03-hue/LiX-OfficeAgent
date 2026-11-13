from types import SimpleNamespace

import pytest

from office_agent.webui.window_chrome import WindowChrome, resize_edge


@pytest.mark.parametrize('point,expected', [((0, 0), 13), ((399, 0), 14), ((0, 299), 16), ((399, 299), 17), ((0, 120), 10), ((399, 120), 11), ((120, 0), 12), ((120, 299), 15), ((200, 150), 0), ((400, 150), 0), ((-1, 0), 0)])
def test_native_resize_hit_bounds_and_negative_monitor_coordinates(point, expected):
    assert resize_edge(*point, (0, 0, 400, 300)) == expected
    assert resize_edge(point[0] - 1200, point[1] - 600, (-1200, -600, -800, -300)) == expected


def test_window_actions_use_actual_native_state_and_reject_unbounded_commands():
    calls = []
    form = SimpleNamespace(InvokeRequired=False, WindowState='Normal')
    window = SimpleNamespace(native=form, maximized=False)
    window.maximize = lambda: (calls.append('maximize'), setattr(form, 'WindowState', 'Maximized'))
    window.restore = lambda: (calls.append('restore'), setattr(form, 'WindowState', 'Normal'))
    window.minimize = lambda: calls.append('minimize')
    chrome = WindowChrome(window)
    form.WindowState = 'Maximized'
    assert chrome.state()['maximized']
    assert not chrome.action('maximize')['maximized']
    chrome.action('minimize')
    with pytest.raises(ValueError, match='不支持'):
        chrome.action('destroy')
    assert calls == ['restore', 'minimize']


def test_drag_and_resize_accept_only_named_operations_and_keep_maximized_boundary():
    window = SimpleNamespace(native=None, maximized=False)
    chrome = WindowChrome(window)
    hits = []
    chrome._move_size = hits.append
    chrome.action('drag')
    chrome.action('resize-bottom-right')
    assert hits == [2, 17]
    window.maximized = True
    with pytest.raises(ValueError, match='还原'):
        chrome.action('drag')
    with pytest.raises(ValueError, match='不支持'):
        chrome.action({'resize': [2000, 2000]})
    assert hits == [2, 17]

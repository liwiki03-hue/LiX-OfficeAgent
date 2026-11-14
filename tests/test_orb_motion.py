import math
from types import SimpleNamespace

import pytest

from office_agent.webui.orb_motion import OrbMotion, Transition
from office_agent.webui.orb_visuals import display_status, orb_image, points


class Clock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def advance(self, seconds):
        self.now += seconds


def test_delayed_ticks_preserve_time_based_phase_instead_of_tick_count():
    first_clock, second_clock = Clock(), Clock()
    first, second = OrbMotion(first_clock), OrbMotion(second_clock)
    for motion in (first, second):
        motion.update(visible=True, status='busy')
    for _ in range(120):
        first_clock.advance(.05)
        first.frame()
    second_clock.advance(6)
    assert first.frame()['phase'] == second.frame()['phase'] == 270
    assert math.isclose(first.angle, second.angle, abs_tol=1e-8)


def test_hide_stops_frames_and_does_not_advance_or_catch_up_on_restore():
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True)
    clock.advance(2)
    visible = motion.frame()
    motion.update(visible=False)
    assert motion.interval() == 0
    clock.advance(900)
    assert motion.frame()['phase'] == visible['phase']
    motion.update(visible=True)
    assert motion.frame()['phase'] == visible['phase']
    assert motion.interval() == 83


@pytest.mark.parametrize('state', ['ready', 'busy', 'opening', 'confirm', 'error', 'stopped', 'uncertain', 'complete'])
def test_reduced_motion_has_a_stable_frame_and_no_timer_in_every_state(state):
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True, status=state, reduced=True, hovered=True)
    first = motion.frame()
    clock.advance(10)
    assert motion.frame() == first
    assert motion.interval() == 0
    motion.update(expanded=True)
    assert motion.frame()['menu'] == 1
    assert not motion.frame()['menu_running']
    motion.update(expanded=False)
    assert motion.frame()['menu'] == 0


def test_hover_and_menu_reverse_from_current_position_without_jump_or_extra_timer():
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True, hovered=True, expanded=True)
    assert motion.frame()['emphasis'] == 0
    clock.advance(.08)
    current = motion.frame()['menu']
    assert 0 < current < 1
    motion.update(expanded=False)
    assert motion.frame()['menu'] == current
    clock.advance(.03)
    closing = motion.frame()['menu']
    assert 0 < closing < current
    motion.update(expanded=True)
    assert motion.frame()['menu'] == closing
    clock.advance(.24)
    assert motion.frame()['menu'] == 1
    assert motion.frame()['emphasis'] == 8
    motion.update(hovered=False, expanded=False)
    clock.advance(.3)
    assert motion.frame()['emphasis'] == 0 and motion.frame()['menu'] == 0


@pytest.mark.parametrize('status', ['confirm', 'stopped', 'error', 'uncertain', 'complete'])
def test_wait_and_terminal_states_settle_without_residual_busy_animation(status):
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True, status='busy')
    clock.advance(1)
    motion.update(status=status)
    angle = motion.frame()['phase']
    clock.advance(.8)
    assert motion.interval() == 0
    assert motion.frame()['phase'] == angle
    assert motion.frame()['pulse'] == 0


def test_dragging_pauses_visual_motion_without_changing_input_geometry():
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True, dragging=True)
    before = motion.frame()
    clock.advance(6)
    assert motion.frame() == before
    assert motion.interval() == 0
    motion.update(dragging=False)
    assert motion.interval() == 83


def test_status_never_claims_success_from_ready_or_running_state():
    assert display_status('verified', False, False, False) == 'complete'
    assert display_status('verified', True, False, False) == 'busy'
    assert display_status('drafting', False, False, False) == 'ready'
    assert display_status('verified', False, False, True) == 'confirm'


def test_dot_geometry_is_immutable_cached_and_dynamic_pixels_remain_inside_hit_circle():
    assert isinstance(points(), tuple) and points() is points()
    for size in (64, 80, 128):
        for pulse in (-5, 0, 10):
            image = orb_image(size, 'ready', 18, True, True, False, pulse, 8)
            assert all(image.getpixel(xy)[3] == 0 for xy in ((0, 0), (size - 1, 0), (0, size - 1), (size - 1, size - 1)))
            assert image.getbbox()[0] > 0 and image.getbbox()[2] < size
    assert orb_image(64, phase=1).tobytes() != orb_image(64, phase=6).tobytes()
    assert orb_image.cache_info().maxsize == 128


def test_hidden_native_visual_stops_timer_and_actions_without_redraw_or_recursion():
    from office_agent.webui.orb_morph import MorphVisual
    class Timer:
        def __init__(self):
            self.enabled = True

        def Stop(self):
            self.enabled = False
    hidden = []
    owner = SimpleNamespace(orb=SimpleNamespace(Visible=True), hidden=True, exiting=False,
                            tooltip=SimpleNamespace(Hide=lambda *args: None, SetToolTip=lambda *args: None),
                            _moved=False, _drag=None, _accept=False, _status='busy', observation=lambda *args: None)
    visual = MorphVisual.__new__(MorphVisual)
    visual.owner, visual.disposed, visual.expanded = owner, False, True
    visual.motion, visual.timer = OrbMotion(), Timer()
    visual.reduced_motion = visual.hovered = visual.focused = False
    visual.decoration = SimpleNamespace(Hide=lambda: hidden.append('decoration'))
    visual.buttons = [(SimpleNamespace(Hide=lambda: hidden.append('action')), None, None, None)]
    visual.redraw = lambda: pytest.fail('Hidden orb was redrawn')
    visual.refresh()
    assert not visual.timer.enabled and not visual.expanded
    assert hidden == ['decoration', 'action']
    assert visual.motion.interval() == 0


def test_user_cannot_override_a_system_reduction_and_preference_refreshes_immediately():
    from office_agent.webui.orb_morph import MorphVisual
    visual = MorphVisual.__new__(MorphVisual)
    visual.system_reduced = True
    visual.owner = SimpleNamespace(observation=lambda *args: None)
    refreshed = []
    visual.refresh = lambda: refreshed.append(True)
    visual.set_reduced_motion(False)
    assert visual.reduced_motion and not visual.user_reduced
    visual.system_reduced = False
    visual.set_reduced_motion(True)
    assert visual.reduced_motion and visual.user_reduced
    visual.set_reduced_motion(False)
    assert not visual.reduced_motion and len(refreshed) == 3


@pytest.mark.parametrize('animation,query,power,saving,expected', [(0, 1, 1, 0, 'system-client'), (1, 0, 1, 0, 'unavailable'),
    (1, 1, 0, 0, 'unavailable'), (1, 1, 1, 1, 'battery-saver'), (1, 1, 1, 0, 'enabled')])
def test_dedicated_windows_animation_flag_and_power_settings_are_read_only(animation, query, power, saving, expected):
    from office_agent.webui.orb_morph import motion_preference
    calls = []
    def parameters(flag, ignored, pointer, unused):
        assert flag == 0x1042 and ignored == unused == 0
        pointer._obj.value = animation
        calls.append('read-animation')
        return query
    def status(pointer):
        pointer._obj.saving = saving
        calls.append('read-power')
        return power
    user, kernel = SimpleNamespace(SystemParametersInfoW=parameters), SimpleNamespace(GetSystemPowerStatus=status)
    reduced, reason = motion_preference(SimpleNamespace(HighContrast=False, IsMenuAnimationEnabled=True), user, kernel)
    assert reason == expected and reduced == (expected != 'enabled')
    assert calls[0] == 'read-animation'


@pytest.mark.parametrize('contrast,menu,reason', [(True, True, 'high-contrast'), (False, False, 'system-menu')])
def test_system_accessibility_reduction_does_not_query_or_change_other_settings(contrast, menu, reason):
    from office_agent.webui.orb_morph import motion_preference
    reduced, actual = motion_preference(SimpleNamespace(HighContrast=contrast, IsMenuAnimationEnabled=menu), SimpleNamespace(), SimpleNamespace())
    assert reduced and actual == reason


def test_native_menu_reversal_reuses_visible_forms_without_showing_them_twice():
    from office_agent.webui.orb_morph import MorphVisual
    clock = Clock()
    motion = OrbMotion(clock)
    motion.update(visible=True, expanded=True)
    clock.advance(.08)
    motion.update(expanded=False)
    visual = MorphVisual.__new__(MorphVisual)
    activated = []
    def show_again(owner):
        pytest.fail('An already visible action window was shown again')
    window = SimpleNamespace(Visible=True, Show=show_again, Activate=lambda: activated.append(True))
    visual.owner = SimpleNamespace(orb=object(), observation=lambda *args: None,
                                   tooltip=SimpleNamespace(Hide=lambda *args: None, SetToolTip=lambda *args: None))
    visual.expanded, visual.buttons, visual.motion = False, [(window, None, None, None)] * 4, motion
    visual._place_actions = lambda: None
    visual._paint_action = lambda *args: None
    visual._animate = lambda *args: None
    visual.refresh = lambda: None
    visual.toggle_actions()
    assert visual.expanded and len(visual.buttons) == 4 and activated == [True]

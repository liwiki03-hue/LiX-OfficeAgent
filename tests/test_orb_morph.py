import math

import pytest

from office_agent.webui.orb_visuals import action_image, action_positions, display_status, orb_image


@pytest.mark.parametrize('scale', [1, 1.25, 1.5, 2])
@pytest.mark.parametrize('corner', ['top-left', 'top-right', 'bottom-left', 'bottom-right'])
def test_actions_stay_inside_work_area_without_overlap_at_corners_and_negative_monitors(scale, corner):
    area = (-1920, -200, -640, 824)
    x = area[0] if corner.endswith('left') else area[2] - 64 * scale
    y = area[1] if corner.startswith('top') else area[3] - 64 * scale
    positions = action_positions((x, y), scale, area)
    size = round(40 * scale)
    assert len(positions) == len(set(positions)) == 4
    for px, py in positions:
        assert area[0] <= px <= area[2] - size and area[1] <= py <= area[3] - size
    for index, first in enumerate(positions):
        for second in positions[index + 1:]:
            assert math.dist(first, second) >= size + 7 * scale
        assert math.dist((first[0] + size / 2, first[1] + size / 2), (x + 32 * scale, y + 32 * scale)) > (size + 64 * scale) / 2


def test_tiny_work_area_refuses_instead_of_overlapping_actions():
    with pytest.raises(ValueError, match='可用区域'):
        action_positions((0, 0), 1, (0, 0, 70, 70))


def test_orb_alpha_keeps_corners_transparent_and_allows_safe_circle_hit_target():
    normal = orb_image(64)
    assert normal.mode == 'RGBA' and normal.size == (64, 64)
    assert all(normal.getpixel(point)[3] == 0 for point in [(0, 0), (63, 0), (0, 63), (63, 63)])
    assert normal.getpixel((32, 32))[3] > 0
    assert len(set(normal.getchannel('A').tobytes())) > 80
    assert normal.tobytes() != orb_image(64, 'busy', 9).tobytes()
    assert action_image(40, 'folder-open').getpixel((0, 0))[3] == 0
    assert action_image(40, 'settings').tobytes() != action_image(40, 'history').tobytes()


def test_stop_error_and_uncertainty_are_not_running_animation_states():
    assert display_status('stopped', True, False, False) == 'stopped'
    assert display_status('open_failed', False, False, False) == 'error'
    assert display_status('uncertain', True, False, True) == 'uncertain'
    assert display_status('idle', False, True, False) == 'opening'
    assert display_status('drafting', True, False, False) == 'busy'
    assert display_status('idle', False, False, True) == 'confirm'

from office_agent.layout_checks import character_geometry, compare_layout


def test_page_edge_observation_separates_existing_and_new_and_unknown():
    class Text:
        def get_text_range(self, index, count, errors):
            return ['字', ' ', '字', '字'][index]

        def get_charbox(self, index):
            return [(10, 10, 20, 20), (0, 0, 0, 0), (-2, 10, 4, 20), (float('nan'), 0, 1, 1)][index]

    geometry = character_geometry(Text(), 4, 100, 100)
    assert geometry['outside_page_characters'] == 1
    assert geometry['checked_characters'] == 2 and geometry['unknown_boxes'] == 1
    source = {'pages': [{'text': '字', 'geometry': geometry}]}
    result = compare_layout(source, source)
    assert result['source_counts']['outside_page_characters'] == 1 and not result['increased_observations']
    result = compare_layout(source, {'pages': source['pages'] + [{'text': '', 'geometry': geometry}]})
    assert result['increased_observations'] == ['outside_page_characters', 'native_text_empty_pages']
    assert result['not_a_failure_verdict'] and result['pagination_change_allowed']

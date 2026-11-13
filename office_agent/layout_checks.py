import math


def character_geometry(textpage, count, width, height):
    outside, examples, checked, unknown = 0, [], 0, 0
    for index in range(count):
        value = textpage.get_text_range(index, 1, errors='strict')
        if not value or value.isspace() or all(ord(char) < 32 for char in value):
            continue
        box = textpage.get_charbox(index)
        if len(box) != 4 or not all(math.isfinite(value) for value in box):
            unknown += 1
            continue
        checked += 1
        left, bottom, right, top = box
        if left < -1 or bottom < -1 or right > width + 1 or top > height + 1:
            outside += 1
            if len(examples) < 4:
                examples.append({'character_index': index, 'box_points': [round(value, 2) for value in box]})
    return {'checked_characters': checked, 'unknown_boxes': unknown, 'outside_page_characters': outside,
            'examples': examples, 'tolerance_points': 1, 'coordinate_system': 'PDFium bottom-left page points'}


def compare_layout(before, after):
    def summarize(snapshot):
        return [{'page': index + 1, 'native_text_empty': not page['text'].strip(),
                 'geometry': page.get('geometry', {'status': 'not_measured'})}
                for index, page in enumerate(snapshot['pages'])]
    sides = [summarize(snapshot) for snapshot in (before, after)]
    counts = [{'outside_page_characters': sum(row['geometry'].get('outside_page_characters', 0) for row in side),
               'native_text_empty_pages': sum(row['native_text_empty'] for row in side)} for side in sides]
    return {'method': 'pdfium-character-page-boundaries-v1', 'source': sides[0], 'output': sides[1],
            'source_counts': counts[0], 'output_counts': counts[1],
            'increased_observations': [key for key in counts[0] if counts[1][key] > counts[0][key]],
            'coverage': 'native character boxes against page edge only; empty native text can be an image page; source observations are recorded separately',
            'unverified': ['table-cell overflow within page', 'image overlap or drift', 'heading orphaning', 'header/footer/page-number semantics', 'blank-page cause'],
            'not_a_failure_verdict': True, 'pagination_change_allowed': True}

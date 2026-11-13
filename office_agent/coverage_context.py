import json


def coverage_context(tool, result, arguments, *, limit=6000):
    def dump(value):
        return json.dumps(value, ensure_ascii=False, default=str)

    keys = ('source_sha256', 'read_only', 'truncated', 'next_offset', 'next_text_offset', 'next_page',
            'next_char_offset', 'slide_index', 'slide_count', 'item_count', 'title', 'title_truncated',
            'page_count', 'coordinate_system', 'ocr_performed', 'document_count', 'candidate_count',
            'semantic_available', 'retrieval', 'corpus_truncated', 'coverage_complete', 'skipped_document_count',
            'backend', 'cached_values_verified')
    view = {key: result[key] for key in keys if key in result}
    view.update(context_truncated=False, note='Only the included evidence is visible; omitted content is not empty.')

    def admit(name, values):
        view[name] = []
        count = 0
        for value in values:
            proposed = {**view, name: view[name] + [value]}
            if len(dump(proposed)) > limit - 400:
                break
            view[name].append(value)
            count += 1
        view['context_truncated'] = count < len(values)
        return count

    if tool == 'read_pptx_slide':
        items = result.get('items', [])
        count = admit('items', items)
        if count < len(items):
            view.update(next_offset=items[count]['item_index'], next_text_offset=items[count].get('text_offset', 0))
    elif tool == 'read_pptx_structure':
        count = admit('slides', result.get('slides', []))
        if view['context_truncated']:
            view['next_offset'] = arguments.get('offset', 0) + count
    elif tool == 'search_directory_documents':
        admit('hits', result.get('hits', []))
        view['inventory_omitted_from_context'] = True
        view['note'] += ' Refine the query if results were omitted; directory inventory is in the full user result.'
    elif tool == 'read_pdf':
        view['pages'] = []
        for page in result.get('pages', []):
            current = {key: page[key] for key in ('page', 'size_points', 'char_offset', 'total_chars',
                                                'text_truncated', 'text_layer_empty')}
            if page.get('runs'):
                current.update(runs=[], text_provided_in_runs=True)
                view['pages'].append(current)
                for run in page['runs']:
                    current['runs'].append(run)
                    if len(dump(view)) > limit - 400:
                        current['runs'].pop()
                        view.update(context_truncated=True, next_page=page['page'], next_char_offset=run['char_index'])
                        return dump(view)
            else:
                current['text'] = page['text']
                view['pages'].append(current)
                if len(dump(view)) > limit - 400:
                    view['pages'].pop()
                    view.update(context_truncated=True, next_page=page['page'], next_char_offset=page['char_offset'],
                                retry_max_chars=2000)
                    break
    elif tool == 'com_read' and result.get('backend') == 'libreoffice_read_only_conversion':
        if 'paragraphs' in result:
            admit('paragraphs', result['paragraphs'])
            view['tables_omitted_from_context'] = 'tables' in result
        else:
            cells = [{'sheet': sheet['sheet'], **cell} for sheet in result.get('sheets', []) for cell in sheet['cells']]
            admit('cells', cells)
        view['note'] += ' This is a bounded converted legacy snapshot, not complete document coverage or a fidelity claim.'
    else:
        return None
    encoded = dump(view)
    return encoded if len(encoded) <= limit else dump({'context_truncated': True, 'note': 'Request a smaller read; metadata exceeds context bounds.'})

from pathlib import Path

from office_agent.catalogue_metrics import catalogue_metrics


def test_real_host_initial_schema_cost_and_count_below_half(tmp_path):
    result = catalogue_metrics(tmp_path/'selected.docx')
    assert result['below_half'] is True
    assert result['baseline_tool_count'] == 48
    word = result['initial_packs']['word']
    sheet = result['initial_packs']['spreadsheet']
    assert word['cost_ratio'] < .5 and sheet['cost_ratio'] < .5
    assert word['visible_names'] == sorted({'inspect','com_read','search','batch_documents','summarize_word_document',
        'read_word_text','read_word_region','read_word_table_region','read_word_table_structure','read_word_structure',
        'read_document_image','review_document','ocr_docx_image','request_write_tools','request_read_tools'})
    assert sheet['visible_names'] == sorted({'inspect','com_read','search','batch_documents','read_excel_region',
        'summarize_excel_column','filter_sort_excel_view','read_excel_structure','read_document_image','review_document',
        'request_write_tools','request_read_tools'})
    assert all(not name.startswith('plan_') for pack in (word,sheet) for name in pack['visible_names'])

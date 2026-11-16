from office_agent.fidelity import pdf_snapshot
from office_agent.range_alignment import locate,normalized
from tests.test_fidelity import pdf_bytes


def test_real_pdf_stream_order_is_recovered_only_by_exact_visible_glyphs(tmp_path):
    source=tmp_path/'stream-order.pdf'
    source.write_bytes(pdf_bytes(text=') Tj 9 Tc (ACD) Tj ET BT /F1 12 Tf 0 Tc 29 100 Td (B'))
    page=pdf_snapshot(source,reading_order=True)['pages'][0]
    assert normalized(page['text'])!='ABCD'
    assert normalized(page['reading_text'])=='ABCD'
    match=locate([page],{'text':'ABCD'})
    assert match['status']=='unique' and match['pages']==[1] and match['method']=='pdfium-glyph-origin-v1'
    assert locate([page],{'text':'ABCE'})['status']=='not_located'


def test_glyph_order_cannot_override_ambiguity_or_remove_word_boundaries():
    context={'text':'same context'}
    pages=[{'text':'same context same context','reading_text':'same context'}]
    assert locate(pages,context)['status']=='ambiguous'
    pages=[{'text':'stream records','reading_text':'same context same context'}]
    assert locate(pages,context)['status']=='ambiguous' and not locate(pages,context)['pages']
    assert locate([{'text':'records','reading_text':'therapist'}],{'text':'the rapist'})['status']=='not_located'


def test_rotated_text_does_not_offer_horizontal_order(tmp_path):
    source=tmp_path/'rotated.pdf'
    source.write_bytes(pdf_bytes(text=') Tj ET BT /F1 12 Tf 0 1 -1 0 60 70 Tm (Rotate'))
    assert pdf_snapshot(source,reading_order=True)['pages'][0]['reading_text'] is None
    source.write_bytes(pdf_bytes(text='AC) Tj ET BT /F1 12 Tf 29 100 Td (B'))
    assert pdf_snapshot(source,reading_order=True)['pages'][0]['reading_text'] is None


def test_ordered_paragraph_context_keeps_its_method_without_overriding_ambiguity():
    context={'text':'target','prefix':'before ','suffix':' after','paragraph':{'prefix':'before ','suffix':' after'}}
    pages=[{'text':'stream records','reading_text':'before target after'}]
    assert locate(pages,context)=={'pages':[1],'status':'unique','matches':1,
        'method':'host-paragraph-context-fallback','text_order':'pdfium-glyph-origin-v1'}
    pages=[{'text':'stream records','reading_text':'before target after before target after'}]
    assert locate(pages,context)['status']=='ambiguous'


def test_disjoint_runs_in_the_same_ink_band_keep_horizontal_reading_order(tmp_path):
    source=tmp_path/'mixed-baselines.pdf'
    source.write_bytes(pdf_bytes(text=') Tj ET BT /F1 11 Tf 100 103.2 Td (right) Tj ET BT /F1 11 Tf 20 100 Td (left '))
    page=pdf_snapshot(source,reading_order=True)['pages'][0]
    assert normalized(page['reading_text'])=='left right'
    assert locate([page],{'text':'left right'})['status']=='unique'
    source.write_bytes(pdf_bytes(text=') Tj ET BT /F1 11 Tf 100 103.2 Td (right) Tj ET BT /F1 11 Tf 20 100 Td (left words '))
    assert normalized(pdf_snapshot(source,reading_order=True)['pages'][0]['reading_text'])=='left words right'
    source.write_bytes(pdf_bytes(text=') Tj ET BT /F1 11 Tf 20 103.2 Td (right) Tj ET BT /F1 11 Tf 20 100 Td (left '))
    assert [value.strip() for value in pdf_snapshot(source,reading_order=True)['pages'][0]['reading_text'].splitlines()]==['right','left']

import pytest

from office_agent.pdf_layout import unit_review


def parsed(rows):
    return {'rows':rows,'cells':[{'row':ri,'column':ci,'raw':value,'box':[ci/len(row),ri/len(rows),(ci+1)/len(row),(ri+1)/len(rows)],
        'char_indices':[ri*100+ci]} for ri,row in enumerate(rows) for ci,value in enumerate(row)]}


def test_symbols_do_not_inherit_unit_column_or_become_dimensionless():
    source = parsed([['Measurand','Units','w','k','U'],['S','%','0.3334','2.19','0.0078']])
    candidates, review = unit_review(source)
    assert candidates == []
    assert [row['value_column'] for row in review] == [3,4,5]
    assert all(row['status']=='no_confirmed_unit_scope' and not row['confirmed'] for row in review)
    assert all(row['options'][0]['status']=='ambiguous_unit_scope' for row in review)
    assert review[1]['header_cell'] == source['cells'][3]
    assert review[1]['explicit_header_units']==[]


@pytest.mark.parametrize('header,unit,status,suggested',[
    ('Latency (ms)','ms','header_and_row_unit_agreement',True),
    ('Latency (ms)','s','header_unit_conflict',False),
    ('Factor (1)','%','header_unit_conflict',False),
    ('Latency (ms) (kg)','ms','header_unit_conflict',False),
    ('Latency (ms)','','unsupported_or_missing_unit',False),
    ('Latency (ms)','unknown','unsupported_or_missing_unit',False),
])
def test_explicit_header_agreement_conflict_and_missing_values(header,unit,status,suggested):
    candidates, review = unit_review(parsed([[header,'Units'],['20',unit]]))
    assert bool(candidates) is suggested
    assert review[0]['options'][0]['status']==status
    if suggested:
        assert candidates == [{'value_column':1,'unit_column':2,'confirmed':False,'kind':'header_and_row_unit_agreement'}]


def test_multiple_unit_columns_and_missing_numeric_do_not_select_or_fill():
    candidates, review = unit_review(parsed([['Value (ms)','Units','Units'],['20','ms','ms']]))
    assert not candidates and review[0]['status']=='multiple_unit_sources_require_review'
    candidates, review = unit_review(parsed([['Value (ms)','Units'],['','ms']]))
    assert candidates==review==[]

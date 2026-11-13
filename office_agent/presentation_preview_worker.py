import json
from pathlib import Path
import sys
import unicodedata

stage='initializing'

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))


def main():
    global stage
    import pypdfium2 as pdfium
    root=Path(sys.argv[1])
    spec=json.loads((root/'spec.json').read_text(encoding='utf8'))
    pdf=root/'pdf'/'presentation.pdf'
    stage='rasterizing'
    checks=[]
    compact=lambda value:''.join(unicodedata.normalize('NFKC',value).split())
    with pdfium.PdfDocument(str(pdf)) as book:
        if len(book)!=len(spec['slides']):raise ValueError('页数不符')
        for index in range(len(book)):
            page=book[index]
            try:
                textpage=page.get_textpage()
                try:text=textpage.get_text_range()
                finally:textpage.close()
                slide=spec['slides'][index]
                expected=[slide['title'],*slide['points'],*slide['right'],*[cell for row in slide['rows'] for cell in row if cell.strip()]]
                if slide['layout']=='chart':
                    from office_agent.presentation_charts import chart_values,unit_text
                    values=chart_values(slide['rows'])
                    expected=[slide['title'],*slide['rows'][0],*[row[0] for row in slide['rows'][1:]],*[format(value.normalize(),'f') if value else '0' for value in values],unit_text(slide['chart_unit'])]
                matched=sum(compact(value) in compact(text) for value in expected)
                checks.append({'page':index+1,'expected_blocks':len(expected),'matched_blocks':matched})
                if matched!=len(expected):
                    stage='text_check_page_'+str(index+1)+'_matched_'+str(matched)+'_of_'+str(len(expected))
                    raise ValueError('渲染后文字缺失或无法核对')
                bitmap=page.render(scale=1024/page.get_width())
                try:
                    with bitmap.to_pil() as image:image.save(root/f'page-{index+1}.png',format='PNG')
                finally:bitmap.close()
            finally:page.close()
        (root/'report.json').write_text(json.dumps({'pages':len(book),'preview_width':1024,'text_checks':checks,'visual_review_required':True,'coverage':'real PPTX reopen, PDF text presence and fixed renderer pages; text fit estimates are not PowerPoint proof'},ensure_ascii=False),encoding='utf8')


if __name__=='__main__':
    try:main()
    except Exception as error:
        import traceback
        detail={'stage':stage,'type':type(error).__name__,'code':getattr(error,'code',None),'frames':[{'function':frame.name,'line':frame.lineno} for frame in traceback.extract_tb(error.__traceback__)]}
        if isinstance(error,ModuleNotFoundError):detail['module']=error.name
        (Path(sys.argv[1])/'failure.json').write_text(json.dumps(detail),encoding='utf8')
        raise SystemExit(1)

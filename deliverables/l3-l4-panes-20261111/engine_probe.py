import base64
import io
import json
import os
from pathlib import Path
import sys


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0,str(ROOT))
sys.path.insert(1,str(ROOT/'tests'))


def worker():
    runtime=(HERE/'probe-runtime').resolve()
    runtime.relative_to(HERE.resolve())
    runtime.mkdir(exist_ok=True)
    os.environ['TEMP']=os.environ['TMP']=str(runtime)
    from docx import Document
    from docx.shared import Inches
    from playwright.sync_api import sync_playwright
    from tests.test_paged_browser import serve_static
    from tests.test_professional_browser import EDGE
    book = Document();book.add_paragraph('列宽诊断正文，数字120不变。')
    book.sections[0].header.paragraphs[0].text='自建列宽诊断页眉'
    table = book.add_table(rows=2,cols=3);table.autofit=False
    for column in table.columns:
        column.width=Inches(1.5)
    for i,row in enumerate(table.rows):
        for j,cell in enumerate(row.cells):
            cell.text=f'行{i+1}列{j+1}';cell.width=Inches(1.5)
    stream=io.BytesIO();book.save(stream);original=stream.getvalue()
    (HERE/'probe-source.docx').write_bytes(original)
    with serve_static() as url,sync_playwright() as pw:
        browser=pw.chromium.launch(executable_path=str(EDGE),headless=True,args=['--disable-background-networking','--disable-gpu','--disable-extensions','--renderer-process-limit=1'])
        try:
            page=browser.new_page();page.set_default_timeout(15000)
            page.route('**/*',lambda route:route.continue_() if route.request.url.startswith(url) else route.abort())
            page.goto(url+'paged_editor_frame.html')
            result=page.evaluate('''async(data)=>{
              const blob=new Blob([Uint8Array.from(atob(data),c=>c.charCodeAt(0))]);
              const Type=typeof SuperDoc==='function'?SuperDoc:SuperDoc.SuperDoc;
              window.probeEngine=await new Promise((resolve,reject)=>{let instance;const timer=setTimeout(()=>reject(Error('open deadline')),12000);instance=new Type({selector:'#editor',document:blob,documentMode:'editing',telemetry:{enabled:false},onReady:()=>{clearTimeout(timer);resolve(instance);},onContentError:()=>reject(Error('content rejected'))});});
              const editor=probeEngine.activeEditor;let table,position;
              editor.state.doc.descendants((node,pos)=>{if(node.type.name==='table'){table=node;position=pos;return false;}});
              const before=table.toJSON(),selected=OfficeTableColumnLayout.capture(editor.state.doc,position,0,0),size=OfficeTableColumnLayout.validate(editor,selected,0,120);let receipt,error,proof;
              try{receipt=editor.doc.tables.setColumnWidth({nodeId:table.attrs.sdBlockId,columnIndex:0,widthPt:120});}catch(e){error=String(e.message);}
              if(receipt?.success)proof=OfficeTableColumnLayout.verify(selected,editor.state.doc,size);
              const after=editor.state.doc.nodeAt(position).toJSON();
              const output=await probeEngine.export({triggerDownload:false});
              const bytes=await new Promise(resolve=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result.split(',')[1]);reader.readAsDataURL(output);});
              return {before,after,receipt,error,proof,data:bytes};
            }''',base64.b64encode(original).decode())
            candidate=base64.b64decode(result.pop('data'));(HERE/'probe-candidate.docx').write_bytes(candidate)
            reopened=Document(io.BytesIO(candidate))
            result['export_widths']=[[cell.width.twips if cell.width else None for cell in row.cells] for row in reopened.tables[0].rows]
            result['original_widths']=[[cell.width.twips if cell.width else None for cell in row.cells] for row in book.tables[0].rows]
            result['texts_preserved']=[p.text for p in reopened.paragraphs]==[p.text for p in book.paragraphs] and [[c.text for c in r.cells] for r in reopened.tables[0].rows]==[[c.text for c in r.cells] for r in table.rows]
            result['header_preserved']=reopened.sections[0].header.paragraphs[0].text==book.sections[0].header.paragraphs[0].text
            result['source_unchanged']=(HERE/'probe-source.docx').read_bytes()==original
            page.goto(url+'paged_editor_frame.html')
            page.evaluate('''async(data)=>{
              let sequence=0;const pending=new Map(),channel=new MessageChannel();window.frameRevision=0;
              const connected=new Promise(resolve=>channel.port1.onmessage=event=>{const value=event.data;if(value.type==='connected')resolve();if(value.type==='changed')window.frameRevision=value.value.revision;const request=pending.get(value.id);if(request){clearTimeout(request.timer);pending.delete(value.id);value.ok?request.resolve(value.result):request.reject(Error(value.error));}});
              channel.port1.start();window.postMessage({type:'office-editor-connect'},'*',[channel.port2]);await connected;
              window.frameRequest=(action,value={})=>new Promise((resolve,reject)=>{const id=++sequence;const timer=setTimeout(()=>reject(Error('frame deadline')),12000);pending.set(id,{resolve,reject,timer});channel.port1.postMessage({id,action,value});});
              await frameRequest('open',{data,revision:0,styles:[{id:'Normal',name:'正文'}],fonts:{fonts:[]},theme:'dark'});await frameRequest('enable',{editable:true});
            }''',base64.b64encode(original).decode())
            page.get_by_text('行2列1',exact=True).first.click();page.keyboard.press('End');page.keyboard.insert_text('人工草稿')
            baseline=page.evaluate("async()=>await frameRequest('snapshot',{revision:frameRevision})")
            assert baseline['revision']==1
            baseline_book=Document(io.BytesIO(base64.b64decode(baseline['data'])))
            page.get_by_text('行1列1',exact=True).first.click()
            page.locator('#table-column-width').click()
            page.get_by_role('spinbutton',name='表格列宽（磅）').fill('150')
            page.get_by_role('dialog').get_by_role('button',name='取消',exact=True).click()
            assert page.evaluate('frameRevision')==1
            page.locator('#table-column-width').click()
            page.get_by_role('spinbutton',name='表格列宽（磅）').fill('120')
            page.get_by_role('dialog').get_by_role('button',name='采用到草稿',exact=True).click()
            page.get_by_role('dialog').wait_for(state='detached')
            edited=page.evaluate("async()=>await frameRequest('snapshot',{revision:frameRevision})")
            assert edited['revision']==2
            page.evaluate("frameRequest('undo')")
            undone=page.evaluate("async()=>await frameRequest('snapshot',{revision:frameRevision})")
            page.evaluate("frameRequest('redo')")
            redone=page.evaluate("async()=>await frameRequest('snapshot',{revision:frameRevision})")
            frame_outputs={}
            for name,value in [('baseline',baseline),('candidate',edited),('undo',undone),('redo',redone)]:
                data=base64.b64decode(value['data']);(HERE/('probe-ui-'+name+'.docx')).write_bytes(data)
                reopened_frame=Document(io.BytesIO(data))
                frame_outputs[name]={'revision':value['revision'],'widths':[[cell.width.twips for cell in row.cells] for row in reopened_frame.tables[0].rows],
                                     'header_preserved':reopened_frame.sections[0].header.paragraphs[0].text==book.sections[0].header.paragraphs[0].text,
                                     'texts_preserved':[p.text for p in reopened_frame.paragraphs]==[p.text for p in baseline_book.paragraphs] and [[c.text for c in r.cells] for r in reopened_frame.tables[0].rows]==[[c.text for c in r.cells] for r in baseline_book.tables[0].rows]}
            assert frame_outputs['baseline']['widths']==frame_outputs['undo']['widths']==[[2160,2160,2160],[2160,2160,2160]]
            assert frame_outputs['candidate']['widths']==frame_outputs['redo']['widths']==[[2400,2160,2160],[2400,2160,2160]]
            assert all(value['texts_preserved'] and value['header_preserved'] for value in frame_outputs.values())
            result['iframe_code_checks']=frame_outputs
            result['iframe_cancel_unchanged']=True
            result['preceding_manual_draft_preserved']=True
            page.evaluate("frameRequest('close')")
            result['kind']='isolated headless engine code diagnostic, not native/Office acceptance'
            (HERE/'ENGINE_PROBE.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
            assert result['receipt']['success'] and result['texts_preserved'] and result['header_preserved'] and result['source_unchanged']
            assert result['export_widths']==[[2400,2160,2160],[2400,2160,2160]]
        finally:
            browser.close()


def run_probe():
    from office_agent.diagram_process import _run_owned, _WindowsJob
    class ProbeJob(_WindowsJob):
        def __init__(self):
            super().__init__()
            info=self.job.QueryInformationJobObject(self.handle,self.job.JobObjectExtendedLimitInformation)
            info['JobMemoryLimit']=2*1024*1024*1024
            info['BasicLimitInformation']['ActiveProcessLimit']=32
            self.job.SetInformationJobObject(self.handle,self.job.JobObjectExtendedLimitInformation,info)
    result=_run_owned((sys.executable,str(Path(__file__).resolve()),'worker'),timeout_ms=30000,_job_factory=ProbeJob)
    record={key:getattr(result,key) for key in ('code','exit_code','empty_confirmed')}
    record['limits']={'memory_bytes':2*1024*1024*1024,'processes':32,'timeout_ms':30000,'scope':'this headless diagnostic only; production limits unchanged'}
    (HERE/'ENGINE_PROCESS.json').write_text(json.dumps(record,indent=2),encoding='utf8')
    if result.empty_confirmed:
        import shutil
        runtime=(HERE/'probe-runtime').resolve()
        runtime.relative_to(HERE.resolve())
        if runtime.is_dir():
            shutil.rmtree(runtime)
    return record


if __name__=='__main__' and sys.argv[1:]==['worker']:
    try:
        worker()
    except Exception:
        import traceback
        (HERE/'ENGINE_ERROR.txt').write_text(traceback.format_exc()[-6000:],encoding='utf8')
        raise
elif __name__=='__main__':
    print(json.dumps(run_probe()))

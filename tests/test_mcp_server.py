import asyncio
import json
from pathlib import Path
import subprocess
import sys
import time

from docx import Document
import pytest

from office_agent.auth import LocalAuth
from office_agent.mcp_host import McpHost, TOOLS, HostError

ROOT = Path(__file__).resolve().parents[1]


def make(root, monkeypatch):
    auth = LocalAuth(root / 'accounts')
    assert auth.register('testuser', 'Testpass123')['ok']
    assert auth.login('testuser', 'Testpass123')['authenticated']
    source = root / 'owned.docx'
    document = Document()
    document.add_paragraph('PRIVATE-ORIGINAL original observation.')
    document.save(source)
    host = McpHost(auth)
    host.authorize(source)
    from test_professional_edits import fake_render
    monkeypatch.setattr('office_agent.advanced_writes.render_gate', fake_render)
    return host, source


def test_read_metadata_plan_no_write_and_local_confirm_cancel_undo(tmp_path, monkeypatch):
    host, source = make(tmp_path, monkeypatch)
    original = source.read_bytes()
    before = set(tmp_path.iterdir())
    assert host.call('inspect_document', {})['result']['kind'] == 'word'
    assert host.call('word_structure', {})['result']['paragraph_count'] == 1
    found = host.call('search_current_document', {'query': 'observation'})
    assert found['result']['hit_count'] == 1 and not found['result']['content_returned']
    planned = host.call('plan_word_replace', {'paragraph_index': 0, 'old': 'original', 'new': 'revised'})
    assert planned['result']['wrote_nothing'] and planned['result']['confirmation_required']
    assert 'PRIVATE-ORIGINAL' not in json.dumps([found, planned]) and str(tmp_path) not in json.dumps(planned)
    assert set(tmp_path.iterdir()) == before and source.read_bytes() == original
    identifier = planned['result']['plan_id']
    assert host.call('confirm', {'plan_id': identifier, 'approved': True})['reason'] == 'invalid_arguments'
    assert host.confirm_local(identifier, False)['status'] == 'cancelled'
    assert set(tmp_path.iterdir()) == before
    planned = host.call('plan_word_replace', {'paragraph_index': 0, 'old': 'original', 'new': 'revised'})
    identifier = planned['result']['plan_id']
    saved = host.confirm_local(identifier, True)
    output = Path(saved['output'])
    assert output.exists() and 'revised observation' in Document(output).paragraphs[0].text
    assert source.read_bytes() == original
    host.rollback_local(identifier, True)
    assert not output.exists() and source.read_bytes() == original


@pytest.mark.parametrize('tool,args', [
    ('inspect_document', {'path': 'C:/private/document.docx'}),
    ('plan_word_replace', {'paragraph_index': True, 'old': 'original', 'new': 'new'}),
    ('search_current_document', {'query': 'x' * 201}),
    ('search_current_document', {'query': {'nested': 'argument'}}),
    ('search_current_document', {'query': 'api_key=private-key-test-value'}),
    ('execute', {'confirmed': True}),
])
def test_scalar_schema_privacy_and_remote_authority_refused(tmp_path, monkeypatch, tool, args):
    host, source = make(tmp_path, monkeypatch)
    result = host.call(tool, args)
    assert result == {'ok': False, 'reason': 'invalid_arguments'}
    assert len(list(tmp_path.glob('*.docx'))) == 1 and not host.pending


def test_auth_epoch_expiry_stale_source_and_plan_capacity(tmp_path, monkeypatch):
    host, source = make(tmp_path, monkeypatch)
    start = host.clock()
    now = [start]
    host.clock = lambda: now[0]
    ids = []
    for index in range(8):
        ids.append(host.call('plan_word_replace', {'paragraph_index': 0, 'old': 'original', 'new': str(index)})['result']['plan_id'])
    assert host.call('plan_word_replace', {'paragraph_index': 0, 'old': 'original', 'new': 'ninth'})['reason'] == 'plan_capacity'
    now[0] += 301
    with pytest.raises(HostError, match='plan_expired'):
        host.preview(ids[0])
    assert not host.pending
    now[0] += 600
    assert host.call('inspect_document', {})['reason'] == 'authorization_expired'
    host.authorize(source)
    source.write_bytes(b'changed')
    assert host.call('inspect_document', {})['reason'] == 'source_changed'
    host.logout()
    assert host.call('inspect_document', {})['reason'] == 'not_authenticated'
    assert host.editor is None and not host.pending and not host.history


def test_excel_plan_uses_host_sheet_index_and_rejects_formula(tmp_path, monkeypatch):
    from openpyxl import Workbook
    host, _ = make(tmp_path, monkeypatch)
    source = tmp_path / 'owned.xlsx'
    workbook = Workbook()
    workbook.active.title = 'PRIVATE-SHEET'
    workbook.active['A1'] = 'original'
    workbook.save(source)
    workbook.close()
    host.authorize(source)
    assert not host.call('plan_excel_set', {'sheet_index': 0, 'cell': 'A1', 'value': '=1+2'})['ok']
    result = host.call('plan_excel_set', {'sheet_index': 0, 'cell': 'A1', 'value': 'replacement'})
    assert result['result']['wrote_nothing'] and 'PRIVATE-SHEET' not in json.dumps(result)
    host.logout()
    assert not host.pending


def test_uncertain_local_write_is_sticky_and_never_remote_retry(tmp_path, monkeypatch):
    host, source = make(tmp_path, monkeypatch)
    identifier = host.call('plan_word_replace', {'paragraph_index': 0, 'old': 'original', 'new': 'revised'})['result']['plan_id']
    def failed(*args, **kwargs):
        raise RuntimeError('PRIVATE-ORIGINAL secret detail')
    monkeypatch.setattr(host.editor, 'execute', failed)
    with pytest.raises(HostError, match='outcome_uncertain'):
        host.confirm_local(identifier, True)
    host.revoke()
    host.authorize(source)
    assert host.call('inspect_document', {})['reason'] == 'outcome_uncertain'


def test_headless_real_stdio_requires_human_and_never_exposes_write(tmp_path):
    async def probe():
        from mcp import ClientSession, StdioServerParameters
        from mcp.client.stdio import stdio_client
        with (tmp_path / 'stderr.txt').open('w') as sink:
            async with stdio_client(StdioServerParameters(command=sys.executable,
                    args=['-m','office_agent.mcp_server','--headless'], cwd=str(ROOT)), errlog=sink) as streams:
                async with ClientSession(*streams) as session:
                    await session.initialize()
                    listed = await session.list_tools()
                    assert {tool.name for tool in listed.tools} == set(TOOLS)
                    for tool in listed.tools:
                        reply = await session.call_tool(tool.name, {})
                        assert reply.isError and json.loads(reply.content[0].text)['reason'] == 'not_authenticated'
    asyncio.run(asyncio.wait_for(probe(), timeout=20))
    assert 'PRIVATE-' not in (tmp_path / 'stderr.txt').read_text()


def test_authenticated_real_stdio_metadata_plan_and_cancel_fixture(tmp_path):
    transcript = []
    worker = tmp_path / 'owned_worker.py'
    worker.write_text('''import asyncio,sys,json
from pathlib import Path
sys.path.insert(0,sys.argv[1])
from office_agent.mcp_host import McpHost
from office_agent.mcp_server import serve
from office_agent.auth import LocalAuth
from docx import Document
root=Path(sys.argv[2])
auth=LocalAuth(root/'auth-fixture')
auth.register('testuser','Testpass123')
auth.login('testuser','Testpass123')
path=root/'owned.docx'
document=Document()
document.add_paragraph('PRIVATE-ORIGINAL original observation.')
document.save(path)
host=McpHost(auth)
host.authorize(path)
asyncio.run(serve(host))
''', encoding='utf8')
    async def probe():
        from mcp import ClientSession, StdioServerParameters
        from mcp.client.stdio import stdio_client
        with (tmp_path / 'stderr.txt').open('w') as sink:
            async with stdio_client(StdioServerParameters(command=sys.executable, args=[str(worker),str(ROOT),str(tmp_path)], cwd=str(ROOT)), errlog=sink) as streams:
                async with ClientSession(*streams) as session:
                    await session.initialize()
                    reply = await session.call_tool('word_structure', {})
                    transcript.append({'tool':'word_structure','reply':json.loads(reply.content[0].text)})
                    assert json.loads(reply.content[0].text)['result']['paragraph_count'] == 1
                    reply = await session.call_tool('plan_word_replace', {'paragraph_index':0,'old':'original','new':'revised'})
                    data = json.loads(reply.content[0].text)
                    transcript.append({'tool':'plan_word_replace','reply':data})
                    assert data['result']['wrote_nothing'] and data['result']['confirmation_required']
                    reply = await session.call_tool('confirm', {'plan_id':data['result']['plan_id'],'confirmed':True})
                    transcript.append({'tool':'confirm','reply':json.loads(reply.content[0].text)})
                    assert reply.isError
                    assert 'PRIVATE-ORIGINAL' not in str(reply.content) and str(tmp_path) not in str(reply.content)
    asyncio.run(asyncio.wait_for(probe(), timeout=20))
    assert not (tmp_path / '.office_agent').exists()
    assert len(list(tmp_path.glob('*.docx'))) == 1
    assert Document(tmp_path / 'owned.docx').paragraphs[0].text == 'PRIVATE-ORIGINAL original observation.'
    (tmp_path / 'MCP_TRANSCRIPT.json').write_text(json.dumps(transcript,indent=2),encoding='utf8')


@pytest.mark.parametrize('payload', [b'PRIVATE-ORIGINAL' * 1000 + b'\n', b'\xff\n'])
def test_malformed_or_oversize_stdio_closes_without_echo(payload):
    result = subprocess.run([sys.executable, '-m', 'office_agent.mcp_server', '--headless'], input=payload,
                            capture_output=True, cwd=ROOT, timeout=10)
    assert result.returncode == 0 and b'PRIVATE-ORIGINAL' not in result.stdout + result.stderr


@pytest.mark.skipif(sys.platform != 'win32', reason='Owned native control acceptance uses the existing Windows Tk; no runtime is installed')
def test_real_tk_control_can_login_revoke_and_close_owned_hidden_window(tmp_path, monkeypatch):
    import threading
    from office_agent.mcp_server import control_window
    host, source = make(tmp_path, monkeypatch)
    host.logout()
    root = control_window(host, threading.Event())
    root.withdraw()
    try:
        def descendants(widget):
            return [child for value in widget.winfo_children() for child in [value, *descendants(value)]]
        children = descendants(root)
        entries = [value for value in children if value.winfo_class() == 'TEntry']
        entries[0].insert(0, 'testuser')
        entries[1].insert(0, 'Testpass123')
        buttons = {value.cget('text'): value for value in children if value.winfo_class() == 'TButton'}
        buttons['登录'].invoke()
        assert host.auth.status()['authenticated'] and entries[1].get() == ''
        host.authorize(source)
        buttons['退出并撤销授权'].invoke()
        assert not host.auth.status()['authenticated'] and host.path is None
        root.update_idletasks()
    finally:
        root.destroy()

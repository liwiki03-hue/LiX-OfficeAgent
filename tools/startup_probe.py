"""Exact offline BAT diagnostic, with a self-contained synthetic controller hook.

Real app entry/runtime/HTML/WebView2; NOT actual accounts, provider or business
acceptance. A suspended child enters a private Job before any code executes.
The hook must install or the app process exits: no unhooked business fallback.
"""
from pathlib import Path
import json
import os
import shutil
import sys
import tempfile
import time
import uuid

ROOT = Path(__file__).resolve().parents[1]
HOOK_PREFIX = '.next-p1-hook-'
# Only prepare_hook's mkdtemp creations enter this private, process-local registry.
_OWNED_HOOKS = {}
HOOK = r'''
import importlib, json, os, sys, threading, time

def install():
    original = list(sys.orig_argv)
    if '-m' not in original:
        return
    index = original.index('-m')
    module = original[index + 1]
    if module != 'office_agent.app':
        return  # Runtime readiness command is deliberately not replaced.
    if original[index + 2:] != ['--ui', 'web']:
        raise RuntimeError('Unexpected app invocation')
    bridge = importlib.import_module('office_agent.webui.bridge')
    class EmptyApi:
        def __init__(self, app): pass
        def auth_status(self): return {'ok': True, 'authenticated': False}
        def poll(self): return []
    class EmptyApp:
        def __init__(self, **kwargs): pass
        def attach(self, window):
            finished, lock = threading.Event(), threading.Lock()
            def finish(data):
                with lock:
                    if finished.is_set(): return
                    try:
                        data.update(synthetic_controller=True, hook_applied=True,
                            task_submitted=False, document_selected=False, provider_called=False)
                        with open(os.environ['OFFICE_AGENT_STARTUP_PROBE_REPORT'], 'x', encoding='utf-8') as stream:
                            json.dump(data, stream, ensure_ascii=False)
                    finally:
                        finished.set()
                        window.destroy()
            def loaded():
                try:
                    deadline = time.monotonic() + 8
                    while not finished.is_set():
                        signal, values = threading.Event(), []
                        def done(value): values.append(value); signal.set()
                        window.evaluate_js("""(()=>{
                            const auth=document.getElementById('auth-screen'), workspace=document.getElementById('workspace');
                            return {ready:document.readyState,input:!!document.getElementById('prompt'),
                              css:document.styleSheets.length>0,bridge:!!window.pywebview?.api,
                              auth_visible:!!auth&&!auth.hidden&&getComputedStyle(auth).display!=='none',
                              workspace_visible:!!workspace&&!workspace.hidden,
                              width:innerWidth,height:innerHeight};
                        })()""", callback=done)
                        if not signal.wait(max(0, deadline-time.monotonic())):
                            raise TimeoutError('DOM callback deadline')
                        result = values[0]
                        if (type(result) is dict and result.get('ready') == 'complete'
                            and all(result.get(key) is True for key in ('input','css','bridge','auth_visible'))
                            and result.get('workspace_visible') is False):
                            finish({'ok': True, **result}); return
                        if time.monotonic() >= deadline: raise TimeoutError('DOM readiness deadline')
                        finished.wait(0.05)
                except BaseException:
                    finish({'ok': False, 'code': 'startup_dom_probe_failed'})
            def watchdog():
                if not finished.wait(15): finish({'ok':False, 'code':'startup_window_deadline'})
            window.events.loaded += loaded
            threading.Thread(target=watchdog, daemon=True).start()
    bridge.WebviewApp, bridge.Api = EmptyApp, EmptyApi

try:
    install()
except BaseException:
    os._exit(97)  # sitecustomize errors otherwise only print and continue unhooked.
'''


def prepare_hook(base):
    base = Path(base)
    if base.is_symlink() or getattr(base, 'is_junction', lambda: False)():
        raise RuntimeError('Hook base must not be a link')
    base = base.resolve()
    base.mkdir(exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=HOOK_PREFIX, dir=base)).resolve()
    if temporary.parent != base or not temporary.name.startswith(HOOK_PREFIX):
        raise RuntimeError('Unexpected hook location')
    info = temporary.lstat()
    _OWNED_HOOKS[temporary] = (info.st_dev, info.st_ino)
    try:
        (temporary / 'sitecustomize.py').write_text(HOOK, encoding='utf-8')
    except BaseException:
        remove_hook(temporary, base)
        raise
    return temporary


def remove_hook(temporary, base):
    temporary, base = Path(temporary), Path(base).resolve()
    if (temporary.is_symlink() or getattr(temporary, 'is_junction', lambda: False)()
            or temporary.resolve().parent != base
            or not temporary.name.startswith(HOOK_PREFIX)):
        raise RuntimeError('Refuse cleanup outside the exact owned hook directory')
    resolved = temporary.resolve()
    if resolved not in _OWNED_HOOKS:
        raise RuntimeError('Refuse cleanup of an unregistered hook directory')
    info = resolved.lstat()
    if (info.st_dev, info.st_ino) != _OWNED_HOOKS[resolved]:
        raise RuntimeError('Refuse cleanup of a replaced hook directory')
    # Re-verify the absolute intended owned target immediately BEFORE deletion.
    # These checks are not OS-atomic or protection against same-privilege attackers.
    if (resolved != temporary.absolute() or resolved.resolve() != resolved
            or resolved.parent != base or resolved.is_symlink()
            or getattr(resolved, 'is_junction', lambda: False)()
            or (resolved.lstat().st_dev, resolved.lstat().st_ino) != _OWNED_HOOKS[resolved]):
        raise RuntimeError('Refuse cleanup of a changed hook directory')
    shutil.rmtree(resolved)
    removed = not resolved.exists()
    if removed:
        del _OWNED_HOOKS[resolved]
    return removed


def close_child_handle(child, report):
    """Record handle-close failure without blocking hook cleanup or evidence output."""
    try:
        if child is not None:
            child.close()
        report['child_handle_closed'] = True
    except BaseException as exc:
        report['child_handle_closed'] = False
        report['child_close_error'] = f'{type(exc).__name__}: {exc}'
        report['ok'] = False


if __name__ == '__main__':
    sys.path.insert(0, str(ROOT/'examples'))
    from office_owned_launcher import NativeJob
    target = ROOT/'deliverables'/('NEXT_P1-'+uuid.uuid4().hex[:10]+'.worker.json')
    logdir = ROOT/'.startup-logs'
    before = set(logdir.glob('*.log'))
    temporary = prepare_hook(ROOT/'deliverables')
    environment = {'PYTHONPATH':str(temporary),
        'OFFICE_AGENT_STARTUP_PROBE_REPORT':str(target),'OFFICE_AGENT_NO_PAUSE':'1','UV_OFFLINE':'1',
        'OFFICE_AGENT_REAL_COM':'0','OFFICE_WEBVIEW_GUI':'edgechromium','PYTHONNET_RUNTIME':'netfx'}
    previous = {key:os.environ.get(key) for key in environment}
    direct = '--direct-empty-window' in sys.argv
    report = {'ok':False,'exact_entry':None if direct else '一键启动.bat',
        'entry':'existing interpreter -m office_agent.app' if direct else 'one-click BAT',
        'offline_enforced':True,'synthetic_controller':True,'real_authentication':False,
        'user_tasks':False,'user_documents':False,'real_office':False,'pixel_visual_acceptance':False}
    job = None
    child = None
    started = time.monotonic()
    try:
        job = NativeJob()
        os.environ.update(environment)
        if direct:
            import importlib.metadata
            report['runtime']={'ok':True,'versions':{'pythonnet':importlib.metadata.version('pythonnet')},
                'evidence':'Parent environment metadata only; actual native surface checked separately.'}
            command = [sys.executable,'-m','office_agent.app','--ui','web']
        else:
            command = [str(Path(os.environ['SYSTEMROOT'])/'System32'/'cmd.exe'),
                '/d','/c','call',str(ROOT/'一键启动.bat')]
        child = job.spawn_worker(command)
        report['owned_pid'] = child.pid
        report['exit_code'] = child.wait(timeout=40)
        if target.exists(): report['window'] = json.loads(target.read_text(encoding='utf-8'))
        logs = set(logdir.glob('*.log'))-before
        report['startup_logs'] = [str(p) for p in sorted(logs)]
        if len(logs) == 1:
            text = next(iter(logs)).read_text(encoding='utf-8',errors='replace')
            offset = text.find('{')
            if offset >= 0: report['runtime'] = json.JSONDecoder().raw_decode(text[offset:])[0]
        report['ok'] = (report.get('exit_code') == 0 and report.get('window',{}).get('ok') is True
            and report.get('window',{}).get('hook_applied') is True
            and report.get('runtime',{}).get('ok') is True
            and report['runtime']['versions']['pythonnet'] == '3.0.5')
    except BaseException as exc:
        report['error'] = f'{type(exc).__name__}: {exc}'
    finally:
        for key,value in previous.items():
            if value is None: os.environ.pop(key,None)
            else: os.environ[key] = value
        try:
            if job is not None: report['cleanup'] = job.close()
        except BaseException as exc:
            report['cleanup_error'] = f'{type(exc).__name__}: {exc}'
            report['ok'] = False
        finally:
            close_child_handle(child, report)
        report['empty_confirmed'] = job is None or job.empty_confirmed
        try:
            if not report['empty_confirmed']:
                raise RuntimeError('Retain hook until owned cleanup is confirmed')
            report['hook_directory_removed'] = remove_hook(temporary,ROOT/'deliverables')
        except BaseException as exc:
            report['hook_cleanup_error'] = f'{type(exc).__name__}: {exc}'
            report['hook_directory_removed'] = False
        report['ok'] = report['ok'] and report['empty_confirmed'] and report['hook_directory_removed']
        report['elapsed_seconds'] = round(time.monotonic()-started,2)
    output = target.with_name(target.stem.replace('.worker','')+'.json')
    with output.open('x',encoding='utf-8') as stream: json.dump(report,stream,ensure_ascii=False,indent=2)
    print(json.dumps({'report':str(output),**report},ensure_ascii=True),flush=True)
    raise SystemExit(0 if report['ok'] else 1)

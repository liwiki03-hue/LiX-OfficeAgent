"""One-click branch coverage with deny-only uv stub; no app/Office is launched by these tests."""
import os
from pathlib import Path
import subprocess
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
BAT = ROOT / '一键启动.bat'


def test_native_entry_preserves_fallback_and_uses_locked_offline_cache():
    source = BAT.read_text(encoding='utf-8')
    raw = BAT.read_bytes()
    assert raw.startswith(b'@echo off\r\n')
    assert raw.count(b'\r\n') == raw.count(b'\n')  # CMD + UTF-8 requires CRLF, not LF.
    assert 'call uv sync' in source
    assert 'call uv run --offline --isolated --with "pythonnet==3.0.5" python -m office_agent.webui.runtime' in source
    assert 'call uv run --offline --isolated --with "pythonnet==3.0.5" python -m office_agent.app --ui web' in source
    assert source.count('--offline --isolated --with "pythonnet==3.0.5"') == 2
    assert 'OFFICE_WEBVIEW_GUI=edgechromium' in source and 'PYTHONNET_RUNTIME=netfx' in source
    for forbidden in ('powershell', 'Set-ExecutionPolicy', 'taskkill', 'curl ', 'Invoke-WebRequest'):
        assert forbidden.lower() not in source.lower().replace('no powershell', '')
    assert 'application_exit=%launch_exit%' in source
    assert 'endlocal & exit /b %launch_exit%' in source
    assert (ROOT / 'Start-OfficeAgent.cmd').is_file()
    assert (ROOT / 'Start-OfficeAgent.ps1').is_file()
    assert "pythonnet==3.0.5; sys_platform == 'win32'" in (ROOT / 'pyproject.toml').read_text(encoding='utf-8')


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows CMD launcher branches require Windows')
@pytest.mark.parametrize('has_venv,fail_stage,expected', [(False, '', 0), (True, '', 0),
    (False, 'sync', 39), (True, 'readiness', 40), (True, 'app', 41), (True, 'missing', 2), (True, 'log', 3)])
def test_cmd_branches_with_chinese_space_path_and_no_real_uv(tmp_path, has_venv, fail_stage, expected):
    project = tmp_path / '中文 project space'; project.mkdir()
    (project / BAT.name).write_bytes(BAT.read_bytes())
    if has_venv:
        scripts=project/'.venv'/'Scripts'; scripts.mkdir(parents=True)
        (scripts/'python.exe').write_bytes(b'not executable; never invoked')
    if fail_stage=='log':(project/'.startup-logs').write_text('owned blocker fixture',encoding='ascii')
    bin_dir=tmp_path/'stub bin'; bin_dir.mkdir()
    trace=tmp_path/'trace.txt'
    (bin_dir/'uv.cmd').write_text('''@echo off
>>"%TRACE%" echo %*
>>"%TRACE%" echo cwd=%CD%
>>"%TRACE%" echo gui=%OFFICE_WEBVIEW_GUI%;runtime=%PYTHONNET_RUNTIME%
if "%FAIL_STAGE%"=="sync" if "%~1"=="sync" exit /b 39
if "%FAIL_STAGE%"=="readiness" if "%~8"=="office_agent.webui.runtime" exit /b 40
if "%FAIL_STAGE%"=="app" if "%~8"=="office_agent.app" exit /b 41
exit /b 0
''', encoding='ascii')
    env=os.environ.copy()
    env.update(OFFICE_AGENT_NO_PAUSE='1', TRACE=str(trace), FAIL_STAGE=fail_stage,
               OFFICE_WEBVIEW_GUI='sentinel-old', PYTHONNET_RUNTIME='sentinel-old', PYTHONUTF8='0')
    system32=Path(env['SYSTEMROOT'])/'System32'
    env['PATH']=str(system32) if fail_stage=='missing' else str(bin_dir)+os.pathsep+str(system32)
    wrapper=tmp_path/'caller.cmd'
    wrapper.write_text('''@echo off
call "%ENTRY%"
set "rc=%ERRORLEVEL%"
if not "%CD%"=="%CALLER_DIR%" exit /b 97
if not "%OFFICE_WEBVIEW_GUI%"=="sentinel-old" exit /b 96
if not "%PYTHONNET_RUNTIME%"=="sentinel-old" exit /b 95
if not "%PYTHONUTF8%"=="0" exit /b 94
exit /b %rc%
''',encoding='ascii')
    env.update(ENTRY=str(project/BAT.name),CALLER_DIR=str(tmp_path))
    run=subprocess.run([str(system32/'cmd.exe'), '/d', '/c', str(wrapper)],
                       cwd=tmp_path, env=env, capture_output=True, timeout=15)
    stdout=run.stdout.decode('utf-8', errors='replace')
    assert run.returncode==expected, (stdout, run.stderr)
    if fail_stage=='log':
        assert not trace.exists() and (project/'.startup-logs').is_file()
        return
    logs=list((project/'.startup-logs').glob('*.log')); assert len(logs)==1
    log=logs[0].read_text(encoding='utf-8', errors='replace')
    assert 'task/provider payloads' in log
    args=trace.read_text(encoding='utf-8', errors='replace') if trace.exists() else ''
    if fail_stage=='missing':
        assert not trace.exists() and 'uv_missing' in log
    else:
        assert ('sync\n' in args)==(not has_venv)
        assert 'gui=edgechromium;runtime=netfx' in args
        if fail_stage=='sync': assert 'office_agent.webui.runtime' not in args
        elif fail_stage=='readiness': assert 'office_agent.app' not in args
        else: assert 'run --offline --isolated --with "pythonnet==3.0.5" python -m office_agent.app --ui web' in args
    assert 'sentinel-old' not in args
    # Child env is private; no real uv/python/app processes exist in this branch test.
    assert env['OFFICE_WEBVIEW_GUI']=='sentinel-old'

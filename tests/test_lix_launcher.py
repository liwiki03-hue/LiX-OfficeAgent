import ctypes
from pathlib import Path
import shutil
import struct
import sys
import time

import pytest

from tools.build_lix_launcher import build


pytestmark = pytest.mark.skipif(sys.platform != 'win32', reason='Windows executable launcher')


@pytest.fixture(scope='module')
def executable(tmp_path_factory):
    return build(tmp_path_factory.mktemp('lix-exe') / 'LiX.exe')


def test_launcher_is_gui_executable_with_actual_brand_icon(executable):
    data = executable.read_bytes()
    assert data[:2] == b'MZ'
    pe = struct.unpack_from('<I', data, 60)[0]
    assert data[pe:pe + 4] == b'PE\0\0'
    assert struct.unpack_from('<H', data, pe + 24 + 68)[0] == 2
    shell = ctypes.WinDLL('shell32')
    shell.ExtractIconExW.argtypes = [ctypes.c_wchar_p, ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint]
    assert shell.ExtractIconExW(str(executable), -1, None, None, 0) == 1


@pytest.mark.parametrize('name', ['LiX 正常目录', 'LiX & 空格 (测试) %PATH%'])
def test_launcher_uses_own_directory_and_does_not_forward_arguments(executable, tmp_path, name):
    from office_agent.diagram_process import _WindowsJob
    folder = tmp_path / name
    folder.mkdir()
    shutil.copyfile(executable, folder / 'LiX.exe')
    (folder / 'Start-LiX.cmd').write_bytes(b'@echo off\r\ncall Start-OfficeAgent.cmd\r\nexit /b %ERRORLEVEL%\r\n')
    (folder / 'Start-OfficeAgent.cmd').write_bytes(b'@echo off\r\n"%ComSpec%" /d /u /c cd > launched.txt\r\necho %* > args.txt\r\nexit /b 0\r\n')
    job = _WindowsJob()
    process = None
    try:
        process = job.spawn((str(folder / 'LiX.exe'), '&', 'echo', 'unsafe', '>', 'unexpected.txt'))
        deadline = time.monotonic() + 15
        while process.poll() is None and time.monotonic() < deadline:
            time.sleep(.05)
        assert process.poll() == 0
        assert (folder / 'launched.txt').read_text('utf-16-le').strip() == str(folder)
        assert not (folder / 'unexpected.txt').exists()
        assert 'unsafe' not in (folder / 'args.txt').read_text('oem')
    finally:
        empty = job.close()
        if process:
            process.close()
        assert empty

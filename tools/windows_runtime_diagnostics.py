import hashlib
import json
import os
from pathlib import Path
import sys


def inspect(destination):
    result = dict(schema=1, windows=sys.platform == 'win32', changes_performed=False,
                  network_requests=0, service_started=False, installer_activated=False)
    if sys.platform != 'win32':
        return result
    import win32api
    import win32con
    import win32job
    import win32security
    import win32service
    process = win32api.GetCurrentProcess()
    token = win32security.OpenProcessToken(process, win32con.TOKEN_QUERY)
    try:
        sid = win32security.ConvertSidToStringSid(win32security.GetTokenInformation(token, win32security.TokenUser)[0])
        result['identity'] = dict(sid_sha256=hashlib.sha256(sid.encode('utf-8')).hexdigest(),
            elevated=bool(win32security.GetTokenInformation(token, win32security.TokenElevation)),
            system_account=sid == 'S-1-5-18', restricted=bool(win32security.IsTokenRestricted(token)))
    finally:
        token.Close()
    result['job'] = dict(attached=bool(win32job.IsProcessInJob(process, None)))
    if result['job']['attached']:
        try:
            limits = win32job.QueryInformationJobObject(None, win32job.JobObjectExtendedLimitInformation)
            result['job'].update(limit_flags=limits['BasicLimitInformation']['LimitFlags'],
                active_process_limit=limits['BasicLimitInformation']['ActiveProcessLimit'],
                memory_limit=limits['JobMemoryLimit'])
        except OSError as error:
            result['job']['query_error'] = error.winerror
    manager = service = None
    try:
        manager = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_CONNECT)
        service = win32service.OpenService(manager, 'msiserver', win32service.SERVICE_QUERY_STATUS | win32service.SERVICE_QUERY_CONFIG)
        status = win32service.QueryServiceStatus(service)
        configuration = win32service.QueryServiceConfig(service)
        result['installer_service'] = dict(accessible=True, state=status[1], start_type=configuration[1])
    except OSError as error:
        result['installer_service'] = dict(accessible=False, query_error=getattr(error, 'winerror', None))
    finally:
        if service:
            win32service.CloseServiceHandle(service)
        if manager:
            win32service.CloseServiceHandle(manager)
    destination = Path(destination).absolute()
    parent = destination
    while not parent.exists() and parent != parent.parent:
        parent = parent.parent
    attributes = parent.stat()
    result['location'] = dict(local=not str(destination).startswith('\\\\'),
        existing_parent_reparse=bool(getattr(attributes, 'st_file_attributes', 0) & 0x400),
        contains_non_ascii=not str(destination).isascii(), contains_spaces=' ' in str(destination),
        destination_exists=destination.exists(), free_bytes=__import__('shutil').disk_usage(parent).free)
    return result


if __name__ == '__main__':
    import threading
    if len(sys.argv) != 2:
        raise SystemExit('A local destination is required')
    watchdog = threading.Timer(20, lambda: os._exit(124))
    watchdog.daemon = True
    watchdog.start()
    try:
        print(json.dumps(inspect(sys.argv[1])))
    finally:
        watchdog.cancel()

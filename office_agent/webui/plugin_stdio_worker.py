import json
import os
import secrets
import subprocess
import sys


def main():
    import win32api
    import win32job
    job = win32job.CreateJobObject(None, 'Local\\LiX.Plugin.' + secrets.token_hex(16))
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info['BasicLimitInformation']['LimitFlags'] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    win32job.AssignProcessToJobObject(job, win32api.GetCurrentProcess())
    keys = {'SYSTEMROOT', 'WINDIR', 'TEMP', 'TMP', 'PATH', 'PATHEXT', 'SYSTEMDRIVE'}
    keys.update(json.loads(sys.argv[1]))
    env = {key: value for key, value in os.environ.items() if key.upper() in keys}
    process = subprocess.Popen(sys.argv[2:], stdin=sys.stdin.buffer, stdout=subprocess.PIPE,
                               stderr=subprocess.DEVNULL, env=env, creationflags=subprocess.CREATE_NO_WINDOW)
    with process.stdout:
        while line := process.stdout.readline(256 * 1024 + 1):
            if len(line) > 256 * 1024 or not line.endswith(b'\n'):
                return 1
            sys.stdout.buffer.write(line)
            sys.stdout.buffer.flush()
    return process.wait()


if __name__ == '__main__':
    os._exit(main())

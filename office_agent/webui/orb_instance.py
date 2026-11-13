import hashlib
import re


class SingleInstance:
    def __init__(self, namespace='OfficeAgentOrb'):
        if not re.fullmatch(r'[A-Za-z0-9-]{1,64}', namespace):
            raise ValueError('Invalid instance namespace')
        import pywintypes
        import win32api
        import win32con
        import win32event
        import win32security
        token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_QUERY)
        try:
            sid = win32security.ConvertSidToStringSid(win32security.GetTokenInformation(token, win32security.TokenUser)[0])
        finally:
            token.Close()
        attributes = pywintypes.SECURITY_ATTRIBUTES()
        attributes.SECURITY_DESCRIPTOR = win32security.ConvertStringSecurityDescriptorToSecurityDescriptor(
            'D:P(A;;GA;;;' + sid + ')', win32security.SDDL_REVISION_1)
        name = 'Local\\' + namespace + '-' + hashlib.sha256(sid.encode()).hexdigest()[:24]
        self.wake = self.mutex = None
        try:
            self.wake = win32event.CreateEvent(attributes, True, False, name + '-wake')
            self.mutex = win32event.CreateMutex(attributes, False, name)
            self.primary = win32api.GetLastError() != 183
            if not self.primary:
                win32event.SetEvent(self.wake)
        except BaseException:
            self.close()
            raise

    def poll_wake(self):
        import win32event
        if self.wake is not None and win32event.WaitForSingleObject(self.wake, 0) == win32event.WAIT_OBJECT_0:
            win32event.ResetEvent(self.wake)
            return True
        return False

    def close(self):
        for name in ('mutex', 'wake'):
            handle = getattr(self, name, None)
            if handle is not None:
                handle.Close()
                setattr(self, name, None)

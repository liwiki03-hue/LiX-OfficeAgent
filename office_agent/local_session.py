import getpass
import hashlib
import os
from pathlib import Path


class LocalSession:
    def __init__(self, root=None):
        if root is None:
            root = ((Path(os.environ.get('LOCALAPPDATA') or Path.home() / 'AppData/Local') / 'OfficeAgent')
                    if os.name == 'nt' else Path.home() / '.local/share/OfficeAgent')
        self.root = Path(os.path.abspath(os.fspath(root)))
        self._key = hashlib.sha256(('single-user:' + getpass.getuser() + ':' + os.path.normcase(str(self.root))).encode()).hexdigest()

    def status(self):
        return {'ok': True, 'authenticated': True, 'mode': 'single_user', 'user': '本机工作区'}

    def account_key(self):
        return self._key

    def logout(self):
        return self.status()

    def register(self, username, password):
        return {'ok': False, 'reason': 'single_user', 'error': '单机模式不需要注册或登录'}

    def reset_accounts(self, confirmed=False):
        return {'ok': False, 'reason': 'single_user', 'error': '单机模式不管理账户；原账户数据保留'}

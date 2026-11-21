import sqlite3


class PluginApi:
    def _plugin_call(self, function, *args):
        if not self._auth.status().get('authenticated'):
            return {'ok': False, 'error': '请先打开本机工作区。'}
        try:
            result = function(*args)
            if not self._auth.status().get('authenticated'):
                return {'ok': False, 'error': '工作区已关闭。'}
            return {'ok': True, **result}
        except ValueError as error:
            return {'ok': False, 'error': str(error)[:300]}
        except (OSError, sqlite3.Error):
            return {'ok': False, 'error': '插件本机配置暂时无法访问，请重新打开设置。'}
        except Exception:
            return {'ok': False, 'error': '插件操作未完成，请核对设置后重试。'}

    def plugin_list(self):
        return self._plugin_call(lambda: {'items': self._plugins.list()})

    def plugin_pick_program(self):
        def pick():
            from webview import FileDialog
            files = self._app.window.create_file_dialog(FileDialog.OPEN, file_types=('Windows 程序 (*.exe)',))
            return {'path': files[0] if files else ''}
        return self._plugin_call(pick)

    def plugin_credentials(self, id, revision):
        def state():
            from .plugin_transport import credential_available
            with self._plugins.database() as db:
                row = self._plugins.get(db, id, revision)
            return {'available': bool(credential_available(self._plugins.runtime(row)))}
        return self._plugin_call(state)

    def plugin_search(self, text, cursor=''):
        return self._plugin_call(self._plugins.search, text, cursor)

    def plugin_save(self, config, id=None, revision=None, separate=False):
        def save():
            row, duplicate = self._plugins.save(config, id, revision, separate)
            return {'item': row, 'duplicate': duplicate}
        return self._plugin_call(save)

    def plugin_probe(self, id, revision):
        return self._plugin_call(lambda: {'item': self._plugins.probe(id, revision)})

    def plugin_login(self, id, revision):
        return self._plugin_call(lambda: {'item': self._plugins.probe(id, revision, login=True)})

    def plugin_enable(self, id, revision, allowed):
        return self._plugin_call(lambda: {'item': self._plugins.change(id, revision, 'enable', allowed)})

    def plugin_disable(self, id, revision):
        return self._plugin_call(lambda: {'item': self._plugins.change(id, revision, 'disable')})

    def plugin_disconnect(self, id, revision):
        return self._plugin_call(lambda: {'item': self._plugins.change(id, revision, 'disconnect')})

    def plugin_remove(self, id, revision):
        return self._plugin_call(lambda: {'removed': self._plugins.change(id, revision, 'remove')['id']})

    def plugin_cancel(self, id):
        def cancel():
            self._plugins.cancel(id)
            return {}
        return self._plugin_call(cancel)


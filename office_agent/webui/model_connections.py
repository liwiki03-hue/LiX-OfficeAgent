import json
import threading
import secrets
import time
import hmac
import hashlib

from ..model_connections import ConnectionState, from_values, plan_catalog
from ..settings_store import SettingsError, SettingsStore
from ..chat_policy import reply_policy, output_parameter, continuation_chain, public_continuations, safe_usage
from ..visible_answer import LeadingThoughtFilter, visible_answer


CHAT_REPLY_TOKENS = 512


class ChatBudgetError(SettingsError):
    def __init__(self, budget):
        super().__init__('model_context_budget')
        self.budget = budget


def _http_client():
    import httpx2
    return httpx2.Client(follow_redirects=False, trust_env=False, timeout=60)


class ModelConnections:
    def __init__(self, settings, app, authenticated):
        self.settings, self.app, self.authenticated = settings, app, authenticated
        self.lock = settings._lock
        self.state = ConnectionState()
        self.chat_epoch = 0
        self.chat_messages = []
        self.managed = False
        self.local = threading.local()
        self.conversations = None
        self.chat_running = None
        self.chat_owner = None
        self.standalone_root = None
        self.image_tickets = {}
        self.model_lists = {}
        self._probe_secret = secrets.token_bytes(32)

    def invalidate(self):
        with self.lock:
            self.state.invalidate()
            self.chat_epoch += 1
            self.chat_messages.clear()
            self.chat_running = None
            self.image_tickets.clear()
            self.model_lists.clear()
            if self.conversations is not None:
                self.conversations.invalidate()

    def _catalog(self):
        if not self.authenticated():
            raise SettingsError('not_authenticated')
        if self.settings._disabled:
            raise SettingsError('settings_uncertain')
        if self.settings._store is None:
            self.settings._store = SettingsStore(self.settings._root)
        store = self.settings._store
        if store._fused:
            raise SettingsError('settings_uncertain')
        values = store._values(store._read(store.root / '.env', True))
        self.managed = 'OFFICE_MODEL_CONNECTIONS' in values
        return from_values(values)

    def snapshot(self):
        with self.lock:
            try:
                catalog = self._catalog()
                store = self.settings._store
                source = store._read(store.root / '.env', True)
                catalog = from_values(store._values(source))
                revision = secrets.token_hex(16)
                store._connection_snapshots[revision] = source
                store._connection_snapshots = dict(list(store._connection_snapshots.items())[-8:])
                from ..settings_definition import PROVIDER_DEFAULTS
                return {'ok': True, 'catalog': self.state.public(catalog), 'revision': revision,
                        'provider_defaults': {key: {'base_url': value['base_url']} for key, value in PROVIDER_DEFAULTS.items()}}
            except SettingsError as exc:
                return {'ok': False, 'reason': exc.code}
            except Exception:
                return {'ok': False, 'reason': 'settings_file_invalid'}

    def plan(self, payload, revision):
        return self.settings._call(lambda store, session: {'plan': plan_catalog(store, payload, session, revision)})

    def select(self, selection):
        with self.lock:
            try:
                catalog = self._catalog()
                if type(selection) is not dict or set(selection) != {'connection_id', 'model_id', 'capability', 'revision'}:
                    raise SettingsError('model_selection_invalid')
                if any(type(v) is not str for v in selection.values()) or selection['capability'] not in ('chat', 'agent', 'image'):
                    raise SettingsError('model_selection_invalid')
                cap = selection['capability']
                previous = self.state.selected.get(cap)
                self.state.selected[cap] = dict(selection)
                try:
                    route = self.state.resolve(catalog, cap)
                except Exception:
                    if previous is None:
                        self.state.selected.pop(cap, None)
                    else:
                        self.state.selected[cap] = previous
                    raise
                self.state.mode = cap
                return {'ok': True, 'selection': route.public(), 'next_request_only': bool(self.app.busy)}
            except SettingsError as exc:
                return {'ok': False, 'reason': exc.code}

    def mode(self, capability):
        with self.lock:
            if not self.authenticated():
                return {'ok': False, 'reason': 'not_authenticated'}
            if capability not in ('chat', 'agent', 'image'):
                return {'ok': False, 'reason': 'model_selection_invalid'}
            self.state.mode = capability
            return {'ok': True}

    def capture(self, capability):
        with self.lock:
            return self.state.resolve(self._catalog(), capability)

    def active(self):
        with self.lock:
            catalog = self._catalog()
            return self.managed or bool(catalog['connections']) or bool(self.state.selected)

    def record(self, route, status):
        with self.lock:
            if route.epoch != self.state.epoch or not self.authenticated():
                return
            try:
                row = next(r for r in self._catalog()['connections'] if r['id'] == route.connection_id)
                if self.state.identity(row) != route.revision:
                    return
                self.state.outcomes[(route.connection_id, route.model_id, route.capability, route.revision)] = status
                from ..diagnostics import record
                record('model', 'request', status, correlation=route.connection_id)
            except (SettingsError, StopIteration):
                return

    def build_model(self, route=None):
        route = route or getattr(self.local, 'route', None) or self.capture('agent')
        from langchain_core.callbacks import BaseCallbackHandler
        owner = self

        class Outcome(BaseCallbackHandler):
            def on_llm_end(self, response, **kwargs):
                owner.record(route, 'verified')

            def on_llm_error(self, error, **kwargs):
                if isinstance(error, GeneratorExit):
                    return
                owner.record(route, 'failed')

        callbacks = [Outcome()]
        from ..network import options
        proxy = options(self.settings._root).route(route.base_url)
        if route.protocol == 'ollama':
            from langchain_ollama import ChatOllama
            model = ChatOllama(model=route.model_id, base_url=route.base_url, temperature=0, callbacks=callbacks,
                               **({'num_ctx': route.context_window} if route.capability == 'chat' else {}),
                               client_kwargs={'timeout': 60, 'follow_redirects': False, 'trust_env': False, **({'proxy': proxy} if proxy else {})})
            for client in (model._client, model._async_client):
                client._client.headers.pop('authorization', None)
                if route.api_key:
                    client._client.headers['Authorization'] = 'Bearer ' + route.api_key
            return model
        from langchain_openai import ChatOpenAI
        from pydantic import SecretStr
        class ChatOutputModel(ChatOpenAI):
            def _get_request_payload(self, *args, **kwargs):
                payload = super()._get_request_payload(*args, **kwargs)
                if output_parameter(route) == 'max_tokens' and 'max_completion_tokens' in payload:
                    payload['max_tokens'] = payload.pop('max_completion_tokens')
                return payload
        cls = ChatOutputModel if route.capability == 'chat' else ChatOpenAI
        if proxy:
            import httpx2
            client = httpx2.Client(proxy=proxy, follow_redirects=False, trust_env=False, timeout=60)
        else:
            client = _http_client()
        return cls(model=route.model_id, base_url=route.base_url, api_key=SecretStr(route.api_key),
                          timeout=60, max_retries=0, callbacks=callbacks,
                          http_client=client, http_socket_options=())

    def research_model(self):
        route = self.capture('agent')
        from ..research_model import ResearchModel, ResearchModelConfig
        owner = self

        class RoutedResearchModel(ResearchModel):
            @property
            def identity(self):
                return {**super().identity, 'route': route.public()}

            def _request(self, *args, **kwargs):
                try:
                    result = super()._request(*args, **kwargs)
                    owner.record(route, 'verified')
                    return result
                except Exception:
                    owner.record(route, 'failed')
                    raise

        return RoutedResearchModel(ResearchModelConfig('ollama' if route.protocol == 'ollama' else 'deepseek', route.model_id,
                                  route.base_url, route.context_window, route.api_key))

    def image_config(self, route):
        from ..imagegen import ImageGenConfig, ImageGenError
        if route is None:
            raise ImageGenError('image_configuration_missing')
        store = self.settings._store
        if store is None or self.settings._disabled or store._fused or route.epoch != self.state.epoch or not self.authenticated():
            raise ImageGenError('image_configuration_changed')
        try:
            catalog = from_values(store._values(store._read(store.root / '.env', True)))
            row = next((r for r in catalog['connections'] if r['id'] == route.connection_id), None)
        except Exception:
            raise ImageGenError('image_configuration_changed') from None
        if route.epoch != self.state.epoch or not self.authenticated() or row is None or not row['enabled'] or self.state.identity(row) != route.revision:
            raise ImageGenError('image_configuration_changed')
        return ImageGenConfig('openai_compatible', route.api_key, route.base_url, route.model_id)

    def plan_image(self, question, options=None):
        from ..imagetransactions import ImageGenerationService
        from ..imagegen import ImageRequest, ImageGenError
        import uuid
        with self.lock:
            if not self.authenticated():
                return {'ok': False, 'reason': 'not_authenticated'}
            if self.app.busy:
                return {'ok': False, 'reason': 'busy'}
            if type(question) is not str:
                return {'ok':False,'reason':'image_request_invalid'}
            try:
                route = self.capture('image')
                from ..settings_store import _private_values
                store = self.settings._store
                if any(secret and secret in question for secret in _private_values(store._values(store._read(store.root / '.env',True)))):
                    return {'ok':False,'reason':'image_request_sensitive'}
                from ..clarification import image_request_count
                if image_request_count(question) != 1:
                    return {'ok': False, 'reason': 'image_request_invalid'}
                options = options or {}
                if type(options) is not dict or not set(options) <= {'size', 'style'}:
                    return {'ok': False, 'reason': 'image_request_invalid'}
                request = ImageRequest(prompt=question, **options)
                session = uuid.uuid4().hex
                if self.app.path and self.app.editor:
                    service = ImageGenerationService(self.app.editor.space.root, self.app.path, session,
                                                     config_loader=lambda: self.image_config(route))
                else:
                    from ..auth import LocalAuth
                    if self.standalone_root is None:
                        return {'ok': False, 'reason': 'image_context_unavailable'}
                    root = self.standalone_root
                    LocalAuth._check_path(root)
                    root.mkdir(parents=True, exist_ok=True)
                    LocalAuth._check_path(root)
                    service = ImageGenerationService(root, None, session, standalone=True,
                                                     config_loader=lambda: self.image_config(route),
                                                     reuse_metadata=self.settings._store._values(self.settings._store._read(self.settings._store.root / '.env', True)).get('OFFICE_IMAGE_REUSE_METADATA','false').lower() == 'true')
                plan = service.plan(request)
                plan['reuse_metadata_saved'] = getattr(service, '_reuse_metadata', False)
                plan['route'] = route.public()
                old = self.app._drop_image_session()
                if old is not None:
                    old.invalidate()
                self.app.pending_id = None
                self.app.model_agent = None
                self.app._pending_clarification = None
                self.app._image_pending_ids.clear()
                self.app._image_plan_in_run = False
                self.app._history_dir = None
                self.app._history_events = []
                service._model_route = route
                self.app._image_session = (session, service)
                self.app._push({'type': 'task', 'text': question})
                self.app._push({'type': 'log', 'tag': 'meta', 'text': '本次生图计划接口与模型：' + json.dumps(route.public(), ensure_ascii=False)})
                self.app.emit('image_plan', plan)
                return {'ok': True, 'route': route.public(), 'generation_not_started': True}
            except SettingsError as exc:
                return {'ok': False, 'reason': exc.code}
            except (ImageGenError, ValueError):
                return {'ok': False, 'reason': 'image_configuration_invalid'}

    def probe(self, connection_id, revision):
        from ..settings_probe import probe_connection
        prepared = None

        def prepare(store, session):
            nonlocal prepared
            catalog = self._catalog()
            row = next((r for r in catalog['connections'] if r['id'] == connection_id), None)
            if row is None or not row['enabled'] or self.state.identity(row) != revision:
                raise SettingsError('settings_source_changed')
            model = row['models'][0]['id'] if row['models'] else ''
            prepared = ('ollama' if row['protocol'] == 'ollama' else 'deepseek', row['base_url'], model, row['api_key'])
            return {'prepared': True}

        with self.lock:
            if self.settings._probing:
                return {'ok': False, 'reason': 'busy'}
            result = self.settings._call(prepare)
            if not result['ok']:
                return result
            self.settings._probing = True
            epoch = self.state.epoch
        try:
            outcome = probe_connection(prepared, include_models=True)
            with self.lock:
                if not self.authenticated() or epoch != self.state.epoch:
                    return {'ok': False, 'reason': 'not_authenticated'}
                row = next((r for r in self._catalog()['connections'] if r['id'] == connection_id), None)
                if row is None or self.state.identity(row) != revision:
                    return {'ok': False, 'reason': 'settings_source_changed'}
                self.state.lists[(connection_id, revision)] = outcome['status']
                return {'ok': True, 'result': outcome}
        finally:
            prepared = None
            with self.lock:
                self.settings._probing = False

    def probe_draft(self, draft, refresh=False):
        from ..settings_probe import probe_connection
        from ..settings_store import _validate
        from urllib.parse import urlsplit
        from ..network import endpoint_identity
        prepared, key_id, source_identity = None, None, None
        def prepare(store, session):
            nonlocal prepared, key_id, source_identity
            if (not isinstance(draft, dict) or set(draft) - {'id', 'protocol', 'base_url', 'key_action', 'api_key'}
                    or type(refresh) is not bool or draft.get('protocol') not in ('openai', 'ollama')):
                raise SettingsError()
            protocol, url = draft['protocol'], draft.get('base_url', '')
            _validate({'base_url': url})
            parsed = urlsplit(url)
            if not (parsed.scheme == 'https' or protocol == 'ollama' and parsed.hostname in {'localhost', '127.0.0.1', '::1'}):
                raise SettingsError()
            saved = next((row for row in self._catalog()['connections'] if row['id'] == draft.get('id')), None)
            action = draft.get('key_action', 'keep')
            if action not in ('keep', 'replace', 'clear') or action != 'replace' and draft.get('api_key'):
                raise SettingsError()
            key = draft.get('api_key', '') if action == 'replace' else ''
            if action == 'replace' and not key:
                raise SettingsError('settings_invalid', field='api_key')
            if action == 'keep' and saved:
                if endpoint_identity(saved['base_url']) != endpoint_identity(url) or saved['protocol'] != protocol:
                    raise SettingsError('settings_invalid', field='api_key')
                key = saved['api_key']
                source_identity = self.state.identity(saved)
            _validate({'api_key': key})
            if protocol != 'ollama' and not key:
                raise SettingsError('settings_invalid', field='api_key')
            prepared = ('ollama' if protocol == 'ollama' else 'deepseek', endpoint_identity(url), '', key)
            key_id = hmac.new(self._probe_secret, json.dumps([draft.get('id'), source_identity, prepared], ensure_ascii=True).encode(), hashlib.sha256).hexdigest()
            return {'prepared': True}
        with self.lock:
            if self.settings._probing:
                return {'ok': False, 'reason': 'busy'}
            result = self.settings._call(prepare)
            if not result['ok']:
                return result
            cached = self.model_lists.get(key_id)
            if not refresh and cached and time.monotonic() - cached[0] < 300:
                return {'ok': True, 'result': cached[1], 'cached': True}
            self.settings._probing = True
            epoch = self.state.epoch
        try:
            outcome = probe_connection(prepared, include_models=True)
            with self.lock:
                if not self.authenticated() or epoch != self.state.epoch:
                    return {'ok': False, 'reason': 'settings_source_changed'}
                if source_identity:
                    saved = next((row for row in self._catalog()['connections'] if row['id'] == draft.get('id')), None)
                    if not saved or self.state.identity(saved) != source_identity:
                        return {'ok': False, 'reason': 'settings_source_changed'}
                self.model_lists[key_id] = (time.monotonic(), outcome)
                self.model_lists = dict(list(self.model_lists.items())[-16:])
                return {'ok': True, 'result': outcome, 'cached': False}
        finally:
            prepared = None
            with self.lock:
                self.settings._probing = False

    def image_history(self):
        import os
        import re
        from ..auth import LocalAuth
        from .bridge import _read_bounded_json
        with self.lock:
            if not self.authenticated():
                return {'ok': False, 'reason': 'not_authenticated'}
            if self.standalone_root is None:
                return {'ok': False, 'reason': 'image_context_unavailable'}
            root = self.standalone_root / '.office_agent' / 'imagegen'
            try:
                self._catalog()
                store = self.settings._store
                reuse_enabled = store._values(store._read(store.root / '.env',True)).get('OFFICE_IMAGE_REUSE_METADATA','false').lower() == 'true'
                LocalAuth._check_path(root)
                if not root.exists():
                    return {'ok': True, 'items': [], 'limited': False, 'unreadable': 0,'reuse_metadata_enabled':reuse_enabled}
                items, unreadable, examined = [], 0, 0
                self.image_tickets.clear()
                with os.scandir(root) as entries:
                    for entry in entries:
                        examined += 1
                        if examined > 64:
                            break
                        if not re.fullmatch('[a-f0-9-]{36}', entry.name):
                            unreadable += 1
                            continue
                        path = root / entry.name / 'audit.json'
                        try:
                            LocalAuth._check_path(path)
                            body = _read_bounded_json(path, 16384)
                            if body.get('kind') != 'image_generation' or body.get('transaction_id') != entry.name or body.get('source') is not None:
                                raise ValueError()
                            output = body.get('saved_output') or {}
                            config = body.get('configuration') or {}
                            from ..rag_evidence import _read
                            raw, receipt = _read(path)
                            if json.loads(raw) != body:
                                raise ValueError()
                            self.image_tickets[entry.name] = (self.state.epoch, path, receipt)
                            items.append({'id': entry.name, 'status': str(body.get('status', 'unknown'))[:40],
                                          'created': path.stat().st_mtime, 'model': str(config.get('model', ''))[:128],
                                          'connection': str(config.get('base_url', ''))[:2048],
                                          'path': str(output.get('path', ''))[:2048],
                                          'sha256': str(output.get('sha256', ''))[:64],
                                          'saved': body.get('saved') is True,
                                          'reusable': type(body.get('reusable')) is dict,
                                          'request_state': str(body.get('provider_request_state', 'unknown'))[:40]})
                        except Exception:
                            unreadable += 1
                return {'ok': True, 'items': sorted(items, key=lambda item: item['created'], reverse=True),
                        'reuse_metadata_enabled':reuse_enabled,
                        'limited': examined > 64, 'unreadable': unreadable,
                        'scope': '本机独立生图审计摘要；不读图片、不恢复生成/保存/撤销批准，记录不代表远端实时状态。'}
            except Exception:
                return {'ok': False, 'reason': 'image_history_unavailable'}

    def _history_enabled(self):
        self._catalog()
        store = self.settings._store
        values = store._values(store._read(store.root / '.env', True))
        raw = values.get('OFFICE_CHAT_HISTORY', 'false')
        if type(raw) is not str:
            raise SettingsError('settings_file_invalid')
        raw = raw.lower()
        if raw not in {'true', 'false'}:
            raise SettingsError('settings_file_invalid')
        return raw == 'true'

    def conversation(self, operation='snapshot', identifier=None, confirmed=False, title=None):
        from .conversations import ConversationError
        with self.lock:
            try:
                enabled = self._history_enabled()
                book = self.conversations
                if book is None:
                    raise ConversationError('chat_account_required')
                book.ensure(enabled)
                if operation in {'new', 'select', 'delete', 'clear', 'reload', 'rename'}:
                    if self.app.busy and self.chat_running is None:
                        return {'ok': False, 'reason': 'busy'}
                    if operation == 'reload':
                        if self.app.busy:
                            return {'ok': False, 'reason': 'busy'}
                        selected = book.current
                        book.invalidate()
                        book.ensure(enabled)
                        if book.row(selected) is not None:
                            book.current = selected
                    elif operation == 'new':
                        previous = book.current
                        book.new()
                        outcome = book.persist(enabled)
                        if outcome == 'not_saved':
                            book.rows.pop()
                            book.current = previous
                        if outcome in {'not_saved', 'uncertain'}:
                            raise ConversationError(book.warning)
                    elif operation == 'select':
                        if book.row(identifier) is None:
                            raise ConversationError('chat_session_missing')
                        book.current = identifier
                    elif operation == 'rename':
                        if self.app.busy:
                            return {'ok': False, 'reason': 'busy'}
                        from ..settings_store import _private_values
                        self._catalog()
                        store = self.settings._store
                        if type(title) is str and any(value and value in title for value in _private_values(store._values(store._read(store.root / '.env', True)))):
                            raise ConversationError('chat_sensitive_input')
                        book.rename(identifier, title, enabled)
                    elif operation == 'delete':
                        if confirmed is not True:
                            raise ConversationError('confirmation_required')
                        if self.chat_running and self.chat_running[1] == identifier:
                            return {'ok': False, 'reason': 'busy'}
                        book.remove(identifier, enabled)
                    else:
                        if self.app.busy:
                            return {'ok': False, 'reason': 'busy'}
                        book.clear_saved(confirmed)
                        book.ensure(enabled)
                    if operation != 'rename' and not self.app.busy:
                        self.app.pending_id = None
                        self.app._read_authorized_path = None
                        old = self.app._drop_image_session()
                        if old is not None:
                            old.invalidate()
                    if operation != 'rename':
                        self.chat_messages = []
                    self.app._push({'type': 'chat_changed'})
                elif operation == 'save':
                    if not enabled:
                        raise ConversationError('chat_history_disabled')
                    if book.persist(True) != 'saved':
                        raise ConversationError(book.warning)
                elif operation != 'snapshot':
                    raise ConversationError('chat_request_invalid')
                result = book.public(enabled)
                store = self.settings._store
                result['reply_policy'] = reply_policy(store._values(store._read(store.root / '.env', True)))
                result['continuations'] = public_continuations(book.row())
                return {'ok': True, 'result': result}
            except (ConversationError, SettingsError) as error:
                reply = {'ok': False, 'reason': error.code}
                if error.code in {'chat_history_write_failed', 'chat_history_conflict', 'chat_history_uncertain'}:
                    try:
                        reply['result'] = book.public(enabled)
                    except Exception:
                        pass
                    reply['outcome_uncertain'] = error.code == 'chat_history_uncertain'
                return reply
            except Exception:
                return {'ok': False, 'reason': 'chat_history_unavailable'}

    def _chat_context(self, question, continuation=None):
        import hashlib
        from ..token_chunks import token_count
        if type(question) is not str or not question.strip():
            raise SettingsError('empty')
        if len(question) > 8000:
            raise SettingsError('empty')
        route = self.capture('chat')
        enabled = self._history_enabled()
        from ..settings_store import _private_values
        store = self.settings._store
        policy = reply_policy(store._values(store._read(store.root / '.env', True)), route)
        reserve = policy['reply_token_reserve']
        credentials = tuple(sorted(set([route.api_key, *_private_values(store._values(store._read(store.root / '.env', True)))]), key=len, reverse=True))
        credentials = tuple(value for value in credentials if value)
        if any(value in question for value in credentials):
            raise SettingsError('chat_sensitive_input')
        def redact(text):
            for value in credentials:
                text = text.replace(value, '[已隐藏密钥]')
            return text
        book = self.conversations
        row = None
        if book is not None:
            book.ensure(enabled)
            row = book.row()
            if len(row['turns']) >= 100 or len(json.dumps(book.rows, ensure_ascii=False).encode('utf8')) + len(question.encode('utf8')) + 132000 > 4 * 1024 * 1024:
                raise SettingsError('chat_history_capacity')
        completed = [turn for turn in row['turns'] if turn['status'] == 'complete'] if row else []
        mandatory = []
        if continuation is not None:
            _, mandatory = continuation_chain(row, continuation)
            completed = [turn for turn in completed if turn not in mandatory] + mandatory
        history = [message for turn in completed for message in ({'role': 'user', 'content': turn['question']}, {'role': 'assistant', 'content': turn['answer']})] if row else list(self.chat_messages)
        messages = [{**message, 'content': redact(message['content'])} for message in history] + [{'role': 'user', 'content': question}]
        omitted = 0
        while len(messages) > 1 + len(mandatory) * 2 and token_count(json.dumps(messages, ensure_ascii=False)) + reserve > route.context_window:
            messages = messages[2:]
            omitted += 1
        tokens = token_count(json.dumps(messages, ensure_ascii=False))
        budget = {'route':route.public(), 'context_window':route.context_window, 'estimated_input_tokens':tokens,
                  **policy,
                  'continuation_required': bool(mandatory),
                  'estimated_excess_tokens':max(0,tokens+reserve-route.context_window),
                  'ollama_num_ctx':route.context_window if route.protocol == 'ollama' else None,
                  'runtime_verified':False}
        if budget['estimated_excess_tokens']:
            raise ChatBudgetError(budget)
        revision = hashlib.sha256(json.dumps({'route': route.public(), 'session': row['id'] if row else None,
            'messages': messages, 'policy':policy, 'continuation':continuation,
            'turns': [(turn['id'], turn['status'], turn['finish_reason']) for turn in row['turns']] if row else []},
            ensure_ascii=False, sort_keys=True).encode('utf8')).hexdigest()
        transfer = bool(completed and (completed[-1]['route']['connection_id'] != route.connection_id or completed[-1]['route']['revision'] != route.revision))
        return {'route':route, 'enabled':enabled, 'credentials':credentials, 'redact':redact, 'book':book, 'row':row,
                'completed':completed, 'messages':messages, 'omitted':omitted, 'tokens':tokens,
                'revision':revision, 'transfer':transfer, 'budget':budget, 'continuation':continuation}

    def preview_chat(self, question, options=None):
        with self.lock:
            if not self.authenticated():
                return {'ok':False, 'reason':'not_authenticated'}
            if isinstance(question, str) and len(question) > 8000:
                return {'ok': False, 'reason': 'empty', 'error_code': 'too_long', 'characters': len(question), 'limit': 8000}
            try:
                if options is not None and (type(options) is not dict or not set(options) <= {'continuation'}):
                    raise SettingsError('chat_request_invalid')
                value = self._chat_context(question, (options or {}).get('continuation'))
                def summary(turn):
                    return {'id':turn['id'], 'question':value['redact'](turn['question'])[:120],
                            'answer_characters':len(value['redact'](turn['answer'])), 'status':turn['status'],
                            'finish_reason':turn['finish_reason'], 'route':turn['route']}
                completed, omitted, row = value['completed'], value['omitted'], value['row']
                return {'ok':True, 'result':{'route':value['route'].public(), 'session':row['id'] if row else None,
                    'context_revision':value['revision'], **value['budget'],
                    'context_turns':(len(value['messages'])-1)//2, 'omitted_turns':omitted,
                    'included':[summary(turn) for turn in completed[omitted:]],
                    'omitted':[summary(turn) for turn in completed[:omitted]],
                    'excluded':[summary(turn) for turn in row['turns'] if turn['status'] != 'complete'] if row else [],
                    'transfer_required':value['transfer'], 'read_only':True,
                    'scope':'内置tokenizer估算，含角色封装；服务商计量可能不同。仅预览，不调用模型、不保存消息，也不授予文档权限。'}}
            except ChatBudgetError as error:
                return {'ok':False, 'reason':error.code, 'budget':error.budget, 'request_started':False}
            except SettingsError as error:
                return {'ok':False, 'reason':error.code}
            except Exception:
                return {'ok':False, 'reason':'chat_history_unavailable'}

    def chat(self, question, options=None):
        import uuid
        with self.lock:
            if not self.authenticated():
                return {'ok': False, 'reason': 'not_authenticated'}
            if self.app.busy:
                return {'ok': False, 'reason': 'busy'}
            if type(question) is not str or not question.strip():
                return {'ok': False, 'reason': 'empty'}
            if len(question) > 8000:
                return {'ok': False, 'reason': 'too_long', 'characters': len(question), 'limit': 8000}
            options = options or {}
            if type(options) is not dict or not set(options) <= {'stream', 'transfer_to', 'context_revision', 'continuation'} or type(options.get('stream', True)) is not bool:
                return {'ok': False, 'reason': 'chat_request_invalid'}
            try:
                value = self._chat_context(question, options.get('continuation'))
                route, enabled, credentials = value['route'], value['enabled'], value['credentials']
                redact, book, row = value['redact'], value['book'], value['row']
                messages, omitted = value['messages'], value['omitted']
                if 'context_revision' in options and options['context_revision'] != value['revision']:
                    return {'ok':False, 'reason':'chat_context_changed'}
                if value['transfer'] and options.get('transfer_to') != route.public():
                    return {'ok': False, 'reason': 'chat_transfer_required', 'route': route.public(), 'history_turns': len(value['completed'])}
            except ChatBudgetError as exc:
                return {'ok':False, 'reason':exc.code, 'budget':exc.budget, 'request_started':False}
            except SettingsError as exc:
                return {'ok': False, 'reason': exc.code}
            except Exception:
                return {'ok': False, 'reason': 'chat_history_unavailable'}
            self.chat_epoch += 1
            epoch = self.chat_epoch
            token = uuid.uuid4().hex
            self.chat_running = (token, row['id'] if row else None)
            self.chat_owner = token
            turn = {'id': token, 'question': redact(question),
                    'answer': '', 'route': route.public(), 'status': 'interrupted', 'created': __import__('time').time(),
                    'context_turns': (len(messages)-1)//2, 'omitted_turns': omitted, 'finish_reason': 'unknown'}
            if value['budget']['configured'] or value['continuation']:
                turn['budget'] = {key:value['budget'][key] for key in ('reply_token_reserve','context_window','estimated_input_tokens','output_parameter')}
            if value['continuation']:
                turn['continuation_of'] = dict(value['continuation'])
            if row:
                row['turns'].append(turn)
                if len(row['turns']) == 1 and row['title'] == '新会话':
                    row['title'] = row['turns'][0]['question'][:60]
                book.persist(enabled)
            if hasattr(self.app, '_history_dir'):
                self.app._history_dir = None
                self.app._history_events = []
                self.app.pending_id = None
                self.app.model_agent = None
                self.app._pending_clarification = None
                self.app._image_pending_ids.clear()
                self.app._image_plan_in_run = False
                old = self.app._drop_image_session()
                if old is not None:
                    old.invalidate()
                self.app._stop_task.clear()
                self.app._model_task_active = True
            self.app.busy = True
            self.app._push({'type': 'task', 'text': question, 'channel': 'chat', 'session_id': row['id'] if row else None, 'turn_id': token})
            self.app._push({'type': 'busy', 'busy': True})
            self.app._push({'type': 'status', 'level': 'info', 'text': f'已受理普通问答；本轮发送{(len(messages)-1)//2}组历史，较早{omitted}组不发送。停止仅停止本地显示，远端可能继续处理或计费。', 'channel': 'chat', 'session_id': row['id'] if row else None, 'turn_id': token})
            self.app._push({'type': 'log', 'tag': 'meta', 'text': '本轮实际接口与模型：' + json.dumps(route.public(), ensure_ascii=False)})

        def visible():
            return epoch == self.chat_epoch and route.epoch == self.state.epoch and self.authenticated() and (row is None or book.current == row['id'])

        def stopped():
            value = getattr(self.app, '_stop_task', None)
            return bool(value and value.is_set())

        def run():
            model, stream = None, None
            content = ''
            try:
                model = self.build_model(route)
                if hasattr(model, 'max_tokens'):
                    model.max_tokens = value['budget']['reply_token_reserve']
                if hasattr(model, 'num_predict'):
                    model.num_predict = value['budget']['reply_token_reserve']
                if options.get('stream', True) and callable(getattr(model, 'stream', None)):
                    thought_filter = LeadingThoughtFilter()
                    stream = model.stream(messages)
                    with self.lock:
                        if visible():
                            self.app._push({'type': 'answer_draft_start', 'stream_id': token,'session_id':row['id'] if row else None})
                    emitted = 0
                    started = __import__('time').monotonic()
                    for count, piece in enumerate(stream):
                        if count >= 2048 or __import__('time').monotonic() - started > 90:
                            turn['finish_reason'] = 'local_limit'
                            break
                        if route.epoch != self.state.epoch or stopped():
                            break
                        metadata = getattr(piece, 'response_metadata', {}) or {}
                        reason = metadata.get('finish_reason') or metadata.get('done_reason')
                        if reason in ('stop', 'length', 'content_filter'):
                            turn['finish_reason'] = reason
                        fragment = piece.content
                        if type(fragment) is not str:
                            raise ValueError('chat_response_incompatible')
                        content += thought_filter.feed(fragment)
                        if len(content) > 32000:
                            content = redact(content)[:32000]
                            turn['finish_reason'] = 'local_limit'
                            break
                        if 'budget' in turn and safe_usage(piece):
                            turn['usage'] = safe_usage(piece)
                        safe = redact(content)
                        available = max(0, len(safe) - max((len(value) for value in credentials), default=0))
                        with self.lock:
                            if epoch == self.chat_epoch and route.epoch == self.state.epoch and self.authenticated():
                                turn['answer'] = safe
                            if visible() and available > emitted:
                                self.app._push({'type': 'answer_draft_delta', 'stream_id': token,'session_id':row['id'] if row else None, 'text': safe[emitted:available]})
                                emitted = available
                    content += thought_filter.finish()
                else:
                    reply = model.invoke(messages)
                    content = visible_answer(reply.content)
                    if 'budget' in turn and safe_usage(reply):
                        turn['usage'] = safe_usage(reply)
                    metadata = getattr(reply, 'response_metadata', {}) or {}
                    reason = metadata.get('finish_reason') or metadata.get('done_reason')
                    if reason in ('stop', 'length', 'content_filter'):
                        turn['finish_reason'] = reason
                if type(content) is str and len(content) > 32000:
                    content = redact(content)[:32000]
                    turn['finish_reason'] = 'local_limit'
                if type(content) is not str or not content.strip():
                    raise ValueError('chat_response_incompatible')
                content = redact(content)
                if len(content) > 32000:
                    content = content[:32000]
                    turn['finish_reason'] = 'local_limit'
                with self.lock:
                    if epoch == self.chat_epoch and route.epoch == self.state.epoch and self.authenticated():
                        turn['status'] = 'stopped' if stopped() else 'complete'
                        turn['answer'] = content
                        if not stopped():
                            self.record(route, 'verified')
                            turn['answer'] = content
                            self.chat_messages = [*messages, {'role': 'assistant', 'content': content}][-8:]
                            if visible():
                                self.app._push({'type': 'answer_draft_cancel', 'stream_id': token})
                                self.app._push({'type': 'log', 'tag': 'answer', 'text': content})
                        if book:
                            book.persist(enabled)
            except Exception as error:
                with self.lock:
                    if epoch == self.chat_epoch and route.epoch == self.state.epoch and self.authenticated():
                        turn['status'] = 'stopped' if stopped() else 'failed'
                        turn['answer'] = redact(content)[:32000] if type(content) is str else ''
                        if not stopped():
                            self.record(route, 'failed')
                        if book:
                            book.persist(enabled)
                        if visible():
                            status = getattr(error, 'status_code', None)
                            name = type(error).__name__.lower()
                            code = 'authentication_failed' if status in (401, 403) else 'model_not_found' if status == 404 else 'rate_limited' if status == 429 else 'redirect_refused' if status in (301, 302, 303, 307, 308) else 'service_failed' if status and status >= 500 else 'timeout' if 'timeout' in name else 'address_unreachable' if 'connect' in name else 'response_incompatible'
                            messages_by_code = {'authentication_failed':'鉴权失败，请修复当前接口密钥。','model_not_found':'模型或接口路径不存在，请核对模型与地址。','rate_limited':'接口限流，请稍后主动重试。','redirect_refused':'接口返回重定向，为保护凭据未跟随；请填写最终服务地址。','service_failed':'模型服务异常，请核对服务状态。','timeout':'等待超时，服务端可能已受理，请先核对。','address_unreachable':'接口地址不可达，请检查地址和网络。','response_incompatible':'响应格式不兼容，可关闭流式后主动重试。'}
                            self.app._push({'type': 'badge', 'level': 'warn', 'label': '问答失败'})
                            self.app._push({'type': 'status', 'level': 'warn', 'text': messages_by_code[code] + '未切换接口或自动重试，输入可恢复。', 'channel':'chat', 'session_id':row['id'] if row else None, 'turn_id':token})
                            self.app._push({'type': 'chat_error', 'code': code, 'prompt': turn['question'],'session_id':row['id'] if row else None, 'turn_id':token, 'request_state': 'may_have_been_received'})
                            self.app._push({'type': 'log', 'tag': 'warn', 'text': '本轮模型请求失败，未切换接口或自动重试；请检查连接与账单。'})
            finally:
                try:
                    if stream is not None and callable(getattr(stream, 'close', None)):
                        stream.close()
                    client = getattr(model, 'http_client', None)
                    if client is not None:
                        client.close()
                    elif getattr(model, '_client', None) is not None:
                        model._client._client.close()
                except Exception:
                    pass
                with self.lock:
                    if self.chat_owner == token:
                        self.chat_running = None
                        self.chat_owner = None
                        self.app._model_task_active = False
                        if epoch == self.chat_epoch and route.epoch == self.state.epoch:
                            if visible():
                                self.app._push({'type': 'answer_draft_cancel', 'stream_id': token})
                                if turn['status'] == 'complete':
                                    uncertain = turn['finish_reason'] == 'unknown'
                                    partial = turn['finish_reason'] in ('length', 'content_filter', 'local_limit')
                                    self.app._push({'type':'badge','level':'warn' if partial or uncertain else 'ok','label':'回复完整性未确认' if uncertain else '回复可能未完整' if partial else '问答已返回'})
                                    self.app._push({'type':'status','level':'warn' if partial or uncertain else 'ok','text':('服务过滤了回复，不提供自动续写；请核对服务规则。' if turn['finish_reason'] == 'content_filter' else '达到本地字符、分块或等待限制，已保留收到的内容；可核对后主动续写。' if turn['finish_reason'] == 'local_limit' else '服务未报告结束原因，回复完整性未确认；请核对后决定是否继续。' if uncertain else '回复达到输出预算，可能不完整；可核对后主动续写。' if partial else '本轮模型回复已返回；可以复制回答、继续问答或新建会话。') + (' 本地历史未保存，请核对问答区提示。' if book and book.warning else ''), 'channel':'chat', 'session_id':row['id'] if row else None, 'turn_id':token})
                                elif turn['status'] == 'stopped':
                                    self.app._push({'type':'badge','level':'warn','label':'已停止本地显示'})
                                    self.app._push({'type':'status','level':'warn','text':'本地已停止，当前请求资源已释放；远端状态与费用需核对，可主动重新发送。', 'channel':'chat', 'session_id':row['id'] if row else None, 'turn_id':token})
                            self.app._push({'type': 'chat_changed'})
                            self.app.emit('idle', None)
                        else:
                            self.app.busy = False

        threading.Thread(target=run, daemon=True).start()
        return {'ok': True, 'route': route.public(), 'turn_id':token, 'context_turns': (len(messages)-1)//2, 'omitted_turns': omitted,
                'budget':value['budget']}

"""Host-only embeddings; no environment reads, automatic downloads or tool authority."""
from dataclasses import dataclass, field
import hashlib
import json
import math
import os
from pathlib import Path
import stat
import time
from urllib.parse import urlsplit

import httpx2

from .rag import SECRET


class EmbeddingError(ValueError):
    def __init__(self, code='embedding_invalid'):
        self.code = code
        super().__init__(code)


def fail(code='embedding_invalid'):
    raise EmbeddingError(code)


@dataclass(frozen=True)
class EmbeddingConfig:
    channel: str
    model: str
    dimension: int
    api_key: str = field(default='',repr=False)
    base_url: str = ''
    timeout: float = 20
    budget: float = 30

    @classmethod
    def from_values(cls, values):
        """Trusted host passes canonical .env values; never consult ambient process env."""
        try:
            return cls(values['OFFICE_EMBEDDING_CHANNEL'],values['OFFICE_EMBEDDING_MODEL'],
                       int(values['OFFICE_EMBEDDING_DIMENSION']),values.get('DASHSCOPE_API_KEY','') if values['OFFICE_EMBEDDING_CHANNEL'] == 'api' else '',
                       values.get('DASHSCOPE_BASE_URL','') if values['OFFICE_EMBEDDING_CHANNEL'] == 'api' else '',float(values.get('OFFICE_EMBEDDING_TIMEOUT','20')),
                       float(values.get('OFFICE_EMBEDDING_BUDGET','30')))
        except Exception:
            fail('embedding_not_configured')

    def __post_init__(self):
        if (self.channel not in ('api','local') or type(self.model) is not str or not self.model.strip() or
                len(self.model) > 200 or '\x00' in self.model or SECRET.search(self.model) or
                type(self.dimension) is not int or not 1 <= self.dimension <= 4096 or
                type(self.timeout) not in (int,float) or not 1 <= self.timeout <= 20 or
                type(self.budget) not in (int,float) or not 1 <= self.budget <= 60):
            fail()
        if self.channel == 'api':
            try:
                parsed = urlsplit(self.base_url)
                safe = (parsed.scheme == 'https' and parsed.hostname in ('dashscope.aliyuncs.com','dashscope-intl.aliyuncs.com') and
                        parsed.port in (None,443) and not parsed.username and not parsed.password and not parsed.query and
                        not parsed.fragment and parsed.path.rstrip('/') == '/compatible-mode/v1')
            except Exception:
                safe = False
            if (not safe or self.model not in ('text-embedding-v3','text-embedding-v4') or
                    self.dimension not in (64,128,256,512,768,1024,1536,2048) or
                    (self.model == 'text-embedding-v3' and self.dimension > 1024) or
                    type(self.api_key) is not str or not self.api_key or len(self.api_key) > 512 or
                    any(char.isspace() or ord(char) < 32 for char in self.api_key)):
                fail('embedding_api_configuration')
        elif not Path(self.model).is_absolute() or self.model.startswith('\\\\'):
            fail('embedding_local_directory')


def validate_vectors(vectors, count, dimension):
    if type(vectors) is not list or len(vectors) != count:
        fail('embedding_response')
    for vector in vectors:
        if (type(vector) is not list or len(vector) != dimension or
                any(type(value) not in (int,float) or not math.isfinite(value) or abs(value) > 1e6 for value in vector) or
                not any(value != 0 for value in vector)):
            fail('embedding_response')
    return vectors


def _deadline(deadline):
    if time.monotonic() > deadline:
        fail('embedding_timeout')


def _unique(pairs):
    result = {}
    for key,value in pairs:
        if key in result:
            fail('embedding_response')
        result[key] = value
    return result


class Embedder:
    def __init__(self, config, *, transport_factory=None, local_factory=None):
        if type(config) is not EmbeddingConfig:
            fail()
        self.config = config
        self._transport = transport_factory or (lambda:httpx2.HTTPTransport(retries=0,trust_env=False))
        self._local_factory = local_factory
        self._model = None
        self._manifest_id = None
        self._disabled = False
        self.last_run = {}

    @property
    def identity(self):
        if self._disabled:
            fail('embedding_failed')
        model = self.config.model
        if self.config.channel == 'local':
            fingerprint = self._manifest(time.monotonic()+self.config.budget)
            if self._manifest_id is not None and fingerprint != self._manifest_id:
                fail('embedding_local_changed')
            self._manifest_id = fingerprint
            model = 'local:sha256:'+fingerprint
        return {'channel':self.config.channel,'model':model,'dimension':self.config.dimension}

    def embed(self, texts):
        return self._embed(texts, False)

    def embed_queries(self, texts):
        return self._embed(texts, True)

    def _embed(self, texts, query_mode):
        if self._disabled:
            fail('embedding_failed')
        if (type(texts) is not list or not 1 <= len(texts) <= 500 or
                any(type(text) is not str or not text.strip() or len(text) > 8000 or '\x00' in text or SECRET.search(text) for text in texts) or
                sum(map(len,texts)) > 1_000_000):
            fail('embedding_input')
        deadline = time.monotonic()+self.config.budget
        try:
            if self.config.channel == 'api':
                return self._api(texts,deadline)
            fingerprint = self._manifest(deadline)
            if self._manifest_id is not None and fingerprint != self._manifest_id:
                fail('embedding_local_changed')
            self._manifest_id = fingerprint
            if (Path(self.config.model) / 'research_model.json').is_file():
                from .research_onnx import run_owned
                self.last_run = {}
                try:
                    rows = run_owned(self.config.model, 'embedding', texts, min(self.config.timeout, self.config.budget), query_mode=query_mode, telemetry=self.last_run)
                except Exception as error:
                    if str(error) == 'onnx_cleanup_failed' or getattr(error, 'code', '') == 'renderer_cleanup_failed':
                        self._disabled = True
                    raise
                return validate_vectors(rows, len(texts), self.config.dimension)
            if self._model is None:
                factory = self._local_factory
                if factory is None:
                    try:
                        from sentence_transformers import SentenceTransformer
                        factory = SentenceTransformer
                    except ImportError:
                        fail('embedding_local_not_installed')
                self._model = factory(self.config.model,local_files_only=True,trust_remote_code=False,
                                      model_kwargs={'use_safetensors':True,'local_files_only':True,'trust_remote_code':False},
                                      tokenizer_kwargs={'local_files_only':True,'trust_remote_code':False})
            _deadline(deadline)
            vectors = []
            for start in range(0,len(texts),10):
                _deadline(deadline)
                rows = self._model.encode(texts[start:start+10],batch_size=10,normalize_embeddings=True,
                                          convert_to_numpy=True,show_progress_bar=False)
                if hasattr(rows,'tolist'):
                    rows = rows.tolist()
                vectors.extend(validate_vectors(rows,len(texts[start:start+10]),self.config.dimension))
                _deadline(deadline)
            if self._manifest(deadline) != fingerprint:
                fail('embedding_local_changed')
            return vectors
        except EmbeddingError:
            self._model = None
            raise
        except Exception:
            self._model = None
            raise EmbeddingError('embedding_failed') from None

    def _api(self, texts, deadline):
        vectors = []
        with httpx2.Client(transport=self._transport(),trust_env=False,follow_redirects=False) as client:
            for start in range(0,len(texts),10):
                _deadline(deadline)
                batch = texts[start:start+10]
                payload = {'model':self.config.model,'input':batch,'dimensions':self.config.dimension,'encoding_format':'float'}
                headers = {'Authorization':'Bearer '+self.config.api_key,'Content-Type':'application/json',
                           'Accept':'application/json','Accept-Encoding':'identity'}
                timeout = min(self.config.timeout,max(0.001,deadline-time.monotonic()))
                with client.stream('POST',self.config.base_url.rstrip('/')+'/embeddings',json=payload,headers=headers,timeout=timeout) as response:
                    _deadline(deadline)
                    if response.status_code != 200 or response.headers.get('content-encoding','identity').lower() != 'identity':
                        fail('embedding_response')
                    length = response.headers.get('content-length')
                    if length is not None and (not length.isdecimal() or int(length) > 2_000_000):
                        fail('embedding_response')
                    body = bytearray()
                    for chunk in response.iter_raw():
                        _deadline(deadline)
                        if len(body)+len(chunk) > 2_000_000:
                            fail('embedding_response')
                        body.extend(chunk)
                    _deadline(deadline)
                data = json.loads(body,object_pairs_hook=_unique)
                if type(data) is not dict or data.get('model',self.config.model) != self.config.model:
                    fail('embedding_response')
                rows = data.get('data')
                if type(rows) is not list or len(rows) != len(batch):
                    fail('embedding_response')
                indexed = {}
                for row in rows:
                    if (type(row) is not dict or set(row)-{'object','index','embedding'} or type(row.get('index')) is not int or
                            not 0 <= row['index'] < len(batch) or row['index'] in indexed):
                        fail('embedding_response')
                    indexed[row['index']] = row.get('embedding')
                vectors.extend(validate_vectors([indexed[i] for i in range(len(batch))],len(batch),self.config.dimension))
        return vectors

    def _manifest(self, deadline):
        try:
            return self._manifest_impl(deadline)
        except EmbeddingError:
            raise
        except ImportError:
            raise EmbeddingError('embedding_local_not_installed') from None
        except Exception:
            raise EmbeddingError('embedding_local_directory') from None

    def _manifest_impl(self, deadline):
        """Bounded explicit offline bundle only; reject links, executable modules and pickle weights."""
        root = Path(self.config.model)
        if (root / 'research_model.json').is_file():
            from .research_onnx import bundle_identity
            fingerprint, metadata = bundle_identity(self.config.model, 'embedding', deadline)
            if metadata['dimension'] != self.config.dimension:
                fail('embedding_response')
            return fingerprint
        def stamp(path, directory=False):
            value = path.lstat()
            if (stat.S_ISLNK(value.st_mode) or getattr(value,'st_file_attributes',0)&0x400 or
                    (directory and not stat.S_ISDIR(value.st_mode)) or
                    (not directory and (not stat.S_ISREG(value.st_mode) or value.st_nlink != 1))):
                fail('embedding_local_directory')
            return (value.st_dev,value.st_ino,value.st_size,value.st_mtime_ns)
        for ancestor in (root,*root.parents):
            stamp(ancestor,True)
        pending, count, total, config_total, weights = [root], 0, 0, 0, False
        digest = hashlib.sha256()
        while pending:
            _deadline(deadline)
            directory = pending.pop()
            before = stamp(directory,True)
            for path in sorted(directory.iterdir()):
                _deadline(deadline)
                count += 1
                if count > 512:
                    fail('embedding_local_capacity')
                entry = path.lstat()
                if stat.S_ISDIR(entry.st_mode):
                    stamp(path,True)
                    if len(path.relative_to(root).parts) > 4:
                        fail('embedding_local_capacity')
                    pending.append(path)
                    continue
                original = stamp(path)
                total += entry.st_size
                if total > 2_000_000_000 or path.suffix.lower() in {'.py','.bin','.pkl','.pickle','.pt','.pth'}:
                    fail('embedding_local_capacity')
                is_weight = path.suffix.lower() == '.safetensors'
                weights |= is_weight
                if not is_weight:
                    config_total += entry.st_size
                    if entry.st_size > 8_000_000 or config_total > 32_000_000:
                        fail('embedding_local_capacity')
                digest.update(str(path.relative_to(root)).encode('utf8'))
                digest.update(repr(original).encode('ascii'))
                content = bytearray() if path.suffix.lower() == '.json' else None
                with path.open('rb') as stream:
                    descriptor = os.fstat(stream.fileno())
                    if (descriptor.st_dev,descriptor.st_ino,descriptor.st_size,descriptor.st_mtime_ns) != original:
                        fail('embedding_local_changed')
                    while chunk := stream.read(65536):
                        _deadline(deadline)
                        digest.update(chunk)
                        if content is not None:
                            content.extend(chunk)
                if stamp(path) != original:
                    fail('embedding_local_changed')
                if content is not None:
                    configuration = json.loads(content,object_pairs_hook=_unique)
                    nodes = [configuration]
                    seen = 0
                    while nodes:
                        node = nodes.pop()
                        seen += 1
                        if seen > 4000:
                            fail('embedding_local_capacity')
                        if type(node) is dict:
                            for key,value in node.items():
                                if type(value) is str and ('file' in key.lower() or 'path' in key.lower()):
                                    if (Path(value).is_absolute() or ':' in value or
                                            '..' in value.replace('\\','/').split('/')):
                                        fail('embedding_local_modules')
                            nodes.extend(node.values())
                        elif type(node) is list:
                            nodes.extend(node)
                if path.name == 'modules.json':
                    modules = configuration
                    allowed = {'sentence_transformers.models.Transformer','sentence_transformers.models.Pooling','sentence_transformers.models.Normalize'}
                    if type(modules) is not list or len(modules) > 8:
                        fail('embedding_local_modules')
                    for module in modules:
                        if (type(module) is not dict or module.get('type') not in allowed or type(module.get('path')) is not str or
                                any(part in ('..','.') for part in module['path'].replace('\\','/').split('/')) or
                                Path(module['path']).is_absolute() or ':' in module['path']):
                            fail('embedding_local_modules')
            if stamp(directory,True) != before:
                fail('embedding_local_changed')
        if not weights:
            fail('embedding_local_weights')
        return digest.hexdigest()

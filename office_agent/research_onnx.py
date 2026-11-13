import hashlib
from importlib.metadata import version
import json
import os
from pathlib import Path
import stat
import sys
import tempfile
import time


def bundle_identity(directory, task, deadline, *, file_digests=None):
    def stamp(entry):
        return entry.st_dev, entry.st_ino, entry.st_size, entry.st_mtime_ns, entry.st_mode, entry.st_nlink

    root = Path(directory)
    if not root.is_absolute() or str(root).startswith('\\\\'):
        raise ValueError('onnx_directory')
    parents = {}
    for parent in (root, *root.parents):
        entry = parent.lstat()
        if not stat.S_ISDIR(entry.st_mode) or getattr(entry, 'st_file_attributes', 0) & 0x400:
            raise ValueError('onnx_directory')
        parents[parent] = (entry.st_dev, entry.st_ino)
    paths = sorted(root.iterdir())
    if {path.name for path in paths} != {'model.onnx', 'tokenizer.json', 'config.json', 'MODEL_CARD.md', 'research_model.json'}:
        raise ValueError('onnx_bundle_files')
    digest = hashlib.sha256()
    digest.update(repr(parents[root]).encode())
    metadata = None
    for path in paths:
        before = path.lstat()
        if (not stat.S_ISREG(before.st_mode) or before.st_nlink != 1 or getattr(before, 'st_file_attributes', 0) & 0x400 or
                not 0 < before.st_size <= (125_000_000 if path.suffix == '.onnx' else 20_000_000)):
            raise ValueError('onnx_bundle_capacity')
        digest.update(path.name.encode())
        digest.update(repr((before.st_dev, before.st_ino, before.st_size, before.st_mtime_ns)).encode())
        body = bytearray() if path.name == 'research_model.json' else None
        content_digest = hashlib.sha256()
        if body is not None and before.st_size > 16000:
            raise ValueError('onnx_metadata')
        with path.open('rb') as stream:
            opened = os.fstat(stream.fileno())
            if stamp(opened) != stamp(before):
                raise ValueError('onnx_changed')
            while chunk := stream.read(65536):
                if time.monotonic() > deadline:
                    raise ValueError('onnx_timeout')
                digest.update(chunk)
                content_digest.update(chunk)
                if body is not None:
                    body.extend(chunk)
        if stamp(path.lstat()) != stamp(before):
            raise ValueError('onnx_changed')
        if body is not None:
            metadata = json.loads(body)
        if file_digests is not None:
            file_digests[path.name] = content_digest.hexdigest()
    expected = {'format', 'task', 'repo', 'revision', 'license', 'dimension', 'max_tokens', 'normalization', 'pooling', 'query_prefix', 'document_prefix', 'artifacts'}
    if (type(metadata) is not dict or set(metadata) != expected or metadata['format'] != 'office-onnx-v1' or metadata['task'] != task or
            metadata['max_tokens'] != 512 or metadata['dimension'] != (384 if task == 'embedding' else 1) or
            metadata['pooling'] != ('attention_mean' if task == 'embedding' else 'logit') or
            metadata['normalization'] != ('l2' if task == 'embedding' else 'none') or
            metadata['query_prefix'] != ('query: ' if task == 'embedding' else '') or
            metadata['document_prefix'] != ('passage: ' if task == 'embedding' else '')):
        raise ValueError('onnx_metadata')
    for dependency in ('onnxruntime', 'onnx', 'tokenizers', 'numpy'):
        digest.update((dependency + '=' + version(dependency)).encode())
    for parent, identity in parents.items():
        current = parent.lstat()
        if (current.st_dev, current.st_ino) != identity or getattr(current, 'st_file_attributes', 0) & 0x400:
            raise ValueError('onnx_changed')
    digest.update(b'office-onnx-inference-v2:immutable-bytes:cpu2:reject-overflow:mean-l2:pair-logit')
    return digest.hexdigest(), metadata


def run_owned(directory, task, texts, timeout, query=None, query_mode=False, telemetry=None):
    from .diagram_process import _run_owned
    from .fidelity import _temporary_root

    fingerprint, metadata = bundle_identity(directory, task, time.monotonic() + timeout)
    uncertain = False
    try:
        with tempfile.TemporaryDirectory(prefix='office-onnx-', dir=_temporary_root()) as raw:
            root = Path(raw)
            request, response = root / 'input.json', root / 'output.json'
            request.write_text(json.dumps({'model': directory, 'task': task, 'texts': texts, 'query': query, 'query_mode': query_mode}), encoding='utf8')
            try:
                result = _run_owned((sys.executable, str(Path(__file__).with_name('research_onnx_worker.py')), str(request), str(response)), timeout_ms=int(timeout * 1000))
            except Exception as error:
                uncertain = getattr(error, 'code', None) == 'renderer_cleanup_failed'
                raise
            if result.empty_confirmed is not True:
                uncertain = True
                raise ValueError('onnx_cleanup_failed')
            if result.code == 'render_timeout':
                raise ValueError('onnx_timeout')
            if result.exit_code != 0 or not response.is_file() or response.stat().st_size > 12_000_000:
                raise ValueError('onnx_unavailable')
            data = json.loads(response.read_bytes())
            if data.get('fingerprint') != fingerprint or bundle_identity(directory, task, time.monotonic() + timeout)[0] != fingerprint:
                raise ValueError('onnx_changed')
            if telemetry is not None:
                telemetry.update(elapsed_seconds=result.elapsed_seconds, empty_confirmed=result.empty_confirmed, process_code=result.code,
                                 peak_working_set_bytes=data['peak_working_set_bytes'], peak_commit_bytes=data['peak_commit_bytes'])
            return data['values']
    except Exception:
        if uncertain:
            raise ValueError('onnx_cleanup_failed') from None
        raise


def infer(directory, task, texts, query=None, query_mode=False):
    import numpy as np
    import onnx
    import onnxruntime as ort
    from tokenizers import Tokenizer

    files = {}
    fingerprint, metadata = bundle_identity(directory, task, time.monotonic() + 10, file_digests=files)
    root = Path(directory)
    model_bytes = (root / 'model.onnx').read_bytes()
    tokenizer_bytes = (root / 'tokenizer.json').read_bytes()
    if hashlib.sha256(model_bytes).hexdigest() != files['model.onnx'] or hashlib.sha256(tokenizer_bytes).hexdigest() != files['tokenizer.json']:
        raise ValueError('onnx_changed')
    graph = onnx.load_model_from_string(model_bytes)
    pending = [graph.graph]
    while pending:
        part = pending.pop()
        tensors = list(part.initializer)
        for sparse in part.sparse_initializer:
            tensors.extend((sparse.values, sparse.indices))
        for node in part.node:
            if node.domain not in ('', 'ai.onnx', 'ai.onnx.ml', 'com.microsoft'):
                raise ValueError('onnx_domain')
            for attribute in node.attribute:
                tensors.extend(attribute.tensors)
                if attribute.HasField('t'):
                    tensors.append(attribute.t)
                pending.extend(attribute.graphs)
                if attribute.HasField('g'):
                    pending.append(attribute.g)
        if any(tensor.data_location == onnx.TensorProto.EXTERNAL or tensor.external_data for tensor in tensors):
            raise ValueError('onnx_external_data')
    del graph
    tokenizer = Tokenizer.from_str(tokenizer_bytes.decode('utf8'))
    del tokenizer_bytes
    tokenizer.no_truncation()
    tokenizer.enable_padding(pad_id=1, pad_token='<pad>')
    options = ort.SessionOptions()
    options.intra_op_num_threads = 2
    options.inter_op_num_threads = 1
    options.execution_mode = ort.ExecutionMode.ORT_SEQUENTIAL
    options.log_severity_level = 4
    session = ort.InferenceSession(model_bytes, sess_options=options, providers=['CPUExecutionProvider'])
    names = {entry.name for entry in session.get_inputs()}
    if not names <= {'input_ids', 'attention_mask', 'token_type_ids'} or not {'input_ids', 'attention_mask'} <= names:
        raise ValueError('onnx_inputs')
    results = []
    for offset in range(0, len(texts), 4):
        batch = texts[offset:offset + 4]
        if task == 'embedding':
            prefix = metadata['query_prefix'] if query_mode else metadata['document_prefix']
            encodings = tokenizer.encode_batch([prefix + text for text in batch])
        else:
            encodings = tokenizer.encode_batch([(query, text) for text in batch])
        if any(len(row.ids) > 512 for row in encodings):
            raise ValueError('onnx_token_capacity')
        arrays = {'input_ids': np.array([row.ids for row in encodings], dtype=np.int64),
                  'attention_mask': np.array([row.attention_mask for row in encodings], dtype=np.int64),
                  'token_type_ids': np.array([row.type_ids for row in encodings], dtype=np.int64)}
        output = session.run(None, {name: arrays[name] for name in names})[0]
        if task == 'embedding':
            if output.ndim != 3 or output.shape[2] != 384:
                raise ValueError('onnx_output')
            mask = arrays['attention_mask'][..., None]
            pooled = (output * mask).sum(axis=1) / mask.sum(axis=1)
            norms = np.linalg.norm(pooled, axis=1, keepdims=True)
            if np.any(norms <= 0):
                raise ValueError('onnx_output')
            values = pooled / norms
        else:
            if output.shape != (len(batch), 1):
                raise ValueError('onnx_output')
            values = output[:, 0]
        if not np.isfinite(values).all():
            raise ValueError('onnx_output')
        results.extend(values.tolist())
    return {'values': results, 'fingerprint': fingerprint}

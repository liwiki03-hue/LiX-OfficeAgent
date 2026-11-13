import json
from pathlib import Path
import sys
import tempfile
import time

from .diagram_process import _run_owned
from .embeddings import Embedder, EmbeddingConfig
from .fidelity import _temporary_root


class OfflineReranker:
    def __init__(self, model, timeout=8, *, runner=_run_owned):
        self.model, self.timeout, self.runner = model, timeout, runner
        self.last_run = {}

    @property
    def identity(self):
        if (Path(self.model) / 'research_model.json').is_file():
            from .research_onnx import bundle_identity
            fingerprint = bundle_identity(self.model, 'reranker', time.monotonic() + 4)[0]
        else:
            fingerprint = Embedder(EmbeddingConfig('local', self.model, 1))._manifest(time.monotonic() + 4)
        return {'model_sha256': fingerprint, 'inference_version': 'offline-pair-logit-v1', 'device': 'cpu', 'threads': 2}

    def rank(self, question, hits):
        original = list(hits)
        cleanup_uncertain = False
        try:
            if type(self.timeout) not in (int, float) or not 1 <= self.timeout <= 30 or len(hits) > 20:
                return original, 'configuration_invalid'
            if (Path(self.model) / 'research_model.json').is_file():
                from .research_onnx import run_owned
                import math
                self.last_run = {}
                scores = run_owned(self.model, 'reranker', [hit['text'] for hit in hits], self.timeout, query=question, telemetry=self.last_run)
                if len(scores) != len(hits) or any(type(value) not in (int, float) or not math.isfinite(value) for value in scores):
                    return original, 'invalid_response'
                return [hits[index] for index in sorted(range(len(scores)), key=lambda index: (-scores[index], index))], 'completed'
            validator = Embedder(EmbeddingConfig('local', self.model, 1))
            fingerprint = validator._manifest(time.monotonic() + 4)
            with tempfile.TemporaryDirectory(prefix='office-rerank-', dir=_temporary_root()) as raw:
                root = Path(raw)
                request, response = root / 'input.json', root / 'output.json'
                payload = json.dumps({'model': self.model, 'query': question, 'texts': [hit['text'] for hit in hits]}, ensure_ascii=False).encode('utf-8')
                if len(payload) > 131072:
                    return original, 'capacity'
                request.write_bytes(payload)
                result = self.runner((sys.executable, str(Path(__file__).with_name('research_reranker_worker.py')), str(request), str(response)), timeout_ms=int(self.timeout * 1000))
                if result.empty_confirmed is not True:
                    cleanup_uncertain = True
                    return original, 'cleanup_failed'
                if result.code == 'render_timeout':
                    return original, 'timeout'
                if result.exit_code != 0 or not response.is_file() or response.stat().st_size > 8192:
                    return original, 'unavailable'
                data = json.loads(response.read_bytes())
                order = data.get('order')
                if (type(order) is not list or any(type(value) is not int for value in order) or
                        sorted(order) != list(range(len(hits))) or validator._manifest(time.monotonic() + 4) != fingerprint):
                    return original, 'invalid_response'
                return [hits[index] for index in order], 'completed'
        except Exception as error:
            return original, 'cleanup_failed' if cleanup_uncertain or getattr(error, 'code', None) == 'renderer_cleanup_failed' or str(error) == 'onnx_cleanup_failed' else 'timeout' if str(error) == 'onnx_timeout' else 'unavailable'

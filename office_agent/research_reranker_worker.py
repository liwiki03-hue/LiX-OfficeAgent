import contextlib
import json
import math
import os
from pathlib import Path
import sys


def main():
    os.environ.update(HF_HUB_OFFLINE='1', TRANSFORMERS_OFFLINE='1', TOKENIZERS_PARALLELISM='false')
    with open(os.devnull, 'w') as sink, contextlib.redirect_stdout(sink), contextlib.redirect_stderr(sink):
        request = Path(sys.argv[1])
        if request.stat().st_size > 131072:
            return 2
        data = json.loads(request.read_bytes())
        from sentence_transformers import CrossEncoder
        import torch
        torch.set_num_threads(2)
        model = CrossEncoder(data['model'], device='cpu', local_files_only=True, trust_remote_code=False,
                             model_kwargs={'use_safetensors': True, 'local_files_only': True, 'trust_remote_code': False})
        scores = model.predict([(data['query'], text) for text in data['texts']], batch_size=4, show_progress_bar=False)
        scores = [float(score) for score in scores]
        if len(scores) != len(data['texts']) or any(not math.isfinite(score) for score in scores):
            return 3
        order = sorted(range(len(scores)), key=lambda index: (-scores[index], index))
        Path(sys.argv[2]).write_text(json.dumps({'order': order}), encoding='utf-8')
        return 0


if __name__ == '__main__':
    try:
        code = main()
    except Exception:
        code = 1
    sys.exit(code)

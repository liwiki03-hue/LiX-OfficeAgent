import contextlib
import json
import os
from pathlib import Path
import sys


def main():
    os.environ.update(HF_HUB_OFFLINE='1', TRANSFORMERS_OFFLINE='1', TOKENIZERS_PARALLELISM='false',
                      OPENBLAS_NUM_THREADS='2', OMP_NUM_THREADS='2', MKL_NUM_THREADS='2')
    request, response = map(Path, sys.argv[1:3])
    if request.stat().st_size > 1_200_000:
        return 2
    data = json.loads(request.read_bytes())
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    with open(os.devnull, 'w') as sink, contextlib.redirect_stdout(sink), contextlib.redirect_stderr(sink):
        from office_agent.research_onnx import infer
        result = infer(data['model'], data['task'], data['texts'], data.get('query'), data.get('query_mode', False))
        import win32api
        import win32process
        memory = win32process.GetProcessMemoryInfo(win32api.GetCurrentProcess())
        result['peak_working_set_bytes'] = memory['PeakWorkingSetSize']
        result['peak_commit_bytes'] = memory['PeakPagefileUsage']
    response.write_text(json.dumps(result, allow_nan=False), encoding='utf8')
    return 0


if __name__ == '__main__':
    try:
        code = main()
    except Exception:
        code = 1
    sys.exit(code)

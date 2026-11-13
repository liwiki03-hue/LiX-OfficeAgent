PROVIDER_DEFAULTS = {
    'deepseek': {'base_url': 'https://api.deepseek.com/v1', 'model': 'deepseek-chat'},
    'dashscope': {'base_url': 'https://dashscope.aliyuncs.com/compatible-mode/v1', 'model': 'qwen-plus'},
    'ollama': {'base_url': 'http://127.0.0.1:11434', 'model': 'qwen2.5:7b'},
}

LIMITS = {
    'context_window': [1024, 2097152, True],
    'embedding_dimension': [1, 4096, True],
    'embedding_timeout': [1, 20, False],
    'embedding_budget': [1, 60, False],
}

MODEL_PATTERN = r'[A-Za-z0-9][A-Za-z0-9._:/@+-]{0,199}'

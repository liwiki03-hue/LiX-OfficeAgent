import asyncio
import importlib
from types import SimpleNamespace


def load_entry(monkeypatch):
    monkeypatch.setattr('dotenv.load_dotenv', lambda *args, **kwargs: None)
    return importlib.import_module('main')


def test_no_question_never_constructs_or_calls_model(monkeypatch, capsys):
    module = load_entry(monkeypatch)
    monkeypatch.setattr(module, 'ARGS', [])
    calls = []
    monkeypatch.setattr(module, 'build_llm', lambda *args: calls.append(args))
    asyncio.run(module.main())
    assert calls == []
    assert '显式提供问题后才会请求模型' in capsys.readouterr().out


def test_explicit_question_keeps_credentials_and_raw_failure_private(monkeypatch, capsys):
    module = load_entry(monkeypatch)
    monkeypatch.setattr(module, 'ARGS', ['deepseek', '仅测试输入'])
    secret = 'synthetic-sensitive-example'
    monkeypatch.setattr(module, 'read_api_key', lambda *args: secret)
    monkeypatch.setattr(module, 'build_llm', lambda *args: SimpleNamespace(model_name='fixture'))
    monkeypatch.setattr(module, 'build_agent', lambda *args: object())
    calls = []

    async def fail(agent, question):
        calls.append(question)
        raise RuntimeError('provider response: ' + secret)

    monkeypatch.setattr(module, 'ask', fail)
    asyncio.run(module.main())
    output = capsys.readouterr()
    assert calls == ['仅测试输入']
    assert secret not in output.out + output.err
    assert 'provider response:' not in output.out + output.err
    assert 'key=已配置' in output.out

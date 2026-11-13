"""
最小可运行 AI Agent（对应课程 3-7 / 3-8 / 4-1 / 4-2 / 6-3）

包含：
  1. 三模型切换（Ollama 本地 / 阿里云百炼 Qwen / DeepSeek 官网 API）
  2. @tool 装饰器定义工具 + args_schema 精确控制入参
  3. LangChain 1.x / LangGraph 兼容的建 Agent 方式
  4. 流式输出 + 工具调用轨迹打印 + 异常兜底 + 防死循环

运行：uv run python main.py
"""
from __future__ import annotations

import asyncio
import os
import sys

from pathlib import Path

from dotenv import load_dotenv
from office_agent.settings_definition import PROVIDER_DEFAULTS
from pydantic import BaseModel, Field, SecretStr

from langchain_core.messages import AIMessageChunk, ToolMessage
from langchain_core.tools import tool

# 显式指定 .env 路径 + override=True：
#   1) 无论从哪个目录启动都能读到同一个 .env
#   2) 即使系统里已存在同名环境变量，也以 .env 为准（load_dotenv 默认不覆盖，容易踩坑）
ENV_PATH = Path(__file__).with_name(".env")
load_dotenv(ENV_PATH, override=True)

# Windows 控制台默认 GBK，模型一旦返回 emoji / 生僻符号会 UnicodeEncodeError 直接崩：
# 保留原编码，只把无法编码的字符替换掉。
for _stream in (sys.stdout, sys.stderr):
    _reconfigure = getattr(_stream, "reconfigure", None)
    if callable(_reconfigure):
        _reconfigure(errors="replace")

SYSTEM_PROMPT = (
    "你是一个严谨的助手，始终用简体中文回答。"
    "需要外部信息时必须调用工具获取，严禁编造工具没有返回的信息；"
    "不要在回答正文前输出英文过渡句；工具报错时先分析原因，再决定是否重试。"
)

DEFAULT_QUESTIONS = [
    "北京今天天气怎么样？适合跑步吗？",
    "年假有多少天？顺便看看上海的天气。",
    "帮我写一首关于秋天的五言诗。",
]

# 命令行参数：provider 名用来临时切模型，其余参数当作要提问的问题
PROVIDER_CHOICES = ("ollama", "dashscope", "deepseek")

ARGS = [a for a in sys.argv[1:] if a.strip()]


def resolve_provider() -> str:
    for arg in ARGS:
        if arg.lower() in PROVIDER_CHOICES:
            return arg.lower()
    return os.getenv("MODEL_PROVIDER", "dashscope").strip().lower()


def resolve_questions() -> list[str]:
    questions = [a for a in ARGS if a.lower() not in PROVIDER_CHOICES]
    return questions


# 云端 provider 配置表：provider -> (Key 环境变量, 默认 base_url, 默认模型, 去哪申请 Key)
# dashscope 与 deepseek 都是 OpenAI 兼容接口，所以共用 ChatOpenAI，只换 base_url + key + model
PROVIDERS = {
    "dashscope": (
        "DASHSCOPE_API_KEY",
        PROVIDER_DEFAULTS['dashscope']['base_url'],
        PROVIDER_DEFAULTS['dashscope']['model'],
        "阿里云百炼控制台 https://bailian.console.aliyun.com/ → 右上角「API-KEY」"
        "（注意：要的是百炼/Model Studio 的 Key，不是阿里云 AccessKey ID/Secret）",
    ),
    "deepseek": (
        "DEEPSEEK_API_KEY",
        PROVIDER_DEFAULTS['deepseek']['base_url'],
        PROVIDER_DEFAULTS['deepseek']['model'],
        "DeepSeek 开放平台 https://platform.deepseek.com/api_keys → 创建 API Key",
    ),
}


def key_help(provider: str) -> str:
    env_name, _, _, where = PROVIDERS[provider]
    return f"""\
[配置缺失] {env_name} 未填写或仍是占位符，无法调用 {provider} 云端模型。

  1) 打开本目录下的 .env（没有就从 .env.example 复制一份）
  2) 去这里创建并复制 Key：{where}
  3) 写入 {env_name}=sk-你的真实Key（不要加引号、行尾不要留空格）
  4) 重新运行：uv run python main.py {provider}

  也可以直接用本地模型试跑（无需 Key）：uv run python main.py ollama
"""


def read_api_key(provider: str) -> str | None:
    """读取并清洗指定 provider 的 API Key；返回 None 表示未配置或仍是占位符。"""
    env_name = PROVIDERS[provider][0]
    raw = (os.getenv(env_name) or "").strip().strip('"').strip("'").strip()
    if not raw or len(raw) < 20 or not raw.startswith("sk-"):
        return None
    low = raw.lower()
    if any(h in low for h in ("xxx", "your", "changeme", "填入", "replace")):
        return None
    return raw


def auth_hint(exc: Exception) -> str:
    """把常见 HTTP 错误翻译成可执行的排查步骤。"""
    text = str(exc)
    low = text.lower()
    if "401" in text or "invalid_api_key" in low or "incorrect api key" in low or "authentication" in low:
        return (
            "\n[提示] 401 = Key 无效。排查顺序：\n"
            "  1) 确认 .env 里当前 provider你5研讨会 的 Key 是真实值（不能是 sk-xxx 占位符）：\n"
            "     dashscope → DASHSCOPE_API_KEY（百炼/Model Studio 控制台，不是阿里云 AccessKey ID/Secret）\n"
            "     deepseek  → DEEPSEEK_API_KEY（platform.deepseek.com/api_keys）\n"
            "  2) Key 前后不要有引号/空格/换行；改完 .env 必须重新运行程序；\n"
            "  3) 先运行 uv run python check_key.py 验证（能区分 Key 错还是网络/模型错）。"
        )
    if "402" in text or "insufficient balance" in low or "insufficient_quota" in low:
        return "\n[提示] 402/额度不足：该账号余额或免费额度已用尽，去对应控制台充值（DeepSeek 常见）。"
    if "404" in text or "model_not_found" in low or "model not exist" in low:
        return (
            "\n[提示] 模型名不存在：检查 DASHSCOPE_MODEL / DEEPSEEK_MODEL。\n"
            "  dashscope 可用：qwen-plus / qwen-max / qwen-turbo\n"
            "  deepseek  可用：deepseek-chat / deepseek-reasoner"
        )
    if "429" in text or "throttling" in low or "quota" in low:
        return "\n[提示] 限流或额度不足：降低并发，或到对应控制台确认余额。"
    if "connection" in low or "timeout" in low or "timed out" in low:
        return "\n[提示] 网络不通：公司代理请设置 HTTPS_PROXY；本地模型请确认 ollama serve 已启动。"
    return ""


# --------------------------------------------------------------------------- #
# 1. 模型接入                                                                  #
# --------------------------------------------------------------------------- #
def build_llm(provider: str | None = None):
    provider = (provider or resolve_provider()).strip().lower()

    if provider == "ollama":
        from langchain_ollama import ChatOllama

        return ChatOllama(
            model=os.getenv("OLLAMA_MODEL", PROVIDER_DEFAULTS['ollama']['model']),
            base_url=os.getenv("OLLAMA_BASE_URL", PROVIDER_DEFAULTS['ollama']['base_url']),
            temperature=0,
        )

    if provider not in PROVIDERS:
        raise SystemExit(f"未知 provider: {provider}，可选 {' / '.join(PROVIDER_CHOICES)}")

    _, base_url, default_model, _ = PROVIDERS[provider]
    key = read_api_key(provider)
    if key is None:
        raise SystemExit(key_help(provider))

    from langchain_openai import ChatOpenAI

    model = os.getenv(f"{provider.upper()}_MODEL", default_model)
    kwargs = dict(
        model=model,
        # SecretStr 让 key 在日志/repr 中显示为 **********，避免泄漏
        api_key=SecretStr(key),
        base_url=os.getenv(f"{provider.upper()}_BASE_URL", base_url),
        timeout=60,
        max_retries=2,
    )
    # 推理模型（deepseek-reasoner / qwen3-thinking）不接受 temperature，传了会报错
    if "reason" not in model and "thinking" not in model:
        kwargs["temperature"] = 0

    return ChatOpenAI(**kwargs)


# --------------------------------------------------------------------------- #
# 2. 工具定义                                                                  #
# --------------------------------------------------------------------------- #
_FAKE_WEATHER = {
    "北京": "晴，26℃，湿度 40%，北风 2 级",
    "上海": "多云，29℃，湿度 65%，东南风 3 级",
    "杭州": "小雨，24℃，湿度 80%，东风 2 级",
}


@tool
def get_weather(city: str) -> str:
    """查询指定城市的实时天气。当用户询问天气、气温、湿度、是否下雨时调用。

    Args:
        city: 城市中文名，例如「北京」。
    """
    return _FAKE_WEATHER.get(city, f"暂未收录 {city} 的天气数据")


class SearchInput(BaseModel):
    """知识库检索入参（用 args_schema 精确约束，对应课程 3-8）"""

    keyword: str = Field(description="搜索关键词，必须是名词短语，例如「年假制度」")
    limit: int = Field(default=3, ge=1, le=10, description="返回条数，1~10")


_DOCS = [
    "员工年假制度：入职满 1 年享 5 天带薪年假，每满 1 年增加 1 天，上限 15 天。",
    "报销流程：填写报销单 → 直属主管审批 → 财务复核 → 3 个工作日内到账。",
    "远程办公：每周最多 2 天，需提前一天在 OA 系统提交申请。",
]


@tool(args_schema=SearchInput)
def search_knowledge(keyword: str, limit: int = 3) -> str:
    """搜索公司内部知识库文档。当问题涉及公司制度、流程、规定时调用。"""
    try:
        hits = [d for d in _DOCS if keyword[:2] in d][:limit] or _DOCS[:limit]
        return "\n".join(f"{i + 1}. {h}" for i, h in enumerate(hits))
    except Exception as exc:  # 工具内部兜底：把错误当文本返回，让模型有机会自我修正
        return f"[工具错误] search_knowledge 执行失败：{exc}"


TOOLS = [get_weather, search_knowledge]


# --------------------------------------------------------------------------- #
# 3. 创建 Agent（兼容 LangChain 1.x 与旧版 LangGraph）                         #
# --------------------------------------------------------------------------- #
def build_agent(llm, tools, system_prompt: str):
    try:
        # LangChain 1.x 官方入口
        from langchain.agents import create_agent

        return create_agent(model=llm, tools=tools, system_prompt=system_prompt)
    except ImportError:
        # 旧环境兜底（视频里的写法，LangGraph 侧）
        from langgraph.prebuilt import create_react_agent

        return create_react_agent(llm, tools, prompt=system_prompt)


# --------------------------------------------------------------------------- #
# 4. 流式运行 + 轨迹打印                                                        #
# --------------------------------------------------------------------------- #
async def ask(agent, text: str) -> None:
    print(f"\n{'=' * 62}\n[用户] {text}\n{'-' * 62}")

    msg_id: str | None = None
    buffer: list[str] = []
    has_tool_call = False
    seen_tc: set[str] = set()

    def flush() -> None:
        """一条 AI 消息结束时才输出：如果这条消息只是在调工具，
        它前面的过渡语（DeepSeek 常见的 "I'll check ... for you."）直接丢弃。"""
        nonlocal buffer, has_tool_call
        if buffer and not has_tool_call:
            print("".join(buffer), flush=True)
        buffer.clear()
        has_tool_call = False

    async for chunk, _meta in agent.astream(
        {"messages": [{"role": "user", "content": text}]},
        stream_mode="messages",
        # 防止 ReAct 循环失控烧 token
        config={"recursion_limit": int(os.getenv("RECURSION_LIMIT", "25"))},
    ):
        if isinstance(chunk, AIMessageChunk):
            if chunk.id and chunk.id != msg_id:
                flush()  # 上一条 AI 消息结束
                msg_id = chunk.id
            for tc in chunk.tool_call_chunks or []:
                name, tc_id = tc.get("name"), tc.get("id") or ""
                if name and tc_id not in seen_tc:
                    seen_tc.add(tc_id)
                    has_tool_call = True
                    print(f"\n[调用工具] {name}")
            if chunk.content:
                buffer.append(chunk.content)
        elif isinstance(chunk, ToolMessage):
            flush()
            preview = str(chunk.content).replace("\n", " ")[:140]
            print(f"\n[工具返回] {chunk.name} → {preview}")
    flush()
    print()


async def main() -> None:
    questions = resolve_questions()
    if not questions:
        print('这是课程命令行示例；桌面应用请使用 Start-OfficeAgent.cmd。显式提供问题后才会请求模型。')
        return
    provider = resolve_provider()
    key = None if provider == "ollama" else read_api_key(provider)

    # 启动即校验，避免发了请求才拿到 401
    if provider != "ollama" and key is None:
        print(key_help(provider), file=sys.stderr)
        return

    llm = build_llm(provider)
    model_name = getattr(llm, "model_name", None) or getattr(llm, "model", "?")

    print(f"[环境] .env={ENV_PATH} | provider={provider} | model={model_name}")
    print(
        f"[环境] tools={[getattr(t, 'name', '?') for t in TOOLS]}"
        + (" | key=已配置" if key else " | key=本地模型无需 Key")
    )

    agent = build_agent(llm, TOOLS, SYSTEM_PROMPT)

    for q in questions:
        try:
            await ask(agent, q)
        except Exception as exc:
            print(f"\n[异常] {type(exc).__name__}；模型请求失败，请核对所选接口。", file=sys.stderr)
            hint = auth_hint(exc)
            if hint:
                print(hint, file=sys.stderr)


if __name__ == "__main__":
    asyncio.run(main())

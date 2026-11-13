# 贡献与回归

先阅读适用项目指令和能力边界。保持原生 DOM、Python 宿主与安全事务；模型不能获得 shell、任意 XML 或无界文件权限。新增实现不添加解释性注释/docstring，保留既有注释、版权、许可证和工具指令。

```powershell
uv lock --check
uv sync --locked --python 3.12.14
uv run python tools/install_frontend_assets.py
uv run python tools/ci_preflight.py
uv run pytest -q
```

先做定向检查，最终冻结候选仍需全量。诊断有限时且只清理自有进程，不并行干扰正式窗口或其他 Office/字体状态。按 ci_preflight 返回的 Playwright Node 路径加入 PATH；Edge/渲染器/Node 缺失应失败，不能静默跳过。

新增公开文件须审核 release/PUBLIC_FILES.json，用 release_source.py manifest/check/export 核对并导出。秘密模式例外必须绑定确切文件和匹配摘要并说明合成来源；扫描不等于 Git 历史无秘密。

PR 说明问题、行为、验证范围和未验事项；不删有效断言。项目自有代码采用 AGPL-3.0-only；提交贡献前确认有权按此许可提供代码，保留第三方来源和许可。不会仅凭 PR 推断贡献者拥有第三方版权。

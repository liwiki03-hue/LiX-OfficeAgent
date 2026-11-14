# LiX

**[下载 Q 修复预览 · Windows x64](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/download/v0.7.11-q-repair-20261114-preview/LiX-Setup-0.7.11-q-final-20261114-renderer-preview-x64.exe)** · [发行说明、对应源码与校验](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-q-repair-20261114-preview) · [旧预览版](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-preview)

本次 q-final-20261114 包含固定原版渲染器，其余环境首次准备仍需联网。安装器未签名，最终安装和完整首次准备尚未验收；前一诊断候选出现 uv 缓存访问拒绝。本机源码完整回归已通过，GitHub CI 以运行页面为准；旧预览的安装证据不代表本候选。

项目仓库：[liwiki03-hue/LiX-OfficeAgent](https://github.com/liwiki03-hue/LiX-OfficeAgent)。本仓库提供0.7.11开发源码，具体导入范围见[源码复核](docs/SOURCE_IMPORT_REVIEW.md)。技能安装、参考与更新的使用方式见[工作区技能](docs/SKILLS.md)。

LiX 是原 Office Agent 的正式名称，面向 Windows 的本机文档助手。已有配置、历史和文件安全链沿用；内部 Python 包名和旧启动入口保持兼容。

Windows 桌面文档编辑助手：打开文档，在本机编辑或提出修改要求，查看实际差异与验证副本。

当前公开提供 **0.7.11 Q 修复预览包及准确对应源码**。构建 q-final-20261114 已独立发布，标签绑定提交 `1beb85d` 的754项源码；其后 main 的发布说明更新不改变安装包内容。Q系列覆盖异步确认、配置候选、有界表格协作、网络和诊断。项目自有代码采用 AGPL-3.0-only，SuperDoc编译JS/CSS在准备时使用固定上游归档。详见[本次发行说明](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-q-repair-20261114-preview)、[验证范围](docs/VALIDATION.md)及[第三方来源](docs/SUPERDOC_SOURCE.md)。

## 三步上手

Windows 联网修复预览安装包已在[Releases](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-q-repair-20261114-preview)提供，默认按用户安装并由用户主动准备运行环境；最终安装、首次准备和干净机器验收仍未完成。构建、数据位置与卸载保留规则见[安装包说明](docs/WINDOWS_INSTALLER.md)。下方为源码运行方式。

1. 在项目目录准备依赖（本机基线 Windows 11 x64、Python 3.12）：

   ```powershell
   uv python install 3.12.14
   uv lock --check
   uv sync --locked --python 3.12.14 --no-dev
   Copy-Item .env.example .env
   ```

   仅在 `.env` 不存在时复制，保留原配置；模板不是可用密钥。

2. 准备独立渲染运行时，使用标准入口：

   ```powershell
   uv run python tools/install_frontend_assets.py
   .\tools\Install-DocumentRuntime.ps1
   .\Start-LiX.cmd
   ```

   需要 WebView2、.NET Framework 和可用字体。标准入口固定 Edge/netfx/pythonnet 3.0.5。下载失败或执行策略拦截见[安装指南](docs/INSTALL.md)，不要更改全局策略。

   本机项目文件夹还提供带图标的 `LiX.exe`，双击即可启动，内部仍调用上述标准入口。请保留在项目目录，不要单独移动 EXE；它不包含 Python、uv 或渲染运行时，也不是独立安装包。源码包不携带编译产物，可在 Windows 已有 .NET Framework 和 PowerShell 7 环境下运行 `uv run pwsh -NoProfile -File tools/build_lix_launcher.ps1` 重新生成。

3. 将自己的 DOCX、XLSX 或 PPTX 拖到悬浮球，核对读取范围。手工编辑无需模型密钥。主动发送修改要求后，默认修改草稿并自动生成验证副本；也可选“仅修改草稿”或停止。手动保存、只读简化和撤销仍有相应确认。

## 当前范围

| 文件 | 已接入 | 主要限制 |
| --- | --- | --- |
| DOCX | 分页编辑、普通文本和部分格式、普通表格有界引用与 AI 草稿、验证副本 | 固定 SuperDoc；表格引用限8个非空段落、4000字；复杂域、合并/嵌套及混合格式按能力只读，不等于 Word 全兼容 |
| XLSX | 有界网格、冻结行列、连续矩形选择、多格分析与普通格建议、验证副本 | 可视窗口30×16；逻辑范围最多100格，AI文字总量12000字符；数值逐格明确范围，公式和保护格不因此解除保护 |
| PPTX | 页图、已有普通文字片段、独立副本 | 不是任意对象编辑器；页图与草稿各自核对版本 |

只读DOCX可在支持范围内创建简化可编辑副本，再重新同意读取后编辑。正文和普通表格文字保留，原样式、自动编号和分页不保证保留；保护或复杂正文拒绝转换，不强制解锁原文件。

详见[能力边界](docs/CAPABILITIES.md)。旧 DOC/XLS/PPT 写入和真实 Office 自动化保持暂停。LibreOffice/PDFium 证据不等于 Microsoft Office 人工验收。

## AI 与数据

设置中配置连接和模型；问答、文档 Agent、生图各自选择。选择影响后续请求，已发出的任务保留原配置。获取模型列表不等于调用通过。

主动发送时获准的相关内容进入所选服务，可能计费；单机应用不代表 AI 离线。模型不能直接执行代码、XML 或文件写入。没有有效模型配置仍可手工编辑。真实 AI 主故事尚待当前候选上的独立验收。

## 文件保护

主动发送要求 → 宿主绑定当前文档授权 → 校验目标和旧值 → 修改草稿 → 备份 → 写新副本 → 重开/渲染校验 → 审计。默认自动存副本，“仅修改草稿”不发布文件；手动发布和撤销沿用相应确认。源变化、旧要求、旧计划或结果不确定会阻止重复执行，停止与历史恢复不会续用旧授权。原稿保持不变。

配置使用 `.env`。本机草稿、备份、审计与用户材料不进入源码包；不要上传密钥、正文或完整日志。请先用自建材料试用并保留重要源文件的独立备份。

## 开发与发行状态

[安装](docs/INSTALL.md) · [故障处理](docs/TROUBLESHOOTING.md) · [贡献](CONTRIBUTING.md) · [安全](SECURITY.md) · [变更](CHANGELOG.md) · [合成样例](examples/README.md)

最终完整门禁为 `uv run pytest -q`。Q修复源码在本机标准隔离环境已完成完整回归，R5历史成绩不替代本轮结果；实际范围和限制见[验证记录](docs/VALIDATION.md)与[本次发行说明](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-q-repair-20261114-preview)。GitHub CI 单独核对，安装器仍为预览候选。10 MiB代码更新ZIP是另一条交付路径；应用内自动更新的仓库和公钥仍未配置。

网络、官方PyPI候选锁与缓存隔离方式见[网络与诊断](docs/NETWORK.md)。默认保留原锁和现有环境，不自动切换来源或安装依赖。

项目自有代码采用 [AGPL-3.0-only](LICENSE)，[许可决定](docs/LICENSING.md)记录范围；[第三方声明](THIRD_PARTY_NOTICES.md)和 [NOTICE](NOTICE)保留第三方原许可。许可选定不代表发行验收完成。

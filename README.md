# LiX

**[下载 Windows x64 联网预览版](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/download/v0.7.11-preview/LiX-Setup-0.7.11-preview-x64.exe)** · [发行说明、源码与校验](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-preview)

首次使用需联网准备环境。此预览未签名，完整 CI 尚未通过；当前机器已跑通安装、首次准备、悬浮球与工作台。

项目仓库：[liwiki03-hue/LiX-OfficeAgent](https://github.com/liwiki03-hue/LiX-OfficeAgent)。本仓库提供0.7.11开发源码，具体导入范围见[源码复核](docs/SOURCE_IMPORT_REVIEW.md)。技能安装、参考与更新的使用方式见[工作区技能](docs/SKILLS.md)。

LiX 是原 Office Agent 的正式名称，面向 Windows 的本机文档助手。已有配置、历史和文件安全链沿用；内部 Python 包名和旧启动入口保持兼容。

Windows 桌面文档编辑助手：打开文档，在本机编辑或提出修改要求，审阅差异后生成可验证的新副本。

当前提供 **0.7.11 Windows 联网预览安装包及对应开发源码**，完整发行验收仍未完成。最新改动完成选定基础检查；完整自动化的历史成绩不代表当前提交已全量通过。原生、Office及独立质量验收仍有缺项。项目自有代码采用 AGPL-3.0-only；本仓库不捆绑SuperDoc编译JS/CSS，安装时从上游固定归档准备，见[第三方来源](docs/SUPERDOC_SOURCE.md)。详见[验证范围](docs/VALIDATION.md)。

## 三步上手

Windows 联网预览安装包已在[Releases](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-preview)提供，默认按用户安装并由用户主动准备运行环境；干净机器验收仍未完成。构建、数据位置与卸载保留规则见[安装包说明](docs/WINDOWS_INSTALLER.md)。下方为源码运行方式。

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

3. 将自己的 DOCX、XLSX 或 PPTX 拖到悬浮球，核对并同意读取。手工编辑无需模型密钥；保存先审阅并独立确认，副本验证成功才显示完成。

## 当前范围

| 文件 | 已接入 | 主要限制 |
| --- | --- | --- |
| DOCX | 分页编辑、普通文本和部分格式、受控 AI 草稿、独立副本 | 固定 SuperDoc；复杂域、修订、表格结构按能力检查拒绝，不等于 Word 全兼容 |
| XLSX | 有界网格、冻结行列、部分格式、同表受限公式、独立副本 | 不是全函数 Excel；当前窗口预算 30×16 |
| PPTX | 页图、已有普通文字片段、独立副本 | 不是任意对象编辑器；页图与草稿各自核对版本 |

只读DOCX可在支持范围内创建简化可编辑副本，再重新同意读取后编辑。正文和普通表格文字保留，原样式、自动编号和分页不保证保留；保护或复杂正文拒绝转换，不强制解锁原文件。

详见[能力边界](docs/CAPABILITIES.md)。旧 DOC/XLS/PPT 写入和真实 Office 自动化保持暂停。LibreOffice/PDFium 证据不等于 Microsoft Office 人工验收。

## AI 与数据

设置中配置连接和模型；问答、文档 Agent、生图各自选择。选择影响后续请求，已发出的任务保留原配置。获取模型列表不等于调用通过。

主动发送时获准的相关内容进入所选服务，可能计费；单机应用不代表 AI 离线。模型不能直接执行代码、XML 或文件写入。没有有效模型配置仍可手工编辑。真实 AI 主故事尚待当前候选上的独立验收。

## 文件保护

计划 → 独立确认 → 备份 → 写新副本 → 重开/渲染校验 → 审计 → 独立撤销。源变化、旧计划或结果不确定会阻止重复执行。接受 AI 建议只更新草稿，不等于批准写文件；历史不恢复旧批准。

配置使用 `.env`。本机草稿、备份、审计与用户材料不进入源码包；不要上传密钥、正文或完整日志。请先用自建材料试用并保留重要源文件的独立备份。

## 开发与发行状态

[安装](docs/INSTALL.md) · [故障处理](docs/TROUBLESHOOTING.md) · [贡献](CONTRIBUTING.md) · [安全](SECURITY.md) · [变更](CHANGELOG.md) · [合成样例](examples/README.md)

最终完整门禁为 `uv run pytest -q`。R5历史冻结候选为 **4126 passed、5 skipped、165 subtests passed**，不适用于之后的产品改动。当前源码按维护者要求只运行选定基础检查，不能称为完整验收。范围与未验项见[验证记录](docs/VALIDATION.md)和[源码复核](docs/SOURCE_IMPORT_REVIEW.md)。没有发布安装包，10 MiB代码更新ZIP不是安装器；应用更新仓库和公钥未配置。

项目自有代码采用 [AGPL-3.0-only](LICENSE)，[许可决定](docs/LICENSING.md)记录范围；[第三方声明](THIRD_PARTY_NOTICES.md)和 [NOTICE](NOTICE)保留第三方原许可。许可选定不代表发行验收完成。

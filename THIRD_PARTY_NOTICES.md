# 第三方来源与许可

保留原始版权、许可和工具指令，不把第三方代码重新授权为项目许可。完整 Python 锁定清单是 uv.lock；源码候选不含 .venv、MSI 或已安装运行时。

| 资产 | 固定版本/来源 | 许可及适配 |
| --- | --- | --- |
| SuperDoc | npm superdoc@1.47.2；上游源候选 66c22186c7c0735360ca4fd94f7499898415e903；release/SUPERDOC_SOURCE.json | AGPL-3.0，Git仅保留原LICENSE和来源摘要，不捆绑编译JS/CSS；由tools/install_frontend_assets.py从官方固定归档准备相同字节。1,762项内嵌项目源内容匹配固定提交，但完整对应源码、版本处理和重建仍未收口，见docs/SUPERDOC_SOURCE.md |
| Mermaid | 11.12.2，static/vendor/mermaid.manifest.json | MIT，保留 mermaid.LICENSE 和资产摘要 |
| Lucide 图标及本机 PNG | lucide-icons/lucide 提交 53fab88ff9d1b48f4704b5b5b560d6b4d560bf4c；release/LUCIDE_SOURCE.json | ISC 及部分 Feather 来源的 MIT，原 LICENSE 完整保留；27 个 SVG 与固定源码逐字节一致，5 个 PNG 从本地 SVG 经 QtSvg 派生 |
| IEEE CSL | citation-style-language/styles 提交 6ba2e9c6aa6e0dab87d3a20b833ed8ff6500f698 | CC-BY-SA-3.0，未经修改；作者/来源保留于 CSL info 及 assets/csl/NOTICE.md |
| tiktoken 数据 | cl100k_base，运行库 0.14.0 | MIT，保留 assets/TIKTOKEN_LICENSE.txt；OpenAI、Shantanu Jain |
| PySide6/Qt | 6.11.2 | 元数据 LGPL-3.0-only OR GPL-2.0-only OR GPL-3.0-only；具体模块许可随安装包保留 |
| pythonnet / python-docx / python-pptx / openpyxl | 3.0.5 / 1.2.0 / 1.0.2 / 3.1.5 | MIT，安装包保留上游通知 |
| grill-me 离线测试材料 | mattpocock/skills，a6bdfd9fed6c17d21b306aeb7ff6a0de8b72ef3c，grill-me/SKILL.md 与根 LICENSE | MIT，Copyright (c) 2026 Matt Pocock；原始文件、来源摘要与许可保存在 tests/fixtures/public_skills，指令正文未改，只转换 LiX 单文件封装；列入固定历史精选，不是默认安装或上游最新版本 |
| 精选来源元数据 | office_agent/skill_catalog.json：Anthropic skills固定2c5958ecac7cee2b583967a2c4981d0f4a75c579的internal-comms/theme-factory/docx/xlsx/pptx及上述grill-me | 对应许可文件路径/摘要逐项记录；五项为Apache-2.0，grill-me为MIT。随应用只提供名称、来源、摘要与自行编写的中文简介，不捆绑新增技能正文/图片或静默安装。六项不是LiX官方原创技能 |
| OpenAI插件补充目录协议参考 | openai/plugins固定d420d13051bffe4543ed39e30fa46b87148e47dd：.agents/plugins/marketplace.json及plugins/notion、plugins/superpowers/.codex-plugin/plugin.json | 只读核对来源结构，自行实现适配器，不复制插件实现或随发行捆绑本轮下载文件。插件声明的MIT不代替具体Skill许可；本轮两个检查样例未安装。来源/固定版本/拒绝原因见本地交付证据 |
| pywebview | 6.2.1 | BSD-3-Clause，标准入口为 Edge/netfx |
| pypdfium2 | 5.14.0 | BSD-3-Clause、Apache-2.0 及 PDFium 依赖许可，完整依赖通知不可省略 |
| citeproc-py / Playwright | 0.11.1 / 1.63.0 | BSD-2-Clause-Views / Apache-2.0，后者为开发组 |
| LibreOffice | 官方归档 26.2.6.2 x64 MSI | 外部下载，不打包进源码；保留下载器摘要和原安装包许可 |

安装器候选使用 Inno Setup 7.1.1（is-7_1_1）官方编译器，下载归档和许可摘要固定在 release/WINDOWS_INSTALLER.json。原安装引擎通知保留，自行编写 .iss 配置，不复制第三方界面实现。随候选携带的 uv.exe 为官方 0.12.18 x64 归档，提交 01cb90c1a4f88af09906cb60de9766d2add4a062，原 Apache-2.0/MIT 两份许可完整保存在安装目录 licenses。来源及具体文件见[安装包说明](docs/WINDOWS_INSTALLER.md)；用户首次准备下载的 Python、锁定依赖、SuperDoc 与 LibreOffice 不包含在安装器中，原许可与捆绑发行缺项继续适用。

本轮新脚本自行实现，未粘贴第三方实现。CI 参考 uv 官方锁定同步设计；actions/checkout 固定 11bd71901bbe5b1630ceea73d27597364c9af683（v4.2.2），astral-sh/setup-uv 固定 37802adc94f370d6bfd71619e3f0bf239e1f3b78（v7.6.0）。Action 不打包进应用。

来源：[SuperDoc 固定归档](https://registry.npmjs.org/superdoc/-/superdoc-1.47.2.tgz)、[源码仓库](https://github.com/superdoc/docx-editor)、[CSL 固定源码](https://github.com/citation-style-language/styles/tree/6ba2e9c6aa6e0dab87d3a20b833ed8ff6500f698)。不能用当前分支冒充无法核对的固定源码。

安装元数据只是清点依据，不是完整兼容性意见。发行前仍需审核根许可、贡献归属、传递依赖和实际随包资产。历史基线回归资产同样保留第三方许可。

逐包版本、安装元数据和原许可证位置见 release/DEPENDENCY_NOTICES.json；未安装的可选/平台依赖明确标为未核对，不自动安装大型 extras。

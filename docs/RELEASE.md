# 发布与源码候选

## 2026-11-14 Q续接本地身份

本次只读核对公开 main 仍为80bdc6bd2b063a271dafb58212d2cac3a77fe758，CI run43442448484仍失败；公开 v0.7.11-preview 主安装附件未被本轮替换。旧 Q-repair 源码ZIP原摘要保留。

本地本轮以 q-final-20261114 区分最终候选，正式二进制必须经原系统临时构建流水线，源码使用显式公开白名单及逐成员摘要。最终标准隔离环境完整回归已通过：4266 passed, 2 skipped, 165 subtests passed in 3297.90s (0:54:57)，exit 0，自有 Job 清空；诊断候选的打包成功不等于实际首次准备通过，当前仍有依赖缓存访问拒绝。最终包/源码/清单及实际结果见本机 deliverables/issue-repair-completion-20261114/CANDIDATE.json 和 VALIDATION.json。无签名、360判定、干净机及完整对应源码重建证据时不记稳定发行，不上传未知新包或移动旧标签。

## 2026-11-14 · 身份分别绑定

本轮只读核对：公开main为 `80bdc6bd2b063a271dafb58212d2cac3a77fe758`；`v0.7.11-preview`标签仍指向 `e9c70535c63975e458d40023a6b52e10fe3c38f2`。主安装附件376103211字节，SHA256 `6fc575243bd52418988d21a7975ba95e24954864fe0488ea0887af9a78881471`；其准确renderer源码附件为 `LiX-Source-0.7.11-renderer-preview.zip`，SHA256 `3c7517a777ff8594a47bc60816fa81335841570de4666b39907bfe16adaecfe2`。

本地Q系列工作目录不是Git仓库，仍含更晚代码。不能从main或旧标签推断附件对应源码，更不能把旧安装包当作本轮改动已发布。本轮仅复核远端元数据，没有重新完整下载安装附件、移动标签或上传新包。新的本地源码审阅候选按公开白名单/秘密模式扫描/ZIP重开核对；没有新安装器、未通过完整回归时均不具备发行条件。

本轮本地构建身份为 `Q-repair-20261114`，产品版本仍为0.7.11，没有新标签。一次完整集合4196通过、62失败、5跳过、165子测试通过；其中60项失败经修复产品或更新当前流程断言后定向通过，项目CLR加载和CS0016编译拒绝仍未解决，没有再次运行最终同源码全量。最新机器状态在[验证摘要](../release/VALIDATION_SUMMARY.json)的current_unreleased，历史R5数据保留且不能用于当前源文件。源码审阅ZIP不是安装包或公开发行。

## 2026-11-13 · Windows 联网预览已公开

维护者曾授权上传安装包，[v0.7.11-preview](https://github.com/liwiki03-hue/LiX-OfficeAgent/releases/tag/v0.7.11-preview)已公开。该段描述首次轻量预览构建期；后续renderer主附件替换与准确源码身份以上方2026-11-14记录为准，不将标签730项旧源码称作当前附件准确源码。

当前机器上的安装、完整首次准备、原生悬浮球和工作台已验证，30 项基础检查通过。安装器未签名，上一源码完整 CI 失败；干净 Windows、标准快捷方式重启、安装版文档副本及其他 Office/服务/物理验收仍开放。公开预览不代表 R4/R6 整体完成，也不启用自动更新。

<details>
<summary>历史：此前源码候选准备记录</summary>

# 源码候选准备

2026-11-13维护者新授权：最终确认后向既有LiX仓库导入开发源码。此授权取代下方历史记录中的“未授权源码上传”，不创建标签/Release、安装器或自动更新。当前导入范围和基础验证见[源码复核](SOURCE_IMPORT_REVIEW.md)，SuperDoc编译JS/CSS不进入Git，使用原版本固定摘要主动准备。下方R5/R6与原生/对应源码条件继续作为独立发行缺项，不因源码导入自动关闭。

2026-11-13更新：维护者已建立[liwiki03-hue/LiX-OfficeAgent](https://github.com/liwiki03-hue/LiX-OfficeAgent)，本轮只读核对main为 `d9bb877b03ac45c2b15dba81116245bae32773a5`，14个说明/配置文件，无应用源码或Release。对应[导入规则](GITHUB_SETUP.md)已按原字节纳入本地；远端历史与模板需在独立候选中保留。未执行推送、建标签或发行。本地仍无.git，当前工作树新增S系列产品/测试后不能沿用下方R5/R6全量成绩；本轮仅基础检查。应用更新渠道仍为空，缺签名公钥与签名发行资产，不能仅填仓库名即宣称可更新。

当前为 0.7.11 本地源码预览候选。R5 自动化已通过，R6 已进入本地发行资料准备；没有公开上传或生成安装器。维护者已选择 AGPL-3.0-only；SuperDoc 完整对应源码交付、其他智能体负责的原生/真实服务验收和公开下载复验仍未完成。见[验证范围](VALIDATION.md)和[固定源码核对](SUPERDOC_SOURCE.md)。

```powershell
uv run python tools/release_source.py manifest --manifest SOURCE_MANIFEST.json
uv run python tools/release_source.py check --manifest SOURCE_MANIFEST.json
uv run python tools/release_source.py export --manifest SOURCE_MANIFEST.json --target ..\office-agent-candidate
uv run python tools/release_source.py archive --manifest SOURCE_MANIFEST.json --target ..\office-agent-source-review.zip
```

目标须不存在。精确清单为 release/PUBLIC_FILES.json，拒绝私有路径、链接、大小写冲突、超限、摘要变化与未审查秘密模式。失败候选保留，不自动删除或继续发布。复制并复核后生成 SOURCE_MANIFEST.json；源码 ZIP 重开逐项核对内容并返回包 SHA256，它不是应用更新包或安装器。

新目录锁定同步和渲染准备，不复制开发环境。同机新目录不是新 Windows 机器，缓存复用不是离线安装包。当前无 .git，Git 索引/历史秘密检查未验。公开清单不是许可证权属证明。

独立候选完整命令发现漏带两个旧测试依赖的 tools 模块、PowerShell 入口和桌面仍引用的 main.py；已保留失败并补齐清单。main.py 无显式问题时不再自动发样例，也不再打印密钥前缀或原始服务错误。另将项目外的生图 DOM 探针迁入 tests/probes，使用 Playwright 自带 Node，避免因开发机路径不存在而跳过原有断言。上轮开发目录加载路径可掩盖漏带文件，因此历史收集结果不作为新用户可运行证明。

CI 草案为 Windows 2022/Python 3.12.14，Action 固定提交、contents:read、无秘密与发布任务。未在 GitHub 运行不能称 CI 通过；全量日志和确定性 skip 保留，Edge/Node/渲染器缺失在前置检查失败。

首发仍安排 11 月 13 日 23:59（Asia/Shanghai），预留至少 6 小时冻结/验收，缺条件明确报告，不自动顺延。正式发布需确认 owner/repo、许可和具体动作；上线后还须公开下载复验。

R6 的文档和来源记录更新不改变 R5 产品/测试字节。新源码清单包含更新后的说明，完整自动化证据仍绑定其 R5 原清单；只有逐项证明产品、依赖、入口、资产和测试一致，才可沿用该结果，不能将它表述为本轮再次全量通过。公开链接、文件清单、源码归档和重开核对分别重新执行。

发布应从独立候选目录建立 Git 索引，只加入 release/PUBLIC_FILES.json 中列出的文件，执行 `tools/release_source.py check --require-git` 并补历史秘密检查。生成的 SOURCE_MANIFEST.json 不在 Git 索引中，作为发行附件并随源码 ZIP 提供。不得在开发目录直接 git add -A。目标仓库与具体发布动作尚未授权；准备命令不代表执行推送。原生、真实 AI 和对应源码条件未满足时保持本地候选。

</details>

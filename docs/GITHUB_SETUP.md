# GitHub 配置与后续源码导入

日期：2026-11-13。仓库为 `liwiki03-hue/LiX-OfficeAgent`，产品展示名称为 **LiX — AI 文档编辑助手**。

## 当前上传范围

当前只有项目首页、许可与第三方清点、忽略规则、Issue/PR 模板和发行说明分类；应用源码、构建依赖、测试、安装包和本机材料均未上传。没有配置自动发布、没有创建发行版本。

| 文件 | 用途 |
| --- | --- |
| README.md | 项目定位、使用流程和当前仓库状态 |
| .gitignore | 排除秘密、本机配置、临时文件和产物；不能代替白名单与历史检查 |
| .gitattributes | 禁用自动换行转换，保留启动脚本和固定资产的既有字节 |
| LICENSE | 既有 AGPL-3.0 许可证全文；项目自有代码选用 AGPL-3.0-only |
| NOTICE、THIRD_PARTY_NOTICES.md | 项目许可和本地依赖清点，不宣称第三方交付完成 |
| .github/ISSUE_TEMPLATE/ | 故障和功能建议表单 |
| .github/pull_request_template.md | 要求区分测试、浏览器、原生、Office 和真实服务证据 |
| .github/release.yml | 发行说明分类，不负责打包或发布 |
| SECURITY.md、CONTRIBUTING.md | 反馈与贡献边界 |

## 页面设置建议

这些是供维护者选择的建议，不表示已自动修改 GitHub 设置。

- 默认分支 main；确认实际分支与工作流检查存在后再设置分支规则。
- Issues 用于问题与建议。没有维护需求时可以关闭 Wiki 和 Projects。
- 按需要启用自动删除已合并分支和 Squash 合并。
- 建立 bug、enhancement、ui、documentation、dependencies、build、breaking-change 标签以使用发行说明分类；配置文件不会自动创建标签。
- 若准备接收私密漏洞报告，先在 GitHub 启用该能力，再在安全文档中明确入口。
- Topics 可使用 windows、desktop-app、ai-assistant、document-editor、docx、xlsx、pptx、python、pywebview。

## 后续导入源码

本地项目已有严格的 `release/PUBLIC_FILES.json` 和 `tools/release_source.py`。正式导入必须使用重新核对的白名单和摘要，不能直接上传原始开发目录。

相对原有公开清单，本次新增文件包括 `.gitattributes`、`.github/ISSUE_TEMPLATE/feature_request.yml`、`.github/pull_request_template.md`、`.github/release.yml` 与 `docs/GITHUB_SETUP.md`；合入时同步清单，重新生成 manifest。其他说明文件也有内容变化，旧摘要不能复用。

保留当前 Git 历史，通过普通提交导入，不强制覆盖远端。上传前排除 `.env`、个人文档、数据库、备份、草稿、完整日志、浏览器配置、模型权重和运行时。忽略规则不会自动移除已经提交的秘密，也不能证明历史安全。

本地已有 `.github/workflows/verify.yml`，依赖完整源码、公开清单、浏览器和渲染器；本次没有上传它以免运行缺少源码的任务。源码齐备后复用既有工作流，最终命令仍为 `uv run pytest -q`，不能用配置校验代替产品回归。

创建仓库不自动启用应用更新。更新地址、公钥和发布渠道仍按应用现有设置机制配置；未核对的新机器安装、第三方对应源码、重建及真实服务验收不自动记为通过。

## 配置方法参考

- [GitHub 问题表单](https://docs.github.com/en/communities/using-templates-to-encourage-useful-issues-and-pull-requests/configuring-issue-templates-for-your-repository)
- [GitHub 发行说明分类](https://docs.github.com/en/repositories/releasing-projects-on-github/automatically-generated-release-notes)

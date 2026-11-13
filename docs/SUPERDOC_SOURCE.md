# SuperDoc 固定源码核对

2026-11-13源码导入调整：Git公开清单不包含`superdoc.min.js`和`style.css`编译产物，保留原LICENSE、固定版本与摘要。新用户通过`uv run python tools/install_frontend_assets.py`主动从npm官方固定归档准备原字节资源；不自动安装上游构建依赖，不宣称下述完整对应源码/重建缺项已完成。开发机已有资源保持原字节，不因公开清单调整被删除。后续若分发含这些构建产物的安装器/归档，仍须先收口对应源码交付条件。

当前使用原封不动的 npm `superdoc@1.47.2` JS、CSS 和 LICENSE。固定包摘要及原许可见 [机器记录](../release/SUPERDOC_SOURCE.json)。本次只补来源材料，没有更换引擎或编译产品资产。

## 已确定的来源

此前只查默认分支得到 2.x，不能据此认为 1.x 没有公开源码。沿上游 v1 分支发布时间取得固定提交 [66c22186c7c0735360ca4fd94f7499898415e903](https://github.com/superdoc/docx-editor/tree/66c22186c7c0735360ca4fd94f7499898415e903)，提交时间为 2026-10-07 17:00:49 UTC。

npm 源映射共 2,386 条。排除 node_modules 和带查询参数的转换条目后，可按仓库路径检查 1,764 条：1,762 条 UTF-8 内容的 Git blob 与固定提交一致；一条是 Vite 的空 browser-external 虚拟模块，另一条 SVG 转为 data URI，解码后 XML 结构与仓库 SVG 一致。此检查没有证明全部依赖和转换输入相同。

固定构建入口包括：

- [根 package.json](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/package.json)、[pnpm-lock.yaml](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/pnpm-lock.yaml)、[workspace](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/pnpm-workspace.yaml)。该源码声明 Node 22.21.1、pnpm 10.25.0，不是 Office Agent 新增运行依赖。
- [Vite 配置](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/packages/superdoc/vite.config.js)、[CDN 配置](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/packages/superdoc/vite.config.cdn.js)、[打包脚本](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/packages/superdoc/scripts/pack-output.cjs)。CDN 配置从 package.json 读取版本，内嵌运行依赖。
- [原 LICENSE](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/LICENSE)与[第三方通知](https://github.com/superdoc/docx-editor/blob/66c22186c7c0735360ca4fd94f7499898415e903/THIRD_PARTY_LICENSES.md)。保留原文，尚未将通知存在等同完整分发审核。

14 个构建及通知文件共 1,835,200 字节已定向取得，已知 Git blob 逐项核对。catalog 依赖声明与 npm 展开值一致。完整仓库归档超过本次 128 MiB 下载预算后停止，残留文件只作失败证据，不作为完整源码归档。

## 仍需完成

1. 源码 package.json 为 1.46.2，npm 为 1.47.2；需要核对发布时版本处理和工作区依赖展开，不能直接替换版本后宣布复现。
2. 取得并验证完整构建材料，在独立目录按该版本流程构建，比较产物并解释差异。本次未安装上游 Node/pnpm 依赖或执行构建脚本。
3. 核对实际内嵌依赖通知及完整对应源码的交付路径。AGPL-3.0-only 已获维护者确认，不需要重新选择项目许可。

结论是“已定位内嵌项目模块的固定源候选”；完整对应源码交付及可重复构建仍未完成。不得以主分支、近似标签、源映射数量或本文件代替这些条件。

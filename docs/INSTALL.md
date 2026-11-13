# 安装与运行

当前本机准备基线：Windows 11 x64、CPython 3.12.14、uv 0.12.18。`requires-python >=3.10` 不代表所有版本或操作系统已验证。

使用独立源码目录，按 README 锁定同步；不复制开发者的 .venv、.env、.runtime、浏览器配置或模型缓存。python-pptx 已移入正常依赖，PPT 功能不再依赖安装测试组。默认不安装研究/OCR/本地大模型 extras。

系统需要 Microsoft Edge WebView2、.NET Framework、可用中文字体及下载/解压空间。LiX 入口 Start-LiX.cmd 转接既有 Start-OfficeAgent.cmd，使用隔离依赖和 pythonnet 3.0.5；单纯导入成功不能证明标准入口可用。两入口使用相同数据位置与单实例机制。

## 固定前端资源

源码仓库保留SuperDoc版本、摘要及原许可，不捆绑其编译JS/CSS。运行：

```powershell
uv run python tools/install_frontend_assets.py
```

脚本从npm官方固定`superdoc@1.47.2`归档下载，先核对整个归档和三个文件摘要，只准备原版本JS/CSS/LICENSE，不执行npm脚本、不安装Node依赖、不切换编辑引擎。下载和读取有界，不自动重试或更新版本；目标存在且摘要不符时保留原文件并拒绝覆盖。网络失败时不会将缺少编辑器资源记为可用。已全部核对的资源可离线复用，缓存位于`.runtime/downloads`，不进入Git。

## 渲染运行时

下载器固定官方归档 LibreOffice **26.2.6.2** x64 MSI：

[官方归档](https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.6.2/win/x86_64/LibreOffice_26.2.6.2_Win_x86-64.msi)

SHA256：`788ce7d4b56460357f57552cb8cd848e8f2254f28f6e51fa61e0c65a18096373`

tools/Install-DocumentRuntime.ps1 校验后在新 .runtime/libreoffice 执行 MSI 管理解压，不覆盖已有运行时；生成 manifest，初始 rendering_verified=false。旧 stable/26.2.6 地址失效后改用此固定归档，不修改当前开发目录已有运行时。

下载固定大小为 373293056 字节，每次最多 600 秒；中断仅保留 .partial。下次主动运行时需服务端确认准确的 Range 才能续传，完整 SHA256 核对通过后才改为正式归档名并解压；不自动循环重试。校验失败保留文件供检查，不作为可用运行时。解压另有 180 秒上限。

执行策略拦截时，不设置 Bypass 或修改全局策略。可以在 PowerShell 逐条执行脚本中审阅过的下载、摘要核对和 msiexec /a 管理解压命令，保持版本、目标目录、超时和存在性检查；也可由本机管理员安排本地脚本授权。摘要不符不得解压，已有目录不得删改，不使用来源不明的文件。

```powershell
uv run python examples/eval_fidelity.py --samples .runtime/render-samples --output .runtime/render-check.json
```

样例/输出位置必须尚未存在。该检查覆盖 DOCX 内容/样式差异与 XLSX 等价渲染，不是 Office 验收。缺少渲染器时保存闸门可能拒绝，不用空白图冒充通过。

## 升级

只在 .env 不存在时复制模板，密钥通过设置录入。无本地登录前置，外部服务账户另行配置。

升级保留原目录/配置与源文件，在新目录按锁文件准备；不覆盖正在运行的目录。回退到原目录后重新读取，不复用旧计划或批准。更新源未配置，同机新目录准备、全新电脑启动、公开下载复验分别记录。

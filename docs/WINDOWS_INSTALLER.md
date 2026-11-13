# Windows 安装包候选

本轮推进 R4 的 Windows 分发，两个目标为按用户安装及首次运行环境准备。版本仍为 0.7.11，通道为联网预览；构建脚本不是正式发行通过的证明。实际产物及验证结果以本地交付记录为准，尚未上传 Release。

## 使用流程

运行 `LiX-Setup-0.7.11-preview-x64.exe`，阅读许可并选择安装位置；默认当前用户的 `%LOCALAPPDATA%\Programs\LiX`，不请求管理员权限。开始菜单提供 LiX 和卸载入口，桌面快捷方式可选。

首次启动显示“首次准备”，由用户点击“准备并启动”后联网安装固定 Python 3.12.14、锁定生产依赖、固定 SuperDoc 1.47.2 资源及 LibreOffice 26.2.6.2。程序包携带核对过的 uv，不要求预装 Python、uv、Git、Node 或 PowerShell 7。下载不调用模型，不读取或搬迁开发目录配置，不生成文档写入批准。

需要 Windows 10 1809 或更新的 x64 环境、.NET Framework 4.8、WebView2 Runtime、中文字体及下载空间。首轮预留约 3 GB；这不是离线包。WebView2 缺失时应先从 [Microsoft 官方网站](https://developer.microsoft.com/microsoft-edge/webview2/)安装，检测不能用 Edge 浏览器存在代替。

准备按四个真实阶段显示，不模拟百分比。失败保留下载和诊断，用户主动重试；没有循环重试、自动更换来源或自动模型请求。CLI `LiX.exe --check` 只核对文件与准备状态，不下载或启动应用；`LiX.exe --prepare-only` 是明确的联网准备命令，不打开工作台或提交文档。

## 配置和卸载

设置仍只使用安装位置 `app\.env`，首次建立为空，不带示例密钥。应用读写、模型选择和独立确认仍来自原宿主代码。原开发入口不改，安装版入口使用相同 Edge/netfx/pythonnet 3.0.5 与 `office_agent.app --ui web`，使用独立锁定生产环境。

安装版 Python、环境、缓存和准备日志存放在 `%LOCALAPPDATA%\LiX`；文档渲染器及固定前端资源在安装目录的 `app` 下按现有相对位置准备。历史/本机会话仍沿用现有 `%LOCALAPPDATA%\OfficeAgent`。安装位置和本机配置路径不发送到 GitHub 或模型。

卸载只移除安装器登记的程序、源码和快捷方式；保留 `.env`、设置备份、本机历史及首次准备产生的缓存/运行时，并保留原安装清单以识别这一数据目录、允许主动重装。不存在通配删除用户材料的操作。应用运行中必须先正常退出，安装器不强杀文档进程。应用版本升级与自动更新签名不在本轮实现范围。

## 构建

固定工具、归档摘要和来源见 `release/WINDOWS_INSTALLER.json`。使用官方 Inno Setup 7.1.1 x64 和 Windows .NET Framework 编译器，构建工具不加入 Python 依赖。可将 Inno 以官方 `/PORTABLE=1 /CURRENTUSER` 方式准备到本机独立构建目录。所有临时诊断有期限并核对自有 Job 清理；不更改全局执行策略、权限或缓存。

准备固定 uv 归档及其原始 `LICENSE-APACHE`、`LICENSE-MIT`，按清单摘要核对。构建调用：

```powershell
./tools/build_windows_installer.ps1 -Payload <新的暂存目录> -Output <新的产物目录> -UvArchive <固定uv.zip> -LicenseRoot <原许可目录> -InnoSetup <固定Inno安装归档.exe>
```

构建只复制 `release/PUBLIC_FILES.json` 明确列出的公开源码，再加入经核对的 uv 与原许可。不从 `.env`、`.venv`、用户文档、运行时或开发日志收集文件。安装完整性清单摘要嵌入启动器；缺文件或摘要变化时拒绝启动。输出安装 EXE、SHA256SUMS 和构建回执，不自动推送或发布。

PowerShell 入口经 `uv run` 只读检查和暂存公开源码，在系统临时目录的独立子目录中准备经摘要核对的便携 Inno、编译启动器及安装器，再同卷移动最终二进制。这样不把开发构建目录的安全标签继承到分发 EXE。不会修改 ACL、完整性级别、全局策略或缓存权限；临时目录与输出不在同卷时拒绝构建。Python 模块入口仅提供 `--stage-only`，不能再直接产生发行 EXE。用户截图中的 Temp 错误 5 及原失败包摘要在本轮记录中保留，不能将原候选记为安装通过。

## 复用与发行边界

Inno Setup 按官方安装器许可使用，保留工具自带通知；项目只自行编写 `.iss` 配置，没有复制第三方界面实现。uv 固定 0.12.18、提交 `01cb90c1a4f88af09906cb60de9766d2add4a062`，只随包携带 `uv.exe` 及两份原许可。下载、校验与渲染器准备沿用项目原有固定来源，不增加包管理器或更换文档引擎。

本候选未代码签名；全新 Windows、普通用户权限、物理 IME/DPI、多屏、真实 Office 和真实模型计费分别验收。SuperDoc 编译资源不随安装器分发，用户准备时主动下载；捆绑二进制的完整对应源码与重建缺项仍开放。本机安装成功、导入检查或源码 ZIP 都不能代表以上验证已经通过。

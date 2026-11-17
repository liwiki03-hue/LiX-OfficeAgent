# LiX 离线安装内部预览

本渠道在构建机采集固定输入，在用户最终路径使用随包 Python、uv 与本地 wheel 装配环境。用户端准备命令包含 `offline`、`no-index`、`require-hashes`、`only-binary` 和禁止 Python 下载；缺项失败，不联网补资源。普通虚拟环境不会从开发机复制或在创建后搬移。

## 构建

1. 使用 Windows x64 / Python 3.12.14，通过 `uv run --isolated python -m tools.offline_installer collect --destination <新目录>` 采集生产闭包。锁文件和官方资源摘要是输入身份；可复用的文件必须重新校核。固定源包 `proxy-tools` 在隔离构建环境生成 wheel，工具约束和来源/产物摘要保存在 BUILT_WHEELS.json。
2. 在构建机提取 WINDOWS_INSTALLER.json 固定 MSI，生成 RENDERER.json 全文件清单。客户端不调用 MSI 行政提取。单独采集微软 WebView2 Standalone x64 与 .NET4.8 完整安装器，核验微软签名和 OFFLINE_INSTALLER.json 中的固定摘要；小型在线 Bootstrapper 不在本渠道范围内。
3. 使用 `uv run` 执行现有 tools/build_windows_installer.ps1，并提供 OfflineInputs、OfflineRenderer、OfflinePrerequisites、BuildId。BuildDirectory 可指定具有足够空间的独立构建目录，不改变用户缓存或系统权限。联网输入采集与后续只用本地输入组装分开。
4. SuperDoc 对应源码及重建尚未收口，默认构建拒绝首次含编译资产的公开分发。`InternalOnly` 只允许生成内部核验候选；这不解除发布阻断，也不是发行授权。

## 安装与准备

- Inno原生向导在.NET缺失时处理随包官方完整安装器，只提升该先决进程；3010/1641作为需要重启处理。已有可用WebView2复用，缺少时运行官方完整当前用户安装器，应用不提升。
- 程序安装完成后，启动器检查全部组件，再在既定用户数据目录的构建/安装身份专用最终位置建环境。环境部分完成不写就绪；重试只采用匹配本轮所有权标记的目录。不会清理其他环境。
- `LiX.exe --check` 只读核对文件和就绪状态。未准备为42，完整性/访问错误为1，成功为0；日志目录不可写不会让此体检创建目录或异常崩溃。
- 生产依赖版本、组件清单及环境文件库存分别核对。就绪记录绑定安装路径、锁文件、组件根清单及环境库存；目录迁移后重新装配，不沿用旧绝对路径。
- 云端AI、检索目录和更新仍需联网及配置。本地回环用于现有宿主通信。离线安装不能证明第三方组件没有后台外联尝试；这需要网络隔离实测。

## 当前边界

内部预览拒绝覆盖不同构建的已安装目录。完整跨版本升级、事务切换/回滚和匹配介质修复尚未交付，不能宣传这三项已经支持。卸载沿现有程序文件清单，默认保留配置、历史、草稿、用户副本和应用外部运行数据；不卸载系统共享WebView2/.NET。

WebView2最低能力、无组件/重启/取消、干净机断外网、中文路径、安装版三格式编辑、签名、防护软件及Microsoft Office必须分别验收。内部候选可运行、库可导入、文件存在均不能代替这些实测。验收入口为 [离线验收规划](../OFFLINE_INSTALLER_ACCEPTANCE_PLAN.md)，本轮证据见 deliverables/offline-installer-20261115。

## 来源

实际运行资源为固定 Astral CPython 3.12.14+20260929 install_only、uv0.12.18、锁定生产wheel、SuperDoc1.47.2、LibreOffice26.2.6.2及微软完整先决组件。准确URL、大小、hash和签名身份见 release/OFFLINE_INSTALLER.json、组件内部通知及构建输入清单。

设计参考 Tauri `06374a902a50d2bd8b8d85593623ad16ac32325a` 的 crates/tauri-bundler/src/bundle/windows/nsis/mod.rs 及其 Windows 辅助模块，采用“完整离线WebView2与按需安装”的生命周期思路，没有复制Rust代码或引入Tauri依赖。uv参数按本机0.12.18帮助及官方CLI核对；Python虚拟环境不可随意搬移、微软.NET需要管理员权限及重启处理按官方文档执行。第三方许可与SuperDoc对应源码缺口保持独立登记。

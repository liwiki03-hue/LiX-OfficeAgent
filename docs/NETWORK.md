# 网络与诊断

共用网络设置复用本机 `.env`：`OFFICE_NETWORK_MODE=direct|auto|manual`、`OFFICE_NETWORK_PROXY`，设置中的本机网络与诊断组沿用配置计划和确认。应用请求与首次准备默认直连；用户明确保存自动或手动模式后使用所选方式。代理凭据不进入日志。

自动支持Windows显式代理和标准代理环境变量，不解释PAC。localhost、回环及NO_PROXY匹配地址直连，不自动推断全部内网。手动代理支持HTTP(S)。不全局开启trust_env，不关闭TLS校验，不将API密钥跨主机转发；API地址或协议变化后需要匹配的新密钥。

共用解析已接入模型、模型列表、嵌入、研究模型、图片、文献、技能目录、前端、渲染器与更新下载。首次准备由 C# 启动器有界读取 `.env` 的网络字段，按同一模式为 uv 设置显式代理和 NO_PROXY，再准备 Python/锁定依赖；不导出密钥或配置正文。自动模式优先采用已有代理环境，未配置时读取 Windows 显式代理，不解释 PAC。四条受控 C# 策略检查通过；当前安装候选在依赖缓存访问拒绝处停止，不能记作完整联网准备通过。PAC、企业 HTTPS 代理和全部真实服务未验。

固定资源最多三次连接/超时/429/5xx尝试，总预算不重置，Retry-After最多8秒。连接/读取最长20秒，前端下载120秒，渲染器600秒；外层覆盖提取与清理。续传校验固定URL、版本、摘要、长度与ETag，最后核对SHA-256。取消不重放写入、模型请求或整个安装流程。

失败缓存保留。设置“检查下载缓存”先预览固定前端/渲染器缓存，用户选择后移到同目录唯一隔离文件，不删除、不联网、不清理其他文件。应用不能打开时可用已有Python环境运行宿主工具；第一条只预览，第二条明确隔离，之后在安装器选择重新准备。

```powershell
uv run python tools/repair_download_cache.py --kind frontend
uv run python tools/repair_download_cache.py --kind frontend --confirm-quarantine
```

运维日志 `.runtime/diagnostics/operations.jsonl` 采用5 MiB×3轮转、128项有界队列，与文档审计分开。只收运维状态、脱敏错误码和固定下载摘要；不接受正文、配置、请求头或响应。日志失败不阻塞文档事务。设置可预览后导出JSON，不捆绑.env、会话或草稿，不等于完整网络DLP。

## 依赖来源候选

主锁继续使用阿里云，不无声更新。可在项目内新目录生成官方PyPI或阿里云候选，再用现有uv复核；工具不覆盖当前pyproject、uv.lock、.env和.venv，不安装候选。

```powershell
uv run python tools/prepare_dependency_source.py --source pypi --target .runtime/dependency-source-pypi
uv lock --project .runtime/dependency-source-pypi --offline --python .venv/Scripts/python.exe
```

候选切换单一索引和发行文件地址，保留全部包版本及artifact哈希；uv解析后如版本/哈希改变则拒绝候选。本轮190条包记录通过离线uv复核，完整官方源下载/安装未验。嵌套uv若被拒绝访问缓存，保留错误，不修改权限或路径；本轮通过上述直接uv命令完成复核。

MSI1601、编译写入失败和 uv 缓存访问失败分别记录。相同固定 MSI 在标准隔离运行环境中实际提取成功，原 `.venv` 环境仍返回1601；Windows Installer 手动/停止状态不能单独确定根因。标准构建流水线已生成新候选，首次依赖准备仍受本机缓存访问拒绝阻断。启动器持续排空并有界保留错误流，只输出白名单错误分类；原始输出、配置和凭据不进入准备日志。本轮未修改服务、ACL、缓存位置、系统策略或安全软件。

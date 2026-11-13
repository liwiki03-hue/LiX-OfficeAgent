@echo off
rem Native Windows entry. No PowerShell, policy changes, or model/Office actions in checks.
setlocal
set "launch_exit=0"
set "launch_log="
set "old_codepage="
for /f "tokens=2 delims=:" %%A in ('chcp') do set "old_codepage=%%A"
chcp 65001 >nul
title Office Agent - Local Web
pushd "%~dp0"
if errorlevel 1 goto directory_failed
set "OFFICE_WEBVIEW_GUI=edgechromium"
set "PYTHONNET_RUNTIME=netfx"
set "PYTHONUTF8=1"
if not exist ".startup-logs\." mkdir ".startup-logs" >nul 2>nul
if not exist ".startup-logs\." goto log_failed
set "launch_log=%CD%\.startup-logs\launch-%RANDOM%-%RANDOM%.log"
if exist "%launch_log%" goto log_failed
>"%launch_log%" echo Office Agent startup diagnostics; no task/provider payloads.
if not exist "%launch_log%" goto log_failed
if errorlevel 1 goto log_failed
>>"%launch_log%" echo Started: %DATE% %TIME%
where uv >nul 2>nul
if errorlevel 1 goto uv_missing
if exist ".venv\Scripts\python.exe" goto readiness_check
echo 首次启动：正在用 uv sync 准备项目依赖。首次下载需要网络。
call uv sync
set "launch_exit=%ERRORLEVEL%"
>>"%launch_log%" echo sync_exit=%launch_exit%
if not "%launch_exit%"=="0" goto sync_failed
:readiness_check
echo 正在检查本地 Edge、pythonnet 和 .NET 运行环境……
rem This host's ordinary venv still fails CLR loading even at pythonnet 3.0.5.
rem Offline isolation uses cached packages, not a recurring network download.
echo 当前使用经实机验证的缓存隔离环境；启动检查和界面运行均禁止联网拉包。
call uv run --offline --isolated --with "pythonnet==3.0.5" python -m office_agent.webui.runtime >>"%launch_log%" 2>&1
set "launch_exit=%ERRORLEVEL%"
>>"%launch_log%" echo readiness_exit=%launch_exit%
if not "%launch_exit%"=="0" goto readiness_failed
echo 依赖检查通过，正在打开 Office Agent。此检查不代表窗口已验收。
echo 启动诊断日志：%launch_log%
rem Do not persist application stdout/stderr: it could contain task data.
call uv run --offline --isolated --with "pythonnet==3.0.5" python -m office_agent.app --ui web
set "launch_exit=%ERRORLEVEL%"
>>"%launch_log%" echo application_exit=%launch_exit%
if not "%launch_exit%"=="0" goto app_failed
goto finish
:uv_missing
set "launch_exit=2"
>>"%launch_log%" echo uv_missing
echo 未找到 uv。请先安装 uv 并加入 PATH，然后重新双击；本脚本不下载未知安装器。
goto failed
:sync_failed
echo 依赖准备失败。请检查网络和 uv sync 的错误，保留当前窗口信息；没有启动应用。
goto failed
:readiness_failed
echo 界面运行环境检查失败。请查看诊断日志，核对 WebView2、.NET Framework 以及锁定的 pythonnet 3.0.5。
echo 如果提示缓存缺包，请在项目目录联网执行 uv sync；随后重新双击。没有自动联网重试或切换后端。
echo 可以尝试原有 Start-OfficeAgent.cmd 隔离回退；本脚本不自动切换后端或反复启动。
goto failed
:app_failed
echo 应用未正常退出，错误码 %launch_exit%。请保留当前窗口中的错误并核对启动诊断。
goto failed
:log_failed
set "launch_exit=3"
echo 无法创建项目启动诊断日志，请检查项目目录权限；未启动应用。
goto failed
:directory_failed
set "launch_exit=4"
echo 无法进入项目目录；未启动应用。
goto finish_without_pop
:failed
if defined launch_log echo 启动诊断日志：%launch_log%
echo 未修改系统执行策略，没有终止其他程序。
if not "%OFFICE_AGENT_NO_PAUSE%"=="1" pause
:finish
popd
:finish_without_pop
if defined old_codepage chcp %old_codepage% >nul
endlocal & exit /b %launch_exit%

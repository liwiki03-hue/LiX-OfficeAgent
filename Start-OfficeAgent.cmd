@echo off
rem Native Windows entry: no PowerShell script or execution-policy change.
setlocal
pushd "%~dp0"
if errorlevel 1 exit /b 1
set "OFFICE_WEBVIEW_GUI=edgechromium"
set "PYTHONNET_RUNTIME=netfx"
call uv run --isolated --with pythonnet==3.0.5 python -m office_agent.app --ui web
set "launch_exit=%ERRORLEVEL%"
popd
endlocal & exit /b %launch_exit%

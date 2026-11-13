# Windows standard entry: preserve the verified isolated Edge/netfx combination.
# Does not change execution policy, project dependencies, or existing user windows.
$ErrorActionPreference = 'Stop'
$oldGui = $env:OFFICE_WEBVIEW_GUI
$oldRuntime = $env:PYTHONNET_RUNTIME
Push-Location -LiteralPath $PSScriptRoot
try {
    $env:OFFICE_WEBVIEW_GUI = 'edgechromium'
    $env:PYTHONNET_RUNTIME = 'netfx'
    & uv run --isolated --with 'pythonnet==3.0.5' python -m office_agent.app --ui web
    if ($LASTEXITCODE -ne 0) {
        throw "Office Agent startup failed (exit $LASTEXITCODE). No automatic retry or fallback."
    }
}
finally {
    $env:OFFICE_WEBVIEW_GUI = $oldGui
    $env:PYTHONNET_RUNTIME = $oldRuntime
    Pop-Location
}

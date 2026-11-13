"""Web startup configuration and read-only runtime diagnostics; never a model tool."""
from pathlib import Path
import os
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]


def configure_runtime() -> str:
    """Select the backend before importing webview/clr; preserve explicit environment."""
    from dotenv import load_dotenv
    load_dotenv(PROJECT_ROOT / ".env", override=False)
    gui = os.getenv("OFFICE_WEBVIEW_GUI", "edgechromium" if sys.platform == "win32" else "qt").strip()
    if not gui:
        raise ValueError("OFFICE_WEBVIEW_GUI must not be empty")
    if sys.platform == "win32" and gui == "edgechromium":
        os.environ.setdefault("PYTHONNET_RUNTIME", "netfx")
    return gui


def startup_help(gui: str) -> str:
    if sys.platform == "win32":
        return ("Windows 已验证回退：在项目目录运行 .\\Start-OfficeAgent.cmd；"
                "该原生入口不调用 PowerShell 脚本或更改执行策略，使用隔离 Edge + pythonnet 3.0.5 + netfx，不修改项目依赖。"
                "可选 .\\Start-OfficeAgent.ps1 仅在系统允许脚本时使用。"
                "若脚本受系统执行策略限制，不要绕过策略；手动设置 "
                "OFFICE_WEBVIEW_GUI=edgechromium、PYTHONNET_RUNTIME=netfx，再运行 "
                "uv run --isolated --with 'pythonnet==3.0.5' python -m office_agent.app --ui web。"
                "可选 Tk：uv run python -m office_agent.app --ui tk。")
    return "请核对所选后端依赖；可选 Tk：uv run python -m office_agent.app --ui tk。"


def diagnostic_report() -> dict:
    """Import-only check. No window, COM activation, user files, or model/API access."""
    import importlib.metadata as metadata
    report = {"python": sys.version.split()[0], "executable": sys.executable,
              "platform": sys.platform, "versions": {}, "ok": False}
    for name in ("pywebview", "pythonnet", "clr-loader", "PySide6", "qtpy"):
        try:
            report["versions"][name] = metadata.version(name)
        except metadata.PackageNotFoundError:
            report["versions"][name] = None
    try:
        report["gui"] = configure_runtime()
        report["requested_runtime"] = os.getenv("PYTHONNET_RUNTIME")
        if report["gui"] != "edgechromium" or sys.platform != "win32":
            report["note"] = "This diagnostic verifies only Windows Edge imports, not GUI visibility."
            return report
        import clr  # noqa: F401
        import pythonnet
        report["runtime"] = str(pythonnet.get_runtime_info())
        import webview.platforms.winforms  # noqa: F401
        report["ok"] = True
        report["note"] = "CLR and WinForms imports passed; GUI visibility is not verified."
    except Exception as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"
    return report


if __name__ == "__main__":
    import json
    import threading

    def timeout():
        print('{"ok": false, "error": "Import diagnostic exceeded 20 seconds"}', flush=True)
        os._exit(124)  # Only this diagnostic process; no child/GUI/Office is launched.

    watchdog = threading.Timer(20, timeout)
    watchdog.daemon = True
    watchdog.start()
    try:
        result = diagnostic_report()
        print(json.dumps(result, ensure_ascii=True, indent=2))
    finally:
        watchdog.cancel()
    raise SystemExit(0 if result["ok"] else 1)

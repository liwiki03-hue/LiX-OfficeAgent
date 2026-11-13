"""Explicitly opt-in offline native acceptance host; never a product entrypoint.

Run only with a fresh private acceptance root and OFFICE_AGENT_REAL_COM=0.
Replaces *only* the model factory with fixed read replies or opt-in write previews;
real pywebview, authentication, file chooser, document reader and UI remain intact.
"""
from __future__ import annotations

import os
from pathlib import Path

from openpyxl import Workbook, load_workbook


class OfflineReadModel:
    def __init__(self, path: Path):
        self.path = path
        self.calls = 0

    def bind_tools(self, schemas):
        assert "read_excel_region" in {item["function"]["name"] for item in schemas}
        return self

    def invoke(self, messages):
        self.calls += 1
        preview = "P0_PREVIEW_DISCARD_8000" in str(messages)
        if preview and self.calls == 1:
            # Unlock preview schemas through the normal write gate; still no write.
            return {"tool_calls": [{"name": "request_write_tools", "id": "offline-preview-gate", "args": {}}]}
        if preview and self.calls == 2:
            return {"tool_calls": [{"name": "plan_excel", "id": "offline-preview",
                                    "args": {"path": str(self.path), "sheet": "验收数据",
                                             "cell": "B2", "value": 8000}}]}
        if not preview and self.calls == 1:
            return {"tool_calls": [{"name": "read_excel_region", "id": "offline-read",
                                    "args": {"path": str(self.path), "sheet": "验收数据", "row": 2,
                                             "column": 1, "rows": 1, "columns": 2}}]}
        if preview:
            return {"content": "离线模型：只提交了合成单元格的修改预览；未批准也未执行写入。"}
        # Never pretend a rejected read succeeded. Actual read events remain host-owned.
        tool_text = " ".join(str(item.get("content", "")) for item in messages
                             if isinstance(item, dict) and item.get("role") == "tool")
        if "7319" in tool_text:
            return {"content": "离线模型：已从授权区域读取到 7319；这是合成数据。"}
        return {"content": "离线模型：读取未获授权，未取得单元格值。"}


def main() -> None:
    root_raw = os.environ.get("OFFICE_ACCEPTANCE_ROOT", "")
    root = Path(root_raw).resolve(strict=True) if root_raw else None
    if (os.getenv("OFFICE_ACCEPTANCE_SYNTHETIC_MODEL") != "1"
            or os.getenv("OFFICE_AGENT_REAL_COM") != "0"
            or root is None or not root.name.startswith("p0-private-")
            or not root.is_dir() or not (root / ".offline-acceptance").is_file()
            or not os.environ.get("OFFICE_WEBVIEW_DEBUG_PORT")):
        raise SystemExit("Refusing to run outside an explicit isolated offline acceptance root")
    # A high-integrity WebView with a localhost CDP listener is unsafe to expose.
    import win32api
    import win32con
    import win32security
    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_QUERY)
    sid, _ = win32security.GetTokenInformation(token, win32security.TokenIntegrityLevel)
    integrity = int(win32security.ConvertSidToStringSid(sid).rsplit("-", 1)[1])
    if integrity > 8192:  # Medium integrity is the maximum for this CDP-only probe.
        raise SystemExit("Refusing a high-integrity acceptance host with CDP enabled")
    source = root / "sample.xlsx"
    if source.is_symlink():
        raise SystemExit("Refusing to follow a workbook link")
    if source.exists():
        book = load_workbook(source, read_only=True)
        try:
            if (book.sheetnames != ["验收数据"]
                    or list(book.active.values) != [("项目", "金额"), ("专用合成行", 7319)]):
                raise SystemExit("Refusing an existing non-acceptance workbook")
        finally:
            book.close()
    else:
        book = Workbook()
        sheet = book.active
        sheet.title = "验收数据"
        sheet.append(["项目", "金额"])
        sheet.append(["专用合成行", 7319])
        book.save(source)
    from office_agent.webui.bridge import WebviewApp
    from office_agent.app import _launch_web
    WebviewApp._build_model = lambda self: OfflineReadModel(source)  # process-local, never product code
    _launch_web()


if __name__ == "__main__":
    main()

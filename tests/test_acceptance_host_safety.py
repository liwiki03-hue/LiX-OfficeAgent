"""The opt-in CDP probe must not expose an elevated Windows process."""
import pytest
import win32security

from tools.acceptance import probe_host


def test_probe_host_refuses_high_integrity_before_creating_a_workbook(tmp_path, monkeypatch):
    root = tmp_path / "p0-private-high"
    root.mkdir()
    (root / ".offline-acceptance").touch()
    monkeypatch.setenv("OFFICE_ACCEPTANCE_ROOT", str(root))
    monkeypatch.setenv("OFFICE_ACCEPTANCE_SYNTHETIC_MODEL", "1")
    monkeypatch.setenv("OFFICE_AGENT_REAL_COM", "0")
    monkeypatch.setenv("OFFICE_WEBVIEW_DEBUG_PORT", "9253")
    monkeypatch.setattr(win32security, "GetTokenInformation", lambda *args: (object(), 0))
    monkeypatch.setattr(win32security, "ConvertSidToStringSid", lambda sid: "S-1-16-12288")

    with pytest.raises(SystemExit, match="high-integrity acceptance host"):
        probe_host.main()
    assert not (root / "sample.xlsx").exists()


def test_offline_probe_plan_is_explicit_and_never_executes(tmp_path):
    model = probe_host.OfflineReadModel(tmp_path / "sample.xlsx")
    assert model.invoke([{"role": "user", "content": "普通读取"}])["tool_calls"][0]["name"] == "read_excel_region"
    model = probe_host.OfflineReadModel(tmp_path / "sample.xlsx")
    task = [{"role": "user", "content": "P0_PREVIEW_DISCARD_8000"}]
    gate = model.invoke(task)["tool_calls"][0]
    assert gate["name"] == "request_write_tools" and gate["args"] == {}
    plan = model.invoke(task)["tool_calls"][0]
    assert plan["name"] == "plan_excel"
    assert plan["args"]["cell"] == "B2" and plan["args"]["value"] == 8000
    assert "未执行写入" in model.invoke(task)["content"]
    assert not (tmp_path / "sample.xlsx").exists()

"""Memory UI is mounted in the account settings pane, not a model tool."""
from pathlib import Path

STATIC = Path(__file__).resolve().parents[1] / 'office_agent' / 'webui' / 'static'


def test_account_settings_memory_controls_and_untrusted_text_render():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    js = (STATIC / 'memory_preferences.js').read_text(encoding='utf-8')
    account = html.split('id="settings-pane-account"', 1)[1].split('</section>', 1)[0]
    assert 'id="memory-panel"' in account
    for identity in ('memory-load', 'memory-toggle', 'memory-save', 'memory-delete', 'memory-clear'):
        assert f'id="{identity}"' in account
    assert '<script src="memory_preferences.js"></script>' in html
    assert 'radio.type = "radio"' in js and 'label.append(radio, document.createTextNode(' in js
    assert 'innerHTML' not in js and 'eval(' not in js
    assert 'OfficeConfirm' in js
    assert 'window.confirm(' not in js
    assert 'OfficeAuth.wrapApi(window.pywebview?.api)' in js
    assert 'officeauth:unauthenticated' in js
    assert 'await service.confirm_memory_change(planId, approved === true)' in js


def test_manual_memory_is_grouped_and_rendered_as_text_only():
    html = (STATIC / 'index.html').read_text(encoding='utf-8')
    js = (STATIC / 'memory_preferences.js').read_text(encoding='utf-8')
    for identity in ('memory-manual-kind', 'memory-manual-text', 'memory-manual-save'):
        assert f'id="{identity}"' in html
    for kind in ('persona', 'work', 'personal', 'writing', 'skill'):
        assert f'value="{kind}"' in html
    assert 'change("set_manual", {kind, text}' in js
    assert 'manualText.value = ""' in js
    assert 'document.createTextNode(' in js and 'innerHTML' not in js
    assert '不自动创建技能' in html and '删除本地项不追溯清理' in html

"""Contracts for the compact, plain-language history sidebar."""
import json
import shutil
import subprocess
from html.parser import HTMLParser
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / "office_agent" / "webui" / "static"
APP = STATIC / "app.js"
BUNDLED_NODE = Path.home() / ".dsh/dsh-runtimes/dsh-primary-runtime/dependencies/node/bin/node.exe"
NODE = shutil.which("node") or (str(BUNDLED_NODE) if BUNDLED_NODE.is_file() else None)


def _render_history_body(source):
    start = source.index("function renderHistory()")
    end = source.index("function invalidateHistory(", start)
    return source[start:end]


def test_no_metadata_line_in_sidebar_cards():
    body = _render_history_body(APP.read_text(encoding="utf-8"))
    for phrase in ("時間：", "狀態：", "状态：", "可查看记录", "仅摘要", "条记录"):
        assert phrase not in body


def test_metadata_never_conditionally_appended():
    body = _render_history_body(APP.read_text(encoding="utf-8"))
    assert "available.length > 1" not in body
    assert "button.append(title)" in body
    assert "button.append(title, meta)" not in body


def test_title_uses_css_clamp_not_slice():
    source = APP.read_text(encoding="utf-8")
    body = _render_history_body(source)
    css = (STATIC / "theme.css").read_text(encoding="utf-8")
    assert "title.textContent" in body and ".slice(" not in body
    assert 'title.className = "history-item-title"' in body
    assert ".history-item-title" in css and "-webkit-line-clamp: 2" in css
    assert ".history-question" not in css and ".history-item-meta" not in css


def test_search_has_no_visible_long_label():
    source = APP.read_text(encoding="utf-8")
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    assert "筛选历史（指令" not in source and "筛选历史（指令" not in html
    assert 'id="history-search"' in html and 'aria-label="搜索任务指令、文件名或状态"' in html
    assert 'id="history-search-toggle"' in html and 'aria-expanded="false"' in html


@pytest.mark.skipif(NODE is None, reason="Node unavailable for the pure JavaScript translation probe")
def test_json_summary_translator(tmp_path):
    probe = tmp_path / "history_translation_probe.js"
    probe.write_text(r'''
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const source = fs.readFileSync(process.argv[2], 'utf8');
const start = source.indexOf('function translateHistoryText(');
const end = source.indexOf('function appendHistorySource(', start);
assert(start >= 0 && end > start, 'pure translator must stay independently loadable');
const context = {}; vm.createContext(context); vm.runInContext(source.slice(start, end), context);
const translated = context.translateHistoryText('{"classification":"unclear","skipped":false}');
assert(!translated.includes('classification'));
assert.strictEqual(context.translateHistoryText('正常中文指令'), '正常中文指令');
assert.strictEqual(context.translateHistoryText('指令 {"a":1} 后续'), '指令 （当时的内部信息已省略） 后续');
console.log(JSON.stringify({translated, plain: context.translateHistoryText('正常中文指令'), embedded: context.translateHistoryText('指令 {"a":1} 后续')}));
''', encoding="utf-8")
    result = subprocess.run([NODE, str(probe), str(APP)], capture_output=True,
                            text=True, encoding="utf-8", timeout=20)
    assert result.returncode == 0, result.stderr
    data = json.loads(result.stdout)
    assert data["translated"] == "这条记录当时没能确认需求"
    assert data["plain"] == "正常中文指令"
    assert data["embedded"] == "指令 （当时的内部信息已省略） 后续"


class _IdsAndParents(HTMLParser):
    def __init__(self, text):
        super().__init__()
        self.stack = []
        self.parents = {}
        self.feed(text)

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        identity = attrs.get("id")
        if identity:
            self.parents[identity] = tuple(self.stack)
        if tag not in {"input", "img", "meta", "link", "br", "hr"}:
            self.stack.append(identity or tag)

    def handle_endtag(self, tag):
        if self.stack:
            self.stack.pop()


def test_history_region_has_own_group():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    dom = _IdsAndParents(html)
    for identity in ("history-search-panel", "history-list"):
        assert "history-view" in dom.parents[identity]
    source = APP.read_text(encoding="utf-8")
    assert 'historyRegion.id = "history-region"' in source
    assert 'nav.insertBefore(historyRegion' in source
    assert 'historyRegion.append(el("history-view"))' in source


def test_history_card_text_is_translated_in_detail_and_transcript():
    source = APP.read_text(encoding="utf-8")
    assert "translateHistoryText(String(session.question" in _render_history_body(source)
    assert 'appendHistoryText(description, value' in source
    assert 'appendHistoryText(body, copy' in source
    assert 'details.className = "history-source"' in source
    assert 'summary.textContent = "查看当时的内部信息"' in source

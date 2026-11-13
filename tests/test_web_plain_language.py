"""全面去技术化契约：图表与生图都必须走对话驱动，界面上不再出现技术术语。

这些断言是「用户可见文案」的回归闸门：任何把 Mermaid / SVG / 回滚 / 检查点 /
授权 之类的词重新写回界面的改动都会在这里失败。
"""
import re
from pathlib import Path

import pytest

STATIC = Path(__file__).resolve().parents[1] / "office_agent" / "webui" / "static"

# 只检查用户能看见的文案所在文件；vendor/ 与测试无关。
UI_FILES = ("index.html", "app.js", "imagegen.js", "diagram_export.js", "diagram_preview.js",
            "settings.js", "skills.js", "rag.js")

# 这些词属于「实现细节」，不应出现在用户可见的字符串里。
FORBIDDEN_IN_UI = ("Mermaid", "SVG 副本", "ASCII", "embedding 通道", "token 预算",
                   "检查点", "指纹", "独立确认", "回滚本次", "API Key", "SHA256")


def _user_visible_strings(source: str) -> list[str]:
    """抽出双引号 / 单引号字符串字面量与 HTML 文本，尽量排除注释与标识符。"""
    without_block_comments = re.sub(r"/\*.*?\*/", "", source, flags=re.S)
    without_line_comments = re.sub(r"(?m)^\s*//.*$", "", without_block_comments)
    quoted = re.findall(r'"([^"\n\\]*(?:\\.[^"\n\\]*)*)"|\'([^\'\n\\]*(?:\\.[^\'\n\\]*)*)\'',
                        without_line_comments)
    return [a or b for a, b in quoted]


@pytest.mark.parametrize("name", UI_FILES)
def test_no_implementation_jargon_in_user_facing_strings(name):
    source = (STATIC / name).read_text(encoding="utf-8")
    hits = []
    for text in _user_visible_strings(source):
        # CJK 文案才可能是给用户看的；纯 ASCII 的通常是选择器 / 类名 / 协议字段。
        if not re.search(r"[\u4e00-\u9fff]", text):
            continue
        for term in FORBIDDEN_IN_UI:
            if term in text:
                hits.append((term, text[:60]))
    assert not hits, f"{name} 里仍有技术术语出现在用户文案中：{hits}"


def test_diagram_is_conversation_driven_with_no_form_controls():
    html = (STATIC / "index.html").read_text(encoding="utf-8")
    # 旧表单控件必须彻底消失，只保留说明入口。
    for control in ("diagram-card", "diagram-kind", "diagram-goal", "diagram-generate",
                    "diagram-cancel", "diagram-source", "diagram-image", "diagram-save",
                    "diagram-rollback", "diagram-status", "diagram-export-status",
                    "diagram-render", "diagram-copy", "imagegen-card", "imagegen-body",
                    "imagegen-state"):
        assert control not in html, f"表单控件 {control} 应已移除"
    # 触发方式必须写在界面上，否则用户不知道能直接说。
    assert "画一个" in html and "生成一张" in html


def test_imagegen_renders_inside_the_conversation_flow():
    app = (STATIC / "app.js").read_text(encoding="utf-8")
    js = (STATIC / "imagegen.js").read_text(encoding="utf-8")
    # app.js 必须把消息挂进对话流，而不是留给侧栏工具面板。
    assert "function mountImageMessage(" in app
    assert "mountImageMessage);" in app
    assert 'window.OfficeImageGen?.handleEvent(ev,' in app
    # 模块必须接受调用方提供的 mount，且不自己找 #imagegen-body。
    assert "function handleEvent(event, call, mount)" in js
    assert 'getElementById("imagegen-body")' not in js
    assert 'getElementById("imagegen-card")' not in js
    assert "mount(card)" in js
    # 结果仍是「渲染不提交」：必须有独立确认对话框且默认焦点在取消。
    assert "showModal" in js and "cancel.focus()" in js


def test_inline_messages_keep_separate_confirmations():
    js = (STATIC / "imagegen.js").read_text(encoding="utf-8")
    for operation in ("confirm_image_generation", "plan_image_save", "confirm_image_save",
                      "plan_image_rollback", "confirm_image_rollback"):
        assert operation in js, operation
    # 退出登录后旧消息必须立刻失去权限。
    assert "officeauth:unauthenticated" in js
    assert ".imagegen-item .imagegen-actions button" in js


def test_inline_messages_have_theme_scoped_styles():
    css = (STATIC / "theme.css").read_text(encoding="utf-8")
    for selector in (".diagram-message", ".diagram-figure", ".imagegen-item",
                     ".imagegen-title", ".imagegen-actions", ".imagegen-status"):
        assert selector in css, selector

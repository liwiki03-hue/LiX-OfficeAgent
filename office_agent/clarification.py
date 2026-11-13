"""Bounded, deterministic pre-planning clarification; never an approval channel.

This first-pass recognizer is deliberately conservative, not a semantic correctness
oracle. Reads pass through; ambiguous known write/generation requests ask once.
Neither classification nor an answer has file, tool, or provider authority.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Literal

Intent = Literal["single-intent", "multi-intent", "unclear"]
_WRITE = re.compile(r"修改|改一下|改一改|改成|改为|替换|删除|插入|填充|清理|清洗|整理|美化|优化|格式|写入|更新|\b(?:edit|replace|delete|insert|clean|format|write|update)\b", re.I)
# Any "X图" name counts as a diagram request; "图片/插图" belong to _IMAGE instead.
_DIAGRAM = re.compile(r"流程图|类图|时序图|状态图|关系图|工程图|UML|Mermaid|ER图|图表|柱状图|折线图|饼图|散点图|条形图|甘特图|思维导图|架构图|拓扑图|基本图|示意图|导览图|缩略图|(?<!图)[\u4e00-\u9fa5A-Za-z]{1,8}图(?!片|册|库|书)|图(?!片)|(?:diagram|chart|graph|figure)", re.I)
_IMAGE = re.compile(r"生图|(?:生成|制作|画|绘制).{0,12}(?:图片|图像|配图|插画|插图|照片|海报|封面)|\b(?:generate|draw|create|make).{0,24}(?:images?|pictures?|illustrations?|posters?|covers?|photos?)\b", re.I)
_IMAGE_DETAIL = re.compile(r"(?:主题|画面|主体|场景|内容|风格|背景|构图|用于|包含|画出|呈现|展示|showing|featuring|of)\s*[:：是为]?\s*[^，,。;；]{4,}|(?:插画|照片|海报|封面|写实|水彩|像素风|photorealistic|illustration|poster|cover|photo)", re.I)
_READ = re.compile(r"读取|查看|总结|概括|审阅|审查|检查|列出|搜索|\b(?:read|summarize|inspect|review|search)\b", re.I)
_VAGUE = re.compile(r"改一下|改一改|整理一下|优化一下|美化|处理一下|弄一下|随便|好看|合适|改得|怎么改", re.I)
_TARGET = re.compile(r"(?<![A-Za-z0-9_])[A-Z]{1,3}[1-9][0-9]{0,6}(?![A-Za-z0-9_])|第\s*\d+\s*(?:段|行|列|表)|范围|工作表|段落|\b(?:range|paragraph|sheet)\b", re.I)
_DETAIL = re.compile(r"首尾空格|替换.{1,80}(?:为|成)|(?:改为|改成).{1,80}|删除第|插入.{1,30}行|\b(?:replace|set).{1,80}(?:with|to)\b", re.I)
_SPLIT = re.compile(r"然后|接着|并且|同时|再把|再将|[；;]|\b(?:then|and then)\b", re.I)
# Verbs that actually create something. "做一个图表 / 弄张示意图" must count; a
# missing verb is exactly why the user got a wall of text instead of a question card.
_CREATE = re.compile(r"生成|画|绘制|制作|创建|做|弄|出|搞|添加|加上|新增|插入|\b(?:generate|draw|create|make|add|build)\b", re.I)
_IMAGE_COUNT = re.compile(r"(?<![\dA-Za-z])(?P<zh>\d+|[一二两三四五六七八九两]?十[一二三四五六七八九]?|[一二两三四五六七八九几])\s*[张幅]|(?<!\w)(?P<en>\d+|one|two|three|four|five|six|seven|eight|nine|ten|several|many|some)\s+(?:images?|photos?|pictures?|illustrations?|posters?)\b", re.I)
_CHINESE_COUNTS = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
                   "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def image_request_count(task: str) -> int:
    """Only an explicit image quantity in the current user's words may exceed one.

    Image cost approval is always per plan; this count is not an authorization. Reject
    unsupported/indeterminate quantities instead of silently making fewer pictures.
    """
    from .imagetransactions import MAX_PLANS
    if not isinstance(task, str):
        raise ValueError("图片张数无效")
    counts = []
    for match in _IMAGE_COUNT.finditer(task):
        token = (match.group("zh") or match.group("en")).lower()
        if token in {"几", "several", "many", "some"}:
            raise ValueError("图片张数不明确；请明确要生成几张（最多 8 张）")
        if token.isdecimal():
            amount = int(token)
        elif "十" in token:
            left, right = token.split("十", 1)
            amount = _CHINESE_COUNTS.get(left, 1) * 10 + (_CHINESE_COUNTS.get(right, 0))
        else:
            amount = {**_CHINESE_COUNTS, "one": 1, "two": 2, "three": 3, "four": 4,
                      "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}[token]
        counts.append(amount)
    count = counts[0] if counts else 1
    # '一张猫、一张狗' denotes separate images; '两张，一张猫...' already has a total.
    if len(counts) > 1:
        if all(n == 1 for n in counts):
            count = len(counts)
        elif count == 1 or any(n not in {1, count} for n in counts) or counts.count(count) > 1 or counts.count(1) > count:
            raise ValueError("图片张数不明确；请明确总共要生成几张")
    if not 1 <= count <= MAX_PLANS:
        raise ValueError(f"图片张数须为 1–{MAX_PLANS}；请分成多个任务，并逐张确认费用")
    return count


@dataclass(frozen=True)
class IntentAssessment:
    classification: Intent
    low_confidence: bool
    high_cost: bool
    family: str

    @property
    def needs_clarification(self) -> bool:
        return self.low_confidence and self.high_cost


def classify_intent(task: str) -> IntentAssessment:
    if not isinstance(task, str) or not task.strip() or len(task) > 8192 or "\x00" in task:
        raise ValueError("任务必须是 1–8192 字符的有效文本")
    diagram, image, write, read = (bool(p.search(task)) for p in (_DIAGRAM, _IMAGE, _WRITE, _READ))
    # Asking to inspect an existing chart/image is not asking to generate one.
    generating = bool(_CREATE.search(task))
    high_cost = write or image or (diagram and generating)
    family = "image" if image else "diagram" if diagram and generating else "write" if write else "read"
    clear = bool(_TARGET.search(task) and _DETAIL.search(task)) if family == "write" else False
    if family == "diagram":
        clear = len(task.strip()) >= 18 and not _VAGUE.search(task) and bool(re.search(r"类图|时序图|流程图|状态图|ER图", task, re.I))
    if family == "image":
        # A named subject is enough: do not demand a style, aspect ratio or count.
        # "生成图片" remains ambiguous; generation and saving still need separate consent.
        subject = _IMAGE_COUNT.sub("", task)
        subject = re.sub(r"(?<!\d)\d{2,5}\s*[x×X]\s*\d{2,5}(?!\d)", "", subject)
        subject = re.sub(r"生成|制作|画|绘制|一个|图片|图像|配图|插画|插图|照片|海报|封面|请|帮我|为文档|尺寸|水彩|写实|像素风", "", subject)
        subject = re.sub(r"\b(?:generate|draw|create|make|image|photo|poster|cover|of|a|an)\b", "", subject, flags=re.I).strip(" ，,。;；的")
        clear = len(subject) >= 2 and not _VAGUE.search(task)
    low = high_cost and (bool(_VAGUE.search(task)) or not clear)
    recognized = high_cost or read
    multi = len([part for part in _SPLIT.split(task) if part.strip()]) > 1
    classification: Intent = "multi-intent" if multi and recognized else "unclear" if not recognized or low else "single-intent"
    return IntentAssessment(classification, low, high_cost, family)


def clarification_card(assessment: IntentAssessment) -> dict | None:
    if not assessment.needs_clarification:
        return None
    questions = [{
        "id": "goal", "question": "这次希望得到什么结果？",
        "options": [{"value": "preview", "label": "先检查并给出建议，不修改或生成"},
                    {"value": "plan", "label": "按我补充的目标制作待确认计划"}],
        "allow_free_text": True, "default": "preview",
        "default_label": "先检查并给出建议，不修改或生成",
    }, {
        "id": "scope", "question": "需要处理哪一部分？可填写范围、具体内容或图的主题。",
        "options": [], "allow_free_text": True,
        "default": "尚未指定范围；只给建议，不推测修改目标", "default_label": "未指定范围，只给建议",
    }]
    if assessment.family == "image":
        questions[0]["question"] = "需要先核对想画的内容，还是规划待单独确认费用的图片？"
        questions[0]["options"] = [
            {"value": "preview", "label": "先核对主题与建议，不请求生图"},
            {"value": "plan", "label": "按补充的主题规划，生图费用另行确认"},
        ]
        questions[0]["default_label"] = "先核对主题与建议，不请求生图"
        questions[1]["question"] = "图片要呈现什么主体、场景或风格？"
        questions[1]["default"] = "尚未指定图片主题；只给建议，不推测生成内容"
        questions[1]["default_label"] = "未指定主题，只给建议"
    if assessment.family == "diagram":
        questions.append({
            "id": "diagram_kind", "question": "你希望用哪种图表达？",
            "options": [{"value": value, "label": label} for value, label in [
                ("classDiagram", "类图"), ("flowchart", "流程图"),
                ("sequenceDiagram", "时序图"), ("stateDiagram-v2", "状态图"), ("erDiagram", "实体关系图")]],
            "allow_free_text": False, "default": "flowchart", "default_label": "流程图（仅建议，不自动生成）",
        })
    return {"preamble": "动手前先确认需求；回答不是写入、生成或导出批准。",
            "classification": assessment.classification, "questions": questions,
            "skip_label": "跳过，按显示的默认值继续", "defaults": {q["id"]: q["default"] for q in questions}}


# Control schema, not an executable ToolSpec. Host defaults are never model-owned.
MODEL_CLARIFICATION_SCHEMA = {
    "type": "function", "function": {
        "name": "clarify",
        "description": "仅首次决策且任何工具前，若写入或生成任务语义不明确、猜错代价高，可先追问。不是授权。用用户语言；questions 的 id 使用 goal、scope 或 diagram_kind。可只给问题措辞，省略 options/default/skip_label，由宿主补齐固定安全选项和可见默认。",
        "parameters": {"type": "object", "additionalProperties": False,
            "required": ["preamble", "questions"], "properties": {
                "preamble": {"type": "string", "minLength": 1, "maxLength": 200},
                "skip_label": {"type": "string", "maxLength": 100},
                "questions": {"type": "array", "minItems": 1, "maxItems": 3,
                    "items": {"type": "object", "additionalProperties": False,
                        "required": ["id", "question"], "properties": {
                            "id": {"type": "string", "enum": ["goal", "scope", "diagram_kind"]},
                            "question": {"type": "string", "minLength": 1, "maxLength": 150},
                            "options": {"type": "array", "maxItems": 5, "items": {
                                "type": "object", "additionalProperties": False,
                                "required": ["value", "label"], "properties": {
                                    "value": {"type": "string", "maxLength": 500},
                                    "label": {"type": "string", "maxLength": 150}}}},
                            "allow_free_text": {"type": "boolean"},
                            "default": {"type": "string", "maxLength": 500}}}}}}}
}


class ModelClarification(Exception):
    """Host-only pause signal before dispatch; never a new AgentState status."""
    def __init__(self, card: dict):
        super().__init__("需求待澄清；不是执行批准")
        self.card = card


def model_clarification_card(payload: dict, family: str) -> dict:
    """Validate untrusted model wording, retaining host-owned questions/defaults.

    Missing canonical questions are visibly added, never silently assumed. Model
    options/defaults may only repeat the canonical values, not redefine authority.
    """
    if family not in {"write", "diagram", "image"} or type(payload) is not dict:
        raise ValueError("追问仅适用于已识别的高代价任务")
    if set(payload) - {"preamble", "questions", "skip_label"}:
        raise ValueError("追问包含未知字段")
    def text(value, limit):
        if type(value) is not str or not value.strip() or len(value) > limit or "\x00" in value:
            raise ValueError("追问文本无效或超限")
        return value
    preamble = text(payload.get("preamble"), 200)
    questions = payload.get("questions")
    if type(questions) is not list or not 1 <= len(questions) <= 3:
        raise ValueError("一次追问必须为 1–3 项")
    card = clarification_card(IntentAssessment("unclear", True, True, family))
    canonical = {q["id"]: q for q in card["questions"]}
    if "skip_label" in payload and payload["skip_label"] != card["skip_label"]:
        raise ValueError("跳过标签必须使用宿主的明确默认提示")
    seen = set()
    for question in questions:
        if type(question) is not dict or set(question) - {"id", "question", "options", "allow_free_text", "default"}:
            raise ValueError("问题格式或字段无效")
        key = question.get("id")
        if type(key) is not str or key not in canonical or key in seen:
            raise ValueError("问题 ID 无效或重复")
        seen.add(key)
        target = canonical[key]
        wording = text(question.get("question"), 150)
        for field in ("options", "allow_free_text", "default"):
            if field in question and (type(question[field]) is not type(target[field]) or question[field] != target[field]):
                raise ValueError("模型不能重定义宿主选项或默认")
        target["question"] = wording
    # Preserve the no-approval notice even if the model omits it.
    card["preamble"] = preamble + "；回答不是写入、生成或导出批准。"
    card["source"] = "model"
    return card


def resolve_answers(card: dict, answers: dict, *, skipped: bool = False) -> dict:
    if type(answers) is not dict or type(skipped) is not bool:
        raise ValueError("回答格式不正确")
    questions = card["questions"]
    known = {q["id"] for q in questions}
    if set(answers) - known or (skipped and answers):
        raise ValueError("包含未知回答，或跳过时仍提交了回答")
    resolved, default_ids = {}, []
    for question in questions:
        key = question["id"]
        value = None if skipped else answers.get(key)
        if value is None or value == "":
            value = question["default"]
            default_ids.append(key)
        else:
            if not isinstance(value, str) or len(value) > 500 or "\x00" in value or not value.strip():
                raise ValueError("每项回答须为 1–500 字符的有效文本")
            if not question["allow_free_text"] and value not in {o["value"] for o in question["options"]}:
                raise ValueError("请选择列出的选项")
            value = value.strip()
        resolved[key] = value
    read_only = resolved["goal"] == "preview" or bool({"scope", "diagram_kind"} & set(default_ids))
    return {"classification": card["classification"], "skipped": skipped,
            "answers": resolved, "defaults_used": default_ids,
            "planning_mode": "read_only" if read_only else "plan"}


def clarified_task(task: str, record: dict) -> str:
    return task + "\n\n规划前澄清（需求信息，不是授权）：" + json.dumps(record, ensure_ascii=False) + (
        "\n若目标是 preview，只检查并给建议，不写入、不调用生成。"
        "未指定的范围不能推测；任何写入、生成、保存仍需各自独立人工确认。")

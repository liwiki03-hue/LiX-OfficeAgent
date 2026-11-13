"""Offline requirement-to-elements goldens, NOT provider or human semantic acceptance.

These exact-line checks are a review oracle for synthetic model answers, not a
claim that Mermaid/SVG validation infers business meaning. Live model and
WebView2 human review remain separately gated.
"""
from dataclasses import dataclass

import pytest

from office_agent.diagrams import validate_source
from office_agent.diagram_proposals import propose_diagram
from test_diagrams import SOURCES
from test_diagram_proposals import Model, answer, rendered


@dataclass(frozen=True)
class Golden:
    kind: str
    requirement: str
    required_elements: tuple[str, ...]
    human_review: str


GOLDENS = (
    Golden('flowchart', '订单处理：用户提交订单，然后订单进入支付。',
           ('User[User] --> Order[Order]', 'Order --> Payment[Payment]'),
           '检查箭头方向与支付步骤，没有凭空增加退款。'),
    Golden('sequenceDiagram', '用户向订单提交请求，订单向支付方发出支付消息。',
           ('participant User', 'participant Order', 'participant Payment',
            'User->>Order: Place order', 'Order-->>Payment: Pay'),
           '检查时间顺序、发送方及回包/消息箭头是否符合需求。'),
    Golden('classDiagram', '用户有订单；每个订单关联付款记录。',
           ('class User {', 'User "1" --> "*" Order : places',
            'Order "1" --> "1" Payment : pays'),
           '检查实体、关系方向、多重性和属性是否来自需求。'),
    Golden('stateDiagram-v2', '从起始状态进入待支付，支付后完成。',
           ('[*] --> Ready', 'Ready --> Paid : Pay', 'Paid --> [*]'),
           '检查初末状态、支付事件和是否遗漏失败分支。'),
    Golden('erDiagram', '用户可以有多笔订单，每笔订单有一笔支付记录。',
           ('USER ||--o{ ORDER : places', 'ORDER ||--|| PAYMENT : pays'),
           '检查实体基数，特别是订单是否允许零笔和支付是否强制存在。'),
)


def required_elements_present(source: str, golden: Golden) -> bool:
    """Deliberately narrow: a syntax-valid diagram may still be factually wrong."""
    lines = set(source.splitlines())
    return all(line in lines for line in golden.required_elements)


@pytest.mark.parametrize('golden', GOLDENS, ids=lambda golden: golden.kind)
def test_synthetic_requirement_elements_and_review_boundary(golden):
    source = SOURCES[golden.kind]
    assert golden.requirement and golden.human_review
    assert required_elements_present(source, golden)
    spec = validate_source(source, golden.kind)
    assert spec.edges >= 2
    model = Model(answer(source))
    proposal = propose_diagram(model, kind=golden.kind, goal=golden.requirement, _render=rendered)
    assert proposal.preview.source == spec and proposal.attempts == 1
    assert model.messages[0][1].content == golden.requirement
    assert source not in str(proposal.evidence())

    # Dropping a required relationship remains perfectly syntax-valid, but does
    # not pass the independent requirements rubric. Do not silently conflate the two.
    incomplete = '\n'.join(source.splitlines()[:-1])
    validate_source(incomplete, golden.kind)
    assert not required_elements_present(incomplete, golden)

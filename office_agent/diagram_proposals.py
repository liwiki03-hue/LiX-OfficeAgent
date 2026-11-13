"""Read-only, bounded diagram proposal/repair controller.

No tools, dispatch, exports, credentials, file reads or automatic .env loading.
An injected configured chat model owns its provider timeout; invoke() is not
preempted here. Each render has its own supervised deadline. Production callers
must bind the cancellation guard to the authenticated session/task epoch.
"""
from dataclasses import dataclass
from typing import Callable

from langchain_core.messages import AIMessage, HumanMessage, SystemMessage

from .diagram_process import OwnedRun
from .diagram_service import RenderedDiagram, render_diagram
from .diagrams import (MAX_REPAIRS, DiagramError, diagram_prompt, repair_prompt,
                       validate_source, validate_svg)

_REPAIRABLE = frozenset({'unsupported_source', 'renderer_failed', 'unsafe_svg', 'render_timeout'})
_TERMINAL = frozenset({'render_cancelled', 'renderer_cleanup_failed', 'diagram_preview_mismatch'})
_GOAL = '根据独立用户需求生成只读图表预览'


@dataclass(frozen=True)
class DiagramProposal:
    preview: RenderedDiagram
    attempts: int
    repair_codes: tuple[str, ...]

    def evidence(self):
        # Never add prompt, prior model output, source or SVG to model evidence.
        return {**self.preview.evidence(), 'attempts': self.attempts,
                'repair_codes': list(self.repair_codes)}


def _source(response):
    if (not isinstance(response, AIMessage) or type(response.content) is not str
            or response.tool_calls or response.invalid_tool_calls
            or not (response.additional_kwargs.get('tool_calls') is None
                    or type(response.additional_kwargs.get('tool_calls')) is list
                    and not response.additional_kwargs['tool_calls'])
            or response.additional_kwargs.get('function_call') is not None
            or response.additional_kwargs.get('refusal') not in (None, '')
            or any(response.response_metadata.get(key) not in (None, 'stop', 'end_turn', 'eos_token')
                   for key in ('finish_reason', 'done_reason'))):
        raise DiagramError('unsupported_source')
    return response.content


def _preview(result, spec):
    if type(result) is not RenderedDiagram or type(result.execution) is not OwnedRun:
        raise DiagramError('diagram_preview_mismatch')
    if result.execution.empty_confirmed is not True:
        raise DiagramError('renderer_cleanup_failed')
    if (result.source != spec or result.execution.code != 'process_completed'
            or type(result.execution.exit_code) is not int or result.execution.exit_code != 0):
        raise DiagramError('diagram_preview_mismatch')
    # Do not infer trust from a renderer adapter's return type or claimed metadata.
    return RenderedDiagram(spec, validate_svg(result.svg), result.execution)


def propose_diagram(model, *, kind: str, goal: str, timeout_ms: int = 8000,
                    cancelled: Callable[[], bool] | None = None,
                    _render: Callable | None = None) -> DiagramProposal:
    """One generation plus at most TWO repairs, yielding only a passive preview.

    Error feedback uses fixed codes and a fresh two-message request, never prior
    source/SVG/exceptions. Answers cannot execute tools or approve export. Model
    calls are synchronous; cancellation discards late returns but cannot stop a
    provider call in flight. The host, not the model, injects model and renderer.
    """
    diagram_prompt(kind, goal)  # Validate the complete user request before calls.
    invoke = getattr(model, 'invoke', None)
    if (not callable(invoke) or type(timeout_ms) is not int or not 100 <= timeout_ms <= 8000
            or cancelled is not None and not callable(cancelled)
            or _render is not None and not callable(_render)):
        raise DiagramError('unsupported_source')
    render = render_diagram if _render is None else _render
    stopped = False
    def is_cancelled():
        nonlocal stopped
        if not stopped and cancelled is not None:
            try:
                stopped = bool(cancelled())
            except Exception:
                stopped = True  # Unreliable ownership guard never permits a result.
        return stopped
    def checkpoint():
        if is_cancelled():
            raise DiagramError('render_cancelled')
    failures = []
    for index in range(MAX_REPAIRS + 1):
        checkpoint()
        policy = (diagram_prompt(kind, _GOAL) if not failures
                  else repair_prompt(kind, _GOAL, failures[-1], index))
        messages = [SystemMessage(content=policy), HumanMessage(content=goal.strip())]
        try:
            response = invoke(messages)
        except Exception:
            checkpoint()
            raise DiagramError('diagram_model_failed') from None
        checkpoint()
        try:
            spec = validate_source(_source(response), kind)
            checkpoint()
            result = render(spec.source, expected_kind=kind, timeout_ms=timeout_ms,
                            cancelled=is_cancelled)
            preview = _preview(result, spec)
        except DiagramError as exc:
            # Cleanup uncertainty wins over cancellation and is never retried.
            if exc.code == 'renderer_cleanup_failed':
                raise DiagramError(exc.code) from None
            checkpoint()
            if exc.code in _TERMINAL:
                raise DiagramError(exc.code) from None
            if exc.code not in _REPAIRABLE:
                raise DiagramError('renderer_failed') from None
            failures.append(exc.code)
            if index == MAX_REPAIRS:
                raise DiagramError('diagram_repair_exhausted') from None
            continue
        except Exception:
            # An unexpected adapter failure gives no cleanup guarantee: no retry.
            raise DiagramError('renderer_failed') from None
        checkpoint()
        return DiagramProposal(preview, index + 1, tuple(failures))
    raise DiagramError('diagram_repair_exhausted')  # Defensive, never optimistic.

from collections import Counter
from copy import deepcopy
import json

from openpyxl import Workbook, load_workbook
import pytest

from office_agent.agent import OfficeAgent
from office_agent.app import _WRITE_GATE, build_dispatcher, build_tools
from office_agent.instruction_runtime import StaleInstruction
from office_agent.writes import Editor, digest


SHEET_A = '延期 A'
SHEET_B = '延期 B'
SHEET_C = '延期 C'
KNOWN_GOAL = '比较延期 A和延期 B的B2延期天数，保留单位并给出差值。'


class ReplayModel:
    def __init__(self, replies):
        self.replies = iter(replies)
        self.messages = []

    def bind_tools(self, schemas):
        return self

    def invoke(self, messages):
        self.messages.append(deepcopy(messages))
        return next(self.replies)


@pytest.fixture
def workbook(tmp_path):
    path = tmp_path / '独立范围.xlsx'
    book = Workbook()
    for index, (name, value) in enumerate(((SHEET_A, 12), (SHEET_B, 7), (SHEET_C, 3))):
        sheet = book.active if index == 0 else book.create_sheet()
        sheet.title = name
        sheet.append(['项目', '延期天数'])
        sheet.append(['样本', value])
        sheet.append(['相同值', value])
    book.save(path)
    book.close()
    return path


def region(path, sheet, identifier, *, row=2, column=2):
    return {'name': 'read_excel_region', 'id': identifier, 'args': {
        'path': str(path), 'sheet': sheet, 'row': row, 'column': column, 'rows': 1, 'columns': 1}}


def response(*calls):
    return {'content': '', 'tool_calls': list(calls)}


def make_agent(workbook, replies, *, confirm=None, hook=None, boundary=None, stopped=None,
               persist=False, decision_limit=10, tools=None, **options):
    model = ReplayModel(replies)
    editor = Editor(workbook.parent)
    actual = build_dispatcher(editor, workbook, expected_version=digest(workbook))
    calls = []
    approvals = []
    def dispatch(name, arguments):
        calls.append((name, deepcopy(arguments)))
        return hook(name, arguments, actual) if hook else actual(name, arguments)
    def approve(plan):
        approvals.append(plan)
        return confirm(plan) if confirm else True
    catalogue = build_tools((workbook,))
    selected = tools if tools is not None else [spec for spec in catalogue if spec.name in {
        'read_excel_region', 'read_excel_structure', 'plan_excel'}]
    agent = OfficeAgent(model, dispatch, selected, workbook.parent / 'batch-state', persist=persist,
                        confirm=approve, runtime_boundary=boundary, stop_requested=stopped,
                        context_tokens=1000000, reply_tokens=512, token_estimator=lambda text: 10,
                        decision_limit=decision_limit, **options)
    return agent, model, calls, approvals


def initial(agent, workbook, goal=KNOWN_GOAL):
    return agent.new_state(goal, file_versions={str(workbook.resolve()): digest(workbook)})


def assert_paired(state):
    requested = Counter(call['id'] for item in state.history if item['role'] == 'assistant'
                        for call in item.get('tool_calls', []))
    receipts = Counter(item['tool_call_id'] for item in state.history if item['role'] == 'tool')
    assert requested == receipts
    assert all(count == 1 for count in requested.values())


def test_two_real_regions_take_two_model_requests_and_preserve_call_ids(workbook):
    before = digest(workbook)
    agent, model, calls, approvals = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'first-a'), region(workbook, SHEET_B, 'first-b')),
        {'content': '延期 A为12天，延期 B为7天，相差5天。'}])
    state = agent.run(initial(agent, workbook))
    assert state.status == 'done' and len(model.messages) == 2
    assert [arguments['sheet'] for _, arguments in calls] == [SHEET_A, SHEET_B]
    assert len(approvals) == 2 and len(state.required.evidence) == 2
    tool_text = '\n'.join(item['content'] for item in model.messages[-1] if item['role'] == 'tool')
    assert all(value in tool_text for value in ('延期 A', '延期 B', '12', '7', 'B2'))
    assert state.answer == '延期 A为12天，延期 B为7天，相差5天。'
    assert agent.metrics.snapshot()['counts']['tool_executions'] == 2
    assert_paired(state)
    assert digest(workbook) == before


@pytest.mark.parametrize('invalid', ['parameter', 'path'])
def test_second_invalid_request_rejects_whole_batch_before_real_dispatch(workbook, invalid):
    first, second = region(workbook, SHEET_A, 'valid'), region(workbook, SHEET_B, 'invalid')
    if invalid == 'parameter':
        second['args']['rows'] = 21
    else:
        second['args']['path'] = str(workbook.parent / '未授权.xlsx')
    agent, model, calls, approvals = make_agent(workbook, [response(first, second), {'content': '请修正范围。'}])
    state = agent.run(initial(agent, workbook))
    assert calls == [] and approvals == [] and not state.required.evidence
    assert state.required.errors and len(model.messages) == 2
    assert agent.metrics.snapshot()['counts']['tool_executions'] == 0
    assert_paired(state)


def test_duplicate_call_ids_do_not_execute_either_read(workbook):
    agent, _, calls, approvals = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'same'), region(workbook, SHEET_B, 'same')),
        {'content': '请提供唯一调用标识。'}])
    state = agent.run(initial(agent, workbook))
    assert not calls and not approvals and not state.required.evidence
    assert any('unique' in error.lower() for error in state.required.errors)
    assert_paired(state)


def test_identical_batch_read_reuses_bytes_but_rechecks_both_approvals(workbook):
    agent, model, calls, approvals = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'repeat-a'), region(workbook, SHEET_A, 'repeat-b')),
        {'content': '两处请求均为延期 A B2：12天。'}])
    state = agent.run(initial(agent, workbook))
    assert len(calls) == 1 and len(approvals) == 2 and len(model.messages) == 2
    counts = agent.metrics.snapshot()['counts']
    assert counts['tool_requests'] == 2 and counts['tool_executions'] == counts['tool_reuse'] == 1
    assert len(state.required.evidence) == 1
    assert any('verified batch reuse' in item['content'] for item in state.history if item['role'] == 'tool')
    assert_paired(state)


def test_second_read_declined_has_honest_receipt_and_no_second_dispatch(workbook):
    attempts = []
    def approve(plan):
        attempts.append(plan)
        return len(attempts) == 1
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'approved'), region(workbook, SHEET_B, 'declined'))], confirm=approve)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'confirmation_required' and len(model.messages) == 1
    assert len(calls) == 1 and len(state.required.evidence) == 1
    receipt = next(item['content'] for item in state.history if item.get('tool_call_id') == 'declined')
    assert 'not executed' in receipt.lower() or 'not_executed' in receipt
    assert_paired(state)


def test_stop_during_confirmation_leaves_no_orphan_or_followup_dispatch(workbook):
    stopped = [False]
    def approve(plan):
        stopped[0] = True
        return True
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'stop-a'), region(workbook, SHEET_B, 'stop-b'))],
        confirm=approve, stopped=lambda: stopped[0])
    state = agent.run(initial(agent, workbook))
    assert state.status == 'limit' and len(model.messages) == 1 and not calls
    assert not state.required.evidence
    assert_paired(state)


def test_stop_during_last_real_read_discards_entire_late_batch(workbook):
    stopped = [False]
    def hook(name, arguments, actual):
        result = actual(name, arguments)
        if arguments.get('sheet') == SHEET_B:
            stopped[0] = True
        return result
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'late-a'), region(workbook, SHEET_B, 'late-b'))],
        hook=hook, stopped=lambda: stopped[0])
    state = agent.run(initial(agent, workbook))
    assert state.status == 'limit' and len(model.messages) == 1 and len(calls) == 2
    assert not state.required.evidence and not state.history
    assert agent.metrics.snapshot()['counts']['roundtrips_avoided'] == 0


def test_source_change_during_real_read_pauses_and_rolls_back_evidence(workbook):
    def hook(name, arguments, actual):
        result = actual(name, arguments)
        book = load_workbook(workbook)
        book[SHEET_C]['B2'] = 99
        book.save(workbook)
        book.close()
        return result
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'change-a'), region(workbook, SHEET_B, 'change-b'))], hook=hook)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'confirmation_required' and len(model.messages) == 1 and len(calls) == 1
    assert not state.required.evidence and not state.history
    assert any('source changed' in error.lower() for error in state.required.errors)


def test_new_requirement_drops_old_batch_and_recollects_new_target(workbook):
    changed = [False]
    owner = [None]
    def boundary(phase, staged):
        if phase == 'after_tool' and not changed[0]:
            changed[0] = True
            owner[0].required.goal = '改为比较延期 A和延期 C的B2，忽略旧的延期 B。'
            raise StaleInstruction('new requirement')
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'old-a'), region(workbook, SHEET_B, 'old-b')),
        response(region(workbook, SHEET_A, 'new-a'), region(workbook, SHEET_C, 'new-c')),
        {'content': '采用最新延期 C：3天；延期 A为12天，相差9天。'}], boundary=boundary)
    owner[0] = initial(agent, workbook)
    state = agent.run(owner[0])
    assert state.status == 'done' and len(model.messages) == 3
    assert [args['sheet'] for _, args in calls] == [SHEET_A, SHEET_A, SHEET_C]
    assert len(state.required.evidence) == 2
    assert all('old-a' not in json.dumps(item, ensure_ascii=False) and 'old-b' not in json.dumps(item, ensure_ascii=False)
               for item in state.history)
    assert '3天' in state.answer and '9天' in state.answer
    assert_paired(state)


@pytest.mark.parametrize('tail', ['write', 'unknown'])
def test_mixed_batch_does_not_dispatch_tail_action(workbook, tail):
    later = {'name': 'plan_excel' if tail == 'write' else 'unknown_remote_tool', 'id': 'tail',
             'args': {'path': str(workbook), 'sheet': SHEET_A, 'cell': 'B2', 'value': 88}}
    before = digest(workbook)
    agent, _, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'first'), later), {'content': '后续动作需另行处理。'}])
    state = agent.run(initial(agent, workbook))
    assert all(name == 'read_excel_region' for name, _ in calls)
    assert len(calls) == (1 if tail == 'write' else 0)
    assert not state.required.pending and digest(workbook) == before
    assert_paired(state)


def test_first_write_gate_in_mixed_batch_keeps_persistent_visibility_flags(workbook):
    catalogue = build_tools((workbook,))
    read = [spec for spec in catalogue if spec.name == 'read_excel_region']
    write = [spec for spec in catalogue if spec.name == 'plan_excel']
    agent, _, calls, _ = make_agent(workbook, [
        response({'name': _WRITE_GATE.name, 'args': {}, 'id': 'gate'}, region(workbook, SHEET_A, 'tail')),
        {'content': '写入工具目录已按明确请求开放。'}], tools=[*read, _WRITE_GATE], write_tools=write,
        write_gate=_WRITE_GATE.name, persist=True)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'done' and state.write_tools_requested is True
    assert agent.load().write_tools_requested is True and not calls
    assert 'plan_excel' in agent.tools and _WRITE_GATE.name not in agent.tools
    assert_paired(state)


def test_second_read_failure_can_retry_only_failed_read_with_complete_pairing(workbook):
    failed = [False]
    def hook(name, arguments, actual):
        if arguments.get('sheet') == SHEET_B and not failed[0]:
            failed[0] = True
            raise TimeoutError('read timeout')
        return actual(name, arguments)
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'good-a'), region(workbook, SHEET_B, 'failed-b')),
        response(region(workbook, SHEET_B, 'retry-b')),
        {'content': '补齐延期 B为7天，差值5天。'}], hook=hook)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'done' and len(model.messages) == 3
    assert [args['sheet'] for _, args in calls] == [SHEET_A, SHEET_B, SHEET_B]
    assert len(state.required.evidence) == 2 and '7天' in state.answer
    receipt = next(item['content'] for item in state.history if item.get('tool_call_id') == 'failed-b')
    assert 'failed' in receipt.lower() and 'fresh result' not in receipt.lower()
    assert_paired(state)


def test_identical_versioned_region_reads_without_new_evidence_stop_unlimited_decisions(workbook):
    replies = [response(region(workbook, SHEET_A, f'repeat-{index}-a'), region(workbook, SHEET_A, f'repeat-{index}-b'))
               for index in range(20)]
    agent, model, calls, approvals = make_agent(workbook, replies, decision_limit=None)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'limit' and 2 <= len(model.messages) < 20
    assert any('no new evidence' in error.lower() for error in state.required.errors)
    assert len(state.required.evidence) == 1 and len(approvals) == len(calls) * 2
    assert_paired(state)


def test_changed_ranges_with_same_returned_values_do_not_trigger_repeat_stop(workbook):
    replies = [response(region(workbook, SHEET_A, f'empty-{row}-a', row=row),
                        region(workbook, SHEET_B, f'empty-{row}-b', row=row)) for row in range(4, 12)]
    replies.append({'content': '已核对八组不同空范围，均无记录。'})
    agent, model, calls, _ = make_agent(workbook, replies, decision_limit=None)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'done' and len(model.messages) == 9 and len(calls) == 16
    assert not any('no new evidence' in error.lower() for error in state.required.errors)
    assert_paired(state)


def test_unknown_sheet_names_cannot_be_guessed_as_an_independent_batch(workbook):
    agent, _, calls, approvals = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'guess-a'), region(workbook, SHEET_B, 'guess-b')),
        {'content': '请先读取工作表结构。'}])
    state = agent.run(initial(agent, workbook, '先找出记录延期的两张工作表，再比较延期天数。'))
    assert not calls and not approvals and not state.required.evidence
    assert any('read_excel_structure' in error for error in state.required.errors)
    assert_paired(state)


def test_successful_real_structure_read_admits_following_batch(workbook):
    agent, model, calls, _ = make_agent(workbook, [
        response({'name': 'read_excel_structure', 'id': 'structure', 'args': {'path': str(workbook)}}),
        response(region(workbook, SHEET_A, 'observed-a'), region(workbook, SHEET_B, 'observed-b')),
        {'content': '按实际表名读取：延期 A 12天，延期 B 7天，差值5天。'}])
    state = agent.run(initial(agent, workbook, '先找出记录延期的工作表，再比较对应单元格。'))
    assert state.status == 'done' and len(model.messages) == 3
    assert [name for name, _ in calls] == ['read_excel_structure', 'read_excel_region', 'read_excel_region']
    assert len(state.required.evidence) == 3 and '5天' in state.answer
    assert_paired(state)


def test_invalid_batch_failure_stops_even_when_every_request_uses_fresh_ids(workbook):
    replies = []
    for index in range(20):
        invalid = region(workbook, SHEET_B, f'invalid-{index}')
        invalid['args']['rows'] = 21
        replies.append(response(region(workbook, SHEET_A, f'valid-{index}'), invalid))
    agent, model, calls, _ = make_agent(workbook, replies, decision_limit=None)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'limit' and 2 <= len(model.messages) < 20
    assert not calls and not state.required.evidence
    assert any('invalid batch' in error.lower() for error in state.required.errors)
    assert_paired(state)


def test_truncated_invoke_response_cannot_execute_a_complete_looking_tool_batch(workbook):
    truncated = response(region(workbook, SHEET_A, 'truncated-a'), region(workbook, SHEET_B, 'truncated-b'))
    truncated['response_metadata'] = {'finish_reason': 'length'}
    agent, model, calls, approvals = make_agent(workbook, [truncated], decision_limit=1)
    state = agent.run(initial(agent, workbook))
    assert state.status == 'limit' and len(model.messages) == 1
    assert not calls and not approvals and not state.required.evidence
    assert agent.metrics.snapshot()['counts']['model_failures'] == 1
    assert_paired(state)


def test_dispatcher_change_discards_inflight_batch_instead_of_reusing_old_result(workbook):
    def hook(name, arguments, actual):
        result = actual(name, arguments)
        agent.dispatcher = lambda *_: {'unexpected': True}
        return result
    agent, model, calls, _ = make_agent(workbook, [
        response(region(workbook, SHEET_A, 'old-a'), region(workbook, SHEET_A, 'old-a-again')),
        {'content': '读取服务已更新，请重新取证。'}], hook=hook)
    state = agent.run(initial(agent, workbook))
    assert state.answer == '读取服务已更新，请重新取证。' and len(model.messages) == 2
    assert len(calls) == 1 and not state.required.evidence
    assert agent.metrics.snapshot()['counts']['tool_reuse'] == 0
    assert not any(item['role'] == 'tool' for item in model.messages[-1])
    assert_paired(state)

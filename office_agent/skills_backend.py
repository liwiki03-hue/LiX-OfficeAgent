"""Host adapter between authenticated Skills facade and standalone storage.

No caller roots, sessions, import paths, execution or filesystem writes here.
A directory argument is only received from the facade's native picker callback.
"""
from pathlib import Path
import json
import re

from .skills_store import SkillsError, SkillsStore, _parse
from .skill_context import SkillContext
from .public_skill_package import unpack, original_metadata, prompt_body
from .skill_references import select_references


class SkillsBackend:
    error_type = SkillsError
    error_codes = frozenset({
        'skills_invalid', 'skills_slug_invalid', 'skills_path_changed',
        'skills_not_found', 'skills_file_invalid', 'skills_source_changed',
        'skills_io_failed', 'skills_frontmatter_invalid', 'skills_roots_overlap',
        'skills_state_invalid', 'skills_id_invalid', 'skills_capacity',
        'skills_name_ambiguous', 'skills_disabled', 'skills_session_invalid',
        'skills_uncertain', 'skills_id_collision', 'skills_exists',
        'skills_builtin_immutable', 'skills_enabled_invalid', 'skills_plan_invalid',
        'confirmation_required', 'skills_denied', 'skills_plan_expired',
        'skills_verification_failed',
    })

    def __init__(self, workspace_root, builtin_root=None, *, on_load=None):
        self._store = SkillsStore(workspace_root, builtin_root)
        self._on_load = on_load

    @staticmethod
    def _identifier(name, source):
        if (type(name) is not str or not re.fullmatch('[a-z0-9]+(?:-[a-z0-9]+)*', name)
                or source not in ('builtin', 'user') or type(source) is not str):
            raise SkillsError('skills_id_invalid')
        return source + ':' + name

    def catalog(self):
        from .webui.office_workflows import NAMES, reviewed_workflow
        rows = self._store.catalog(include_sha=True)
        for row in rows:
            workflow = reviewed_workflow(row)
            if workflow:
                row['display_name'] = NAMES[workflow]
            row.pop('sha256')
        return rows

    def read(self, name, source):
        result = self._store.read(self._identifier(name, source))
        body, origin = unpack(result['body'])
        if origin:
            metadata = original_metadata(result['body'])
            metadata['description'] = metadata['description'][:500]
            origin['local_modified'] |= any(result.get(key) != metadata[key] for key in ('name', 'description'))
        return {**result, 'body':body, 'origin':origin}

    def plan_public(self, candidate, session):
        if not candidate.get('compatible') or not isinstance(candidate.get('_content'), str):
            raise SkillsError('skills_invalid')
        result = self._store.plan_create(candidate['slug'], candidate['_content'], session)
        _, envelope = _parse(candidate['_content'].encode('utf-8'))
        _, origin = unpack(envelope)
        return {**result, 'origin':candidate['origin'], 'scope':'text-reference-only', 'workspace':self._store.root.name,
                'format':'LiX 单文件封装（原始文件、受支持文字资料、展示资源与许可完整保留）'+(' · 按需读取，未加载不等于已读' if origin.get('reference_loading')=='task-matched' else '')}

    def public_snapshot(self, name, source):
        if source != 'user':
            raise SkillsError('skills_builtin_immutable')
        result = self.read(name, source)
        if result['origin'] is None:
            raise SkillsError('skills_invalid')
        return {'name':name, 'source':source, 'sha256':result['sha256'], 'origin':result['origin']}

    def validate_public_snapshot(self, snapshot):
        current = self.public_snapshot(snapshot['name'], snapshot['source'])
        if current['sha256'] != snapshot['sha256']:
            raise SkillsError('skills_source_changed')
        return current

    def plan_public_update(self, snapshot, candidate, session):
        current = self.validate_public_snapshot(snapshot)
        if current['origin']['local_modified'] or not candidate.get('compatible'):
            raise SkillsError('skills_source_changed')
        result = self._store.plan_public_update(self._identifier(snapshot['name'], snapshot['source']),
                                               candidate['_content'], snapshot['sha256'], session)
        return {**result, 'origin':candidate['origin'], 'previous_origin':current['origin'],
                'workspace':self._store.root.name, 'format':'固定版本更新；原始文件、文字资料、展示资源与许可一并替换',
                'changes':candidate['update']['files']}

    def plan_change(self, operation, changes, session):
        if type(operation) is not str or type(changes) is not dict:
            raise SkillsError()
        fields = {
            'create': {'slug','content'}, 'edit': {'name','source','content'},
            'delete': {'name','source'}, 'enabled': {'name','source','enabled'},
        }
        if operation not in fields or set(changes) != fields[operation]:
            raise SkillsError()
        if operation == 'create':
            return self._store.plan_create(changes['slug'],changes['content'],session)
        identifier = self._identifier(changes['name'],changes['source'])
        if operation == 'edit':
            return self._store.plan_edit(identifier,changes['content'],session)
        if operation == 'delete':
            return self._store.plan_delete(identifier,session)
        return self._store.plan_enabled(identifier,changes['enabled'],session)

    def plan_import(self, trusted_directory, session):
        # The facade only obtains this path from its host picker, never JS args.
        directory = Path(trusted_directory)
        return self._store.plan_import(directory.name,directory,session)

    def confirm_change(self, plan_id, session, confirmed):
        return self._store.confirm(plan_id,session,confirmed)

    def plan_rollback(self, save_id, session):
        result = dict(self._store.plan_rollback(save_id,session))
        result['rollback_id'] = result.pop('plan_id')
        return result

    def confirm_rollback(self, rollback_id, session, confirmed):
        return self._store.confirm_rollback(rollback_id,session,confirmed)

    def invalidate(self):
        self._store.invalidate()

    def prompt_catalog(self):
        # Stable slug is the prompt identifier; human display name may be Chinese.
        # Too-long descriptions fail closed in SkillContext, never truncate policy.
        return self._prompt_view()[0]

    def _prompt_view(self):
        admitted, excluded, used = [], [], 0
        for item in sorted(self._store.catalog(), key=lambda item:item['id']):
            if not item['enabled']:
                excluded.append({'id':item['id'],'reason':'已停用'})
                continue
            entry = {'name':item['id'].split(':',1)[1], 'description':item['description'][:300],
                     'source':item['source'], 'enabled':True}
            size = len(json.dumps({key:value for key,value in entry.items() if key!='enabled'},ensure_ascii=False,separators=(',',':'))) + bool(admitted)
            if len(admitted) >= 32 or used + size > 4096:
                excluded.append({'id':item['id'],'reason':'本轮目录预算未纳入'})
                continue
            admitted.append(entry)
            used += size
        return admitted, excluded

    def agent_catalog_status(self):
        admitted, excluded = self._prompt_view()
        return {'available':len(admitted),'excluded':excluded,'description_summary_limit':300,
                'note':'原始描述保留；模型目录使用最多300字符摘要，未纳入不代表未安装。'}

    def prompt_load(self, name, source):
        body, origin = unpack(self._store.load(self._identifier(name,source))['body'])
        if origin:
            return prompt_body(body, origin)
        return body

    def context(self):
        return SkillContext(self.prompt_catalog,self.prompt_load,on_load=getattr(self,'_on_load',None),task_load=self.prompt_for_task)

    def prompt_for_task(self, name, source, goal, budget):
        return self._task_prompt(self._store.load(self._identifier(name,source))['body'], goal, budget)

    def turn_choices(self):
        from .webui.office_workflows import NAMES, reviewed_workflow
        choices = []
        for item in self._store.catalog(include_sha=True):
            if item['enabled']:
                choice = {key: item[key] for key in ('id', 'name', 'description', 'sha256')}
                workflow = reviewed_workflow(choice)
                if workflow:
                    choice['name'] = NAMES[workflow]
                choices.append(choice)
        return choices

    def turn_skill(self, ticket, question):
        from .skill_context import MAX_BODY_CHARS
        if not isinstance(ticket, dict) or set(ticket) != {'id', 'sha256'} or not isinstance(ticket['id'], str):
            raise ValueError('请重新选择本轮技能。')
        source, sep, name = ticket['id'].partition(':')
        item = self._store.load(self._identifier(name, source))
        if item['sha256'] != ticket['sha256']:
            raise ValueError('技能内容已改变，请重新选择。')
        material = self._task_prompt(item['body'], question, MAX_BODY_CHARS)
        body = material['body'] if isinstance(material, dict) else material
        if not body.strip() or len(body) > MAX_BODY_CHARS:
            raise ValueError('所选技能内容超出本轮预算，请移除或选择其他技能。')
        from .webui.office_workflows import NAMES, reviewed_workflow
        workflow = reviewed_workflow(ticket)
        if workflow:
            return {'name': NAMES[workflow], 'instructions': body, 'workflow': workflow, 'ticket': dict(ticket)}
        return {'name': item['name'], 'instructions': body}

    @staticmethod
    def _task_prompt(content, goal, budget):
        body, origin = unpack(content)
        if origin and origin.get('provider')=='lix':
            return {'body':prompt_body(body,origin),'display_name':origin['display_name']}
        if origin and origin.get('reference_loading')=='task-matched':
            prefix = prompt_body('',{**origin,'resources':{},'reference_loading':'complete'})
            result = select_references(body,origin['resources'],goal,min(7800,budget-len(prefix)))
            return {**result,'body':prefix+result['body']}
        return prompt_body(body,origin) if origin else body

    def __repr__(self):
        return '<SkillsBackend host-only>'


def workspace_skill_context(root, *, on_load=None):
    try:
        return SkillsBackend(root, on_load=on_load).context()
    except SkillsError:
        return None

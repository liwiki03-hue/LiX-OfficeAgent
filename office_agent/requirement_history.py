from copy import deepcopy
import hashlib
import json
import uuid


def event_digest(record):
    value={key:item for key,item in record.items() if key not in {'event_sha256','text_sha256','actor'}}
    return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode()).hexdigest()


def project_requirements(records):
    rows={}
    for record in records:
        if type(record.get('revision')) is not int or record['revision']<1:
            raise ValueError('要求历史版本号无效，未恢复内容')
        schema=record.get('schema_version',1)
        if type(schema) is not int or schema not in {1,2}:
            raise ValueError('要求历史版本不支持，未恢复内容')
        if schema==1:
            identifier=record.get('id')
            if not isinstance(identifier,str) or not 1<=len(identifier)<=100 or identifier in rows or not isinstance(record.get('text'),str) or not record['text'].strip() or len(record['text'])>4000 or '\x00' in record['text']:
                raise ValueError('要求历史身份或内容无效，未恢复内容')
            rows[identifier]={'id':identifier,'revision':record['revision'],'version':record['revision'],
                              'text':record['text'],'cancelled':False,'history':[]}
            continue
        if record.get('kind')!='requirement_change' or record.get('event_sha256')!=event_digest(record):
            raise ValueError('要求变更记录已改变，未恢复内容')
        target=record.get('target_id')
        row=rows.get(target) if isinstance(target,str) else None
        if not row or row['cancelled'] or type(record.get('expected_version')) is not int or type(record.get('expected_instruction_revision')) is not int or row['version']!=record.get('expected_version') or record.get('expected_instruction_revision')!=record['revision']-1:
            raise ValueError('要求变更顺序或目标不匹配，未恢复内容')
        action=record.get('action')
        if action not in {'replace','cancel'} or not isinstance(record.get('text'),str):
            raise ValueError('要求变更类型无效，未恢复内容')
        if action=='replace' and (not record['text'].strip() or len(record['text'])>4000 or '\x00' in record['text']):
            raise ValueError('要求修改内容无效，未恢复内容')
        if action=='cancel' and record['text']!='':
            raise ValueError('取消要求记录无效，未恢复内容')
        row['history'].append({'version':row['version'],'text':row['text'],'status':'superseded' if action=='replace' else 'cancelled'})
        row['version']=record['revision']
        if action=='replace':row['text']=record['text']
        else:row['cancelled']=True
    return list(rows.values())


def runtime_requirements(runtime):
    if runtime is None:return []
    with runtime.lock:
        records=sorted([*runtime.messages,*getattr(runtime,'requirement_changes',[])],key=lambda row:row['revision'])
        return project_requirements(records)


def active_requirements(runtime):
    return [row for row in runtime_requirements(runtime) if not row['cancelled']]


def change_requirement(runtime, client_id, target_id, version, instruction_revision, action, text):
    if not isinstance(client_id,str) or not 1<=len(client_id)<=100 or not isinstance(target_id,str) or not 1<=len(target_id)<=100 or not isinstance(action,str) or action not in {'replace','cancel'} or not isinstance(text,str):
        raise ValueError('要求修改参数无效')
    if action=='replace':
        text=text.strip()
        if not text or len(text)>4000 or '\x00' in text:raise ValueError('修改后的要求应为1—4000个字符')
    elif text!='':raise ValueError('取消要求不能附带替换内容')
    identity={'target_id':target_id,'expected_version':version,'expected_instruction_revision':instruction_revision,'action':action,'text':text}
    with runtime.lock:
        if runtime.closed or runtime.uncertain:raise ValueError('本轮已停止或结果不确定，请先核对记录')
        previous=runtime.ids.get(client_id)
        if previous:
            if previous.get('kind')!='requirement_change' or any(previous.get(key)!=value for key,value in identity.items()):
                raise ValueError('相同操作标识不能对应不同修改')
            return deepcopy(previous),True
        if type(instruction_revision) is not int or instruction_revision!=runtime.revision:
            raise ValueError('要求列表已更新，请重新读取后核对，当前编辑内容保留')
        row=next((row for row in runtime_requirements(runtime) if row['id']==target_id),None)
        if not row or row['cancelled'] or type(version) is not int or row['version']!=version:
            raise ValueError('目标要求已修改或取消，请重新核对')
        if action=='replace' and row['text']==text:raise ValueError('内容没有变化，无需保存')
        changes=getattr(runtime,'requirement_changes',[])
        records=[*runtime.messages,*changes]
        if len(records)>=64 or sum(len(item['text']) for item in records)+len(text)>16000:
            raise ValueError('本轮要求与变更记录达到容量，请核对后明确开始新一轮')
        record={'schema_version':2,'kind':'requirement_change','id':uuid.uuid4().hex,'client_id':client_id,
                'session_id':runtime.session_id,'document_version':runtime.document_version,'revision':runtime.revision+1,**identity}
        record['event_sha256']=event_digest(record)
        runtime.persist(record)
        runtime.revision+=1
        runtime.requirement_changes=[*changes,record]
        runtime.requirement_epoch=getattr(runtime,'requirement_epoch',0)+1
        runtime.ids[client_id]=record
        runtime.plans.clear()
        runtime.notify('requirement_changed',{**deepcopy(record),'target_revision':row['revision']})
        return deepcopy(record),False

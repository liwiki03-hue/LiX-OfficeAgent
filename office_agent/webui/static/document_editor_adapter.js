(() => {
  'use strict';
  const Contract = window.OfficeEditorContract;
  class DocumentEditorAdapter extends Contract.Adapter {
    constructor(container, onEvent) {
      super();
      this.closed=false;this.pending=new Map();this.sequence=0;this.onEvent=onEvent;
      this.frame=document.createElement('iframe');this.frame.title='实时分页DOCX编辑器';
      this.frame.setAttribute('sandbox','allow-scripts allow-same-origin');
      this.frame.src='paged_editor_frame.html';this.frame.className='paged-frame';container.append(this.frame);
      this.connected=new Promise((yes,no)=>{
        const timer=setTimeout(()=>no(Error('本机编辑器未能连接，请保留原稿并重新打开')),15000);
        this.frame.onload=()=>{
          if(this.closed)return;
          const channel=new MessageChannel();this.port=channel.port1;
          this.port.onmessage=event=>{
            const message=event.data;
            if(this.closed)return;
            if(message.type==='connected'){clearTimeout(timer);yes();return;}
            if(message.id){const request=this.pending.get(message.id);if(request){clearTimeout(request.timer);this.pending.delete(message.id);message.ok?request.yes(message.result):request.no(Error(message.error));}return;}
            if(message.type==='changed')this.revision=message.value.revision;
            this._emitChange(message.type,message.value);
          };
          this.port.start();this.frame.contentWindow.postMessage({type:'office-editor-connect'},'*',[channel.port2]);
        };
      });
    }
    async request(action,value={}) {
      await this.connected;
      if(this.closed)throw Error('编辑会话已经关闭');
      const id=++this.sequence;
      return new Promise((yes,no)=>{
        const timer=setTimeout(()=>{this.pending.delete(id);no(Error('编辑器未返回此操作结果；请核对当前草稿，不要重复采用或发布'));},action==='open'?65000:30000);
        this.pending.set(id,{yes,no,timer});this.port.postMessage({id,action,value});
      });
    }
    async openDocument(data) {
      this._setIdentity(data);
      const result=await this.request('open',{data:data.data,revision:data.draft_revision,editable:data.editable,styles:data.capabilities?.styles,fonts:data.capabilities?.font_report,theme:document.documentElement.dataset.theme||'dark'});
      this._setCapabilities({...data.capabilities,engine:result?.capabilities?.engine,version:result?.capabilities?.version},{format:'docx'});
      return result;
    }
    open(data){return this.openDocument(data);}
    async getSelectionReference(){const value=await this.request('selection');return {node_id:value.node_id,start:value.start,end:value.end};}
    getSelection(){return this.request('selection');}
    getSnapshot(expectedDraftRevision){
      const revision=expectedDraftRevision??this.revision;
      if(revision!==this.revision)throw Error('草稿已更新，请重新核对');
      return this.exportSnapshot(revision);
    }
    exportSnapshot(expectedDraftRevision){return this.request('snapshot',{revision:expectedDraftRevision});}
    exportCandidate(expectedDraftRevision){return this.getSnapshot(expectedDraftRevision);}
    validateCandidate(snapshot,expected={}){
      if(!snapshot?.data||!Array.isArray(snapshot.manifest))throw Error('候选快照缺少正文或结构清单');
      if(expected.revision!==undefined&&snapshot.revision!==expected.revision)throw Error('候选快照版本已改变');
      return {valid:true,revision:snapshot.revision,manifest:snapshot.manifest.length,authority:'host-sync'};
    }
    subscribeChanges(handler){return super.subscribeChanges(handler);}
    capabilities(){return super.capabilities();}
    contract(){return super.contract();}
    applyApprovedProposal(proposal,revision,instruction,token){return this.request('apply',{proposal,revision,instruction,token});}
    async applyValidatedOperations(operations,context={}){
      const list=Array.isArray(operations)?operations:[operations];
      if(list.length<1||list.length>8)throw Error('批量操作数量必须为1—8项');
      if(typeof context.token!=='string'||!context.token)throw Error('缺少宿主采用令牌');
      const instruction=context.instruction_revision??list[0].instruction_revision;
      const keys=new Set(),groups=new Map();
      for(const operation of list){
        Contract.validateOperation(operation,this._operationContext({instruction_revision:instruction}));
        if(this.appliedOperationKeys.has(operation.idempotency_key)||keys.has(operation.idempotency_key))throw Error('此编辑操作已经采用，未重复应用');
        keys.add(operation.idempotency_key);
        if(!groups.has(operation.target_id))groups.set(operation.target_id,[]);
        groups.get(operation.target_id).push(operation);
      }
      for(const spans of groups.values()){
        spans.sort((a,b)=>a.start-b.start);
        for(let index=1;index<spans.length;index++)if(spans[index].start<spans[index-1].end)throw Error('批量操作区间重叠');
      }
      const result=await this.request('applyBatch',{operations:list,revision:this.revision,instruction,token:context.token});
      for(const key of keys)this.appliedOperationKeys.add(key);
      return result;
    }
    applyTableCommand(command){return this.request('tableCommand',{command});}
    undo(){return this.request('undo');}
    redo(){return this.request('redo');}
    history(){return this.request('history');}
    focusTarget(target){const id=typeof target==='string'?target:target?.target_id;if(!id)throw Error('定位目标无效');return this.request('locate',{id});}
    restoreDraft(){throw Error('恢复草稿由宿主 paged_restore 管理');}
    closeDocument(){if(this.closed)return;this.closed=true;this.port?.close();this.frame.remove();for(const request of this.pending.values()){clearTimeout(request.timer);request.no(Error('编辑文档已切换'));}this.pending.clear();}
    close(){this.closeDocument();}
  }
  window.DocumentEditorAdapter=DocumentEditorAdapter;
})();

(() => {
  'use strict';
  const METHODS = Object.freeze([
    'open', 'close', 'capabilities', 'getSelection', 'getSnapshot', 'subscribeChanges',
    'applyValidatedOperations', 'applyTableCommand', 'undo', 'redo', 'exportCandidate',
    'validateCandidate', 'focusTarget', 'restoreDraft',
  ]);
  const OPERATION_FIELDS = Object.freeze([
    'document_id', 'session_epoch', 'source_version', 'draft_revision', 'instruction_revision',
    'target_id', 'expected_before', 'operation_id', 'idempotency_key', 'action', 'start', 'end', 'after',
  ]);
  const FORMATS = Object.freeze(['docx', 'xlsx', 'pptx']);
  const fail = message => { throw Error(message); };

  function normalizeCapabilities(raw = {}, identity = {}) {
    const reasons = [...(Array.isArray(raw.reasons) ? raw.reasons : [])];
    const editable = raw.editable === true;
    return {
      schema_version: 1,
      format: identity.format || 'docx',
      engine: String(raw.engine || 'SuperDoc'),
      version: String(raw.version || '1.47.2'),
      editable,
      reasons,
      layouts: ['paginated'],
      selection: ['single-segment', 'utf16'],
      operations: {replace_text: editable, table_structure: editable, undo: editable, redo: editable, batch: editable},
      export: {formats: ['docx'], validation: 'host-sync'},
      recovery: 'host-snapshot',
      styles: Array.isArray(raw.styles) ? raw.styles.map(row => ({id: String(row.id), name: String(row.name)})) : [],
      limits: {max_operations: 8, max_document_bytes: 8 * 1024 * 1024},
    };
  }

  function validateOperation(operation, context = {}) {
    if (!operation || typeof operation !== 'object' || Array.isArray(operation)) fail('编辑操作必须是对象');
    for (const field of OPERATION_FIELDS) {
      if (!Object.hasOwn(operation, field)) fail('编辑操作缺少字段：' + field);
    }
    for (const field of ['document_id', 'source_version', 'target_id', 'operation_id', 'idempotency_key', 'expected_before', 'after', 'action']) {
      if (typeof operation[field] !== 'string' || !operation[field]) fail('编辑操作字段无效：' + field);
    }
    if (operation.operation_id.length > 128 || operation.idempotency_key.length > 128 || operation.target_id.length > 100) fail('编辑操作标识超过预算');
    if (operation.session_epoch === undefined || operation.session_epoch === null || operation.session_epoch === '') fail('编辑操作缺少会话纪元');
    for (const field of ['draft_revision', 'instruction_revision', 'start', 'end']) {
      if (!Number.isSafeInteger(operation[field]) || operation[field] < 0) fail('编辑操作版本或区间无效：' + field);
    }
    if (operation.end < operation.start) fail('编辑操作区间无效');
    if (operation.action !== 'replace_text') fail('当前适配器只支持经过校验的文字替换操作');
    if (context.document_id !== undefined && operation.document_id !== context.document_id) fail('编辑操作不属于当前文档');
    if (context.session_epoch !== undefined && operation.session_epoch !== context.session_epoch) fail('编辑操作会话已变化');
    if (context.source_version !== undefined && operation.source_version !== context.source_version) fail('编辑操作来源版本已变化');
    if (context.draft_revision !== undefined && operation.draft_revision !== context.draft_revision) fail('编辑操作草稿版本已变化');
    if (context.instruction_revision !== undefined && operation.instruction_revision !== context.instruction_revision) fail('编辑操作要求版本已变化');
    return true;
  }

  class Adapter {
    constructor() {
      this.changeSubscribers = new Set();
      this.capabilityValue = null;
      this.documentId = '';
      this.sourceVersion = '';
      this.sessionEpoch = '';
      this.revision = 0;
      this.appliedOperationKeys = new Set();
    }

    capabilities() {
      if (!this.capabilityValue) fail('编辑器能力尚未加载');
      return {...this.capabilityValue};
    }

    subscribeChanges(handler) {
      if (typeof handler !== 'function') fail('变更订阅必须是函数');
      this.changeSubscribers.add(handler);
      return () => this.changeSubscribers.delete(handler);
    }

    _emitChange(type, value) {
      for (const handler of [...this.changeSubscribers]) {
        try { handler(type, value); } catch (_) {}
      }
      if (typeof this.onEvent === 'function') this.onEvent(type, value);
    }

    _setCapabilities(raw, identity = {}) {
      this.capabilityValue = normalizeCapabilities(raw, identity);
      return {...this.capabilityValue};
    }

    _setIdentity(data = {}) {
      this.documentId = String(data.document_id || '');
      this.sourceVersion = String(data.source_version || '');
      this.sessionEpoch = data.session_epoch ?? '';
      this.revision = Number.isSafeInteger(data.draft_revision) ? data.draft_revision : 0;
    }

    _operationContext(extra = {}) {
      return {
        document_id: this.documentId,
        session_epoch: this.sessionEpoch,
        source_version: this.sourceVersion,
        draft_revision: this.revision,
        ...extra,
      };
    }

    contract() {
      const missing = METHODS.filter(name => typeof this[name] !== 'function');
      return {
        schema_version: 1,
        format: this.capabilityValue?.format || 'docx',
        methods: [...METHODS],
        missing,
        capabilities: this.capabilityValue ? {...this.capabilityValue} : null,
      };
    }
  }

  window.OfficeEditorContract = Object.freeze({
    schemaVersion: 1,
    methods: METHODS,
    operationFields: OPERATION_FIELDS,
    formats: FORMATS,
    normalizeCapabilities,
    validateOperation,
    Adapter,
  });
})();

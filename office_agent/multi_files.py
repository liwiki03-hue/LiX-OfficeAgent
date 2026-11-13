from __future__ import annotations

import json
from pathlib import Path
import time
import uuid

from .documents import inspect_file
from .rag_sources import SourceExtractor
from .writes import GuardError, digest


class MultiFileScope:
    def __init__(self, editor, anchor, names):
        if type(names) is not list or not 2 <= len(names) <= 5 or any(type(name) is not str for name in names) or len({name.casefold() for name in names}) != len(names):
            raise GuardError('select 2-5 distinct file names')
        self.editor, self.anchor = editor, anchor
        self.paths = []
        for name in names:
            if Path(name).name != name or '/' in name or '\\' in name:
                raise GuardError('only same-directory file names allowed')
            path = editor.space.file(str(anchor.parent/name))
            if path.parent != anchor.parent or path.suffix.lower() not in {'.docx','.xlsx'}:
                raise GuardError('multi-file tasks require same-directory DOCX/XLSX')
            report = inspect_file(path)
            if report.get('risks') or report.get('protected') or report.get('macros') or report.get('external_links'):
                raise GuardError('risky multi-file source refused')
            self.paths.append(path)
        if anchor not in self.paths:
            raise GuardError('selected anchor must be included')
        self.versions = {str(path):digest(path) for path in self.paths}
        self.check()

    def check(self):
        for path in self.paths:
            if self.editor.space.file(str(path)) != path or digest(path) != self.versions[str(path)]:
                raise GuardError('multi-file scope changed; select files again')

    def summary(self):
        self.check()
        extractor = SourceExtractor(self.editor.space)
        documents = []
        for path in self.paths:
            source = extractor.extract(path.name)
            blocks = source['blocks']
            excerpts = []
            for block in blocks[:5]:
                text = block.get('text','')
                excerpts.append({'location':block.get('location'), 'text':text[:500], 'truncated':len(text)>500})
            documents.append({'path':str(path),'source':path.name,'source_sha256':source['fingerprint'],
                              'blocks_total':len(blocks),'excerpts':excerpts,'truncated':len(blocks)>5 or any(b['truncated'] for b in excerpts)})
        self.check()
        return {'kind':'explicit_multi_file_scope','documents':documents,'coverage':'bounded excerpts only; not a complete summary',
                'writes':'each individual file plan needs separate confirmation'}

    def dispatcher(self, on_plan=None, ocr_provider=None, on_image=None):
        from .app import build_dispatcher
        routes = {str(path):build_dispatcher(self.editor,path,on_plan=on_plan,ocr_provider=ocr_provider,
                                            on_image=on_image,expected_version=self.versions[str(path)]) for path in self.paths}
        def dispatch(name,args):
            self.check()
            if name in {'batch_documents','search_directory_documents','image_gen','plan_docx_merge','plan_docx_split'}:
                raise GuardError('directory, generation and multi-source writes unavailable in explicit multi-file tasks')
            path = str(self.editor.space.file(args.get('path','')))
            if path not in routes:
                raise GuardError('file was not explicitly selected')
            result = routes[path](name,args)
            self.check()
            return result
        return dispatch


class MultiFileSelection:
    def __init__(self):
        self.pending = None
        self.scope = None

    def invalidate(self):
        self.pending = self.scope = None

    def plan(self, editor, anchor, names):
        scope = MultiFileScope(editor,anchor,names)
        identifier = uuid.uuid4().hex
        self.pending = (identifier,scope,time.monotonic()+300)
        return {'plan_id':identifier,'files':[path.name for path in scope.paths], 'source_sha256':dict(scope.versions),
                'requires_confirmation':True,'notice':'Read summaries and make these exact files available to subsequent tasks; each write remains separately confirmed'}

    def confirm(self, editor, anchor, identifier, confirmed):
        pending, self.pending = self.pending, None
        if not pending or pending[0] != identifier or confirmed is not True or pending[2] <= time.monotonic():
            raise GuardError('explicit multi-file read confirmation required')
        scope = pending[1]
        if scope.editor is not editor or scope.anchor != anchor:
            raise GuardError('multi-file selection expired')
        summary = scope.summary()
        self.scope = scope
        return summary

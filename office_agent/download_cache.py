from pathlib import Path
import time
import uuid

from .downloads import DownloadError, digest, safe


ASSETS = {
    'frontend': ('superdoc-1.47.2.tgz', '12d9ab6178b8082949bbe3171ab395e38b9eec3594ea28f782e9d8fdc18f2125', 64 * 1024 * 1024),
    'renderer': ('LibreOffice_26.2.6.2_Win_x86-64.msi', '788ce7d4b56460357f57552cb8cd848e8f2254f28f6e51fa61e0c65a18096373', 373293056),
}


class DownloadCache:
    def __init__(self, root):
        self.root = safe(Path(root).absolute())
        self.plans = {}

    def preview(self, kind):
        if not isinstance(kind, str) or kind not in ASSETS:
            raise DownloadError('download_cache_kind')
        name, expected, limit = ASSETS[kind]
        directory = safe(self.root / '.runtime/downloads')
        target = safe(directory / name)
        rows = []
        for path in [target, target.with_name(name + '.partial'), target.with_name(name + '.partial.json')]:
            path = safe(path)
            if path.exists():
                if not path.is_file() or path.stat().st_size > (4096 if path.name.endswith('.json') else limit):
                    raise DownloadError('download_cache_capacity')
                rows.append({'name': path.name, 'bytes': path.stat().st_size, 'sha256': digest(path)})
        complete = next((row for row in rows if row['name'] == name), None)
        status = 'missing' if not rows else 'verified' if complete and complete['sha256'] == expected else 'retained_unverified'
        identifier = uuid.uuid4().hex
        plan = {'id': identifier, 'kind': kind, 'status': status, 'expected_sha256': expected, 'files': rows,
                'repairable': bool(rows) and status != 'verified', 'scope': '.runtime/downloads 中此固定组件的缓存；仅隔离，不删除、不联网'}
        self.plans = {identifier: (time.monotonic(), plan)}
        return plan

    def quarantine(self, identifier, confirmed=False):
        saved = self.plans.pop(identifier, None) if isinstance(identifier, str) else None
        if confirmed is not True:
            return {'cancelled': True, 'network_requests': 0}
        if not saved or time.monotonic() - saved[0] > 300 or not saved[1]['repairable']:
            raise DownloadError('download_cache_plan_expired')
        plan = saved[1]
        current = self.preview(plan['kind'])
        if current['files'] != plan['files']:
            raise DownloadError('download_cache_changed')
        directory = safe(self.root / '.runtime/downloads')
        paths = [(safe(directory / row['name']), safe(directory / (row['name'] + '.invalid-' + identifier))) for row in plan['files']]
        moved = []
        try:
            for source, destination in paths:
                if destination.exists():
                    raise DownloadError('download_cache_collision')
                source.rename(destination)
                moved.append((source, destination))
        except OSError as error:
            for source, destination in reversed(moved):
                if not source.exists():
                    destination.rename(source)
            raise DownloadError('download_cache_quarantine_failed') from error
        self.plans.clear()
        return {'status': 'quarantined', 'files': [destination.name for _, destination in moved], 'network_requests': 0,
                'next_step': '在安装器选择重新准备；仅对只读下载重取，不重放安装或文档操作'}

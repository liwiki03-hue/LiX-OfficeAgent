from office_agent import verified_render_cache as cache


def test_verified_cache_is_bounded_to_exact_identity_and_expires(tmp_path, monkeypatch):
    environment = {'complete':True, 'fonts':'fixed', 'renderer':'fixed'}
    identifier = cache.key(tmp_path, 'input', 'projection', 'runtime', environment, 'validator')
    try:
        assert cache.offer(identifier, b'%PDF-known-verified')
        assert cache.get(identifier) == b'%PDF-known-verified'
        assert cache.get(cache.key(tmp_path, 'changed', 'projection', 'runtime', environment, 'validator')) is None
        assert cache.get(cache.key(tmp_path, 'input', 'projection', 'runtime', environment, 'changed-validator')) is None
        assert cache.key(tmp_path, 'input', 'projection', 'runtime', {'complete':False}, 'validator') is None
        monkeypatch.setattr(cache.time, 'monotonic', lambda:10**12)
        assert cache.get(identifier) is None
    finally:
        cache.clear()


def test_invalid_or_oversized_pdf_is_not_reused():
    try:
        assert not cache.offer('bad', b'not-pdf')
        assert cache.get('bad') is None
        assert not cache.offer('large', b'%PDF-' + b'x'*(16*1024*1024))
    finally:
        cache.clear()

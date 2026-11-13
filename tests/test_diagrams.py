"""Diagram foundation only: no GUI/provider/Office or exported files."""
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from office_agent.diagrams import (
    DiagramError, diagram_prompt, repair_prompt, validate_source, validate_svg,
)

SOURCES = {
    'classDiagram': '''classDiagram
class User {
 +String name
 +placeOrder() Order
}
User "1" --> "*" Order : places
Order "1" --> "1" Payment : pays''',
    'flowchart': 'flowchart TD\nUser[User] --> Order[Order]\nOrder --> Payment[Payment]',
    'sequenceDiagram': 'sequenceDiagram\nparticipant User\nparticipant Order\nparticipant Payment\nUser->>Order: Place order\nOrder-->>Payment: Pay',
    'stateDiagram-v2': 'stateDiagram-v2\n[*] --> Ready\nReady --> Paid : Pay\nPaid --> [*]',
    'erDiagram': 'erDiagram\nUSER ||--o{ ORDER : places\nORDER ||--|| PAYMENT : pays',
}
SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 200"><defs><marker id="end" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto"><path d="M0 0 L6 3 L0 6 Z" fill="#222"/></marker></defs><g><rect x="10" y="10" width="100" height="40" fill="white" stroke="black"/><text x="20" y="35" font-family="Arial" font-size="16">User</text><path d="M110 30 L200 30" marker-end="url(#end)"/></g></svg>'


@pytest.mark.parametrize('kind', SOURCES)
def test_five_restricted_sources_are_counted_without_render_claim(kind):
    spec = validate_source(SOURCES[kind], kind)
    assert spec.kind == kind and 1 <= len(spec.nodes) <= 3
    expected_edges = 3 if kind == 'stateDiagram-v2' else 2
    assert spec.edges == expected_edges and spec.sha256 == hashlib.sha256(spec.source.encode()).hexdigest()
    assert spec.evidence() == {'kind': kind, 'nodes': len(spec.nodes), 'edges': expected_edges, 'source_sha256': spec.sha256, 'rendered': False}
    assert 'source' not in spec.evidence() and 'svg' not in spec.evidence()


@pytest.mark.parametrize('source', [
    'flowchart TD\nUser[User]', 'classDiagram\nclass User',
    'sequenceDiagram\nparticipant User', 'stateDiagram-v2\nstate "Ready" as Ready',
    'erDiagram\nUSER {\n int id PK\n string name\n}',
])
def test_single_node_without_edges_is_valid_restricted_source(source):
    spec = validate_source(source)
    assert len(spec.nodes) == 1 and spec.edges == 0


@pytest.mark.parametrize('source', [
    'flowchart TD\n%%{init: {"securityLevel":"loose"}}%%\nA-->B',
    '---\nconfig: {}\n---\nflowchart TD\nA-->B',
    'flowchart TD\nA[<img src=x onerror=alert(1)>] --> B',
    'flowchart TD\nA --> B\nclick A "https://example.com"',
    'flowchart TD\nA --> B\nclassDef red fill:red',
    'flowchart TD\nA --> B\nstyle A fill:red',
    'flowchart TD\nA[https://example.com] --> B',
    'flowchart TD\nA[&#60;script&#62;] --> B',
    'flowchart TD\nA[\\u003cscript] --> B',
    'flowchart TD\nA --> B\nlinkStyle 0 stroke:red',
    'flowchart TD\nA --> B\nsubgraph Malicious\nend',
    'flowchart TD\nA --> B; click A callback',
    'flowchart TD\nA[中文] --> B',
    'flowchart TD\nA[Hidden\u202e] --> B',
    'flowchart TD\nA\x00 --> B',
    '```mermaid\nflowchart TD\nA --> B\n```',
    'pie\n"Secret": 12',
    'sequenceDiagram\nparticipant A\nA->>Undeclared: Pay',
    'classDiagram\nclass A {\n +run()\n',
    'erDiagram\nUSER {\n int id\n',
])
def test_unsupported_source_is_fail_closed(source):
    with pytest.raises(DiagramError) as caught:
        validate_source(source)
    assert caught.value.code == 'unsupported_source'
    assert source not in str(caught.value)


def test_implicit_nodes_edges_size_and_family_are_bounded():
    with pytest.raises(DiagramError):
        validate_source('flowchart TD\n' + '\n'.join(f'N{i} --> N{i+1}' for i in range(12)))
    with pytest.raises(DiagramError):
        validate_source('flowchart TD\n' + '\n'.join('A --> B' for _ in range(37)))
    for value in (None, [], 'x' * 8193):
        with pytest.raises(DiagramError):
            validate_source(value)
    with pytest.raises(DiagramError):
        validate_source(SOURCES['flowchart'], 'classDiagram')


def test_svg_roundtrip_has_local_marker_and_no_active_features():
    safe = validate_svg(SVG)
    root = ET.fromstring(safe)
    assert root.attrib['width'] == '320' and root.attrib['height'] == '200'
    assert 'url(#end)' in safe and '>User<' in safe
    assert validate_svg(safe) == safe


@pytest.mark.parametrize('fragment', [
    '<script>alert(1)</script>', '<foreignObject><div>bad</div></foreignObject>',
    '<image href="https://example.com/a.png"/>', '<use href="#end"/>',
    '<a href="javascript:alert(1)"><text>Link</text></a>',
    '<animate attributeName="href"/>', '<style>@import "https://evil";</style>',
    '<rect onload="alert(1)"/>', '<rect style="fill:url(https://evil)"/>',
    '<rect fill="url(https://evil)"/>', '<path marker-end="url(#missing)"/>',
    '<text font-family="url(https://evil)">Bad</text>',
    '<path d="M 1e999 0 L1 1"/>', '<g transform="translate(1e999 0)"/>',
    '<rect width="999999"/>', '<g id="end"/>',
    '<x:rect xmlns:x="https://evil"/>', '<text data-source="private">Bad</text>',
])
def test_svg_rejects_external_active_unknown_or_extreme_markup(fragment):
    with pytest.raises(DiagramError):
        validate_svg(SVG.replace('</svg>', fragment + '</svg>'))


@pytest.mark.parametrize('svg', [
    '<!DOCTYPE svg [<!ENTITY x "boom">]>' + SVG,
    '<?xml-stylesheet href="https://evil"?>' + SVG,
    SVG.replace('320 200', '0 200'), SVG.replace('320 200', '4097 200'),
    SVG.replace('320 200', '4096 4096'), SVG.replace('320 200', 'NaN 200'),
    '<svg/>', '<html/>', '', SVG + 'x',
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2 2">' + '<g>' * 33 + '</g>' * 33 + '</svg>',
])
def test_svg_limits_and_xml_guards(svg):
    with pytest.raises(DiagramError):
        validate_svg(svg)


def test_prompt_and_repair_are_bounded_and_never_echo_renderer_markup():
    prompt = diagram_prompt('classDiagram', 'User order payment')
    for required in ('12', 'classDiagram', 'participant', 'classDef', '源码', '基数'):
        assert required in prompt
    for attempt in (1, 2):
        repaired = repair_prompt('flowchart', 'User order payment', 'renderer_failed', attempt)
        assert 'renderer_failed' in repaired and SVG not in repaired
    for attempt in (0, 3, True):
        with pytest.raises(DiagramError):
            repair_prompt('flowchart', 'goal', 'renderer_failed', attempt)
    with pytest.raises(DiagramError):
        repair_prompt('flowchart', 'goal', '<script>private</script>', 1)
    for kind, goal in [('pie', 'goal'), ('flowchart', ''), ('flowchart', 'x' * 1001), ('flowchart', 'bad\x00'), ('flowchart', ' ' * 1000 + 'x')]:
        with pytest.raises(DiagramError):
            diagram_prompt(kind, goal)


@pytest.mark.parametrize('value', [None, [], {}, 42])
def test_repair_error_code_must_be_bounded_host_string(value):
    with pytest.raises(DiagramError):
        repair_prompt('flowchart', 'goal', value, 1)


@pytest.mark.parametrize('svg', [
    SVG.replace('d="M0 0', 'marker-end="url(#end)" d="M0 0'),
    SVG.replace('font-size="16"', 'font-size="129"'),
    SVG.replace('stroke="black"', 'stroke="black" stroke-width="21"'),
    SVG.replace('markerWidth="6"', 'markerWidth="65"'),
    SVG.replace('>User<', '>' + 'x' * 8193 + '<'),
    SVG.replace('</svg>', '<g/>' * 4097 + '</svg>'),
    'x' * 1_048_577, SVG.replace('User', '\ud800'),
], ids=['recursive-marker', 'font-limit', 'stroke-limit', 'marker-limit',
        'text-limit', 'element-limit', 'byte-limit', 'invalid-surrogate'])
def test_svg_additional_resource_and_encoding_bounds(svg):
    with pytest.raises(DiagramError) as caught:
        validate_svg(svg)
    assert caught.value.code == 'unsafe_svg'


def test_offline_vendor_is_fixed_complete_and_licensed_without_execution():
    vendor = Path(__file__).resolve().parents[1] / 'office_agent' / 'webui' / 'static' / 'vendor'
    manifest = json.loads((vendor / 'mermaid.manifest.json').read_text(encoding='utf-8'))
    assert manifest['version'] == '11.12.2' and manifest['license'] == 'MIT'
    expected = {
        'mermaid.min.js': 'd0830a6c05546e9edb8fe20a8f545f3e0dc7c4c3134d584bad9c13a99d7a71e0',
        'mermaid.LICENSE': 'ec9fb67dcb25eccc416ed56e1aab819222c805a2a4bfe4cb19e7556bf2ffde80',
    }
    assert {record['file'] for record in manifest['files']} == set(expected)
    for record in manifest['files']:
        data = (vendor / record['file']).read_bytes()
        assert len(data) == record['bytes']
        assert hashlib.sha256(data).hexdigest() == record['sha256'] == expected[record['file']]
        assert '/mermaid@11.12.2/' in record['url']
    assert b'The MIT License' in (vendor / 'mermaid.LICENSE').read_bytes()

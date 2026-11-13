import json
import re


def scalar(value):
    value = value.strip()
    if value.startswith('"'):
        value = json.loads(value)
    elif value.startswith("'") and value.endswith("'"):
        value = value[1:-1].replace("''", "'")
    elif value[:1] in '!&*|>{[?%@`' or ': ' in value or ' #' in value:
        raise ValueError('unsupported_frontmatter')
    if not isinstance(value, str) or not value.strip() or len(value)>2048 or any(ord(c)<32 for c in value):
        raise ValueError('unsupported_frontmatter')
    return value


def frontmatter(head):
    if len(head)>4096 or '\t' in head:
        raise ValueError('unsupported_frontmatter')
    lines, fields, index = head.splitlines(), {}, 0
    while index<len(lines):
        line = lines[index]
        index += 1
        match = re.fullmatch(r'(name|description|license|compatibility|allowed-tools|metadata):[ ]*(.*)',line)
        if not match or match[1] in fields:
            raise ValueError('unsupported_frontmatter')
        key, value = match[1], match[2]
        if key=='metadata':
            if value:
                raise ValueError('unsupported_frontmatter')
            mapping = {}
            while index<len(lines) and lines[index].startswith('  '):
                entry = re.fullmatch(r'  ([A-Za-z0-9][A-Za-z0-9_.-]{0,63}):[ ]*(.+)',lines[index])
                if not entry or entry[1] in mapping or len(mapping)>=16:
                    raise ValueError('unsupported_frontmatter')
                mapping[entry[1]] = scalar(entry[2])
                index += 1
            if not mapping:
                raise ValueError('unsupported_frontmatter')
            fields[key] = mapping
        elif value in ('|','|-','>','>-'):
            chunks = []
            while index<len(lines) and (lines[index].startswith('  ') or not lines[index]):
                chunks.append(lines[index][2:] if lines[index].startswith('  ') else '')
                index += 1
            if not chunks:
                raise ValueError('unsupported_frontmatter')
            text = ('\n' if value.startswith('|') else ' ').join(chunks).rstrip()
            fields[key] = text if value.endswith('-') else text+'\n'
        else:
            fields[key] = scalar(value)
    return fields


def presentation(data):
    if len(data)>8192:
        raise ValueError('metadata_budget')
    lines = data.decode('utf-8').splitlines()
    if not lines or lines[0]!='interface:':
        raise ValueError('unsupported_display_metadata')
    result = {}
    allowed = {'display_name','short_description','icon_small','icon_large','brand_color','default_prompt'}
    for line in lines[1:]:
        match = re.fullmatch(r'  ([a-z_]+):[ ]*(.+)',line)
        if not match or match[1] not in allowed or match[1] in result:
            raise ValueError('unsupported_display_metadata')
        result[match[1]] = scalar(match[2])
    return result

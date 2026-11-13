from copy import deepcopy
import hashlib

from .professional_edits import W, _canonical


def table_identity(entry):
    paragraph = entry['node']
    cell = paragraph.getparent()
    row = cell.getparent() if cell is not None else None
    table = row.getparent() if row is not None else None
    if cell is None or cell.tag != W + 'tc' or row.tag != W + 'tr' or table.tag != W + 'tbl':
        return None
    if table.getparent().tag != W + 'body':
        return None
    grid = table.find(W + 'tblGrid')
    rows = table.findall(W + 'tr')
    columns = len(grid) if grid is not None else 0
    if not 1 <= columns <= 20 or not rows or sum(len(r.findall(W + 'tc')) for r in rows) > 100:
        return None
    if any(node.tag not in {W + 'tblPr', W + 'tblGrid', W + 'tr'} for node in table):
        return None
    for item in rows:
        cells = item.findall(W + 'tc')
        if len(cells) != columns or any(child.tag not in {W + 'trPr', W + 'tc'} for child in item):
            return None
        for item_cell in cells:
            if any(child.tag not in {W + 'tcPr', W + 'p'} for child in item_cell) or len(item_cell.findall(W + 'p')) != 1:
                return None
            if any(item_cell.find(W + 'tcPr/' + W + name) is not None for name in ('gridSpan', 'vMerge', 'hMerge')):
                return None
    masked = deepcopy(table)
    for item in masked.iter(W + 'p'):
        for child in list(item):
            item.remove(child)
        item.attrib.clear()
    return hashlib.sha256(_canonical(masked)).hexdigest()


def plain_current(entry):
    if table_identity(entry) is None or entry['refusal']:
        return False
    table = entry['node'].getparent().getparent().getparent()
    for paragraph in table.iter(W + 'p'):
        if any(node.tag not in {W + 'pPr', W + 'r'} for node in paragraph):
            return False
        if any(child.tag not in {W + 'rPr', W + 't'} for run in paragraph.findall(W + 'r') for child in run):
            return False
        if any(node.tag in {W + name for name in ('rPrChange', 'pPrChange', 'sectPr', 'numPr')} for node in paragraph.iter()):
            return False
    return True


def eligibility(row, mappings):
    if row['baseline_location'] != row['delivered_location'] or row['current_location'] != row['delivered_location']:
        return 'table_structure_changed'
    entries = [mapping.get(row[key]) for mapping, key in zip(mappings, ('baseline_location', 'delivered_location', 'current_location'))]
    if any(entry is None for entry in entries) or not plain_current(entries[-1]):
        return 'table_structure_unsupported'
    identities = [table_identity(entry) for entry in entries]
    return None if identities[0] is not None and len(set(identities)) == 1 else 'table_structure_changed'

"""Supplemental structural checks; no cell text is printed or persisted."""
import json
import sys
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from inventory_large import bounds, local, sha_file, sheet_events, workbook_parts

source = Path(sys.argv[1])
result = {'source_sha256': sha_file(source), 'sheets': []}
with ZipFile(source) as archive:
    for sheet in workbook_parts(archive):
        item = {'part': sheet['part'], 'non_row_sheet_data_children': [],
                'top_level_metadata': [], 'xml_space_attribute_count': 0}
        ranges = defaultdict(list)
        for kind, node in sheet_events(archive, sheet['part']):
            if kind == 'row':
                item['xml_space_attribute_count'] += sum('{http://www.w3.org/XML/1998/namespace}space' in child.attrib for child in node.iter())
                continue
            name = local(node.tag)
            if name == 'mergeCells':
                for child in node:
                    r1, c1, r2, c2 = bounds(child.get('ref'))
                    if r1 < r2:
                        ranges[(r1, r2)].append([c1, c2])
            elif name == 'sheetData':
                item['sheet_data_attribute_names'] = list(node.attrib)
                item['non_row_sheet_data_children'] = [local(child.tag) for child in node]
                item['sheet_data_remainder_serialized_bytes'] = len(ET.tostring(node))
            else:
                item['top_level_metadata'].append({'tag': name, 'serialized_bytes': len(ET.tostring(node))})
        item['shared_vertical_span_group_count'] = sum(len(cols) > 1 for cols in ranges.values())
        item['shared_vertical_span_examples'] = [{'first_row': span[0], 'last_row': span[1], 'column_ranges': cols}
                                                for span, cols in ranges.items() if len(cols) > 1][:20]
        with archive.open(sheet['part']) as stream:
            namespaces = Counter()
            tokens = {b'<![CDATA[': 0, b'<!DOCTYPE': 0, b'<!--': 0, b'<?': 0}
            # Count complete tokens while retaining an overlap across chunks.
            carry = b''
            for block in iter(lambda: stream.read(1024 * 1024), b''):
                data = carry + block
                cut = max(0, len(data) - 16)
                for token in tokens:
                    start = 0
                    while True:
                        pos = data.find(token, start)
                        if pos < 0 or pos >= cut:
                            break
                        tokens[token] += 1
                        start = pos + len(token)
                carry = data[cut:]
            for token in tokens:
                tokens[token] += carry.count(token)
            item['raw_xml_token_counts'] = {token.decode(): count for token, count in tokens.items()}
        with archive.open(sheet['part']) as stream:
            parser = ET.iterparse(stream, events=('start', 'start-ns'))
            for event, node in parser:
                if event == 'start-ns':
                    namespaces[node[1]] += 1
                else:
                    item['root_tag'] = node.tag
                    item['root_attribute_names'] = list(node.attrib)
                    break
            item['root_namespaces'] = dict(namespaces)
        result['sheets'].append(item)
result['source_unchanged'] = result['source_sha256'] == sha_file(source)
Path(__file__).with_name('structure-details.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))

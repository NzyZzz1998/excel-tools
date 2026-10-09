"""Read-only ZIP/ElementTree streaming inventory; never records cell values."""
import argparse
import hashlib
import json
import posixpath
import re
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile


def sha_stream(stream):
    value = hashlib.sha256()
    for block in iter(lambda: stream.read(1024 * 1024), b''):
        value.update(block)
    return value.hexdigest()


def sha_file(path):
    with path.open('rb') as stream:
        return sha_stream(stream)


def local(tag):
    return tag.rsplit('}', 1)[-1]


def coordinate(ref):
    match = re.fullmatch(r'([A-Za-z]+)([0-9]+)', ref or '')
    if not match:
        raise ValueError('Invalid or missing coordinate')
    column = 0
    for char in match[1].upper():
        column = column * 26 + ord(char) - 64
    return int(match[2]), column


def address(row, column):
    letters = ''
    while column:
        column, rem = divmod(column - 1, 26)
        letters = chr(65 + rem) + letters
    return letters + str(row)


def bounds(ref):
    parts = ref.split(':')
    r1, c1 = coordinate(parts[0])
    r2, c2 = coordinate(parts[-1])
    return r1, c1, r2, c2


def workbook_parts(archive):
    base = 'xl/workbook.xml'
    relations = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    targets = {r.get('Id'): (r.get('Target').lstrip('/') if r.get('Target').startswith('/') else
               posixpath.normpath(posixpath.join('xl', r.get('Target'))))
               for r in relations if r.get('Type', '').endswith('/worksheet')}
    workbook = ET.fromstring(archive.read(base))
    result = []
    for node in workbook.iter():
        if local(node.tag) == 'sheet':
            rid = next(value for key, value in node.attrib.items() if key.endswith('}id'))
            result.append({'sheet_name_sha256': hashlib.sha256(node.get('name', '').encode()).hexdigest(),
                           'state': node.get('state', 'visible'), 'part': targets[rid]})
    return result


class HashingReader:
    """Compute the ZIP member digest during its existing first XML parse."""
    def __init__(self, stream):
        self.stream = stream
        self.hash = hashlib.sha256()
        self.reached_eof = False

    def read(self, size=-1):
        block = self.stream.read(size)
        self.hash.update(block)
        self.reached_eof = self.reached_eof or not block
        return block


def sheet_events(archive, part, digests=None):
    """Retain at most one row, then remove it from sheetData, not just clear it."""
    with archive.open(part) as stream:
        reader = HashingReader(stream) if digests is not None else stream
        stack = []
        for event, node in ET.iterparse(reader, events=('start', 'end')):
            if event == 'start':
                stack.append(node)
                continue
            parent = stack[-2] if len(stack) > 1 else None
            if local(node.tag) == 'row' and parent is not None and local(parent.tag) == 'sheetData':
                yield 'row', node
                parent.remove(node)
                node.clear()
            elif parent is not None and local(parent.tag) == 'worksheet':
                yield 'metadata', node
                parent.remove(node)
                node.clear()
            stack.pop()
        if digests is not None:
            if not reader.reached_eof:
                raise RuntimeError('Worksheet XML reader did not reach EOF')
            digests[part] = reader.hash.hexdigest()


class RowRanges:
    """Independent interval index: row queries avoid cell x all-merge scans."""
    def __init__(self, ranges):
        self.middle = None
        if not ranges:
            return
        self.middle = sorted(item['box'][0] for item in ranges)[len(ranges) // 2]
        self.here = [item for item in ranges if item['box'][0] <= self.middle <= item['box'][2]]
        self.left = RowRanges([item for item in ranges if item['box'][2] < self.middle])
        self.right = RowRanges([item for item in ranges if item['box'][0] > self.middle])

    def at(self, row):
        if self.middle is None:
            return []
        found = [item for item in self.here if item['box'][0] <= row <= item['box'][2]]
        if row < self.middle:
            found.extend(self.left.at(row))
        elif row > self.middle:
            found.extend(self.right.at(row))
        return found


def cell_flags(cell):
    names = [local(child.tag) for child in cell]
    return {'type': cell.get('t', 'n'), 'formula': 'f' in names,
            'metadata': any(name in cell.attrib for name in ('vm', 'cm')) or 'extLst' in names,
            'nonempty': 'f' in names or any(node.text for node in cell.iter() if local(node.tag) in ('v', 't'))}


def inspect_sheet(archive, sheet, digests=None):
    count = Counter()
    types, styles, formula_types, features = Counter(), Counter(), Counter(), Counter()
    merges, row_refs, columns, examples = [], set(), set(), []
    min_row = min_col = 10**12
    max_row = max_col = previous_row = 0
    dimension = None
    for kind, node in sheet_events(archive, sheet['part'], digests):
        if kind == 'metadata':
            name = local(node.tag)
            features[name] += 1
            if name == 'dimension':
                dimension = node.get('ref')
            elif name == 'mergeCells':
                for child in node:
                    box = bounds(child.get('ref'))
                    r1, c1, r2, c2 = box
                    merge_kind = ('rectangle' if r2 > r1 and c2 > c1 else 'vertical_only' if r2 > r1 else
                                  'horizontal_only' if c2 > c1 else 'single_cell')
                    merges.append({'ref': child.get('ref'), 'box': box, 'kind': merge_kind})
            elif name == 'cols':
                count['hidden_column_definitions'] += sum(c.get('hidden') in ('1', 'true') for c in node)
                count['styled_column_definitions'] += sum('style' in c.attrib for c in node)
            continue
        count['physical_rows'] += 1
        count['missing_row_coordinates'] += 'r' not in node.attrib
        row = int(node.get('r', previous_row + 1))
        count['duplicate_row_coordinates'] += row in row_refs
        count['out_of_order_rows'] += row < previous_row
        row_refs.add(row)
        previous_row = row
        count['hidden_rows'] += node.get('hidden') in ('1', 'true')
        count['styled_rows'] += 's' in node.attrib
        previous_col, seen = 0, set()
        for cell in node:
            if local(cell.tag) != 'c':
                continue
            count['physical_cells'] += 1
            ref = cell.get('r')
            if ref is None:
                count['missing_cell_coordinates'] += 1
                cell_row, col = row, previous_col + 1
            else:
                cell_row, col = coordinate(ref)
                count['noncanonical_cell_coordinates'] += address(cell_row, col) != ref
            count['duplicate_cell_coordinates_within_row'] += (cell_row, col) in seen
            count['cell_row_mismatch'] += cell_row != row
            count['out_of_order_cells'] += col < previous_col
            seen.add((cell_row, col))
            previous_col = col
            columns.add(col)
            min_row, max_row = min(min_row, cell_row), max(max_row, cell_row)
            min_col, max_col = min(min_col, col), max(max_col, col)
            types[cell.get('t', 'n')] += 1
            styles[cell.get('s', 'implicit')] += 1
            flags = cell_flags(cell)
            count['formula_cells'] += flags['formula']
            count['metadata_cells'] += flags['metadata']
            count['nonempty_payload_cells'] += bool(flags['nonempty'])
            for child in cell:
                if local(child.tag) == 'f':
                    formula_types[child.get('t', 'normal')] += 1
    print(json.dumps({'progress': 'sheet_counts_complete', 'part': sheet['part'],
                      'physical_rows': count['physical_rows'], 'physical_cells': count['physical_cells'],
                      'merge_count': len(merges)}, ensure_ascii=True), flush=True)
    anchors = {(item['box'][0], item['box'][1]): None for item in merges}
    index = RowRanges(merges)
    conflict_counts, conflict_examples = Counter(), []
    for kind, node in sheet_events(archive, sheet['part']):
        if kind != 'row':
            continue
        row = int(node.get('r', '0'))
        relevant = index.at(row)
        if not relevant and not any(r == row for r, _ in anchors):
            continue
        for cell in node:
            if local(cell.tag) != 'c' or not cell.get('r'):
                continue
            position = coordinate(cell.get('r'))
            flags = cell_flags(cell)
            if position in anchors:
                anchors[position] = flags
            for merged in relevant:
                r1, c1, r2, c2 = merged['box']
                if c1 <= position[1] <= c2 and position != (r1, c1) and (flags['nonempty'] or flags['metadata']):
                    key = 'all' if r1 == r2 else 'both'
                    conflict_counts[key] += 1
                    if len(conflict_examples) < 20:
                        conflict_examples.append({'merge': merged['ref'], 'cell': cell.get('r'), 'modes': key})
    predictions = {}
    for mode in ('default', 'all'):
        chosen = [item for item in merges if item['box'][2] > item['box'][0] or
                  (mode == 'all' and item['box'][3] > item['box'][1])]
        filled = 0
        blockers, anchor_types = Counter(), Counter()
        remaining = len(merges) - len(chosen)
        for item in chosen:
            r1, c1, r2, c2 = item['box']
            anchor = anchors[(r1, c1)]
            anchor_types['missing' if anchor is None else anchor['type']] += 1
            if anchor:
                blockers['formula_anchors'] += anchor['formula']
                blockers['metadata_anchors'] += anchor['metadata']
                if anchor['nonempty']:
                    filled += (r2-r1+1)*(c2-c1+1)-1 if mode == 'all' else r2-r1
            if mode == 'default' and c2 > c1:
                remaining += r2-r1+1
        blockers['hidden_independent_content_cells'] = conflict_counts['both'] + (conflict_counts['all'] if mode == 'all' else 0)
        predictions[mode] = {'selected_merges': len(chosen), 'expected_filled_cells': filled,
                             'remaining_merges': remaining, 'anchor_types': dict(anchor_types),
                             'blocker_counts': dict(blockers)}
    covers = None
    if dimension and max_row:
        a, b, c, d = bounds(dimension)
        covers = a <= min_row <= max_row <= c and b <= min_col <= max_col <= d
    active, overlaps = [], []
    overlap_count = maximum_active = 0
    for item in sorted(merges, key=lambda m: m['box']):
        r1, c1, r2, c2 = item['box']
        active = [other for other in active if other['box'][2] >= r1]
        for other in active:
            if c1 <= other['box'][3] and other['box'][1] <= c2:
                overlap_count += 1
                if len(overlaps) < 20:
                    overlaps.append([other['ref'], item['ref']])
        active.append(item)
        maximum_active = max(maximum_active, len(active))
    merge_structure = {'maximum_simultaneously_active_by_row': maximum_active,
                       'overlapping_merge_pairs': overlap_count, 'overlap_address_examples': overlaps,
                       'maximum_merge_height': max((m['box'][2] - m['box'][0] + 1 for m in merges), default=0),
                       'maximum_merge_width': max((m['box'][3] - m['box'][1] + 1 for m in merges), default=0)}
    return dict(sheet, counts=dict(count), dimension=dimension, dimension_covers_cells=covers,
                actual_bounds=(address(min_row, min_col)+':'+address(max_row, max_col)) if max_row else None,
                max_row=max_row, max_column=max_col, distinct_physical_columns=len(columns),
                cell_types=dict(types), style_indices=dict(styles), formula_types=dict(formula_types),
                worksheet_features=dict(features), merge_count=len(merges),
                merge_types=dict(Counter(item['kind'] for item in merges)), predictions=predictions,
                merge_structure=merge_structure,
                conflict_examples=conflict_examples, merge_address_samples=[item['ref'] for item in merges[:20]])


def inspect_shared_styles(archive, part):
    counts = Counter()
    with archive.open(part) as stream:
        stack = []
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                stack.append(node)
                continue
            name = local(node.tag)
            parent = stack[-2] if len(stack) > 1 else None
            if parent is not None and local(parent.tag) in ('sst', 'cellXfs', 'cellStyleXfs', 'fonts', 'fills', 'borders', 'numFmts'):
                counts[local(parent.tag)+'/'+name] += 1
                parent.remove(node)
                node.clear()
            stack.pop()
    return dict(counts)


def run(source, destination):
    start = time.perf_counter()
    before = sha_file(source)
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(), 'source_name': source.name,
              'source_bytes': source.stat().st_size, 'source_sha256_before': before,
              'method': 'ZIP hashes + ElementTree row streaming; worksheet hash captured in first XML pass; rows removed immediately; second pass stores anchor flags only; row-coordinate set retained for duplicate checks; no engine import.',
              'privacy': 'No business values, formulas, names, custom format strings or external URLs recorded.',
              'script_sha256': sha_file(Path(__file__))}
    with ZipFile(source) as archive:
        sheets = workbook_parts(archive)
        worksheet_parts = {sheet['part'] for sheet in sheets}
        part_digests = {}
        report['parts'] = []
        for info in archive.infolist():
            if info.filename not in worksheet_parts:
                with archive.open(info) as stream:
                    part_digests[info.filename] = sha_stream(stream)
            report['parts'].append({'part': info.filename, 'compressed_bytes': info.compress_size,
                                    'uncompressed_bytes': info.file_size})
        names = archive.namelist()
        report['package_features'] = {key: sum(bool(re.search(pattern, name)) for name in names) for key, pattern in {
            'media': r'^xl/media/', 'drawings': r'^xl/drawings/[^/]+\.xml$', 'charts': r'^xl/charts/[^/]+\.xml$',
            'external_links': r'^xl/externalLinks/', 'macros': r'vbaProject\.bin$', 'metadata': r'metadata\.xml$',
            'rich_data': r'^xl/richData/', 'tables': r'^xl/tables/', 'comments': r'^xl/comments',
            'document_signatures': r'^_xmlsignatures/'}.items()}
        report['shared_strings'] = inspect_shared_styles(archive, 'xl/sharedStrings.xml') if 'xl/sharedStrings.xml' in names else None
        report['styles'] = inspect_shared_styles(archive, 'xl/styles.xml') if 'xl/styles.xml' in names else None
        report['sheets'] = [dict(inspect_sheet(archive, sheet, part_digests), sheet_index=index)
                            for index, sheet in enumerate(sheets, 1)]
        for item in report['parts']:
            item['sha256'] = part_digests[item['part']]
        report['zip_crc_reads'] = 'passed'
    report['source_sha256_after'] = sha_file(source)
    report['source_unchanged'] = before == report['source_sha256_after']
    report['seconds'] = time.perf_counter()-start
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'source_sha256': before, 'seconds': report['seconds'],
                      'sheets': [{'index': s['sheet_index'], 'counts': s['counts'], 'bounds': s['actual_bounds'],
                                  'merges': s['merge_count'], 'predictions': s['predictions']} for s in report['sheets']]}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('--output', type=Path, default=Path(__file__).with_name('input-inventory.json'))
    args = parser.parse_args()
    run(args.source, args.output)

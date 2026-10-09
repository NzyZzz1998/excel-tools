"""Independent streaming OOXML oracle; keeps one row and merge anchors only.

No processing-engine imports. Cell values exist transiently for equality checks,
but reports contain only counts, coordinates, structural issues and hashes.
"""
import argparse
import hashlib
import json
import time
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from inventory_large import RowRanges, address, bounds, coordinate, local, sha_file, sha_stream, workbook_parts

SPACE = '{http://www.w3.org/XML/1998/namespace}space'


def canonical(node, inherited_space='default'):
    space = node.get(SPACE, inherited_space)
    significant = local(node.tag) in ('t', 'v', 'f') or space == 'preserve'
    text = node.text if significant or (node.text and node.text.strip()) else None
    attrs = tuple(sorted((key, val) for key, val in node.attrib.items() if key != SPACE))
    children = tuple((canonical(child, space), child.tail if child.tail and child.tail.strip() else None) for child in node)
    # Effective xml:space on text is compared independently of declaration location.
    return node.tag, attrs, text, space if local(node.tag) == 't' else None, children


def row_stream(archive, part):
    with archive.open(part) as stream:
        stack, spaces = [], []
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                stack.append(node)
                spaces.append(node.get(SPACE, spaces[-1] if spaces else 'default'))
                continue
            parent = stack[-2] if len(stack) > 1 else None
            if local(node.tag) == 'row' and parent is not None and local(parent.tag) == 'sheetData':
                yield node, spaces[-1]
                parent.remove(node)
                node.clear()
            elif parent is not None and local(parent.tag) == 'worksheet':
                parent.remove(node)
                node.clear()
            stack.pop()
            spaces.pop()


def sheet_plan(archive, part):
    result = {'merges': [], 'metadata': [], 'dimension': None}
    with archive.open(part) as stream:
        stack, spaces = [], []
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                stack.append(node)
                spaces.append(node.get(SPACE, spaces[-1] if spaces else 'default'))
                if len(stack) == 1:
                    result['root_tag'] = node.tag
                    result['root_attributes'] = dict(node.attrib)
                if local(node.tag) == 'sheetData' and len(stack) == 2:
                    result['sheet_data_attributes'] = dict(node.attrib)
                    result['sheet_data_non_rows'] = []
                continue
            parent = stack[-2] if len(stack) > 1 else None
            if parent is not None and local(parent.tag) == 'sheetData':
                if local(node.tag) != 'row':
                    result['sheet_data_non_rows'].append(canonical(node, spaces[-2]))
                parent.remove(node)
                node.clear()
            elif parent is not None and local(parent.tag) == 'worksheet':
                name = local(node.tag)
                if name == 'mergeCells':
                    result['merges'] = [{'ref': child.get('ref'), 'box': bounds(child.get('ref'))} for child in node]
                elif name == 'dimension':
                    result['dimension'] = node.get('ref')
                elif name != 'sheetData':
                    result['metadata'].append(canonical(node, spaces[-2]))
                parent.remove(node)
                node.clear()
            stack.pop()
            spaces.pop()
    return result


def content(cell, inherited_space='default'):
    return (tuple(sorted((key, value) for key, value in cell.attrib.items() if key != 'r' and key != SPACE)),
            tuple(canonical(child, cell.get(SPACE, inherited_space)) for child in cell))


def anchor_payload(cell, inherited_space):
    if cell is None:
        return (), (), False
    attrs = tuple(sorted((key, cell.get(key)) for key in ('s', 't') if key in cell.attrib))
    nodes = tuple(canonical(child, cell.get(SPACE, inherited_space)) for child in cell if local(child.tag) in ('v', 'is'))
    nonempty = any(node.text for node in cell.iter() if local(node.tag) in ('v', 't'))
    return attrs, nodes, bool(nonempty)


def union_rows(ranges):
    spans = []
    for first, last in sorted((item['box'][0], item['box'][2]) for item in ranges):
        if spans and first <= spans[-1][1] + 1:
            spans[-1][1] = max(spans[-1][1], last)
        else:
            spans.append([first, last])
    for first, last in spans:
        yield from range(first, last + 1)


def compare_sheet(before, after, part, all_merges, issue):
    source_plan, output_plan = sheet_plan(before, part), sheet_plan(after, part)
    selected = [item for item in source_plan['merges'] if item['box'][2] > item['box'][0] or
                (all_merges and item['box'][3] > item['box'][1])]
    selected_refs = {item['ref'] for item in selected}
    expected_merges = [item['ref'] for item in source_plan['merges'] if item['ref'] not in selected_refs]
    if not all_merges:
        for item in selected:
            r1, c1, r2, c2 = item['box']
            if c2 > c1:
                expected_merges.extend(address(row, c1)+':'+address(row, c2) for row in range(r1, r2+1))
    if sorted(expected_merges) != sorted(item['ref'] for item in output_plan['merges']):
        issue(part, 'remaining merge ranges differ')
    for key in ('root_tag', 'root_attributes', 'sheet_data_attributes', 'sheet_data_non_rows', 'metadata'):
        if source_plan.get(key) != output_plan.get(key):
            issue(part, 'non-target worksheet structure differs: '+key)
    if not selected:
        with before.open(part) as stream:
            source_hash = sha_stream(stream)
        with after.open(part) as stream:
            output_hash = sha_stream(stream)
        if source_hash != output_hash:
            issue(part, 'no-match worksheet is not byte-identical')
    index = RowRanges(selected)
    anchors = {}
    src_iter, dst_iter, target_iter = row_stream(before, part), row_stream(after, part), union_rows(selected)
    src = next(src_iter, None)
    dst = next(dst_iter, None)
    target_row = next(target_iter, None)
    source_cells = output_cells = expected_cells = filled = checked_rows = target_cells = 0
    previous_source = previous_output = 0
    max_row = max_col = 0
    min_row = min_col = 10**12
    while src is not None or target_row is not None:
        src_number = int(src[0].get('r')) if src is not None else None
        row_number = min(number for number in (src_number, target_row) if number is not None)
        has_source = src_number == row_number
        row, row_space = src if has_source else (None, 'default')
        if has_source:
            if row_number <= previous_source:
                issue(part, 'source rows not strictly increasing', address(row_number, 1))
                raise ValueError('Streaming oracle requires source rows in order; inventory must be reviewed.')
            previous_source = row_number
        source_nodes = {}
        if row is not None:
            for cell in row:
                if local(cell.tag) != 'c':
                    continue
                cr, col = coordinate(cell.get('r'))
                if cr != row_number or col in source_nodes:
                    raise ValueError('Source cell coordinates are ambiguous; no output values recorded.')
                source_nodes[col] = cell
            source_cells += len(source_nodes)
        expected = {col: (cell.get('r'), content(cell, row_space)) for col, cell in source_nodes.items()}
        covered = index.at(row_number)
        attrs = dict(row.attrib) if row is not None else {'r': str(row_number)}
        if covered:
            attrs.pop('spans', None)
        for merged in covered:
            r1, c1, r2, c2 = merged['box']
            for column, candidate in source_nodes.items():
                if c1 <= column <= c2 and (row_number, column) != (r1, c1):
                    independent = (any(local(child.tag) in ('f', 'extLst') for child in candidate)
                                   or any(key in candidate.attrib for key in ('vm', 'cm'))
                                   or any(node.text for node in candidate.iter() if local(node.tag) in ('v', 't')))
                    if independent:
                        issue(part, 'source merge hides independent content; processing should reject', candidate.get('r'))
            if row_number == r1:
                anchor = source_nodes.get(c1)
                if anchor is not None and (any(local(child.tag) in ('f', 'extLst') for child in anchor) or
                                           any(key in anchor.attrib for key in ('vm', 'cm'))):
                    issue(part, 'selected anchor has unsupported formula or metadata', address(r1, c1))
                anchors[merged['ref']] = anchor_payload(anchor, row_space)
            payload_attrs, payload_nodes, nonempty = anchors[merged['ref']]
            for col in range(c1, (c2 if all_merges else c1)+1):
                if (row_number, col) == (r1, c1):
                    continue
                expected[col] = (address(row_number, col), (payload_attrs, payload_nodes))
                filled += nonempty
                target_cells += 1
        if dst is None:
            issue(part, 'missing output row', address(row_number, 1))
        else:
            actual_row, actual_space = dst
            actual_number = int(actual_row.get('r'))
            if actual_number <= previous_output:
                issue(part, 'output rows not strictly increasing', address(actual_number, 1))
            previous_output = actual_number
            if actual_number != row_number:
                issue(part, 'output row sequence differs', address(row_number, 1))
                raise ValueError('Output row sequence differs; stopped rather than misalign all following rows.')
            if actual_row.attrib != attrs:
                issue(part, 'row attributes differ', address(row_number, 1))
            source_children = row if row is not None else ()
            if [canonical(child, row_space) for child in source_children if local(child.tag) != 'c'] != [canonical(child, actual_space) for child in actual_row if local(child.tag) != 'c']:
                issue(part, 'non-cell row content differs', address(row_number, 1))
            seen, previous_col = set(), 0
            for cell in actual_row:
                if local(cell.tag) != 'c':
                    continue
                cr, col = coordinate(cell.get('r'))
                if cr != row_number or col in seen or col < previous_col:
                    issue(part, 'output cell coordinates duplicate or out of order', cell.get('r'))
                seen.add(col)
                previous_col = col
                if expected.get(col) != (cell.get('r'), content(cell, actual_space)):
                    issue(part, 'cell value/type/style/XML semantic mismatch', cell.get('r'))
                max_row, max_col = max(max_row, cr), max(max_col, col)
                min_row, min_col = min(min_row, cr), min(min_col, col)
            for col in expected.keys()-seen:
                issue(part, 'missing output cell', address(row_number, col))
            output_cells += len(seen)
            dst = next(dst_iter, None)
        expected_cells += len(expected)
        checked_rows += 1
        if has_source:
            src = next(src_iter, None)
        if target_row == row_number:
            target_row = next(target_iter, None)
    if dst is not None:
        issue(part, 'extra output rows remain', address(int(dst[0].get('r')), 1))
    dimension = output_plan['dimension']
    if selected:
        if not dimension and source_plan['dimension']:
            issue(part, 'processed worksheet missing dimension')
        elif dimension:
            r1, c1, r2, c2 = bounds(dimension)
            if max_row and not (r1 <= min_row <= max_row <= r2 and c1 <= min_col <= max_col <= c2):
                issue(part, 'dimension excludes output cells')
            for ref in expected_merges + ([source_plan['dimension']] if source_plan['dimension'] else []):
                a, b, c, d = bounds(ref)
                if not (r1 <= a <= c <= r2 and c1 <= b <= d <= c2):
                    issue(part, 'dimension excludes retained merge or shrinks original dimension', ref)
    return {'part': part, 'selected_merges': len(selected), 'remaining_merges': len(expected_merges),
            'source_cells_checked': source_cells, 'expected_cells': expected_cells, 'output_cells_checked': output_cells,
            'rows_checked': checked_rows, 'target_cells': target_cells, 'expected_filled_cells': filled,
            'output_dimension': dimension, 'anchors_retained': len(anchors)}


def run(source, output, all_merges, destination):
    started = time.perf_counter()
    initial = {'source': sha_file(source), 'output': sha_file(output)}
    issues, issue_count = [], 0

    def issue(part, reason, address=None):
        nonlocal issue_count
        issue_count += 1
        if len(issues) < 100:
            issues.append({'part': part, 'reason': reason, 'address': address})

    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(), 'source_name': source.name,
              'output_name': output.name, 'source_sha256': initial['source'], 'output_sha256': initial['output'],
              'mode': 'all' if all_merges else 'default', 'script_sha256': sha_file(Path(__file__)),
              'inventory_helpers_sha256': sha_file(Path(__file__).with_name('inventory_large.py')),
              'method': 'Independent ET per-row expected content from original merge anchors, no processing-engine import.',
              'privacy': 'Report omits all cell text/formulas/values; only counts, addresses, structure and hashes.'}
    try:
        with ZipFile(source) as before, ZipFile(output) as after:
            if before.namelist() != after.namelist() or before.comment != after.comment:
                issue(None, 'ZIP member list/order or archive comment differs')
            if workbook_parts(before) != workbook_parts(after):
                issue(None, 'workbook sheets or relationships differ')
            results = []
            for sheet in workbook_parts(before):
                results.append(compare_sheet(before, after, sheet['part'], all_merges, issue))
            changed = {item['part'] for item in results if item['selected_merges']}
            untouched = []
            for part in before.namelist():
                if part in changed:
                    continue
                with before.open(part) as stream:
                    a = sha_stream(stream)
                with after.open(part) as stream:
                    b = sha_stream(stream)
                untouched.append({'part': part, 'sha256': a, 'unchanged': a == b})
                if a != b:
                    issue(part, 'untouched part bytes differ')
            report['sheets'] = results
            report['untouched_parts'] = untouched
    except Exception as error:
        issue(None, 'validator stopped: '+type(error).__name__)
        report['exception_kind'] = type(error).__name__
        # Exception message can include source XML; do not persist it.
    report['source_unchanged_during_validation'] = sha_file(source) == initial['source']
    report['output_unchanged_during_validation'] = sha_file(output) == initial['output']
    report['issues'], report['issue_count'] = issues, issue_count
    report['passed'] = not issue_count and report['source_unchanged_during_validation'] and report['output_unchanged_during_validation']
    report['seconds'] = time.perf_counter()-started
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'passed': report['passed'], 'issues': issue_count, 'seconds': report['seconds'],
                      'sheets': report.get('sheets', [])}, ensure_ascii=False))
    return report['passed']


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('source', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--all-merges', action='store_true')
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    raise SystemExit(0 if run(args.source, args.output, args.all_merges, args.report) else 1)

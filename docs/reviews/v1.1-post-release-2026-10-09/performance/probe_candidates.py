"""Scratch-only candidate probes; no application files or standard-library globals patched."""
import argparse
import cProfile
import hashlib
import inspect
import json
import statistics
from pathlib import Path
import pstats
import time
from xml.dom import Node

from profile_stream import ENGINE, HERE, NS, fixture, load, run


def row_fragments(module):
    # List buffering is restricted to one detached node/row, never the full document.
    source = inspect.getsource(module.serialize_xml)
    source = source.replace('output = StringIO() if stream is None else None',
        'fragments = [] if stream is None and document.nodeType != Node.DOCUMENT_NODE else None\n'
        '    output = StringIO() if stream is None and fragments is None else None')
    source = source.replace('write = output.write if output is not None else lambda text: stream.write(text.encode("utf-8"))',
        'write = fragments.append if fragments is not None else (output.write if output is not None else lambda text: stream.write(text.encode("utf-8")))')
    source = source.replace('    if output is not None:\n        return output.getvalue().encode("utf-8")',
        '    if fragments is not None:\n        return "".join(fragments).encode("utf-8")\n'
        '    if output is not None:\n        return output.getvalue().encode("utf-8")')
    exec(compile(source, '<scratch-row-fragments>', 'exec'), module.__dict__)


def anchor_context(module):
    source = inspect.getsource(module.clone_payload)
    source = source.replace('    cloned = child.cloneNode(True)',
        '    context = getattr(child, "_review_clone_context", None)\n'
        '    if context is None:\n'
        '        context = (namespace_bindings(child), xml_space(child))\n'
        '        child._review_clone_context = context\n'
        '    cloned = child.cloneNode(True)')
    source = source.replace('namespace_bindings(child).items()', 'context[0].items()')
    source = source.replace('source_space = xml_space(child)', 'source_space = context[1]')
    exec(compile(source, '<scratch-anchor-context>', 'exec'), module.__dict__)


def validated_columns(module):
    # The first pass already validated the original row and raw r attributes.
    def column_of_validated_ref(ref):
        col = 0
        for letter in ref.rstrip('0123456789').replace('$', '').upper():
            col = col * 26 + ord(letter) - 64
        return col
    module.column_of_validated_ref = column_of_validated_ref
    source = inspect.getsource(module.transform_sheet_stream)
    source = source.replace('cells = {coordinate(cell.getAttribute("r"))[1]: cell',
                            'cells = {column_of_validated_ref(cell.getAttribute("r")): cell')
    exec(compile(source, '<scratch-validated-columns>', 'exec'), module.__dict__)


def reusable_row_builder(module):
    source = inspect.getsource(module.transform_sheet_stream)
    source = source.replace('            for number in sorted(row_numbers):',
                            '            row_builder = expatbuilder.ExpatBuilderNS()\n'
                            '            for number in sorted(row_numbers):')
    source = source.replace('row_doc = minidom.parseString(prefix + raw + suffix)',
                            'row_doc = row_builder.parseString(prefix + raw + suffix)')
    exec(compile(source, '<scratch-reusable-row-builder>', 'exec'), module.__dict__)


def cases():
    def sheet(body, merge='A1:A3', attributes=''):
        return (f'<worksheet xmlns="{NS}" {attributes}><dimension ref="A1"/><sheetData>'
                + body+'</sheetData><mergeCells><mergeCell ref="'+merge+'"/></mergeCells></worksheet>').encode()
    return {
        'entities': sheet('<row r="1"><c r="A1" t="inlineStr"><is><t>a&#13;b&#10;c&#9;d&amp;&lt;</t></is></c><c r="B1" t="inlineStr"><is><t>outside&#13;x</t></is></c></row>'),
        'inherited_space_and_local_prefix': sheet(f'<row r="1" xml:space="preserve" xmlns:q="{NS}"><c r="A1" t="inlineStr"><q:is><q:t>  x  </q:t></q:is></c></row><row r="2" xml:space="default"/>'),
        'prefixed_root': (f'<s:worksheet xmlns:s="{NS}"><s:sheetData><s:row r="1"><s:c r="A1" t="inlineStr"><s:is><s:t>x</s:t></s:is></s:c></s:row></s:sheetData><s:mergeCells><s:mergeCell ref="A1:A3"/></s:mergeCells></s:worksheet>').encode(),
        'prefix_shadow': sheet(f'<row r="1" xmlns:p="{NS}"><c r="A1" t="inlineStr"><p:is><p:t>x</p:t></p:is></c></row><row r="2" xmlns:p="urn:synthetic-other"/>'),
        'cdata_pi_attribute': sheet('<row r="1" custom="a&#13;b&#10;c&#9;d"><c r="A1" t="inlineStr"><is><t><![CDATA[a&<b>]]><?synthetic data?></t></is></c></row>'),
        'noncanonical': sheet('<row r="1"><c r="$a$1"><v>4</v></c><c r="b1"><v>8</v></c></row>'),
        'hidden_metadata_reject': sheet('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2" vm="1"/></row>'),
        'formula_reject': sheet('<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>2</v></c></row>'),
        'no_match': sheet('<row r="1"><c r="A1"><v>1</v></c></row>', 'A1:B1'),
        'rectangle': sheet('<row r="1"><c r="A1"><v>1</v></c></row>', 'A1:C4'),
    }


VARIANTS = {'baseline': None, 'row_fragments': row_fragments,
            'anchor_context': anchor_context, 'validated_columns': validated_columns,
            'reusable_row_builder': reusable_row_builder}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=4000)
    parser.add_argument('--repeats', type=int, default=5)
    parser.add_argument('--report', type=Path, default=HERE/'candidate-results.json')
    args = parser.parse_args()
    if args.report.exists() or not 1000 <= args.rows <= 12000 or not 1 <= args.repeats <= 5:
        raise ValueError('Preserve evidence and bounded workload')
    raw, merges, fills = fixture(args.rows)
    modules = {}
    for name, install in VARIANTS.items():
        module = load()
        if install:
            install(module)
        modules[name] = module
    edge_results = []
    for label, data in cases().items():
        expected = None
        for name, module in modules.items():
            try:
                result = run(module, data)
                identity = {'counts': result['counts'], 'sha256': result['output_sha256']}
            except Exception as error:
                identity = {'exception_type': type(error).__name__, 'message': str(error)}
            if name == 'baseline':
                expected = identity
            assert identity == expected, (label, name, identity, expected)
            edge_results.append({'case': label, 'candidate': name, 'baseline_identical': True})
    observations = {name: [] for name in modules}
    for module in modules.values():
        run(module, fixture(1000)[0])
    for repeat in range(args.repeats):
        names = list(modules)
        # Rotate order so each implementation occurs in every position once.
        names = names[repeat:] + names[:repeat]
        for name in names:
            result = run(modules[name], raw)
            assert result['counts'] == (merges, fills)
            if observations['baseline']:
                assert result['output_sha256'] == observations['baseline'][0]['output_sha256']
            observations[name].append(result)
    medians = {name: statistics.median(x['seconds'] for x in values) for name, values in observations.items()}
    record = dict(engine_sha256=hashlib.sha256(ENGINE.read_bytes()).hexdigest(),
                  rows=args.rows, columns=22, active_merges_max=4,
                  original_cells=args.rows*18+merges, expected_fills=fills, input_bytes=len(raw),
                  order='five-position rotation; one warmup at 1000 rows per variant; no profiler in comparisons',
                  observations=observations, median_seconds=medians,
                  relative_change_percent={name: (1-value/medians['baseline'])*100 for name, value in medians.items()},
                  edge_equivalence_checks=edge_results,
                  scope='Synthetic forced-stream transform only, not ZIP pipeline or actual monthly file. Scratch module-local replacements are experiments, not application changes or general correctness proofs.')
    args.report.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'medians': medians, 'relative_change_percent': record['relative_change_percent'],
                      'edge_equivalence_checks': len(edge_results), 'output_sha256': observations['baseline'][0]['output_sha256']}, indent=2))


if __name__ == '__main__':
    main()

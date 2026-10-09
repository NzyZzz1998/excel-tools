"""Bounded scratch-only probes for serializer/no-match costs; never edits the engine."""
import argparse
import cProfile
import hashlib
import inspect
import io
import json
from pathlib import Path
import pstats
import re
import statistics
from xml.parsers import expat

from profile_stream import ENGINE, HERE, fixture, load, run
from probe_candidates import cases


def skip_empty_attribute_maps(module):
    source = inspect.getsource(module.serialize_xml)
    source = source.replace('            for attribute in node.attributes.values():\n',
                            '            for attribute in (node.attributes.values() if node.hasAttributes() else ()):\n')
    exec(compile(source, '<scratch-empty-attribute-map>', 'exec'), module.__dict__)


def preflight(module):
    # Scratch prototype. Keeps malformed-bound errors and first mergeCells semantics;
    # it is intentionally a seekable-input wrapper, not a production integration.
    original = module.transform_sheet_stream
    def transformed(source, output, all_merges=False):
        parser = expat.ParserCreate(namespace_separator='\x1f')
        depth, first_group, group_depth = 0, False, None
        class Selected(Exception):
            pass
        def start(name, attrs):
            nonlocal depth, first_group, group_depth
            depth += 1
            local = name.rsplit('\x1f', 1)[-1]
            if depth == 2 and local == 'mergeCells' and not first_group:
                first_group, group_depth = True, depth
            elif depth == 3 and group_depth == 2 and local == 'mergeCell':
                r1, c1, r2, c2 = module.bounds(attrs.get('ref', ''))
                if r2 > r1 or (all_merges and c2 > c1):
                    raise Selected
        def end(name):
            nonlocal depth, group_depth
            if depth == group_depth:
                group_depth = None
            depth -= 1
        parser.StartElementHandler = start
        parser.EndElementHandler = end
        try:
            while True:
                block = source.read(65536)
                parser.Parse(block, not block)
                if not block:
                    return 0, 0
        except Selected:
            source.seek(0)
            return original(source, output, all_merges)
    module.transform_sheet_stream = transformed


def raw_row_upper_bound(module, raw):
    # Fixture-specific oracle supplies original UTF-8 rows outside the timer.
    # This does NOT implement a deployable source-offset extractor or test arbitrary XML.
    rows = {int(match[1]): match[0] for match in re.finditer(rb'<row r="([0-9]+)">.*?</row>', raw)}
    module.synthetic_rows = rows
    source = inspect.getsource(module.RowSpoolBuilder.store_row)
    source = '\n'.join(line[4:] for line in source.splitlines())
    source = source.replace('raw = serialize_xml(node)', 'raw = synthetic_rows[number]')
    exec(compile(source, '<scratch-known-raw-rows-upper-bound>', 'exec'), module.__dict__)
    module.RowSpoolBuilder.store_row = module.store_row


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=4000)
    parser.add_argument('--report', type=Path, default=HERE/'path-results.json')
    args = parser.parse_args()
    if args.report.exists() or not 1000 <= args.rows <= 12000:
        raise ValueError('Preserve evidence and bounded workload')
    raw, selected, filled = fixture(args.rows)
    no_match = re.sub(rb'<mergeCells.*?</mergeCells>', b'', raw)
    variants = {'baseline': load(), 'skip_empty_attribute_maps': load(),
                'preflight': load(), 'raw_row_upper_bound': load()}
    skip_empty_attribute_maps(variants['skip_empty_attribute_maps'])
    preflight(variants['preflight'])
    raw_row_upper_bound(variants['raw_row_upper_bound'], raw)
    edges = []
    for name in ('skip_empty_attribute_maps', 'preflight'):
        for label, data in cases().items():
            expected = None
            for module in (variants['baseline'], variants[name]):
                try:
                    result = run(module, data)
                    identity = ('ok', result['counts'], result['output_sha256'])
                except Exception as error:
                    identity = (type(error).__name__, str(error))
                if expected is None:
                    expected = identity
                else:
                    assert identity == expected, (name, label, identity, expected)
            edges.append(dict(candidate=name, case=label, baseline_identical=True))
    observations = {}
    for scenario, data in [('matched', raw), ('no_selected_merges', no_match)]:
        names = list(variants) if scenario == 'matched' else ['baseline', 'preflight']
        observations[scenario] = {name: [] for name in names}
        for index in range(5):
            order = names[index % len(names):]+names[:index % len(names)]
            for name in order:
                value = run(variants[name], data)
                observations[scenario][name].append(value)
        identity = observations[scenario]['baseline'][0]
        for values in observations[scenario].values():
            for value in values:
                assert value['counts'] == identity['counts']
                assert value['output_sha256'] == identity['output_sha256']
    medians = {scenario: {name: statistics.median(v['seconds'] for v in values)
                         for name, values in variants.items()} for scenario, variants in observations.items()}
    profiler = cProfile.Profile()
    profiler.enable()
    profiled = run(variants['baseline'], no_match)
    profiler.disable()
    stats = pstats.Stats(profiler)
    top = []
    for (file, line, name), (primitive, total, own, cumulative, callers) in stats.stats.items():
        top.append(dict(file=Path(file).name, line=line, function=name, total_calls=total,
                        self_seconds=own, cumulative_seconds=cumulative))
    record = dict(engine_sha256=hashlib.sha256(ENGINE.read_bytes()).hexdigest(), rows=args.rows,
                  columns=22, original_cells=args.rows*18+selected, active_merges_max=4,
                  input_bytes=len(raw), observations=observations, median_seconds=medians,
                  profiler_no_match=profiled,
                  profiler_no_match_top=sorted(top, key=lambda x:x['cumulative_seconds'], reverse=True)[:25],
                  edge_equivalence_checks=edges,
                  limitations=['Synthetic transform timings exclude ZIP inflation/compression.',
                               'Raw-row lookup fixture preparation and memory are excluded: this is an upper-bound counterfactual, not an implementation or promised gain.',
                               'Preflight prototype needs seekable input; its extra real ZIP decompression cost and broad malformed-XML behavior are not validated.'])
    args.report.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'medians': medians, 'edges': len(edges)}, indent=2))


if __name__ == '__main__':
    main()

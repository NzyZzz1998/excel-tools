"""Bounded synthetic profiling; reads no business workbook and edits no application."""
import argparse
import cProfile
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import pstats
import statistics
import subprocess
import sys
import tempfile
import time

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
ENGINE = ROOT / 'excel_unmerge_fill.py'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def load():
    spec = importlib.util.spec_from_file_location('profiled_engine', ENGINE)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def fixture(rows, height=1000):
    chunks = [f'<worksheet xmlns="{NS}"><dimension ref="A1"/><sheetData>']
    merges = []
    for first in range(1, rows+1, height):
        last = min(first+height-1, rows)
        if last > first:
            merges.extend(f'{chr(65+c)}{first}:{chr(65+c)}{last}' for c in range(4))
    for row in range(1, rows+1):
        chunks.append(f'<row r="{row}">')
        if (row-1) % height == 0:
            for col in range(4):
                chunks.append(f'<c r="{chr(65+col)}{row}" t="inlineStr"><is><t>synthetic-anchor-{row}-{col}</t></is></c>')
        for col in range(4, 22):
            addr = f'{chr(65+col)}{row}'
            if col < 7:
                chunks.append(f'<c r="{addr}" t="inlineStr"><is><t>synthetic-text-{row}-{col}</t></is></c>')
            else:
                chunks.append(f'<c r="{addr}" t="n"><v>{row*100+col}</v></c>')
        chunks.append('</row>')
    chunks.append('</sheetData><mergeCells count="'+str(len(merges))+'">')
    chunks.extend('<mergeCell ref="'+ref+'"/>' for ref in merges)
    chunks.append('</mergeCells></worksheet>')
    return ''.join(chunks).encode('utf-8'), len(merges), rows*4-len(merges)


def run(module, raw):
    with tempfile.TemporaryFile(mode='w+b') as output:
        started = time.perf_counter()
        counts = module.transform_sheet_stream(io.BytesIO(raw), output)
        elapsed = time.perf_counter()-started
        output.seek(0)
        digest = hashlib.sha256()
        size = 0
        for block in iter(lambda: output.read(1024*1024), b''):
            digest.update(block)
            size += len(block)
        return dict(seconds=elapsed, counts=counts, output_sha256=digest.hexdigest(), output_bytes=size)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=4000)
    parser.add_argument('--repeats', type=int, default=3)
    parser.add_argument('--report', type=Path, default=HERE/'baseline.json')
    args = parser.parse_args()
    if not (1000 <= args.rows <= 12000) or not 1 <= args.repeats <= 5:
        raise ValueError('Keep review probes bounded')
    if args.report.exists():
        raise FileExistsError('Do not overwrite evidence')
    raw, merges, fills = fixture(args.rows)
    module = load()
    warmup = run(module, fixture(1000)[0])
    observations = [run(module, raw) for _ in range(args.repeats)]
    profiler = cProfile.Profile()
    profiler.enable()
    profiled = run(module, raw)
    profiler.disable()
    stats = pstats.Stats(profiler)
    functions = []
    for (file, line, name), (primitive, total, own, cumulative, callers) in stats.stats.items():
        functions.append(dict(file=Path(file).name, line=line, function=name, primitive_calls=primitive,
                              total_calls=total, self_seconds=own, cumulative_seconds=cumulative))
    for item in observations+[profiled]:
        assert item['counts'] == (merges, fills)
        assert item['output_sha256'] == observations[0]['output_sha256']
    record = dict(head=subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip(),
                  python=sys.version, engine_sha256=hashlib.sha256(ENGINE.read_bytes()).hexdigest(),
                  fixture=dict(rows=args.rows, columns=22, original_cells=args.rows*18+merges,
                               active_merges_max=4, merges=merges, fills=fills, input_bytes=len(raw),
                               business_data=False), warmup=warmup, unprofiled=observations,
                  unprofiled_median_seconds=statistics.median(x['seconds'] for x in observations),
                  profiled=profiled, profiler_total_seconds=stats.total_tt,
                  total_function_calls=stats.total_calls,
                  top_cumulative=sorted(functions, key=lambda x:x['cumulative_seconds'], reverse=True)[:45],
                  top_self=sorted(functions, key=lambda x:x['self_seconds'], reverse=True)[:35],
                  scope='Direct forced streaming transform with bytes input and disk-backed temporary output; excludes fixture construction, ZIP inflation/compression and output hashing from transform timing. Synthetic narrow rows with four active vertical merges; not actual monthly-workbook timing.')
    args.report.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(dict(report=str(args.report), median=record['unprofiled_median_seconds'],
                          profiled=profiled['seconds'], top=record['top_cumulative'][:16]), indent=2))


if __name__ == '__main__':
    main()

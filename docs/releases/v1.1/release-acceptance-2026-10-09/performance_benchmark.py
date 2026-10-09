"""Scale a local daily workbook; publish timings/hashes only, never business values."""
import argparse
import copy
import gc
import hashlib
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

EVIDENCE = Path(__file__).resolve().parent
REPO = EVIDENCE.parents[3]
SOURCE = REPO / 'testfile/期间缺货Top20/每日缺货Top50 (1).xlsx'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def point(ref):
    letters, row = re.fullmatch(r'([A-Z]+)([0-9]+)', ref).groups()
    col = 0
    for letter in letters:
        col = col * 26 + ord(letter) - 64
    return int(row), col


def offset(ref, rows):
    return re.sub(r'[0-9]+', lambda m: str(int(m.group()) + rows), ref)


def semantic(node):
    return (node.tag, tuple(sorted(node.attrib.items())), node.text,
            tuple((semantic(child), child.tail) for child in node))


def cell_value(node):
    return (tuple(sorted((k, v) for k, v in node.attrib.items() if k != 'r')),
            tuple(semantic(child) for child in node))


def expected_template(raw):
    root = ET.fromstring(raw)
    cells = {node.get('r'): cell_value(node) for node in root.findall('.//' + NS + 'c')}
    selected = filled = 0
    remaining = []
    for merge in root.findall('.//' + NS + 'mergeCell'):
        first, last = merge.get('ref').split(':')
        r1, c1 = point(first)
        r2, c2 = point(last)
        if r1 == r2:
            remaining.append(merge.get('ref'))
            continue
        assert c1 == c2, 'This performance cohort contains vertical and horizontal merges only.'
        selected += 1
        for row in range(r1 + 1, r2 + 1):
            cells[offset(first, row - r1)] = cells[first]
            filled += 1
    return cells, selected, filled, remaining


def expanded(source, dest, factor):
    with ZipFile(source) as src, ZipFile(dest, 'w') as out:
        out.comment = src.comment
        for info in src.infolist():
            raw = src.read(info)
            if info.filename == 'xl/worksheets/sheet1.xml' and factor > 1:
                root = ET.fromstring(raw)
                data = root.find(NS + 'sheetData')
                rows = list(data)
                last = max(int(row.get('r')) for row in rows)
                merges = root.find(NS + 'mergeCells')
                refs = [item.get('ref') for item in merges]
                data.clear()
                merges.clear()
                for block in range(factor):
                    for row in rows:
                        clone = copy.deepcopy(row)
                        clone.set('r', str(int(row.get('r')) + last * block))
                        for cell in clone.findall(NS + 'c'):
                            cell.set('r', offset(cell.get('r'), last * block))
                        data.append(clone)
                    for ref in refs:
                        ET.SubElement(merges, NS + 'mergeCell', {'ref': offset(ref, last * block)})
                merges.set('count', str(len(merges)))
                root.find(NS + 'dimension').set('ref', 'A1:AA' + str(last * factor))
                raw = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            out.writestr(copy.copy(info), raw)


def verify(source, output, factor, templates):
    checked = 0
    with ZipFile(source) as src, ZipFile(output) as out:
        assert src.namelist() == out.namelist() and src.comment == out.comment
        for name in src.namelist():
            if name not in templates:
                assert src.read(name) == out.read(name)
        for name, (cells, regions, filled, remaining, span) in templates.items():
            repeat = factor if name.endswith('sheet1.xml') else 1
            count = 0
            previous = (0, 0)
            actual_merges = []
            with out.open(name) as stream:
                for event, node in ET.iterparse(stream, events=('end',)):
                    if node.tag == NS + 'c':
                        row, col = point(node.get('r'))
                        # 取模前限制实际行范围，防止遗漏格被移到额外重复块而仍通过计数。
                        assert 1 <= row <= span * repeat
                        assert (row, col) > previous
                        previous = (row, col)
                        ref = offset(node.get('r'), -((row - 1) // span) * span)
                        assert cell_value(node) == cells[ref]
                        count += 1
                        node.clear()
                    elif node.tag == NS + 'mergeCell':
                        actual_merges.append(node.get('ref'))
                    elif node.tag == NS + 'row':
                        node.clear()
            assert count == len(cells) * repeat
            assert sorted(actual_merges) == sorted(offset(ref, block * span)
                for block in range(repeat) for ref in remaining)
            checked += count
    return checked


def worker(engine, source):
    spec = importlib.util.spec_from_file_location('scale_engine', engine)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    start = time.perf_counter()
    output, stats = mod.process_file(source)
    elapsed = time.perf_counter() - start
    print(json.dumps({'seconds': elapsed, 'output': str(output),
                      'regions': sum(item[1] for item in stats),
                      'filled': sum(item[2] for item in stats)}, ensure_ascii=True), flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--worker', nargs=2)
    parser.add_argument('--revision', default='initial')
    parser.add_argument('--current-only', action='store_true')
    parser.add_argument('--factors', nargs='+', type=int, default=[1, 7, 31])
    args = parser.parse_args()
    if args.worker:
        worker(*args.worker)
        return
    import psutil
    work = REPO / 'testfile/性能验收_v1.1_2026-10-09' / args.revision
    work.mkdir(parents=True, exist_ok=False)
    report_path = EVIDENCE / ('performance-' + args.revision + '.json')
    assert not report_path.exists()
    baseline = work / 'v1.0-engine.py'
    baseline.write_bytes(subprocess.check_output(['git', 'show', 'v1.0.0:excel_unmerge_fill.py'], cwd=REPO))
    candidate = work / 'candidate-engine.py'
    if args.current_only:
        shutil.copyfile(REPO / 'excel_unmerge_fill.py', candidate)
    else:
        candidate.write_bytes(subprocess.check_output(['git', 'show', '96cfbfe:excel_unmerge_fill.py'], cwd=REPO))
    templates = {}
    with ZipFile(SOURCE) as src:
        for name in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'):
            raw = src.read(name)
            span = max(point(n.get('r'))[0] for n in ET.fromstring(raw).findall('.//' + NS + 'c'))
            templates[name] = (*expected_template(raw), span)
    available = psutil.virtual_memory().available
    memory_cap = min(3 * 1024 ** 3, available // 3)
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(),
              'input_sha256_before': sha(SOURCE), 'python': sys.version,
              'memory_cap_bytes': memory_cap, 'worker_timeout_seconds': 120,
              'baseline_sha256': sha(baseline), 'candidate_sha256': sha(candidate),
              'scope': 'One daily sheet repeated by rows; summary/filter sheet unchanged. This is a load simulation, not a real monthly report or SLA.',
              'results': []}
    engines = [('candidate', candidate)] if args.current_only else [('v1.0', baseline), ('candidate', candidate)]
    for factor in args.factors:
        fixture = work / ('daily-x' + str(factor) + '.xlsx')
        expanded(SOURCE, fixture, factor)
        gc.collect()
        for label, engine in engines:
            folder = work / (label + '-x' + str(factor))
            folder.mkdir()
            source = folder / fixture.name
            shutil.copy2(fixture, source)
            before = sha(source)
            start = time.perf_counter()
            proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--worker', str(engine), str(source)],
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8')
            monitor = psutil.Process(proc.pid)
            peak_rss = peak_private = 0
            stop_reason = None
            while proc.poll() is None:
                try:
                    info = monitor.memory_info()
                    peak_rss = max(peak_rss, info.rss)
                    peak_private = max(peak_private, getattr(info, 'private', 0))
                except psutil.NoSuchProcess:
                    break
                if peak_rss > memory_cap or time.perf_counter() - start > 120:
                    stop_reason = 'memory_cap' if peak_rss > memory_cap else 'timeout'
                    proc.kill()
                    break
                time.sleep(0.01)
            stdout, stderr = proc.communicate()
            lifetime = time.perf_counter() - start
            result = {'factor': factor, 'implementation': label, 'first_sheet_rows': 402 * factor,
                      'columns': 27, 'input_bytes': source.stat().st_size,
                      'peak_rss_mib': round(peak_rss / 1024 ** 2, 2),
                      'peak_private_mib': round(peak_private / 1024 ** 2, 2),
                      'process_lifetime_seconds': round(lifetime, 4), 'exit_code': proc.returncode,
                      'stop_reason': stop_reason, 'source_unchanged': sha(source) == before}
            if proc.returncode == 0:
                detail = json.loads(stdout)
                result.update(detail)
                expected_regions = templates['xl/worksheets/sheet1.xml'][1] * factor + templates['xl/worksheets/sheet2.xml'][1]
                expected_filled = templates['xl/worksheets/sheet1.xml'][2] * factor + templates['xl/worksheets/sheet2.xml'][2]
                assert detail['regions'] == expected_regions and detail['filled'] == expected_filled
                result['validated_cells'] = verify(source, Path(detail['output']), factor, templates)
                result['output_sha256'] = sha(Path(detail['output']))
                result['passed'] = result['source_unchanged']
            else:
                result['error_tail'] = stderr[-1500:]
                result['passed'] = False
            report['results'].append(result)
            report['input_sha256_after'] = sha(SOURCE)
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
            print(json.dumps({k: v for k, v in result.items() if k not in ('output', 'error_tail')}, ensure_ascii=True), flush=True)
    assert report['input_sha256_before'] == report['input_sha256_after']


if __name__ == '__main__':
    main()

"""Measure a frozen engine against isolated real-workbook copies; no cell values logged."""
import argparse
import hashlib
import importlib.util
import json
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

EVIDENCE = Path(__file__).resolve().parent
REPO = EVIDENCE.parents[3]
SOURCE = REPO / 'testfile/期间缺货Top20/日明细表 (2).xlsx'


def sha(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def worker(engine, source):
    sys.path.insert(0, str(Path(engine).parent))
    spec = importlib.util.spec_from_file_location('excel_unmerge_fill', engine)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    started = time.perf_counter()
    output, stats = mod.process_file(source)
    print(json.dumps({'seconds': time.perf_counter() - started, 'output': str(output),
                      'sheet_counts': [{'index': i, 'regions': x[1], 'filled': x[2]}
                                       for i, x in enumerate(stats, 1)]}, ensure_ascii=True), flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--worker', nargs=2)
    parser.add_argument('--revision', default='before-streaming')
    parser.add_argument('--source', type=Path, default=SOURCE)
    parser.add_argument('--baseline-only', action='store_true')
    parser.add_argument('--cap-gib', type=float, default=8)
    parser.add_argument('--timeout', type=int, default=240)
    args = parser.parse_args()
    if args.worker:
        worker(*args.worker)
        return
    import psutil
    actual_source = args.source.resolve()
    folder = REPO / 'testfile/大文件验收_v1.1_2026-10-09' / args.revision
    folder.mkdir(parents=True, exist_ok=False)
    engine = folder / 'engine.py'
    if args.baseline_only:
        engine.write_bytes(subprocess.check_output(['git', 'show', 'v1.0.0:excel_unmerge_fill.py'], cwd=REPO))
    else:
        shutil.copy2(REPO / 'excel_unmerge_fill.py', engine)
        for module in REPO.glob('excel_unmerge_*.py'):
            if module.name not in ('excel_unmerge_fill.py', 'excel_unmerge_gui.py'):
                shutil.copy2(module, folder / module.name)
    source = folder / actual_source.name
    shutil.copy2(actual_source, source)
    before = sha(actual_source)
    assert sha(source) == before
    cap = min(int(args.cap_gib * 1024 ** 3), psutil.virtual_memory().available // 3)
    begin = time.perf_counter()
    proc = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--worker', str(engine), str(source)],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8')
    monitor = psutil.Process(proc.pid)
    rss = private = 0
    stop = None
    while proc.poll() is None:
        try:
            info = monitor.memory_info()
            rss, private = max(rss, info.rss), max(private, getattr(info, 'private', 0))
        except psutil.NoSuchProcess:
            break
        if max(rss, private) > cap or time.perf_counter() - begin > args.timeout:
            stop = 'probe_memory_cap' if max(rss, private) > cap else 'probe_timeout'
            proc.kill()
            break
        time.sleep(.02)
    stdout, stderr = proc.communicate()
    record = dict(captured_utc=datetime.now(timezone.utc).isoformat(), revision=args.revision,
                  python=sys.version, engine_sha256=sha(engine), source_sha256=before,
                  source_unchanged=sha(actual_source) == before and sha(source) == before,
                  engine_dependencies={p.name: sha(p) for p in folder.glob('excel_unmerge_*.py')},
                  source_bytes=source.stat().st_size, memory_cap_bytes=cap, timeout_seconds=args.timeout,
                  process_seconds=time.perf_counter()-begin, peak_rss_mib=round(rss/1024**2, 2),
                  peak_private_mib=round(private/1024**2, 2), exit_code=proc.returncode, stop_reason=stop)
    if proc.returncode == 0:
        record.update(json.loads(stdout))
        output = Path(record['output'])
        record.update(output_sha256=sha(output), output_bytes=output.stat().st_size,
                      completed=True, content_validation='pending independent check')
    else:
        record.update(completed=False, diagnostic=stderr[-2000:])
    destination = EVIDENCE / ('measure-' + args.revision + '.json')
    assert not destination.exists()
    destination.write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(record, ensure_ascii=True), flush=True)


if __name__ == '__main__':
    main()

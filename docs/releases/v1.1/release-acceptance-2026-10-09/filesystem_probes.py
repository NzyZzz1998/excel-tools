"""Independent real-Windows filesystem checks using synthetic workbooks only."""
import ctypes
import argparse
import hashlib
import json
import sys
import tempfile
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

EVIDENCE = Path(__file__).resolve().parent
ROOT = EVIDENCE.parents[3]
sys.path[:0] = [str(ROOT), str(ROOT / 'tests')]
from excel_unmerge_fill import process_file
from test_excel_unmerge import c, load, makebook, sheet


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_one(path):
    output, stats = process_file(path)
    return str(output), stats


def check_output(path):
    with load(path) as book:
        assert book.active['A1'].value == 'synthetic group'
        assert book.active['A2'].value == 'synthetic group'
        assert book.active['B1'].value == 17


def fixture(folder):
    folder.mkdir()
    return makebook(folder / '中文 空格.xlsx', [('Synthetic', sheet({
        1: [c('A1', 'synthetic group'), c('B1', 17, 'number')]
    }, ['A1:A2']))])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--revision', default='')
    args = parser.parse_args()
    report = {'recorded_utc': datetime.now(timezone.utc).isoformat(),
              'engine_sha256': sha(ROOT / 'excel_unmerge_fill.py'), 'cases': [],
              'scope': 'Synthetic inputs, actual Windows locks/attributes and concurrent processes.'}
    result_path = EVIDENCE / ('filesystem-results' + ('-' + args.revision if args.revision else '') + '.json')
    if result_path.exists():
        raise RuntimeError('Preserve existing evidence')
    with tempfile.TemporaryDirectory(prefix='excel-tools-release-filesystem-') as name:
        work = Path(name)
        source = fixture(work / 'concurrent')
        before = sha(source)
        previous = source.with_name(source.stem + '_拆分填充.xlsx')
        previous.write_bytes(b'keep an existing result unchanged')
        previous_sha = sha(previous)
        with ProcessPoolExecutor(max_workers=4) as pool:
            outputs = list(pool.map(run_one, [str(source)] * 4))
        paths = [Path(result[0]) for result in outputs]
        assert len(set(paths)) == 4
        assert sha(source) == before and sha(previous) == previous_sha
        assert len(list(source.parent.glob('*.xlsx'))) == 6
        for path, stats in outputs:
            assert stats == [('Synthetic', 1, 1)]
            check_output(Path(path))
        report['cases'].append({'case': 'four_processes_same_input_with_existing_result',
                                'passed': True, 'distinct_new_outputs': 4,
                                'source_and_existing_result_unchanged': True})

        source = fixture(work / 'directory_collision')
        collision = source.with_name(source.stem + '_拆分填充.xlsx')
        collision.mkdir()
        before = sha(source)
        try:
            output, stats = process_file(source)
            assert output.name.endswith('_2.xlsx') and collision.is_dir()
            check_output(output)
            assert sha(source) == before
            directory_auto_suffix = True
        except PermissionError:
            assert collision.is_dir() and sha(source) == before
            assert list(source.parent.glob('*.xlsx')) == [source, collision]
            directory_auto_suffix = False
        report['cases'].append({'case': 'existing_directory_at_first_output_name',
                                'passed': True, 'directory_preserved': True,
                                'source_preserved': True, 'auto_suffix': directory_auto_suffix,
                                'observation': 'Windows reports a folder collision as PermissionError; safely rejected but not automatically suffixed.' if not directory_auto_suffix else 'Collision skipped.'})

        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
        kernel.GetFileAttributesW.restype = ctypes.c_uint32
        kernel.SetFileAttributesW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
        kernel.SetFileAttributesW.restype = ctypes.c_int
        source = fixture(work / 'read_only_source')
        before = sha(source)
        attributes = kernel.GetFileAttributesW(str(source))
        assert attributes != 0xFFFFFFFF
        try:
            assert kernel.SetFileAttributesW(str(source), attributes | 1)
            output, stats = process_file(source)
            check_output(output)
            assert sha(source) == before
            assert kernel.GetFileAttributesW(str(source)) & 1
        finally:
            assert kernel.SetFileAttributesW(str(source), attributes)
        report['cases'].append({'case': 'read_only_source_in_writable_directory',
                                'passed': True, 'source_unchanged': True,
                                'attributes_restored': True})

        source = fixture(work / 'locked_source')
        before = sha(source)
        kernel.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32,
                                      ctypes.c_void_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p]
        kernel.CreateFileW.restype = ctypes.c_void_p
        kernel.CloseHandle.argtypes = [ctypes.c_void_p]
        kernel.CloseHandle.restype = ctypes.c_int
        handle = kernel.CreateFileW(str(source), 0x80000000, 0, None, 3, 0x80, None)
        assert handle != ctypes.c_void_p(-1).value
        try:
            try:
                process_file(source)
            except PermissionError:
                pass
            else:
                raise AssertionError('Exclusive source lock should reject opening')
            assert list(source.parent.glob('*.xlsx')) == [source]
        finally:
            assert kernel.CloseHandle(handle)
        assert sha(source) == before
        output, stats = process_file(source)
        check_output(output)
        assert sha(source) == before
        report['cases'].append({'case': 'real_exclusive_source_lock_then_retry',
                                'passed': True, 'no_output_while_locked': True,
                                'lock_released_and_retry_succeeded': True, 'source_unchanged': True})
    report['temporary_directory_removed'] = not work.exists()
    report['passed'] = all(case['passed'] for case in report['cases'])
    result_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=True, indent=2))


if __name__ == '__main__':
    main()

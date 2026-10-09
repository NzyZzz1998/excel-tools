"""Independent, synthetic v1.1.1 publication probes. Run only after source freeze.

The parent owns every subprocess handle it may terminate. All workbook fixtures
and abandoned .part files remain under this repository's ignored testfile tree.
"""
import argparse
import ctypes
import errno
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import traceback
import uuid
from unittest import mock
from xml.etree import ElementTree as ET
from zipfile import ZIP_STORED, ZipFile

REPO = Path(__file__).resolve().parents[4]
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def load_engine(path):
    spec = importlib.util.spec_from_file_location('frozen_engine', path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def make_book(path):
    """Self-contained valid tiny OOXML package; no application/test fixture helper."""
    types = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="xml" ContentType="application/xml"/>'
             '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
             '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
             '</Types>')
    worksheet = ('<worksheet xmlns="' + NS + '"><dimension ref="A1:B3"/><sheetData>'
                 '<row r="1"><c r="A1" t="inlineStr"><is><t>SYNTHETIC_GROUP</t></is></c>'
                 '<c r="B1" t="inlineStr"><is><t>SYNTHETIC_UNTOUCHED</t></is></c></row>'
                 '<row r="2"><c r="B2"><v>7</v></c></row></sheetData>'
                 '<mergeCells count="1"><mergeCell ref="A1:A3"/></mergeCells></worksheet>')
    parts = {
        '[Content_Types].xml': types,
        '_rels/.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="wb" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': '<workbook xmlns="' + NS + '" xmlns:r="' + REL + '"><sheets><sheet name="SYNTHETIC" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="' + PKG + '"><Relationship Id="rId1" Type="' + REL + '/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': worksheet,
    }
    with ZipFile(path, 'x', ZIP_STORED) as archive:
        archive.comment = b'SYNTHETIC_PUBLICATION_PROBE'
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def verify_result(source, output):
    with ZipFile(source) as before, ZipFile(output) as after:
        assert after.testzip() is None, 'ZIP CRC failure'
        assert before.namelist() == after.namelist(), 'ZIP members/order changed'
        assert before.comment == after.comment, 'ZIP comment changed'
        for name in before.namelist():
            if name != 'xl/worksheets/sheet1.xml':
                assert before.read(name) == after.read(name), 'Non-target part differs'
        sheet = ET.fromstring(after.read('xl/worksheets/sheet1.xml'))
        cells = {cell.get('r'): cell for cell in sheet.iter('{' + NS + '}c')}
        assert list(cells) == ['A1', 'B1', 'A2', 'B2', 'A3']
        for reference in ('A1', 'A2', 'A3'):
            cell = cells[reference]
            assert cell.get('t') == 'inlineStr'
            assert cell.find('{' + NS + '}is/{' + NS + '}t').text == 'SYNTHETIC_GROUP'
        assert cells['B1'].find('{' + NS + '}is/{' + NS + '}t').text == 'SYNTHETIC_UNTOUCHED'
        assert cells['B2'].find('{' + NS + '}v').text == '7'
        assert not list(sheet.iter('{' + NS + '}mergeCell'))
    return {'valid_zip_and_crc': True, 'expected_cells': 5, 'non_target_parts_unchanged': True,
            'sha256': digest(output), 'bytes': output.stat().st_size}


def unique_part(folder, previous_names):
    parts = [p for p in folder.iterdir() if p.name not in previous_names and p.suffix == '.part']
    assert len(parts) == 1, 'Expected exactly one newly owned .part file'
    return parts[0]


def child_writer(engine_path, source, marker, nonce):
    engine = load_engine(engine_path)
    source = Path(source)
    previous_names = {p.name for p in source.parent.iterdir()}
    original_open = engine.ZipFile.open

    class PausedWriter:
        def __init__(self, archive, destination):
            self.archive = archive
            self.destination = destination

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return self.destination.__exit__(*args)

        def write(self, block):
            self.destination.write(block[:64])
            self.archive.fp.flush()
            part = unique_part(source.parent, previous_names)
            checkpoint = {'nonce': nonce, 'pid': os.getpid(), 'part': str(part),
                          'part_bytes': part.stat().st_size, 'engine_sha256': digest(engine_path)}
            temporary_marker = Path(str(marker) + '.writing')
            temporary_marker.write_text(json.dumps(checkpoint), encoding='utf-8')
            temporary_marker.replace(marker)
            # Parent keeps stdin open; it terminates this exact Popen process handle.
            sys.stdin.buffer.read(1)
            raise RuntimeError('Owned hard-stop probe released without termination')

    def pause_inside_zip(archive, name, mode='r', *args, **kwargs):
        destination = original_open(archive, name, mode, *args, **kwargs)
        return PausedWriter(archive, destination) if mode == 'w' else destination

    engine.ZipFile.open = pause_inside_zip
    engine.process_file(source)


def hard_stop(engine, frozen, folder):
    folder.mkdir()
    source = make_book(folder / 'synthetic.xlsx')
    prior, _ = engine.process_file(source)
    prior_verification = verify_result(source, prior)
    before = {p.name: digest(p) for p in (source, prior)}
    existing_finals = {p.name for p in folder.glob('*.xlsx')}
    marker, nonce = folder / 'checkpoint.json', uuid.uuid4().hex
    # Windows venv python.exe may be a redirector whose Popen PID differs from
    # the actual interpreter. The base interpreter has no external dependencies
    # here and gives us one exact process handle matching the checkpoint PID.
    child_python = getattr(sys, '_base_executable', sys.executable)
    command = [child_python, str(Path(__file__).resolve()), '--child', str(frozen),
               str(source), str(marker), nonce]
    proc = subprocess.Popen(command, cwd=REPO, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    checkpoint = None
    try:
        deadline = time.monotonic() + 20
        while proc.poll() is None and not marker.exists() and time.monotonic() < deadline:
            time.sleep(0.02)
        assert marker.exists() and proc.poll() is None, 'Child did not reach live ZIP-write checkpoint'
        checkpoint = json.loads(marker.read_text(encoding='utf-8'))
        part = Path(checkpoint['part']).resolve()
        assert checkpoint['nonce'] == nonce, 'Checkpoint nonce does not belong to this probe'
        assert checkpoint['pid'] == proc.pid, 'Child interpreter differs from the owned Popen process'
        assert checkpoint['engine_sha256'] == digest(frozen), 'Child did not use frozen source'
        assert part.parent == folder.resolve() and part.suffix == '.part' and part.exists(), 'Unexpected temporary path'
        assert {p.name for p in folder.glob('*.xlsx')} == existing_finals, 'Final name visible during write'
        proc.terminate()  # Exact owned process object, never an enumerated PID.
        stdout, stderr = proc.communicate(timeout=5)
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=5)
    assert {p.name for p in folder.glob('*.xlsx')} == existing_finals, 'Hard stop left a new final workbook'
    assert before == {p.name: digest(p) for p in (source, prior)}
    abandoned = {p.name: digest(p) for p in folder.glob('*.part')}
    retry, stats = engine.process_file(source)
    assert retry.name == 'synthetic_拆分填充_2.xlsx', 'Abandoned .part consumed a final result serial'
    verification = verify_result(source, retry)
    assert before == {p.name: digest(p) for p in (source, prior)}
    assert {p.name for p in folder.glob('*.xlsx')} == existing_finals | {retry.name}
    assert all((folder / name).exists() and digest(folder / name) == value for name, value in abandoned.items()), 'Retry changed an earlier abandoned temporary file'
    return {'passed': True, 'child_python': child_python, 'owned_pid': proc.pid, 'owned_exit_code': proc.returncode,
            'checkpoint': checkpoint, 'no_final_name_during_write_or_after_hard_stop': True,
            'source_and_prior_sha_unchanged': True, 'prior_result': prior_verification,
            'abandoned_parts': abandoned, 'retry_name': retry.name, 'retry_stats': stats,
            'retry_result': verification, 'only_one_new_valid_final_after_retry': True,
            'child_stderr': stderr.decode('utf-8', errors='replace')}


def lock_plus_enospc(engine, folder):
    assert sys.platform == 'win32', 'This acceptance requires a genuine Windows read lock'
    folder.mkdir()
    source = make_book(folder / 'locked.xlsx')
    prior, _ = engine.process_file(source)
    verify_result(source, prior)
    before = {p.name: digest(p) for p in (source, prior)}
    previous_names = {p.name for p in folder.iterdir()}
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_void_p,
                                  ctypes.c_ulong, ctypes.c_ulong, ctypes.c_void_p]
    kernel.CreateFileW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle, writes, failure, part = None, 0, None, None
    original_open = engine.ZipFile.open

    def inject(archive, name, mode='r', *args, **kwargs):
        nonlocal handle, writes, part
        if mode == 'w':
            writes += 1
            if writes == 2:
                part = unique_part(folder, previous_names)
                handle = kernel.CreateFileW(str(part), 0x80000000, 0x1 | 0x2, None, 3, 0x80, None)
                if handle == ctypes.c_void_p(-1).value:
                    raise ctypes.WinError(ctypes.get_last_error())
                raise OSError(errno.ENOSPC, 'SYNTHETIC ENOSPC during ZIP output')
        return original_open(archive, name, mode, *args, **kwargs)

    try:
        with mock.patch.object(engine.ZipFile, 'open', new=inject):
            try:
                engine.process_file(source)
            except Exception as error:
                failure = error
    finally:
        if handle and handle != ctypes.c_void_p(-1).value:
            assert kernel.CloseHandle(handle), 'Could not release owned Win32 read handle'
    assert failure is not None and part is not None and writes == 2
    assert type(failure) is OSError and failure.errno == errno.ENOSPC, 'Cleanup replaced original write error'
    assert Path(failure.partial_path).resolve() == part.resolve(), 'Missing/wrong retained partial path'
    assert isinstance(failure.cleanup_error, PermissionError), 'Missing actual cleanup failure details'
    assert str(part) in str(failure), 'Partial path is not visible in the propagated error text'
    assert 'SYNTHETIC ENOSPC' in str(failure), 'Original reason no longer visible'
    assert part.exists(), 'Cleanup-failure evidence unexpectedly absent'
    assert {p.name for p in folder.glob('*.xlsx')} == {source.name, prior.name}, 'Failure left a final .xlsx'
    assert before == {p.name: digest(p) for p in (source, prior)}
    partial_before = digest(part)
    retry, stats = engine.process_file(source)
    verification = verify_result(source, retry)
    assert retry.name == 'locked_拆分填充_2.xlsx'
    assert digest(part) == partial_before, 'Retry silently modified earlier retained partial file'
    assert before == {p.name: digest(p) for p in (source, prior)}
    return {'passed': True, 'write_calls': writes, 'original_error_type': type(failure).__name__,
            'original_errno': failure.errno, 'visible_error': str(failure),
            'partial_path': str(part), 'cleanup_error_type': type(failure.cleanup_error).__name__,
            'cleanup_winerror': getattr(failure.cleanup_error, 'winerror', None),
            'no_new_final_on_failure': True, 'owned_read_handle_closed': True,
            'source_and_prior_sha_unchanged': True, 'retry_stats': stats, 'retry_result': verification,
            'retained_part_unchanged_after_retry': True}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--engine', required=True, type=Path)
    parser.add_argument('--expected-engine-sha', required=True)
    parser.add_argument('--report', required=True, type=Path)
    args = parser.parse_args()
    assert not args.report.exists(), 'Preserve earlier evidence; choose a new report path'
    live = args.engine.resolve()
    expected = args.expected_engine_sha.lower()
    assert digest(live) == expected, 'Source is not the explicitly frozen candidate'
    allowed = (REPO / 'testfile').resolve()
    assert allowed.is_dir()
    work = Path(tempfile.mkdtemp(prefix='v1.1.1-publication-', dir=allowed)).resolve()
    assert work.parent == allowed
    frozen = work / 'frozen_engine.py'
    frozen.write_bytes(live.read_bytes())
    assert digest(frozen) == expected
    engine = load_engine(frozen)
    report = {'engine_sha256': expected, 'live_engine_path': str(live), 'frozen_engine_path': str(frozen),
              'script_sha256': digest(Path(__file__)), 'work_directory': str(work), 'python': sys.version,
              'scope': 'Owned hidden subprocess hard stop and synthetic genuine Windows read lock. No business workbook or GUI/EXE interaction.',
              'started_utc': time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())}
    try:
        report['hard_stop'] = hard_stop(engine, frozen, work / 'hard-stop')
        report['lock_plus_enospc'] = lock_plus_enospc(engine, work / 'lock-failure')
        report['live_source_unchanged'] = digest(live) == expected
        assert report['live_source_unchanged'], 'Live source changed during frozen acceptance'
        report['passed'] = True
    except BaseException as error:
        report['passed'] = False
        report['failure'] = {'type': type(error).__name__, 'message': str(error),
                             'traceback': traceback.format_exc()}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'passed': report['passed'], 'report': str(args.report),
                      'failure': report.get('failure')}, ensure_ascii=True))
    return 0 if report['passed'] else 1


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--child':
        child_writer(*sys.argv[2:])
    else:
        raise SystemExit(main())

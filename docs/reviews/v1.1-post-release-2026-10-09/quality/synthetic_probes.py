"""Small post-release synthetic probes; all workbook/XML fixtures stay ignored."""
import ctypes
import errno
import hashlib
import importlib.util
import json
import pathlib
import subprocess
import sys
import tempfile
import time
from io import BytesIO
from unittest import mock
from zipfile import BadZipFile, ZipFile

REPO = pathlib.Path(__file__).resolve().parents[4]
sys.path.insert(0, str(REPO / 'tests'))
import test_excel_unmerge as fixture


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def lock_during_failure(root):
    source = fixture.makebook(root / 'write-lock.xlsx', [
        ('SYNTHETIC', fixture.sheet({1: [fixture.c('A1', 'SYNTHETIC')]}, ['A1:A2']))])
    old = source.with_name('write-lock_拆分填充.xlsx')
    old.write_bytes(b'SYNTHETIC PREEXISTING RESULT')
    before = {p.name: digest(p) for p in (source, old)}
    new = source.with_name('write-lock_拆分填充_2.xlsx')
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.CreateFileW.argtypes = [ctypes.c_wchar_p, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_void_p,
                                  ctypes.c_ulong, ctypes.c_ulong, ctypes.c_void_p]
    kernel.CreateFileW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle, writes, captured = None, 0, None
    original_open = fixture.mod.ZipFile.open

    def injected(archive, name, mode='r', *args, **kwargs):
        nonlocal handle, writes
        if mode == 'w':
            writes += 1
            if writes == 2:
                # Real Win32 read handle permits read/write, but denies file deletion.
                handle = kernel.CreateFileW(str(new), 0x80000000, 0x1 | 0x2, None, 3, 0, None)
                if handle == ctypes.c_void_p(-1).value:
                    raise ctypes.WinError(ctypes.get_last_error())
                raise OSError(errno.ENOSPC, 'SYNTHETIC disk full during output write')
        return original_open(archive, name, mode, *args, **kwargs)

    try:
        with mock.patch.object(fixture.mod.ZipFile, 'open', new=injected):
            try:
                fixture.mod.process_file(source)
            except BaseException as error:
                captured = {'type': type(error).__name__, 'errno': getattr(error, 'errno', None),
                            'winerror': getattr(error, 'winerror', None),
                            'context_type': type(error.__context__).__name__ if error.__context__ else None,
                            'context_errno': getattr(error.__context__, 'errno', None)}
    finally:
        if handle and handle != ctypes.c_void_p(-1).value:
            kernel.CloseHandle(handle)
    is_workbook = False
    parts = []
    if new.exists():
        try:
            with ZipFile(new) as archive:
                parts = archive.namelist()
                is_workbook = '[Content_Types].xml' in parts and 'xl/workbook.xml' in parts
        except BadZipFile:
            pass
    return {'injected_primary_error': 'OSError ENOSPC', 'write_calls': writes,
            'propagated_exception': captured, 'partial_final_name_retained': new.exists(),
            'partial_bytes': new.stat().st_size if new.exists() else None,
            'partial_is_workbook': is_workbook, 'partial_zip_parts': parts,
            'source_and_existing_result_unchanged': before == {p.name: digest(p) for p in (source, old)},
            'scope': 'Synthetic output-write exception plus genuine externally held Win32 read handle; uncommon combined failure, supporting atomic-publication finding only.'}


def formatting_probe(root):
    results = []
    for rows in (8000, 16000, 32000):
        for formatted in (False, True):
            whitespace = '\n        ' if formatted else ''
            cells = whitespace.join('<row r="%d"><c r="B%d"><v>%d</v></c></row>' % (r, r, r)
                                    for r in range(1, rows + 1))
            raw = ('<worksheet xmlns="' + fixture.NS + '"><sheetData>' + whitespace + cells
                   + whitespace + '</sheetData></worksheet>').encode()
            path = root / ('whitespace-%d-%s.xml' % (rows, formatted))
            path.write_bytes(raw)
            started = time.perf_counter()
            with tempfile.TemporaryFile(mode='w+b', dir=root) as spool:
                builder = fixture.mod.RowSpoolBuilder(spool)
                doc = builder.parseFile(BytesIO(raw))
                data = fixture.mod.children(doc.documentElement, 'sheetData')[0]
                text_chars = sum(len(n.data) for n in data.childNodes if n.nodeType == n.TEXT_NODE)
                result = {'rows': rows, 'formatted': formatted, 'source_bytes': len(raw),
                          'stored_rows': len(builder.rows), 'retained_sheet_data_nodes': len(data.childNodes),
                          'retained_interrow_text_chars': text_chars, 'seconds': time.perf_counter() - started}
                doc.unlink()
                builder.document.unlink()
                builder._parser = None
                results.append(result)
    return results


def main():
    root = pathlib.Path(tempfile.mkdtemp(prefix='post-release-quality-probes-', dir=REPO / 'testfile'))
    report = {'head': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=REPO, text=True).strip(),
              'engine_sha256': digest(REPO / 'excel_unmerge_fill.py'),
              'script_sha256': digest(pathlib.Path(__file__)), 'fixture_root': str(root),
              'fixtures': 'Synthetic only; no real user workbook reads, no EXE interaction.'}
    report['lock_during_write_failure'] = lock_during_failure(root)
    report['formatting_probe'] = formatting_probe(root)
    destination = pathlib.Path(__file__).with_name('synthetic-probe-results.json')
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=True))


if __name__ == '__main__':
    main()

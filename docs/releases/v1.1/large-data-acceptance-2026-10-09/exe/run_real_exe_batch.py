"""Identity-gated real EXE acceptance; business files stay below testfile.

Run prepare only after the release builder supplies both hashes. Each run launches
its own isolated EXE, uses native controls, and requires screenshot review; file
counts alone are never recorded as proof of the GUI's success summary.
"""
import argparse
import ctypes
import hashlib
import json
import shutil
import time
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

import native_gui_controller as c

WORKSPACE = Path(__file__).resolve().parents[5]
EVIDENCE = Path(__file__).resolve().parent
RUN_ROOT = WORKSPACE / 'testfile/大文件验收_v1.1_2026-10-09/exe-check'
SOURCE = WORKSPACE / 'testfile/期间缺货Top20'
IDENTITY = EVIDENCE / 'identity-and-inputs.json'


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest().upper()


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


def require(condition, message):
    if not condition:
        raise AssertionError(message)


def prepare(args):
    package = args.package.resolve()
    require(sha(package) == args.zip_sha.upper(), 'New release ZIP SHA mismatch')
    require(not IDENTITY.exists(), 'Do not overwrite an earlier identity manifest')
    app = RUN_ROOT / 'app'
    app.mkdir(parents=True, exist_ok=True)
    require(not any(app.iterdir()), 'Candidate app directory must be empty')
    with ZipFile(package) as archive:
        for item in archive.infolist():
            target = (app / item.filename).resolve()
            require(target.is_relative_to(app.resolve()), 'Archive path escapes isolated app directory')
        archive.extractall(app)
    exes = list(app.rglob('ExcelTools.exe'))
    require(len(exes) == 1, 'Expected exactly one packaged EXE')
    require(sha(exes[0]) == args.exe_sha.upper(), 'New release EXE SHA mismatch')
    sources = sorted(SOURCE.glob('*.xlsx'))
    require(len(sources) == 6, 'Acceptance requires exactly six original source files')
    items = []
    for path in sources:
        item = dict(source=str(path), name=path.name, bytes=path.stat().st_size,
                    sha256=sha(path), copies={})
        for mode in ('default', 'all'):
            folder = RUN_ROOT / mode
            folder.mkdir(parents=True, exist_ok=True)
            target = folder / path.name
            if not target.exists():
                shutil.copy2(path, target)
            require(sha(target) == item['sha256'], 'Copied input differs from source')
            item['copies'][mode] = str(target)
        items.append(item)
    for mode in ('default', 'all'):
        require(len(list((RUN_ROOT / mode).iterdir())) == 6,
                'Input directory must contain only six clean original copies')
    identity = dict(captured_utc=datetime.now(timezone.utc).isoformat(),
                    package=str(package), package_sha256=args.zip_sha.upper(),
                    exe=str(exes[0]), exe_sha256=args.exe_sha.upper(),
                    expected_title='Excel 数据处理工具 v1.1', inputs=items)
    write_json(IDENTITY, identity)
    print(json.dumps(dict(prepared=True, inputs=6, exe=str(exes[0])), ensure_ascii=False))


def control(name):
    """Existing validated Win32 geometry; fail closed if the layout changes."""
    main = c.main_window()
    rect = c.details(main)['rect']
    candidates = []
    for item in c.children(main):
        x, y, right, bottom = item['rect']
        width, height = right - x, bottom - y
        if name == 'choose' and (width, height) == (155, 27):
            candidates.append(item)
        elif name == 'option' and (width, height) == (239, 23):
            candidates.append(item)
        elif (name == 'start' and (width, height) == (87, 27)
              and x - rect[0] == 26 and y - rect[1] == 324):
            candidates.append(item)
    require(len(candidates) == 1, 'Native control layout mismatch: ' + name)
    return candidates[0]['hwnd']


def file_dialog(timeout=8):
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        matches = [item for item in c.windows()
                   if item['klass'] == '#32770' and item['title'] == '选择要处理的 Excel 文件']
        if matches:
            require(len(matches) == 1, 'Ambiguous owned file-dialog window')
            c.USER.SetWindowPos(matches[0]['hwnd'], None, -18000, -18000, 0, 0, 0x0015)
            return matches[0]
        time.sleep(.05)
    return None


def wait_dialog_closed():
    end = time.monotonic() + 8
    while any(item['klass'] == '#32770' for item in c.windows()) and time.monotonic() < end:
        time.sleep(.05)
    require(not any(item['klass'] == '#32770' for item in c.windows()), 'File dialog did not close')


def select_files(paths):
    c.guarded_mouse_click(control('choose'))
    dialog = file_dialog()
    require(dialog is not None, 'The owned file dialog did not open')
    items = c.children(dialog['hwnd'])
    edit = next(item for item in items if item['klass'] == 'Edit' and item['control_id'] == 1148)
    confirm = next(item for item in items if item['klass'] == 'Button' and item['control_id'] == 1)
    text = ' '.join('"' + str(path) + '"' for path in paths)
    buffer = ctypes.create_unicode_buffer(text)
    require(c.USER.SendMessageW(edit['hwnd'], 0x000C, 0, ctypes.cast(buffer, ctypes.c_void_p).value),
            'Cannot set owned file-dialog paths')
    readback = ctypes.create_unicode_buffer(max(8192, len(text) + 1))
    c.USER.SendMessageW(edit['hwnd'], 0x000D, len(readback), ctypes.cast(readback, ctypes.c_void_p).value)
    require(readback.value == text, 'File-dialog path verification failed')
    c.USER.PostMessageW(confirm['hwnd'], 0x00F5, 0, 0)
    wait_dialog_closed()
    return dict(dialog_hwnd=dialog['hwnd'], dialog_pid=dialog['pid'], filename_text_verified=True)


def confirm_idle():
    # Disabled ttk buttons ignore clicks while a worker is active. Opening and
    # cancelling this dialog proves that the GUI consumed its final done event.
    c.guarded_mouse_click(control('choose'))
    dialog = file_dialog(timeout=1)
    if dialog is None:
        return False
    cancel = next(item for item in c.children(dialog['hwnd'])
                  if item['klass'] == 'Button' and item['control_id'] == 2)
    c.USER.PostMessageW(cancel['hwnd'], 0x00F5, 0, 0)
    wait_dialog_closed()
    return True


def run(args):
    identity = json.loads(IDENTITY.read_text(encoding='utf-8'))
    require(identity['package_sha256'] == args.zip_sha.upper(), 'Requested ZIP differs from prepared release')
    require(identity['exe_sha256'] == args.exe_sha.upper(), 'Requested EXE differs from prepared release')
    require(sha(identity['package']) == args.zip_sha.upper(), 'ZIP changed since prepare')
    c.EXE = Path(identity['exe'])
    require(sha(c.EXE) == args.exe_sha.upper(), 'EXE changed since prepare')
    mode = args.mode
    no_match = mode.endswith('-no-match')
    rule = mode.split('-')[0]
    folder = RUN_ROOT / rule
    c.ROOT = EVIDENCE / mode
    require(not c.ROOT.exists(), 'Do not overwrite previous run evidence')
    c.ROOT.mkdir()
    original_paths = [folder / item['name'] for item in identity['inputs']]
    output_paths = sorted(folder.glob('*_拆分填充*.xlsx'))
    if no_match:
        require(len(output_paths) == 6, 'No-match run requires the six successful outputs')
        paths = output_paths
    else:
        require(not output_paths, 'Refusing to rerun previously successful originals')
        paths = original_paths
    require(len(paths) == 6, 'Exactly six selected files are required')
    before = {str(path): sha(path) for path in folder.glob('*.xlsx')}
    record = dict(captured_utc=datetime.now(timezone.utc).isoformat(), mode=mode,
                  zip_sha256=args.zip_sha.upper(), exe_sha256=args.exe_sha.upper(),
                  selected_files=list(map(str, paths)), timeout_seconds=args.timeout,
                  before_files=before, assertions_passed=False,
                  ui_conclusion='Await fresh screenshot inspection; not inferred from file counts')
    started = None
    try:
        record['startup'] = c.start()
        require(c.details(c.main_window())['title'] == identity['expected_title'], 'Actual window version mismatch')
        record['actual_title'] = c.details(c.main_window())['title']
        record['actual_exe_pid'] = c.details(c.main_window())['pid']
        record['selection'] = select_files(paths)
        if rule == 'all':
            c.guarded_mouse_click(control('option'))
        record['selected_capture'] = c.capture_fresh(c.ROOT / 'selected.png')
        c.guarded_mouse_click(control('start'))
        started = time.monotonic()
        previous = None
        stable_since = None
        next_notice = started + 30
        while time.monotonic() - started < args.timeout:
            require(c.our_pids(), 'Owned candidate exited before completion')
            outputs = sorted(folder.glob('*_拆分填充*.xlsx'))
            sizes = [(str(path), path.stat().st_size) for path in outputs]
            if len(outputs) == 6 and sizes == previous:
                stable_since = stable_since or time.monotonic()
                if time.monotonic() - stable_since >= 3 and confirm_idle():
                    record['ui_idle_confirmed_by_cancelled_file_dialog'] = True
                    break
            else:
                stable_since = None
            previous = sizes
            if time.monotonic() >= next_notice:
                print(json.dumps(dict(mode=mode, elapsed_seconds=round(time.monotonic() - started, 1),
                                      observed_output_files=len(outputs), ui_completion_confirmed=False)), flush=True)
                next_notice += 30
            time.sleep(.5)
        else:
            raise TimeoutError('Candidate batch did not confirm completion before configured timeout')
        record['observed_batch_seconds_including_idle_probe'] = round(time.monotonic() - started, 3)
        record['capture'] = c.capture_fresh(c.ROOT / 'result.png')
        after = {str(path): sha(path) for path in folder.glob('*.xlsx')}
        record['after_files'] = after
        for path, expected in before.items():
            require(after.get(path) == expected, 'An existing workbook changed')
        if no_match:
            require(before == after, 'No-match run generated or modified a workbook')
        else:
            require(len(after) == 12 and len(after.keys() - before.keys()) == 6,
                    'Expected six original copies and exactly six outputs')
        require(not list(folder.glob('*.tmp')), 'Temporary output remains after GUI completion')
        for item in identity['inputs']:
            require(sha(item['source']) == item['sha256'], 'Original source changed during acceptance')
        record['assertions_passed'] = True
    except Exception as error:
        record['failure_type'] = type(error).__name__
        record['failure_message'] = str(error)
        if c.our_pids():
            try:
                record['failure_capture'] = c.capture_fresh(c.ROOT / 'failure.png')
            except Exception as capture_error:
                record['capture_failure_type'] = type(capture_error).__name__
        raise
    finally:
        if c.our_pids():
            try:
                record['close'] = c.close_idle()
            except Exception as close_error:
                record['close_failure_type'] = type(close_error).__name__
        record['remaining_owned_pids'] = sorted(c.our_pids())
        write_json(c.ROOT / 'execution.json', record)
    require(not record['remaining_owned_pids'], 'Owned EXE has not exited; no forced termination attempted')
    print(json.dumps(dict(mode=mode, assertions_passed=True, screenshot=str(c.ROOT / 'result.png'),
                          seconds=record['observed_batch_seconds_including_idle_probe']), ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['prepare', 'run'])
    parser.add_argument('--package', type=Path)
    parser.add_argument('--zip-sha', required=True)
    parser.add_argument('--exe-sha', required=True)
    parser.add_argument('--mode', choices=['default', 'all', 'default-no-match', 'all-no-match'])
    parser.add_argument('--timeout', type=float, default=900)
    args = parser.parse_args()
    if args.action == 'prepare':
        parser.error('--package is required for prepare') if args.package is None else prepare(args)
    else:
        parser.error('--mode is required for run') if args.mode is None else run(args)


if __name__ == '__main__':
    main()

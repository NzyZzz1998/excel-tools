"""Observe existing GUI UX contracts using owned withdrawn Tk and tiny fixtures.

No EXE launch, Win32 enumeration, keyboard/mouse input, or user-window operation.
Application code is imported unchanged; all workbook fixtures live in temp.
"""
import hashlib
import json
import os
import sys
import tempfile
import time
import tkinter as tk
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

WORKSPACE = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(WORKSPACE))
import excel_unmerge_gui as gui


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def workbook(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.active['A1'] = 'synthetic group'
    book.active.merge_cells('A1:A3')
    book.save(path)
    book.close()
    return str(path)


def choose(app, paths):
    with patch.object(gui.filedialog, 'askopenfilenames', return_value=tuple(paths)):
        app.choose()


def finish(root, app):
    until = time.monotonic() + 10
    while app.running and time.monotonic() < until:
        root.update()
        time.sleep(.01)
    assert not app.running, 'Synthetic task did not finish'


def own_app():
    root = tk.Tk()
    root.withdraw()
    errors = []
    root.report_callback_exception = lambda kind, value, trace: errors.append(kind.__name__)
    app = gui.Application(root)
    root.update()
    return root, app, errors


def recovery_selection(folder):
    root, app, errors = own_app()
    try:
        failed = folder / 'old-failed.xlsx'
        failed.write_bytes(b'synthetic invalid ZIP')
        choose(app, (str(failed),))
        app.start()
        finish(root, app)
        assert app.failed_files == (str(failed),)
        workbook(failed)
        new = workbook(folder / 'new-selection.xlsx')
        choose(app, (new,))
        before = dict(visible_file_list=app.file_list.get('1.0', 'end-1c'),
                      selected_files=list(app.files), recovery_files=list(app.failed_files),
                      status=app.status.get(), retry_button=app.retry_button.cget('text'),
                      retry_enabled='disabled' not in app.retry_button.state())
        with patch.object(gui, 'process_file', wraps=gui.process_file) as process:
            app.retry_failed()
            finish(root, app)
            actual = [call.args[0] for call in process.call_args_list]
        after = dict(visible_file_list=app.file_list.get('1.0', 'end-1c'),
                     processed_files=actual, selected_files=list(app.files),
                     status=app.status.get(), successful=app.succeeded,
                     old_output_exists=failed.with_name('old-failed_拆分填充.xlsx').is_file(),
                     new_output_exists=(folder / 'new-selection_拆分填充.xlsx').exists(),
                     log_contains_old_target=str(failed) in app.results.get('1.0', 'end-1c'),
                     callback_errors=errors)
        assert actual == [str(failed)] and after['visible_file_list'] == new
        assert after['old_output_exists'] and not after['new_output_exists'] and errors == []
        return dict(before_retry=before, after_retry=after,
                    interpretation='Existing PRD contract confirmed: new selection and old recovery task coexist; the visible list remains the new selection during old-task recovery.')
    finally:
        root.destroy()


def multiple_directories(folder):
    root, app, errors = own_app()
    try:
        first = Path(workbook(folder / 'directory-a/same-name.xlsx'))
        second = Path(workbook(folder / 'directory-b/same-name.xlsx'))
        choose(app, (str(first), str(second)))
        app.start()
        finish(root, app)
        outputs = [path.with_name('same-name_拆分填充.xlsx') for path in (first, second)]
        assert all(path.is_file() for path in outputs)
        with patch.object(gui.os, 'startfile') as opened:
            app.open_results()
        actual = opened.call_args.args[0]
        log = app.results.get('1.0', 'end-1c')
        assert actual == str(second.parent.resolve()) and app.succeeded == 2 and errors == []
        return dict(successful=app.succeeded, output_paths=list(map(str, outputs)),
                    button_text=app.open_button.cget('text'), opened_directory=actual,
                    first_directory_not_opened=actual != str(first.parent.resolve()),
                    both_paths_are_in_text_log=all(str(path.resolve()) in log for path in outputs),
                    callback_errors=errors,
                    interpretation='All outputs are safe; the single result-directory button opens only the latest successful directory. Explorer launch was mocked.')
    finally:
        root.destroy()


def main():
    before = sha(WORKSPACE / 'excel_unmerge_gui.py')
    with tempfile.TemporaryDirectory(prefix='excel-post-release-ux-') as temporary:
        folder = Path(temporary)
        cases = dict(recovery_selection=recovery_selection(folder),
                     multiple_directories=multiple_directories(folder))
    report = dict(recorded_utc=datetime.now(timezone.utc).isoformat(), own_python_pid=os.getpid(),
                  gui_sha256=before, gui_unchanged=sha(WORKSPACE / 'excel_unmerge_gui.py') == before,
                  probe_sha256=sha(Path(__file__)), scope='Two existing UX contracts; only owned withdrawn Tk, temporary synthetic workbooks; no user EXE inspection or control.',
                  cases=cases, temporary_fixtures_cleaned=not Path(temporary).exists())
    Path(__file__).with_name('probe-results.json').write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(dict(gui_unchanged=report['gui_unchanged'], cases_confirmed=len(cases),
                         callback_errors=0, own_python_pid=os.getpid(),
                         temporary_fixtures_cleaned=report['temporary_fixtures_cleaned'])))


if __name__ == '__main__':
    main()

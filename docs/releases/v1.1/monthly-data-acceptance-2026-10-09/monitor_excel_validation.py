"""Bound an owned Excel validation instance; never infer ownership from a PID difference.

Standard-library Windows monitor. No Excel is launched by --self-test. Resource
stops are acceptance-probe interventions, not product crashes or application OOMs.
"""
import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid
from datetime import datetime, timezone

GIB = 1024 ** 3


class MemoryStatus(ctypes.Structure):
    _fields_ = [('length', wintypes.DWORD), ('load', wintypes.DWORD)] + [
        (name, ctypes.c_ulonglong) for name in
        ('total_physical', 'available_physical', 'total_pagefile',
         'available_pagefile', 'total_virtual', 'available_virtual', 'extended')]


class ProcessMemory(ctypes.Structure):
    _fields_ = [('cb', wintypes.DWORD), ('page_fault_count', wintypes.DWORD)] + [
        (name, ctypes.c_size_t) for name in
        ('peak_working_set', 'working_set', 'quota_peak_paged', 'quota_paged',
         'quota_peak_nonpaged', 'quota_nonpaged', 'pagefile', 'peak_pagefile', 'private')]


def filetime(value):
    return value.dwHighDateTime << 32 | value.dwLowDateTime


class Windows:
    def __init__(self):
        self.kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        self.psapi = ctypes.WinDLL('psapi', use_last_error=True)
        self.kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        self.kernel.OpenProcess.restype = wintypes.HANDLE
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
        self.kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        self.kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        self.kernel.WaitForSingleObject.restype = wintypes.DWORD
        self.kernel.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
        self.kernel.GlobalMemoryStatusEx.argtypes = [ctypes.POINTER(MemoryStatus)]
        self.psapi.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.POINTER(ProcessMemory), wintypes.DWORD]

    def available(self):
        value = MemoryStatus()
        value.length = ctypes.sizeof(value)
        if not self.kernel.GlobalMemoryStatusEx(ctypes.byref(value)):
            raise ctypes.WinError(ctypes.get_last_error())
        return value.available_physical


class NativeProcess:
    """Keep the verified process handle, so PID reuse cannot redirect termination."""
    def __init__(self, api, pid):
        self.api, self.pid = api, pid
        self.handle = api.kernel.OpenProcess(0x100000 | 0x1000 | 0x10 | 0x1, False, pid)
        if not self.handle:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            creation, end, kernel, user = [wintypes.FILETIME() for _ in range(4)]
            if not api.kernel.GetProcessTimes(self.handle, ctypes.byref(creation), ctypes.byref(end), ctypes.byref(kernel), ctypes.byref(user)):
                raise ctypes.WinError(ctypes.get_last_error())
            self.creation = filetime(creation)
            buffer = ctypes.create_unicode_buffer(32768)
            length = wintypes.DWORD(len(buffer))
            if not api.kernel.QueryFullProcessImageNameW(self.handle, 0, buffer, ctypes.byref(length)):
                raise ctypes.WinError(ctypes.get_last_error())
            self.image_name = Path(buffer.value).name.upper()
        except BaseException:
            self.close()
            raise

    def running(self):
        status = self.api.kernel.WaitForSingleObject(self.handle, 0)
        if status == 0xFFFFFFFF:
            raise ctypes.WinError(ctypes.get_last_error())
        return status == 258

    def memory(self):
        info = ProcessMemory()
        info.cb = ctypes.sizeof(info)
        if not self.api.psapi.GetProcessMemoryInfo(self.handle, ctypes.byref(info), info.cb):
            raise ctypes.WinError(ctypes.get_last_error())
        return info.working_set, info.private, info.peak_working_set

    def terminate(self):
        if not self.api.kernel.TerminateProcess(self.handle, 0xE0000001):
            raise ctypes.WinError(ctypes.get_last_error())

    def close(self):
        if self.handle:
            self.api.kernel.CloseHandle(self.handle)
            self.handle = None


def authorized(progress, run_id, producer, candidate):
    return (
        progress.get('run_id') == run_id
        and progress.get('ownership_confirmed') is True
        and progress.get('producer_pid') == producer.pid
        and progress.get('producer_creation_filetime') == producer.creation
        and progress.get('owned_excel_pid') == candidate.pid
        and progress.get('owned_excel_creation_filetime') == candidate.creation
        and candidate.creation >= producer.creation
        and candidate.image_name == 'EXCEL.EXE'
    )


def terminate_confirmed(progress, run_id, producer, candidate):
    if not authorized(progress, run_id, producer, candidate):
        raise ValueError('Excel ownership identity mismatch; termination refused')
    if candidate.running():
        candidate.terminate()
        return True
    return False


def digest(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def read_progress(path):
    try:
        return json.loads(path.read_text(encoding='utf-8-sig'))
    except (OSError, ValueError):
        return None


def self_test():
    class Fake:
        def __init__(self, pid, creation, name):
            self.pid, self.creation, self.image_name = pid, creation, name
            self.terminated = False
        def running(self):
            return True
        def terminate(self):
            self.terminated = True
    producer = Fake(10, 100, 'PWSH.EXE')
    excel = Fake(20, 120, 'EXCEL.EXE')
    progress = dict(run_id='fresh', ownership_confirmed=True, producer_pid=10,
                    producer_creation_filetime=100, owned_excel_pid=20,
                    owned_excel_creation_filetime=120)
    for key, wrong in [('run_id', 'stale'), ('ownership_confirmed', False),
                       ('producer_pid', 11), ('producer_creation_filetime', 101),
                       ('owned_excel_pid', 21), ('owned_excel_creation_filetime', 121)]:
        bad = dict(progress, **{key: wrong})
        try:
            terminate_confirmed(bad, 'fresh', producer, excel)
        except ValueError:
            pass
        else:
            raise AssertionError('Wrong identity was accepted: ' + key)
        assert not excel.terminated
    assert not authorized(progress, 'fresh', producer, Fake(20, 120, 'OTHER.EXE'))
    assert not authorized(dict(progress, owned_excel_creation_filetime=90), 'fresh', producer, Fake(20, 90, 'EXCEL.EXE'))
    assert terminate_confirmed(progress, 'fresh', producer, excel) and excel.terminated
    api = Windows()
    current = NativeProcess(api, os.getpid())
    try:
        assert current.running() and current.creation > 0 and current.memory()[0] > 0
        assert api.available() > 0
    finally:
        current.close()
    print('PASS: identity/nonce/PID-reuse/pre-existing/image rejection; fake confirmed stop; current Python handle/memory read. No Excel or real process was terminated.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--self-test', action='store_true')
    parser.add_argument('--powershell')
    parser.add_argument('--source', type=Path)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--inventory', type=Path, default=Path(__file__).with_name('input-inventory.json'))
    parser.add_argument('--evidence-dir', type=Path)
    args = parser.parse_args()
    if args.self_test:
        self_test()
        return
    if not all((args.powershell, args.source, args.output, args.evidence_dir)):
        parser.error('--powershell, --source, --output and --evidence-dir are required')
    folder = args.evidence_dir.resolve()
    folder.mkdir(parents=True, exist_ok=False)
    progress_path, report_path = folder / 'progress.json', folder / 'excel-client-validation.json'
    script = Path(__file__).with_name('validate_excel_chunked.ps1').resolve()
    run_id = str(uuid.uuid4())
    api = Windows()
    available = api.available()
    cap = min(8 * GIB, available // 3)
    deadline, grace, interval = 900, 60, .1
    command = [args.powershell, '-NoProfile', '-File', str(script), '-Source',
               str(args.source.resolve()), '-DefaultOutput', str(args.output.resolve()),
               '-Inventory', str(args.inventory.resolve()), '-Report', str(report_path),
               '-Progress', str(progress_path), '-RunId', run_id]
    start = time.monotonic()
    owned = producer = proc = None
    confirmed = progress = None
    stop_reason = failure = None
    stop_at = None
    killed = False
    peak_rss = peak_private = samples = 0
    observed_stages = []
    try:
        with (folder / 'stdout.txt').open('w', encoding='utf-8') as out, (folder / 'stderr.txt').open('w', encoding='utf-8') as err:
            proc = subprocess.Popen(command, stdout=out, stderr=err,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
            producer = NativeProcess(api, proc.pid)
            while True:
                now = time.monotonic()
                incoming = read_progress(progress_path)
                if incoming is not None:
                    progress = incoming
                if owned is None and progress and progress.get('owned_excel_pid'):
                    candidate = NativeProcess(api, int(progress['owned_excel_pid']))
                    if not authorized(progress, run_id, producer, candidate):
                        candidate.close()
                        raise ValueError('Progress identity failed verification')
                    owned, confirmed = candidate, dict(progress)
                if owned is not None and owned.running():
                    try:
                        rss, private, os_peak_rss = owned.memory()
                    except OSError:
                        if not owned.running():
                            continue  # Normal Excel exit can race the memory sample.
                        raise
                    peak_rss, peak_private = max(peak_rss, rss, os_peak_rss), max(peak_private, private)
                    samples += 1
                    if progress and authorized(progress, run_id, producer, owned):
                        stage = progress.get('stage')
                        if not observed_stages or observed_stages[-1]['stage'] != stage:
                            observed_stages.append({'seconds': round(now-start, 3), 'stage': stage,
                                                    'checked_cells_in_sheet': progress.get('checked_cells_in_sheet', 0)})
                    if stop_reason is None and max(rss, private) > cap:
                        stop_reason, stop_at = 'probe_excel_memory_cap', now
                if stop_reason is None and now-start > deadline:
                    stop_reason, stop_at = 'probe_timeout', now
                if stop_reason and owned is not None and not killed:
                    killed = terminate_confirmed(confirmed, run_id, producer, owned)
                worker_done = proc.poll() is not None
                excel_done = owned is None or not owned.running()
                if worker_done and excel_done:
                    break
                if worker_done and stop_reason is None:
                    stop_reason, stop_at = 'probe_cleanup_after_worker_exit', now
                    continue
                if stop_at is not None and now-stop_at > grace:
                    break  # Do not terminate PowerShell or any unconfirmed Excel process.
                time.sleep(interval)
    except Exception as error:
        failure = type(error).__name__
        if owned is not None and confirmed is not None:
            try:
                killed = terminate_confirmed(confirmed, run_id, producer, owned) or killed
                stop_reason = stop_reason or 'probe_monitor_failure_cleanup'
                if proc is not None:
                    try:
                        proc.wait(timeout=grace)
                    except subprocess.TimeoutExpired:
                        pass
            except Exception as cleanup_error:
                failure += ';cleanup:' + type(cleanup_error).__name__
    finally:
        excel_exited = owned is not None and not owned.running()
        child_running = proc is not None and proc.poll() is None
        identity = None if owned is None else dict(pid=owned.pid, creation_filetime=owned.creation, image_name=owned.image_name)
        if owned is not None:
            owned.close()
        if producer is not None:
            producer.close()
    client = read_progress(report_path)
    passed = failure is None and stop_reason is None and not child_running and excel_exited and proc.returncode == 0 and client is not None and client.get('passed') is True
    record = dict(recorded_utc=datetime.now(timezone.utc).isoformat(), run_id=run_id,
                  wrapper_sha256=digest(__file__), script_sha256=digest(script),
                  available_bytes_at_start=available, memory_cap_bytes=cap,
                  timeout_seconds=deadline, cleanup_grace_seconds=grace, sample_interval_seconds=interval,
                  monitored_scope='Only COM-confirmed Excel PID with matching exact creation FILETIME, run nonce and producer identity; retained process handle prevents PID reuse.',
                  owned_excel=identity, samples=samples, peak_rss_mib=round(peak_rss/1024**2, 3),
                  peak_sampled_private_mib=round(peak_private/1024**2, 3),
                  seconds=round(time.monotonic()-start, 3), stop_reason=stop_reason,
                  confirmed_excel_terminated=killed, excel_exited=excel_exited,
                  powershell_pid=None if proc is None else proc.pid, powershell_left_running=child_running,
                  exit_code=None if proc is None else proc.returncode, monitor_failure_type=failure,
                  client_report_present=client is not None, client_passed=client is not None and client.get('passed') is True,
                  stages=observed_stages, passed=passed,
                  limitations='Limits belong to this acceptance probe, not the product. A stop is active monitor intervention, not evidence of application OOM or a product crash. No process difference set, process tree, preexisting Excel, or PowerShell process is terminated. If no owned identity arrives, or finally remains stuck after grace, this report requires manual cleanup of the explicitly recorded worker only.')
    (folder / 'monitor.json').write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'passed': passed, 'stop_reason': stop_reason, 'monitor_failure_type': failure,
                      'report': str(folder / 'monitor.json')}, ensure_ascii=True))
    if not passed:
        raise SystemExit(1)


if __name__ == '__main__':
    main()

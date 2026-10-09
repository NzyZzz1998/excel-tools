"""Real v1.1.1 candidate EXE, one monthly input, bounded whole-process-tree telemetry."""
import hashlib
import importlib.util
import json
import shutil
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

import psutil
import native_gui_controller as c

ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parents[4]
LEGACY = WORKSPACE / 'docs/releases/v1.1/large-data-acceptance-2026-10-09/exe/run_real_exe_batch.py'
spec = importlib.util.spec_from_file_location('validated_gui_helpers', LEGACY)
helpers = importlib.util.module_from_spec(spec)
spec.loader.exec_module(helpers)
assert helpers.c is c
RUN_ROOT = WORKSPACE / 'testfile/月度验收_v1.1.1_2026-10-09'
SOURCE = WORKSPACE / 'testfile/期间缺货Top20/天-缺货明细表沉淀数据 (1).xlsx'
PACKAGE = WORKSPACE / 'dist/ExcelTools-v1.1.1-Windows-x64.zip'
ZIP_SHA = '7D3954B5A898C14F42AEF8C8881ACD8F075097D941689204C68E18FAA4C45B52'
EXE_SHA = 'B5ECA2E5837698E677BDB3F6B9BF844275F734BA0CE7643275FD46871CCAA2ED'
sha = helpers.sha
write_json = helpers.write_json
require = helpers.require


def candidate_control(name):
    main = c.main_window()
    origin = c.details(main)['rect']
    matches = []
    for item in c.children(main):
        x, y, right, bottom = item['rect']
        size = right-x, bottom-y
        if name == 'choose' and size == (155, 27):
            matches.append(item)
        elif name == 'start' and size == (87, 27) and x-origin[0] == 26:
            matches.append(item)
        elif name == 'option' and size == (239, 23):
            matches.append(item)
    require(len(matches) == 1, 'Candidate native control match must be unique: '+name)
    return matches[0]['hwnd']


helpers.control = candidate_control


class TreeMonitor:
    """50 ms samples; sum current RSS/private over owned live descendants."""
    def __init__(self):
        self.available_before = psutil.virtual_memory().available
        self.limit = min(4 * 1024**3, self.available_before // 3)
        self.interval = .05
        self.timeout = 1800
        self.owned = {}
        self.snapshot = {}
        self.peak_rss = self.peak_private = 0
        self.samples = 0
        self.max_gap = 0
        self.stop_reason = None
        self.monitor_error = None
        self.progress_write_error = None
        self.terminated = []
        self.finished = threading.Event()
        self.thread = None

    def start(self, process):
        self.started = time.monotonic()
        self.owned[process.pid] = psutil.Process(process.pid).create_time()
        self.thread = threading.Thread(target=self.sample, daemon=True)
        self.thread.start()
        print(json.dumps(dict(event='launched', launcher_pid=process.pid,
                              memory_limit_bytes=self.limit, available_bytes=self.available_before,
                              timeout_seconds=self.timeout, sample_interval_seconds=self.interval)), flush=True)

    def live_processes(self):
        live = []
        for pid, created in list(self.owned.items()):
            try:
                process = psutil.Process(pid)
                if process.create_time() != created:
                    continue
                live.append(process)
                for child in process.children(recursive=True):
                    self.owned[child.pid] = child.create_time()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        for pid, created in list(self.owned.items()):
            if any(process.pid == pid for process in live):
                continue
            try:
                process = psutil.Process(pid)
                if process.create_time() == created:
                    live.append(process)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return live

    def probe_terminate(self, reason):
        self.stop_reason = reason
        # Only the probe's launcher/descendants with unchanged creation times.
        # This is a resource-protection stop, never evidence of application OOM.
        for process in reversed(self.live_processes()):
            try:
                if process.create_time() == self.owned[process.pid]:
                    process.terminate()
                    self.terminated.append(process.pid)
            except psutil.NoSuchProcess:
                pass

    def sample(self):
        last = self.started
        next_trace = self.started
        next_progress = self.started
        try:
            with (ROOT / 'memory-trace.jsonl').open('x', encoding='utf-8') as trace:
                while not self.finished.is_set():
                    now = time.monotonic()
                    self.max_gap = max(self.max_gap, now - last)
                    last = now
                    rss = private = 0
                    cpu = 0.0
                    pids = []
                    for process in self.live_processes():
                        try:
                            memory = process.memory_info()
                            times = process.cpu_times()
                            rss += memory.rss
                            private += memory.private
                            cpu += times.user + times.system
                            pids.append(process.pid)
                        except psutil.NoSuchProcess:
                            pass
                    self.samples += 1
                    self.peak_rss = max(self.peak_rss, rss)
                    self.peak_private = max(self.peak_private, private)
                    self.snapshot = dict(elapsed_seconds=round(now - self.started, 3), pids=pids,
                                         rss_bytes=rss, private_bytes=private, cpu_seconds=round(cpu, 3),
                                         peak_rss_bytes=self.peak_rss, peak_private_bytes=self.peak_private)
                    if now >= next_trace:
                        trace.write(json.dumps(self.snapshot) + '\n')
                        trace.flush()
                        next_trace += 1
                    if now >= next_progress:
                        progress = dict(recorded_utc=datetime.now(timezone.utc).isoformat(),
                                        **self.snapshot, process_tree_alive=bool(pids),
                                        output_files=len(list((RUN_ROOT / 'default').glob('*_拆分填充*.xlsx'))),
                                        memory_limit_bytes=self.limit, probe_stop_reason=self.stop_reason,
                                        ui_completion_confirmed=False)
                        try:
                            temporary = ROOT / 'progress.json.tmp'
                            write_json(temporary, progress)
                            temporary.replace(ROOT / 'progress.json')
                        except OSError as error:
                            self.progress_write_error = type(error).__name__
                        next_progress += 10
                    if self.stop_reason is None:
                        if max(rss, private) > self.limit:
                            self.probe_terminate('probe_memory_limit_exceeded_not_application_oom')
                        elif now - self.started > self.timeout:
                            self.probe_terminate('probe_1800_second_timeout_not_application_crash')
                    self.finished.wait(max(0, self.interval - (time.monotonic() - now)))
        except Exception as error:
            self.monitor_error = type(error).__name__ + ': ' + str(error)
            self.probe_terminate('probe_monitor_failed_no_resource_guarantee')

    def close(self):
        self.finished.set()
        if self.thread:
            self.thread.join(timeout=5)
        return dict(sample_interval_seconds=self.interval, sample_count=self.samples,
                    observed_max_sample_gap_seconds=round(self.max_gap, 4),
                    available_memory_before_bytes=self.available_before, memory_limit_bytes=self.limit,
                    peak_process_tree_rss_bytes=self.peak_rss,
                    peak_process_tree_private_bytes=self.peak_private,
                    lifetime_seconds=round(time.monotonic() - self.started, 3) if self.thread else None,
                    timeout_seconds=self.timeout, owned_pid_creation_times=self.owned,
                    probe_stop_reason=self.stop_reason, probe_terminated_pids=self.terminated,
                    monitor_error=self.monitor_error,
                    progress_write_error=self.progress_write_error,
                    limitations='50ms target sampling can miss instantaneous peaks; RSS/private are separate measures, not additive. Trace persists approximately 1s snapshots; peaks use all samples.')


def run():
    require(not (ROOT / 'execution.json').exists(), 'Do not overwrite monthly evidence')
    require(SOURCE.stat().st_size == 104185338, 'Monthly input size differs')
    require(sha(PACKAGE) == ZIP_SHA, 'Released package identity mismatch')
    source_sha = sha(SOURCE)
    folder = RUN_ROOT / 'default'
    folder.mkdir(parents=True, exist_ok=True)
    require(not any(folder.iterdir()), 'Monthly default folder must be empty')
    source_copy = folder / SOURCE.name
    shutil.copy2(SOURCE, source_copy)
    require(sha(source_copy) == source_sha, 'Input copy SHA mismatch')
    app = RUN_ROOT / 'app'
    app.mkdir(parents=True, exist_ok=True)
    c.EXE = app / 'ExcelTools.exe'
    require(not c.EXE.exists(), 'Refusing to replace an existing monthly EXE')
    with ZipFile(PACKAGE) as archive:
        with archive.open('ExcelTools.exe') as source, c.EXE.open('xb') as destination:
            shutil.copyfileobj(source, destination)
    require(sha(c.EXE) == EXE_SHA, 'Released EXE identity mismatch')
    c.ROOT = ROOT
    identity = dict(source=str(SOURCE), source_bytes=SOURCE.stat().st_size,
                    source_sha256=source_sha, copied_input=str(source_copy),
                    package=str(PACKAGE), package_sha256=ZIP_SHA,
                    exe=str(c.EXE), exe_sha256=EXE_SHA,
                    controller_sha256=sha(ROOT / 'native_gui_controller.py'),
                    runner_sha256=sha(Path(__file__)), reused_helper_sha256=sha(LEGACY))
    write_json(ROOT / 'identity.json', identity)
    record = dict(recorded_utc=datetime.now(timezone.utc).isoformat(), mode='default',
                  identity=identity, assertions_passed=False,
                  ui_conclusion='Pending fresh screenshot inspection')
    monitor = TreeMonitor()
    try:
        record['startup'] = c.start(on_launched=monitor.start)
        record['actual_window_title'] = c.details(c.main_window())['title']
        record['actual_window_pid'] = c.details(c.main_window())['pid']
        record['selection'] = helpers.select_files([source_copy])
        record['selected_capture'] = c.capture_fresh(ROOT / 'selected.png')
        c.guarded_mouse_click(helpers.control('start'))
        batch_started = time.monotonic()
        next_notice = batch_started
        previous = None
        stable_since = None
        captured_processing = False
        while True:
            require(monitor.stop_reason is None, 'Resource protection: ' + str(monitor.stop_reason))
            require(monitor.monitor_error is None, 'Memory monitor failed')
            require(c.our_pids(), 'Owned EXE exited before verified completion')
            outputs = sorted(folder.glob('*_拆分填充*.xlsx'))
            sizes = [(str(path), path.stat().st_size) for path in outputs]
            if len(outputs) == 1 and sizes == previous:
                stable_since = stable_since or time.monotonic()
                if time.monotonic() - stable_since >= 3 and helpers.confirm_idle():
                    record['ui_idle_confirmed_by_cancelled_file_dialog'] = True
                    break
            else:
                stable_since = None
            previous = sizes
            if time.monotonic() >= next_notice:
                print(json.dumps(dict(event='processing', **monitor.snapshot,
                                      batch_seconds=round(time.monotonic() - batch_started, 1),
                                      final_output_files=len(outputs), ui_completion_confirmed=False)), flush=True)
                if time.monotonic() - batch_started > 25 and not captured_processing:
                    record['processing_capture'] = c.capture_fresh(ROOT / 'processing.png')
                    captured_processing = True
                if time.monotonic() - batch_started > 25 and not outputs and helpers.confirm_idle():
                    record['ui_idle_without_output'] = True
                    raise AssertionError('GUI ended without a new result; inspect its fresh screenshot')
                next_notice += 30
            time.sleep(.5)
        record['observed_batch_seconds_including_idle_probe'] = round(time.monotonic() - batch_started, 3)
        record['result_capture'] = c.capture_fresh(ROOT / 'result.png')
        require(len(list(folder.glob('*.xlsx'))) == 2, 'Expected one copied input and one result')
        require(sha(SOURCE) == source_sha, 'Original source changed')
        require(sha(source_copy) == source_sha, 'Input copy changed')
        output = outputs[0]
        record['output'] = dict(path=str(output), bytes=output.stat().st_size, sha256=sha(output))
        record['identical_to_verified_v1_1_output'] = sha(output).lower() == 'c28ee8a7646ac33315fb89f0a9483a51c5d9b7a8eb05195632493c2d7e02c4ae'
        require(record['identical_to_verified_v1_1_output'], 'New output differs from the independently verified v1.1 result; investigate before acceptance')
        require(not list(folder.glob('*.part')), 'Successful run left a temporary result')
        record['source_unchanged'] = record['input_copy_unchanged'] = True
        record['assertions_passed'] = True
    except Exception as error:
        record['failure_type'] = type(error).__name__
        record['failure_message'] = str(error)
        if c.our_pids() and monitor.stop_reason is None:
            try:
                record['failure_capture'] = c.capture_fresh(ROOT / 'failure.png')
            except Exception as capture_error:
                record['capture_failure_type'] = type(capture_error).__name__
        raise
    finally:
        if c.our_pids() and monitor.stop_reason is None:
            try:
                record['close'] = c.close_idle()
            except Exception as close_error:
                record['close_failure_type'] = type(close_error).__name__
        record['monitor'] = monitor.close()
        record['remaining_owned_pids'] = sorted(c.our_pids())
        record['source_sha256_after'] = sha(SOURCE)
        record['source_unchanged'] = record['source_sha256_after'] == source_sha
        write_json(ROOT / 'execution.json', record)
    require(not record['remaining_owned_pids'], 'Owned EXE is still running; inspect evidence')
    print(json.dumps(dict(event='completed', output=record['output'],
                          seconds=record['observed_batch_seconds_including_idle_probe'],
                          monitor=record['monitor']), ensure_ascii=False), flush=True)


if __name__ == '__main__':
    run()

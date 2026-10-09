"""Atomic output publication and bounded, observational progress regression tests."""
import errno
import os
import tempfile
import threading
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from types import SimpleNamespace
from unittest import mock
from zipfile import ZipFile

from test_excel_unmerge import c, makebook, mod, sha, sheet


class OutputProgressTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='excel-atomic-progress-')
        self.root = Path(self.temp.name)
        self.source = makebook(self.root/'source.xlsx', [('data', sheet(
            {1: [c('A1', 'synthetic group'), c('C1', 'ordinary')]}, ['A1:A3']))])
        self.source_hash = sha(self.source)

    def tearDown(self):
        self.temp.cleanup()

    def test_partial_zip_is_only_part_until_closed_publication(self):
        original_open = mod.ZipFile.open
        observed = []
        def inspect(archive, name, mode='r', *args, **kwargs):
            if mode == 'w':
                files = list(self.root.iterdir())
                parts = [p for p in files if p.suffix == '.part']
                self.assertEqual(len(parts), 1)
                self.assertEqual([p for p in files if p.suffix == '.xlsx'], [self.source])
                observed.append(parts[0])
            return original_open(archive, name, mode, *args, **kwargs)
        with mock.patch.object(mod.ZipFile, 'open', inspect):
            output, _ = mod.process_file(self.source)
        self.assertTrue(observed)
        self.assertFalse(any(p.exists() for p in observed))
        with ZipFile(output) as archive:
            self.assertIsNone(archive.testzip())
        self.assertEqual(sha(self.source), self.source_hash)

    def test_progress_does_not_change_zip_payloads_and_has_real_stage_counts(self):
        ordinary, expected = mod.process_file(self.source)
        events = []
        output, actual = mod.process_file(self.source, progress=events.append)
        self.assertEqual(actual, expected)
        self.assertEqual(sha(output), sha(ordinary))
        self.assertEqual({e['phase'] for e in events}, {'reading', 'filling', 'saving'})
        for event in events:
            self.assertEqual(set(event), {'phase', 'sheet', 'completed', 'total', 'unit'})
            if event['total'] is not None:
                self.assertLessEqual(event['completed'], event['total'])
        saved = [e for e in events if e['phase'] == 'saving']
        self.assertEqual(saved[0]['completed'], 0)
        self.assertEqual(saved[-1]['completed'], saved[-1]['total'])
        with ZipFile(output) as archive:
            self.assertEqual(saved[-1]['total'], sum(i.file_size for i in archive.infolist()))
        self.assertEqual(sha(self.source), self.source_hash)

    def partial_write_failure(self, original_error):
        real_open = mod.ZipFile.open
        state = {'entries': 0, 'written': 0}
        def injected(archive, name, mode='r', *args, **kwargs):
            stream = real_open(archive, name, mode, *args, **kwargs)
            if mode == 'w':
                state['entries'] += 1
                if state['entries'] == 2:
                    real_write = stream.write
                    def fail(data):
                        state['written'] += real_write(data[:5])
                        raise original_error
                    stream.write = fail
            return stream
        return injected, state

    def test_partial_write_failure_keeps_old_results_and_foreign_part(self):
        previous, _ = mod.process_file(self.source)
        foreign = self.root/'other-session.part'
        foreign.write_bytes(b'belongs to another session')
        before = {p: sha(p) for p in self.root.iterdir()}
        error = OSError(errno.ENOSPC, 'synthetic disk full')
        injected, state = self.partial_write_failure(error)
        with mock.patch.object(mod.ZipFile, 'open', injected):
            with self.assertRaises(OSError) as raised:
                mod.process_file(self.source)
        self.assertIs(raised.exception, error)
        self.assertGreater(state['written'], 0)
        self.assertEqual(state['entries'], 2)
        self.assertEqual({p: sha(p) for p in self.root.iterdir()}, before)

    def test_cleanup_failure_preserves_original_errno_and_reports_own_part(self):
        error = OSError(errno.ENOSPC, 'synthetic disk full')
        cleanup = PermissionError(errno.EACCES, 'synthetic external read lock')
        injected, state = self.partial_write_failure(error)
        real_unlink = Path.unlink
        def locked(path, *args, **kwargs):
            if path.suffix == '.part':
                raise cleanup
            return real_unlink(path, *args, **kwargs)
        with mock.patch.object(mod.ZipFile, 'open', injected), mock.patch.object(Path, 'unlink', locked):
            with self.assertRaises(OSError) as raised:
                mod.process_file(self.source)
        self.assertIs(raised.exception, error)
        self.assertEqual(error.errno, errno.ENOSPC)
        self.assertGreater(state['written'], 0)
        self.assertIs(error.cleanup_error, cleanup)
        self.assertEqual(error.partial_path.parent, self.root)
        self.assertEqual(error.partial_path.suffix, '.part')
        self.assertTrue(error.partial_path.exists())
        self.assertIn('synthetic disk full', str(error))
        self.assertIn(str(error.partial_path), str(error))
        self.assertEqual(list(self.root.glob('*.xlsx')), [self.source])
        self.assertEqual(sha(self.source), self.source_hash)

    def test_publish_waits_for_valid_closed_zip_and_handles_late_name_race(self):
        real_publish = mod._publish_no_replace
        attempts = []
        raced = self.root/'source_拆分填充.xlsx'
        def late_occupation(part, output):
            with ZipFile(part) as archive:
                self.assertIsNone(archive.testzip())
            attempts.append(output)
            if len(attempts) == 1:
                output.write_bytes(b'another process won this name')
            return real_publish(part, output)
        with mock.patch.object(mod, '_publish_no_replace', late_occupation):
            output, _ = mod.process_file(self.source)
        self.assertEqual(output.name, 'source_拆分填充_2.xlsx')
        self.assertEqual(len(attempts), 2)
        self.assertEqual(raced.read_bytes(), b'another process won this name')
        self.assertFalse(list(self.root.glob('*.part')))

    def test_concurrent_publications_are_distinct_and_do_not_overwrite(self):
        previous, _ = mod.process_file(self.source)
        previous_hash = sha(previous)
        barrier = threading.Barrier(4)
        real_publish = mod._publish_part
        def synchronized(part, source):
            barrier.wait(timeout=10)
            return real_publish(part, source)
        with mock.patch.object(mod, '_publish_part', synchronized):
            with ThreadPoolExecutor(max_workers=4) as pool:
                results = list(pool.map(lambda _: mod.process_file(self.source), range(4)))
        outputs = {path for path, _ in results}
        self.assertEqual(len(outputs), 4)
        self.assertNotIn(previous, outputs)
        self.assertEqual({p.name for p in outputs}, {'source_拆分填充_'+str(i)+'.xlsx' for i in range(2, 6)})
        self.assertTrue(all(sha(p) == previous_hash for p in outputs))
        self.assertEqual(sha(previous), previous_hash)
        self.assertEqual(sha(self.source), self.source_hash)
        self.assertFalse(list(self.root.glob('*.part')))

    def test_real_publish_permission_failure_is_not_a_name_retry(self):
        error = PermissionError(errno.EACCES, 'synthetic publish denied')
        with mock.patch.object(mod, '_publish_no_replace', side_effect=error) as publish:
            with self.assertRaises(PermissionError) as raised:
                mod.process_file(self.source)
        self.assertIs(raised.exception, error)
        self.assertEqual(publish.call_count, 1)
        self.assertEqual(list(self.root.iterdir()), [self.source])

    def test_non_windows_link_publication_cannot_overwrite_existing_name(self):
        part = self.root/'synthetic.part'
        part.write_bytes(self.source.read_bytes())
        output = self.root/'published.xlsx'
        output.write_bytes(b'old result')
        posix = SimpleNamespace(name='posix', link=os.link)
        with mock.patch.object(mod, 'os', posix):
            with self.assertRaises(FileExistsError):
                mod._publish_no_replace(part, output)
            self.assertTrue(part.exists())
            self.assertEqual(output.read_bytes(), b'old result')
            fresh = self.root/'fresh.xlsx'
            mod._publish_no_replace(part, fresh)
        self.assertFalse(part.exists())
        self.assertEqual(sha(fresh), self.source_hash)

    def test_link_success_then_part_unlink_failure_reports_published_output_once(self):
        part = self.root/'synthetic.part'
        part.write_bytes(self.source.read_bytes())
        error = PermissionError(errno.EACCES, 'synthetic unlink denied')
        posix = SimpleNamespace(name='posix', link=os.link)
        real_unlink = Path.unlink
        def locked(path, *args, **kwargs):
            if path == part:
                raise error
            return real_unlink(path, *args, **kwargs)
        with mock.patch.object(mod, 'os', posix), mock.patch.object(Path, 'unlink', locked):
            with self.assertRaises(PermissionError) as raised:
                mod._publish_part(part, self.source)
        self.assertIs(raised.exception, error)
        self.assertEqual(error.published_output.name, 'source_拆分填充.xlsx')
        self.assertEqual(sha(error.published_output), self.source_hash)
        self.assertTrue(part.exists())
        self.assertIn(str(part), str(error))
        self.assertIn(str(error.published_output), str(error))
        self.assertFalse((self.root/'source_拆分填充_2.xlsx').exists())

    def test_callback_exception_is_observational_but_base_exception_cleans_part(self):
        expected, _ = mod.process_file(self.source, False, ['data'])
        def broken(_):
            raise RuntimeError('observer failed')
        actual, _ = mod.process_file(self.source, progress=broken)
        self.assertEqual(sha(actual), sha(expected))
        before = {p: sha(p) for p in self.root.iterdir()}
        class StopObservation(BaseException):
            pass
        def interrupted(detail):
            if detail['phase'] == 'saving' and detail['completed']:
                raise StopObservation('observer interrupted')
        with self.assertRaises(StopObservation):
            mod.process_file(self.source, progress=interrupted)
        self.assertEqual({p: sha(p) for p in self.root.iterdir()}, before)

    def test_stream_progress_is_throttled_and_counts_physical_and_generated_rows(self):
        raw = sheet({row: [c('B'+str(row), row, 'num')] + ([c('A1', 'group')] if row == 1 else [])
                     for row in range(1, 1001)}, ['A1:A1200'])
        source = makebook(self.root/'stream.xlsx', [('stream', raw)])
        events = []
        with mock.patch.object(mod, 'STREAM_THRESHOLD', 0), mock.patch.object(mod.time, 'monotonic', return_value=0):
            result, _ = mod.process_file(source, progress=events.append)
        self.assertEqual(len(events), 6, 'Constant clock must only produce stage boundary events')
        reading = [e for e in events if e['phase'] == 'reading']
        filling = [e for e in events if e['phase'] == 'filling']
        self.assertEqual((reading[0]['completed'], reading[0]['total'], reading[0]['unit']), (0, None, 'rows'))
        self.assertEqual((reading[-1]['completed'], reading[-1]['total']), (1000, 1000))
        self.assertEqual((filling[0]['completed'], filling[0]['total']), (0, 1200))
        self.assertEqual((filling[-1]['completed'], filling[-1]['total']), (1200, 1200))
        clock = [0.0]
        def tick():
            clock[0] += 0.01
            return clock[0]
        timed = []
        with mock.patch.object(mod, 'STREAM_THRESHOLD', 0), mock.patch.object(mod.time, 'monotonic', side_effect=tick):
            compared, _ = mod.process_file(source, progress=timed.append)
        self.assertEqual(sha(compared), sha(result))
        self.assertLess(len(timed), 90)
        for phase in ('reading', 'filling'):
            samples = [e['completed'] for e in timed if e['phase'] == phase]
            self.assertGreater(len(samples), 2)
            self.assertEqual(samples, sorted(samples))

    def test_no_matches_emit_no_saving_and_create_no_part(self):
        source = makebook(self.root/'none.xlsx', [('none', sheet({1: [c('A1', 'header')]}, ['A1:C1']))])
        before = {p: sha(p) for p in self.root.iterdir()}
        for streaming in (False, True):
            events = []
            with mock.patch.object(mod, 'STREAM_THRESHOLD', 0 if streaming else 8*1024*1024), mock.patch.object(mod, 'mkstemp') as create:
                output, stats = mod.process_file(source, progress=events.append)
            self.assertIsNone(output)
            self.assertEqual(stats, [('none', 0, 0)])
            create.assert_not_called()
            self.assertNotIn('saving', {e['phase'] for e in events})
        self.assertEqual({p: sha(p) for p in self.root.iterdir()}, before)

    def test_saving_progress_updates_during_large_zip_payload(self):
        source = makebook(self.root/'large-part.xlsx', [('data', sheet(
            {1: [c('A1', 'group')]}, ['A1:A2']))], {'opaque/large.bin': b'x'*(5*1024*1024)})
        clock = [0.0]
        def tick():
            clock[0] += 0.11
            return clock[0]
        events = []
        with mock.patch.object(mod.time, 'monotonic', side_effect=tick):
            output, _ = mod.process_file(source, progress=events.append)
        saved = [e for e in events if e['phase'] == 'saving']
        self.assertEqual(saved[0]['completed'], 0)
        self.assertEqual(saved[-1]['completed'], saved[-1]['total'])
        self.assertTrue(any(1024*1024 < e['completed'] < e['total'] for e in saved))
        self.assertLess(len(saved), 10)
        with ZipFile(output) as archive:
            self.assertEqual(saved[-1]['total'], sum(i.file_size for i in archive.infolist()))

    def test_fdopen_failure_cleans_own_part_and_preserves_source(self):
        error = OSError('synthetic handle wrapper failure')
        with mock.patch.object(mod.os, 'fdopen', side_effect=error):
            with self.assertRaises(OSError) as raised:
                mod.process_file(self.source)
        self.assertIs(raised.exception, error)
        self.assertEqual(list(self.root.iterdir()), [self.source])
        self.assertEqual(sha(self.source), self.source_hash)


if __name__ == '__main__':
    unittest.main(verbosity=2)
